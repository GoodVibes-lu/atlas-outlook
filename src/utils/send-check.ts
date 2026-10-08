/**
 * send-check.ts — CONTRÔLES AVANT L'ENVOI d'un mail rédigé dans Outlook (agent d'inbox, phase 2,
 * `.claude/BACKLOG-AGENT-INBOX.md`). Bureau et web seulement (pas de rédaction sur mobile).
 *
 * Logique PURE (aucun réseau, aucun Office.js) : testée par `scripts/test-outlook-send-check.mjs`
 * (node pur). Utilisée par l'onglet « 🛡️ Vérifier » du panneau de rédaction (send-check-panel.ts)
 * et, si l'événement est activé dans le manifeste, par le gestionnaire Smart Alerts
 * `atlasOnMessageSend` (commands.ts).
 *
 * Contrôles :
 *  - pièce jointe annoncée dans le texte rédigé (« ci-joint », « en pièce jointe », « attached »,
 *    « anbei »…) mais aucune pièce jointe (les images intégrées, logos de signature, ne comptent pas) ;
 *  - destinataire probablement erroné :
 *      · homonyme : deux contacts ATLAS du même nom avec des adresses différentes ;
 *      · client d'un autre projet que celui du fil (projet lié à la conversation ou « #NNN » dans
 *        l'objet), d'après les fiches contacts / clients ATLAS ;
 *  - « répondre à tous » avec beaucoup de destinataires alors que le message s'adresse à une seule
 *    personne (« Bonjour Marc, ») ;
 *  - tutoiement / vouvoiement incohérent avec le fil cité (français) ;
 *  - Assistant inbox (07/10/2026, « Relire avant envoi ») : tutoiement d'un CLIENT (sans habitude
 *    connue), langue différente de celle d'habitude du correspondant, date citée incohérente avec
 *    les dates du projet du fil, adresse d'un FOURNISSEUR en copie d'un mail client (les
 *    fournisseurs ne sont jamais exposés), « Cordialement » (charte : « Bien à vous ») ;
 *  - moment d'envoi : destinataire hors de ses heures ouvrées (fuseau d'après le domaine) →
 *    conseil d'envoi différé (`momentEnvoi`, copie de src/utils/inbox-reponse-suivi.ts).
 *
 * Les fonctions de texte (ownPart, visibleText, detectLang, tutoiementMarks) sont une COPIE de
 * `src/utils/client-send-check.ts` (racine du dépôt, contrôle avant envoi client d'ATLAS) : le
 * complément est publié depuis son propre dossier et n'importe rien hors de `outlook-addin/`.
 * Toute correction de ces fonctions là-bas doit être recopiée ici.
 *
 * Gravité : « bloquant » = à corriger avant d'envoyer ; « avertissement » = signalé seulement. Dans la
 * boîte Smart Alerts, seul un cas CRITIQUE retient l'envoi (CODES_CRITIQUES) ; le reste laisse « Envoyer
 * quand même » (decisionEnvoi, 08/10/2026). Rien n'est jamais envoyé ni modifié.
 */
import { indexHtmlQuote, indexQuotedReply, stripSignature } from '../shared/email-body-clean';

export type SendCheckGravite = 'bloquant' | 'avertissement';

export interface SendCheckProblem {
  code: 'piece_jointe' | 'homonyme' | 'client_autre_projet' | 'repondre_a_tous' | 'tutoiement' | 'vouvoiement' | 'registre_mixte'
    | 'tutoiement_client' | 'langue' | 'date_projet' | 'fournisseur_en_copie' | 'cordialement';
  message: string;
  gravite: SendCheckGravite;
}

export interface SendCheckRecipient {
  email: string;
  name?: string;
}

export interface SendCheckContact {
  email: string;
  /** Nom de la personne (fiche contact ATLAS, « Personne de contact »). */
  nom: string;
  /** Société (relation) du contact. */
  societe?: string;
}

export interface SendCheckInput {
  subject?: string;
  /** Corps HTML du message en cours (fil cité compris). */
  html?: string;
  /** Corps texte (si le HTML n'est pas disponible). */
  text?: string;
  /** Destinataires À + Cc (+ Cci). */
  to?: SendCheckRecipient[];
  cc?: SendCheckRecipient[];
  bcc?: SendCheckRecipient[];
  /** Adresse de la boîte qui envoie (exclue du compte des destinataires). */
  moi?: string;
  /**
   * Nombre de pièces jointes NON intégrées (fichiers, pièces jointes cloud). `null` / absent =
   * inconnu (API indisponible) : le contrôle de pièce jointe est alors sauté.
   */
  piecesJointes?: number | null;
  /** Réponse (RE: / AW: / fil cité) : active le contrôle « répondre à tous ». */
  estReponse?: boolean;
  /** Contacts ATLAS (homonymes, société des destinataires). */
  contacts?: SendCheckContact[];
  /** Projet du fil : nom et client (société) ; absent = inconnu. */
  projetDuFil?: { nom: string; client: string } | null;
  /** Sociétés clientes d'au moins un projet ATLAS (pour « client d'un autre projet »). */
  clients?: string[];
  /** Clients / tiers connus par domaine (société d'un destinataire sans fiche contact). */
  domainesClients?: Record<string, string>;
  /** Vrai si l'expéditeur tutoie d'habitude ce correspondant (profil de conversation ARGO). */
  tutoiementConnu?: boolean;
  /** Seuil « répondre à tous » (destinataires hors soi-même). Défaut 4. */
  seuilRepondreATous?: number;
  /** Langue habituelle du correspondant principal (profil de conversation : FR, EN, DE, LU). */
  langueHabituelle?: string | null;
  /** Sociétés FOURNISSEURS (tiers ATLAS) et fournisseurs par domaine. */
  fournisseurs?: string[];
  domainesFournisseurs?: Record<string, string>;
  /** Dates du projet du fil ('YYYY-MM-DD'), si connues. */
  datesProjet?: { debut: string; fin?: string } | null;
  /** Date du jour (tests). */
  maintenant?: number;
}

// ── Texte (copie de src/utils/client-send-check.ts) ─────────────────────────

const decode = (s: string) => s
  .replace(/&nbsp;/gi, ' ').replace(/&amp;/gi, '&').replace(/&lt;/gi, '<').replace(/&gt;/gi, '>')
  .replace(/&quot;/gi, '"').replace(/&#39;|&apos;|&rsquo;/gi, "'").replace(/&eacute;/gi, 'é').replace(/&agrave;/gi, 'à');

/**
 * Partie du mail RÉDIGÉE (sans le fil cité d'une réponse) : coupe au premier séparateur de
 * citation (blockquote, « -----Original », « De : … Envoyé : », « On … wrote: », « Am … schrieb »).
 * Le résultat est toujours un PRÉFIXE du HTML d'origine (cf. quotedPart).
 */
export function ownPart(html: string): string {
  let h = html || '';
  // 08/10/2026 : marqueurs partagés avec ATLAS (shared/email-body-clean.ts : conteneur du nouvel Outlook
  // « mail-editor-reference-message-container », gmail_quote, Apple Mail…) : le faux positif « pièce
  // jointe annoncée » venait d'un « attachments » dans le fil cité d'une réponse du nouvel Outlook Mac.
  const i = indexHtmlQuote(h);
  if (i >= 0) h = h.slice(0, i);
  const cuts = [
    /<blockquote[\s\S]*$/i,
    /<div[^>]*id=["']?(divRplyFwdMsg|appendonsend)["']?[\s\S]*$/i,
    /<hr[^>]*>[\s\S]*?(De|From|Von)(\s|&nbsp;| )*:[\s\S]*$/i,
  ];
  for (const re of cuts) h = h.replace(re, '');
  return h;
}

const QUOTE_TEXT_RE = /\n\s*(-{2,}\s*(Original|Message d'origine|Ursprüngliche)|(De|From|Von)\s*:.*\n\s*(Envoyé|Sent|Gesendet)\s*:|On .{4,80} wrote:|Le .{4,80} a écrit\s*:|Am .{4,80} schrieb)/i;

/** Début du fil cité dans un texte : marqueurs locaux ET partagés (indexQuotedReply), le plus tôt l'emporte. */
function indexCitation(t: string): number {
  const a = t.search(QUOTE_TEXT_RE);
  const b = indexQuotedReply(t);
  if (a > 0 && b > 0) return Math.min(a, b);
  return a > 0 ? a : b > 0 ? b : -1;
}

function htmlToText(html: string): string {
  return decode((html || '')
    .replace(/<style[\s\S]*?<\/style>/gi, '').replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<br\s*\/?>/gi, '\n').replace(/<\/(p|div|li|tr|h\d|table)>/gi, '\n')
    .replace(/<[^>]+>/g, ''))
    .replace(/[ \t]+/g, ' ').replace(/\n\s*\n+/g, '\n').trim();
}

/** Texte visible (lignes conservées), sans le fil cité en texte brut. */
export function visibleText(html: string): string {
  const t = htmlToText(html);
  const m = indexCitation(t);
  return (m > 0 ? t.slice(0, m) : t).trim();
}

/**
 * Texte RÉDIGÉ seul : partie propre (sans le fil cité, HTML puis texte) et SANS la signature (locale ou
 * Exclaimer citée) ; la formule de politesse finale est gardée (contrôle « Cordialement »). C'est sur ce
 * texte, et lui seul, que portent tous les contrôles.
 */
export function texteRedige(html: string): string {
  return stripSignature(visibleText(ownPart(html)));
}

/** Texte du fil CITÉ (ce qui suit la partie rédigée), vide pour un nouveau message. */
export function quotedText(html: string): string {
  const h = html || '';
  const own = ownPart(h);
  const quotedHtml = h.slice(own.length);
  // Fil cité en texte brut, resté dans la partie « rédigée » (Outlook en mode texte).
  const t = htmlToText(own);
  const m = indexCitation(t);
  const tail = m > 0 ? t.slice(m) : '';
  return `${tail}\n${htmlToText(quotedHtml)}`.trim();
}

const norm = (s: string) => s.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
const words = (t: string) => (norm(t).match(/[a-zäöüßë]+/g) || []);

type Lang = 'FR' | 'EN' | 'DE' | 'LB';

const LANG_WORDS: Record<Lang, string[]> = {
  FR: ['le', 'la', 'les', 'des', 'est', 'et', 'pour', 'vous', 'votre', 'vos', 'nous', 'avec', 'une', 'sur', 'dans', 'que', 'qui', 'pas', 'merci', 'bonjour', 'bien', 'sont', 'cette', 'notre', 'nos', 'aux', 'du', 'au', 'ce', 'je', 'il', 'tres', 'plus', 'votre'],
  EN: ['the', 'and', 'you', 'your', 'for', 'with', 'is', 'are', 'we', 'our', 'this', 'that', 'thank', 'thanks', 'please', 'will', 'have', 'of', 'to', 'regards', 'kind', 'hello', 'dear', 'be', 'it', 'on', 'as', 'would', 'can', 'find', 'attached'],
  DE: ['der', 'die', 'das', 'und', 'ist', 'sie', 'ihr', 'ihre', 'ihnen', 'wir', 'unser', 'unsere', 'mit', 'fur', 'nicht', 'ein', 'eine', 'den', 'dem', 'zu', 'auf', 'ich', 'vielen', 'dank', 'gruße', 'grusse', 'freundlichen', 'sehr', 'geehrte', 'geehrter', 'auch', 'bitte', 'haben', 'werden'],
  LB: ['mat', 'fir', 'ech', 'iech', 'ar', 'ass', 'sinn', 'dat', 'awer', 'och', 'eis', 'ze', 'vun', 'moien', 'villmools', 'greiss', 'beschte', 'kennt', 'gett', 'hutt', 'eng', 'engem', 'mech', 'gutt', 'zesummen', 'iwwer'],
};
const LANG_SETS: Record<Lang, Set<string>> = {
  FR: new Set(LANG_WORDS.FR), EN: new Set(LANG_WORDS.EN), DE: new Set(LANG_WORDS.DE), LB: new Set(LANG_WORDS.LB),
};

/** Langue détectée (mots fréquents). confiance = part du meilleur score sur le total. */
export function detectLang(text: string): { lang: Lang | null; confiance: number; mots: number } {
  const ws = words(text);
  const score: Record<Lang, number> = { FR: 0, EN: 0, DE: 0, LB: 0 };
  for (const w of ws) for (const l of Object.keys(score) as Lang[]) if (LANG_SETS[l].has(w)) score[l]++;
  const sorted = (Object.keys(score) as Lang[]).sort((a, b) => score[b] - score[a]);
  const best = sorted[0];
  const total = sorted.reduce((s, l) => s + score[l], 0);
  if (score[best] < 3 || !total) return { lang: null, confiance: 0, mots: ws.length };
  return { lang: best, confiance: score[best] / total, mots: ws.length };
}

/**
 * Marques de tutoiement trouvées (prudent) : « tu », « te », « toi », « ta », « tes », « t' »,
 * « ton » (sauf « le ton », « bon ton »…), impératifs fréquents (« n'hésite pas », « dis-moi »…).
 */
export function tutoiementMarks(text: string): string[] {
  const out: string[] = [];
  const re = /(?<![A-Za-zÀ-ÿ'’-])(tu|te|toi|ta|tes|ton|tien|tienne)(?![A-Za-zÀ-ÿ'’])|(?<![A-Za-zÀ-ÿ'’-])(t['’])(?=[a-zàâéèêëîïôûüh])/gi;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text))) {
    const word = m[1] || m[2];
    const idx = m.index;
    const before = text.slice(Math.max(0, idx - 25), idx);
    const after = text.slice(idx + word.length, idx + word.length + 25);
    const lw = word.toLowerCase();
    const capital = word[0] !== word[0].toLowerCase();
    const sentenceStart = before.trim() === '' || /[.!?:\n]\s*$/.test(before);
    if (capital && !sentenceStart) continue;
    if (capital && /^\s+[A-ZÀ-Ý][a-zà-ÿ]/.test(after)) continue;
    if (lw === 'ton' && /\b(le|du|au|un|de|ce|bon|mauvais|même)\s+$/i.test(before)) continue;
    out.push(word);
  }
  const im = text.match(/(?<![A-Za-zÀ-ÿ])(n['’]h[ée]site pas|dis[- ]moi|fais[- ]moi|tiens[- ]moi|envoie[- ]moi|confirme[- ]moi|rappelle[- ]moi|[ée]cris[- ]moi|bien [àa] toi)(?![A-Za-zÀ-ÿ])/gi);
  if (im) out.push(...im);
  return out;
}

// ── Vouvoiement (propre au complément) ───────────────────────────────────────

/** Marques de vouvoiement : « vous », « votre », « vos », « vôtre », « n'hésitez pas »… */
export function vouvoiementMarks(text: string): string[] {
  return (text || '').match(/(?<![A-Za-zÀ-ÿ'’-])(vous|votre|vos|v[ôo]tres?|n['’]h[ée]sitez pas)(?![A-Za-zÀ-ÿ'’])/gi) || [];
}

// ── Pièce jointe annoncée ────────────────────────────────────────────────────

const ANNONCE_PJ_RES: RegExp[] = [
  // FR
  /\bci[- ]joint(e|s|es)?\b/i,
  /\ben pi[èe]ces? jointes?\b/i,
  /\bpi[èe]ces? jointes?\b/i,
  /\ben p\.?\s?j\.?(?![a-z])/i,
  /\b(vous|tu) trouver(ez|as) (ci-dessous |ici )?(en annexe|joint|jointe|attach[ée]e?)/i,
  /\ben annexe\b/i,
  /\bjoint[es]? [àa] (ce|cet|mon|notre) (mail|message|e-mail|courriel)\b/i,
  // EN
  /\battached\b/i,
  /\battachments?\b/i,
  /\benclosed\b/i,
  /\bplease find (enclosed|attached)\b/i,
  // DE
  /\banbei\b/i,
  /\bim anhang\b/i,
  /\bangeh[äa]ngt(e|en)?\b/i,
  /\bbeigef[üu]gt(e|en)?\b/i,
  /\bin der anlage\b/i,
  // LB
  /\ban der unlag\b/i,
  /\bugeh[äa]ng(t|en)?\b/i,
];

/** Négations à ne pas prendre pour une annonce (« sans pièce jointe », « no attachment »). */
const NEGATION_PJ_RE = /\b(sans|pas de|aucune?) pi[èe]ces? jointes?\b|\bno attachments?\b|\bwithout (an )?attachments?\b|\bohne anhang\b/gi;

/** Expressions annonçant une pièce jointe dans le texte RÉDIGÉ (le fil cité est ignoré). */
export function piecesJointesAnnoncees(text: string): string[] {
  const t = (text || '').replace(NEGATION_PJ_RE, ' ');
  const out: string[] = [];
  for (const re of ANNONCE_PJ_RES) {
    const m = t.match(re);
    if (m) out.push(m[0].trim());
  }
  return Array.from(new Set(out.map(s => s.toLowerCase())));
}

export function checkPieceJointe(ownText: string, piecesJointes: number | null | undefined): SendCheckProblem[] {
  if (piecesJointes === null || piecesJointes === undefined || piecesJointes > 0) return [];
  const annonces = piecesJointesAnnoncees(ownText);
  if (!annonces.length) return [];
  return [{
    code: 'piece_jointe',
    gravite: 'bloquant',
    message: `Pièce jointe annoncée (${annonces.slice(0, 2).map(a => `« ${a} »`).join(', ')}) mais aucune pièce jointe n'est ajoutée.`,
  }];
}

// ── Destinataires ────────────────────────────────────────────────────────────

/** Domaines de messagerie grand public : jamais rattachés à une société par le domaine. */
export const GENERIC_DOMAINS = new Set([
  'gmail.com', 'googlemail.com', 'outlook.com', 'hotmail.com', 'hotmail.fr', 'live.com', 'live.fr', 'msn.com',
  'yahoo.com', 'yahoo.fr', 'icloud.com', 'me.com', 'pt.lu', 'gmx.de', 'gmx.net', 'web.de', 'orange.fr', 'free.fr',
]);

const INTERNAL_DOMAIN = 'vibes.lu';

const mail = (s: string | undefined) => (s || '').trim().toLowerCase();
const domainOf = (email: string) => mail(email).split('@')[1] || '';

/** Nom normalisé pour comparer deux personnes : sans accents ni casse, mots triés (« Muller Anne » = « Anne Müller »). */
export function nomNormalise(nom: string | undefined): string {
  const ws = norm(nom || '').replace(/ß/g, 'ss').replace(/[^a-z\s'-]/g, ' ').split(/[\s'-]+/).filter(w => w.length > 1);
  return ws.sort().join(' ');
}

/** Société normalisée (comparaison souple : « GOOD VIBES SA » ~ « Good Vibes »). */
function societeNormalisee(s: string | undefined): string {
  return norm(s || '').replace(/\b(sa|sarl|sas|gmbh|ag|ltd|llc|s\.a\.|s\.a\.r\.l\.|asbl|scs|sc|bv|nv)\b/g, ' ')
    .replace(/[^a-z0-9]+/g, ' ').trim();
}

function memeSociete(a: string | undefined, b: string | undefined): boolean {
  const x = societeNormalisee(a);
  const y = societeNormalisee(b);
  if (!x || !y) return false;
  return x === y || x.includes(y) || y.includes(x);
}

function tousDestinataires(input: SendCheckInput): SendCheckRecipient[] {
  const moi = mail(input.moi);
  const seen = new Set<string>();
  const out: SendCheckRecipient[] = [];
  for (const r of [...(input.to || []), ...(input.cc || []), ...(input.bcc || [])]) {
    const e = mail(r?.email);
    if (!e || e === moi || seen.has(e)) continue;
    seen.add(e);
    out.push({ email: e, name: (r.name || '').trim() });
  }
  return out;
}

/**
 * Homonymes : pour chaque destinataire, les AUTRES contacts ATLAS du même nom (adresse différente).
 * Le nom du destinataire est celui d'Outlook, sinon celui de sa fiche ATLAS.
 */
export function checkHomonymes(destinataires: SendCheckRecipient[], contacts: SendCheckContact[]): SendCheckProblem[] {
  const out: SendCheckProblem[] = [];
  const parNom = new Map<string, SendCheckContact[]>();
  for (const c of contacts || []) {
    const k = nomNormalise(c.nom);
    if (!k || !k.includes(' ')) continue; // prénom seul : trop ambigu
    const list = parNom.get(k) || [];
    if (!list.some(x => mail(x.email) === mail(c.email))) list.push(c);
    parNom.set(k, list);
  }
  for (const d of destinataires) {
    if (domainOf(d.email) === INTERNAL_DOMAIN) continue;
    const fiche = (contacts || []).find(c => mail(c.email) === d.email);
    const nom = d.name && !d.name.includes('@') ? d.name : (fiche?.nom || '');
    const k = nomNormalise(nom);
    if (!k) continue;
    const autres = (parNom.get(k) || []).filter(c => mail(c.email) && mail(c.email) !== d.email);
    if (!autres.length) continue;
    const autre = autres[0];
    const ici = fiche?.societe ? ` (${fiche.societe})` : '';
    const la = autre.societe ? ` (${autre.societe})` : '';
    out.push({
      code: 'homonyme',
      gravite: 'avertissement',
      message: `Homonyme : « ${nom} » envoyé à ${d.email}${ici}, mais un autre contact du même nom existe : ${mail(autre.email)}${la}${autres.length > 1 ? ` (+${autres.length - 1})` : ''}. Vérifie l'adresse.`,
    });
  }
  return out;
}

/**
 * Destinataire client d'un AUTRE projet que celui du fil : sa société (fiche contact, sinon domaine)
 * est cliente d'au moins un projet ATLAS mais n'est pas le client du projet du fil.
 */
export function checkClientAutreProjet(
  destinataires: SendCheckRecipient[],
  projetDuFil: { nom: string; client: string } | null | undefined,
  contacts: SendCheckContact[],
  clients: string[],
  domainesClients: Record<string, string> = {},
): SendCheckProblem[] {
  if (!projetDuFil?.client) return [];
  const out: SendCheckProblem[] = [];
  for (const d of destinataires) {
    const dom = domainOf(d.email);
    if (!dom || dom === INTERNAL_DOMAIN) continue;
    const fiche = (contacts || []).find(c => mail(c.email) === d.email);
    let societe = fiche?.societe || '';
    if (!societe && !GENERIC_DOMAINS.has(dom)) societe = domainesClients[dom] || '';
    if (!societe) continue;
    if (memeSociete(societe, projetDuFil.client)) continue;
    if (!(clients || []).some(c => memeSociete(c, societe))) continue; // pas un client : fournisseur, partenaire…
    out.push({
      code: 'client_autre_projet',
      gravite: 'avertissement',
      message: `${d.name || d.email} (${societe}) est client d'un autre projet ; ce fil concerne « ${projetDuFil.nom} » (${projetDuFil.client}). Mauvais destinataire ?`,
    });
  }
  return out;
}

// ── « Répondre à tous » ──────────────────────────────────────────────────────

const SALUTATION_RE = /^\s*(bonjour|bonsoir|salut|coucou|cher|ch[èe]re|hello|hi|hey|dear|hallo|liebe[rs]?|guten (tag|morgen|abend)|moien|l[ée]iwen?|l[ée]if)\s*,?\s+([^\n,!:;.]{1,60}?)\s*[,!:\n]/i;
const COLLECTIF_RE = /\b(tous|toutes|tout le monde|[àa] tous|vous deux|mesdames|messieurs|[ée]quipe|team|all|everyone|everybody|both|guys|folks|zusammen|allerseits|alle|ihr beiden|beieneen|all zesummen|chers coll[èe]gues|dear (all|colleagues|team))\b|\b(et|and|und|an)\b|&/i;

/**
 * Personne saluée en tête du texte rédigé (« Bonjour Marc, » → « Marc »). null si la salutation est
 * collective (« Bonjour à tous », « Hi all », « Hallo zusammen », « Marc et Julie ») ou absente.
 */
export function personneSaluee(ownText: string): string | null {
  const first = (ownText || '').trim().split('\n').slice(0, 2).join('\n');
  const m = first.match(SALUTATION_RE);
  if (!m) return null;
  const qui = m[3].trim();
  if (!qui || COLLECTIF_RE.test(qui)) return null;
  // Titres seuls (« Bonjour Madame, ») : une personne quand même.
  const mots = qui.split(/\s+/).filter(Boolean);
  if (mots.length > 4) return null;
  return qui;
}

export function checkRepondreATous(input: { estReponse?: boolean; destinataires: SendCheckRecipient[]; ownText: string; seuil?: number }): SendCheckProblem[] {
  const seuil = input.seuil ?? 4;
  if (!input.estReponse || input.destinataires.length < seuil) return [];
  const qui = personneSaluee(input.ownText);
  if (!qui) return [];
  return [{
    code: 'repondre_a_tous',
    gravite: 'avertissement',
    message: `« Répondre à tous » : ${input.destinataires.length} destinataires alors que le message s'adresse à ${qui}. Retirer les personnes inutiles ?`,
  }];
}

// ── Tutoiement / vouvoiement par rapport au fil ──────────────────────────────

export function checkRegistre(input: { ownText: string; threadText: string; tutoiementConnu?: boolean; nbDestinataires: number }): SendCheckProblem[] {
  const own = input.ownText || '';
  const det = detectLang(own);
  if (det.lang !== 'FR') return [];
  const ownTu = tutoiementMarks(own);
  const ownVous = vouvoiementMarks(own);
  const thread = input.threadText || '';
  const threadFr = detectLang(thread).lang === 'FR';
  const threadTu = threadFr ? tutoiementMarks(thread).length : 0;
  const threadVous = threadFr ? vouvoiementMarks(thread).length : 0;
  const filTutoie = !!input.tutoiementConnu || (threadTu >= 2 && threadTu >= threadVous);
  const filVouvoie = !input.tutoiementConnu && threadVous >= 2 && threadTu === 0;
  const liste = (ms: string[]) => Array.from(new Set(ms.map(m => m.toLowerCase()))).slice(0, 3).map(m => `« ${m} »`).join(', ');

  if (ownTu.length && filVouvoie) {
    return [{ code: 'tutoiement', gravite: 'avertissement', message: `Tu tutoies (${liste(ownTu)}) alors que le fil vouvoie.` }];
  }
  if (!ownTu.length && ownVous.length && filTutoie && input.nbDestinataires <= 1) {
    return [{ code: 'vouvoiement', gravite: 'avertissement', message: `Tu vouvoies (${liste(ownVous)}) alors que vous vous tutoyez dans ce fil.` }];
  }
  if (ownTu.length && ownVous.length && input.nbDestinataires <= 1) {
    // « vous » au pluriel est normal avec plusieurs destinataires ; ici une seule personne.
    const vousFort = ownVous.filter(v => !/^vous$/i.test(v)).length > 0 || ownVous.length >= 2;
    if (vousFort) {
      return [{ code: 'registre_mixte', gravite: 'avertissement', message: `Tutoiement et vouvoiement mélangés (${liste(ownTu)} / ${liste(ownVous)}).` }];
    }
  }
  return [];
}

// ── Assistant inbox : relire avant envoi (07/10/2026) ───────────────────────

/** « Cordialement » dans le texte rédigé (charte de l'agence : « Bien à vous »). */
export function checkCordialement(ownText: string): SendCheckProblem[] {
  return /\bcordialement\b/i.test(ownText || '')
    ? [{ code: 'cordialement', gravite: 'avertissement', message: '« Cordialement » : la charte de l\'agence termine par « Bien à vous ».' }]
    : [];
}

/** Société d'un destinataire : fiche contact ATLAS, sinon domaine connu (hors messageries grand public). */
function societeDe(d: SendCheckRecipient, contacts: SendCheckContact[], parDomaine: Record<string, string>): string {
  const fiche = (contacts || []).find(c => mail(c.email) === d.email);
  if (fiche?.societe) return fiche.societe;
  const dom = domainOf(d.email);
  return dom && !GENERIC_DOMAINS.has(dom) ? parDomaine[dom] || '' : '';
}

/** Destinataires clients (société cliente d'au moins un projet, ou client du projet du fil). */
function destinatairesClients(input: SendCheckInput, destinataires: SendCheckRecipient[]): SendCheckRecipient[] {
  const clients = [...(input.clients || []), ...(input.projetDuFil?.client ? [input.projetDuFil.client] : [])];
  return destinataires.filter(d => {
    if (domainOf(d.email) === INTERNAL_DOMAIN) return false;
    const soc = societeDe(d, input.contacts || [], input.domainesClients || {});
    return !!soc && clients.some(c => memeSociete(c, soc));
  });
}

/** Tutoiement d'un client sans habitude connue de se tutoyer (français). */
export function checkTutoiementClient(ownText: string, clientsDest: SendCheckRecipient[], tutoiementConnu?: boolean): SendCheckProblem[] {
  if (!clientsDest.length || tutoiementConnu || detectLang(ownText).lang !== 'FR') return [];
  const tu = tutoiementMarks(ownText);
  if (!tu.length) return [];
  const qui = clientsDest[0].name || clientsDest[0].email;
  return [{ code: 'tutoiement_client', gravite: 'avertissement', message: `Tu tutoies (${Array.from(new Set(tu.map(x => x.toLowerCase()))).slice(0, 3).map(x => `« ${x} »`).join(', ')}) ${qui}, client : la règle est le vouvoiement des clients.` }];
}

const NOMS_LANGUE: Record<string, string> = { FR: 'français', EN: 'anglais', DE: 'allemand', LB: 'luxembourgeois' };
/** Langue du texte différente de la langue habituelle du correspondant (profil de conversation). */
export function checkLangue(ownText: string, habituelle: string | null | undefined): SendCheckProblem[] {
  const h = String(habituelle || '').toUpperCase().replace(/^LU$/, 'LB');
  if (!NOMS_LANGUE[h]) return [];
  const d = detectLang(ownText);
  if (!d.lang || d.confiance < 0.6 || d.mots < 12 || d.lang === h) return [];
  return [{ code: 'langue', gravite: 'avertissement', message: `Mail en ${NOMS_LANGUE[d.lang]}, alors que vous échangez d'habitude en ${NOMS_LANGUE[h]}.` }];
}

const MOIS_FR: Record<string, number> = { janvier: 1, fevrier: 2, mars: 3, avril: 4, mai: 5, juin: 6, juillet: 7, aout: 8, septembre: 9, octobre: 10, novembre: 11, decembre: 12, january: 1, february: 2, march: 3, april: 4, may: 5, june: 6, july: 7, august: 8, september: 9, october: 10, november: 11, december: 12 };

/** Dates citées ('YYYY-MM-DD') : « 15/10 », « 15.10.2026 », « 15 octobre ». Année : celle du projet, sinon l'année en cours. */
export function datesCitees(text: string, anneeDefaut: number): string[] {
  const out = new Set<string>();
  const t = norm(text || '');
  const ajoute = (d: number, m: number, y?: number) => {
    if (m < 1 || m > 12 || d < 1 || d > 31) return;
    out.add(`${y || anneeDefaut}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`);
  };
  let m: RegExpExecArray | null;
  const reNum = /(?<![\d:])(\d{1,2})[/.](\d{1,2})(?:[/.](\d{4}|\d{2}))?(?![\d:])/g;
  while ((m = reNum.exec(t))) { let y = m[3] ? +m[3] : undefined; if (y && y < 100) y += 2000; ajoute(+m[1], +m[2], y); }
  const reNom = new RegExp(`(?<!\\d)(\\d{1,2})(?:er)?\\s+(${Object.keys(MOIS_FR).join('|')})(?:\\s+(\\d{4}))?`, 'g');
  while ((m = reNom.exec(t))) ajoute(+m[1], MOIS_FR[m[2]], m[3] ? +m[3] : undefined);
  return [...out];
}

/**
 * Date citée proche du projet (14 jours autour) mais en dehors de ses dates : probable erreur
 * (« le 14/10 » pour un événement du 15/10 au 16/10).
 */
export function checkDateProjet(ownText: string, dates: { debut: string; fin?: string } | null | undefined, nomProjet?: string): SendCheckProblem[] {
  if (!dates?.debut || !/^\d{4}-\d{2}-\d{2}/.test(dates.debut)) return [];
  const debut = dates.debut.slice(0, 10), fin = (dates.fin || dates.debut).slice(0, 10);
  const t0 = Date.parse(debut), t1 = Date.parse(fin);
  const horsProjet = datesCitees(ownText, +debut.slice(0, 4)).filter(d => {
    const t = Date.parse(d);
    return Number.isFinite(t) && (d < debut || d > fin) && t >= t0 - 14 * 86_400_000 && t <= t1 + 14 * 86_400_000;
  });
  if (!horsProjet.length) return [];
  const fr = (iso: string) => iso.split('-').reverse().join('/');
  return [{ code: 'date_projet', gravite: 'avertissement', message: `Date citée ${horsProjet.slice(0, 2).map(fr).join(', ')} : le projet${nomProjet ? ` ${nomProjet}` : ''} est ${debut === fin ? `le ${fr(debut)}` : `du ${fr(debut)} au ${fr(fin)}`}.` }];
}

/** Adresse d'un FOURNISSEUR en À ou Cc d'un mail adressé à un client (les fournisseurs ne sont jamais exposés). */
export function checkFournisseurEnCopie(input: SendCheckInput, clientsDest: SendCheckRecipient[]): SendCheckProblem[] {
  if (!clientsDest.length) return [];
  const moi = mail(input.moi);
  const visibles = [...(input.to || []), ...(input.cc || [])].map(r => ({ email: mail(r.email), name: (r.name || '').trim() })).filter(r => r.email && r.email !== moi);
  const out: SendCheckProblem[] = [];
  for (const d of visibles) {
    if (domainOf(d.email) === INTERNAL_DOMAIN || clientsDest.some(c => c.email === d.email)) continue;
    const soc = societeDe(d, input.contacts || [], input.domainesFournisseurs || {});
    if (!soc || !(input.fournisseurs || []).some(f => memeSociete(f, soc))) continue;
    out.push({ code: 'fournisseur_en_copie', gravite: 'bloquant', message: `${d.name || d.email} (${soc}, fournisseur) est visible du client : les fournisseurs ne sont jamais exposés. Retire-le ou passe-le en Cci.` });
  }
  return out;
}

// ── Moment d'envoi (copie de src/utils/inbox-reponse-suivi.ts › momentEnvoi) ──

const FUSEAUX_TLD: Record<string, string> = {
  lu: 'Europe/Luxembourg', fr: 'Europe/Paris', be: 'Europe/Brussels', de: 'Europe/Berlin', nl: 'Europe/Amsterdam', ch: 'Europe/Zurich',
  at: 'Europe/Vienna', it: 'Europe/Rome', es: 'Europe/Madrid', pt: 'Europe/Lisbon', uk: 'Europe/London', ie: 'Europe/Dublin',
  pl: 'Europe/Warsaw', cz: 'Europe/Prague', dk: 'Europe/Copenhagen', se: 'Europe/Stockholm', no: 'Europe/Oslo', fi: 'Europe/Helsinki',
  gr: 'Europe/Athens', ro: 'Europe/Bucharest', us: 'America/New_York', ca: 'America/Toronto', br: 'America/Sao_Paulo',
  ae: 'Asia/Dubai', in: 'Asia/Kolkata', sg: 'Asia/Singapore', cn: 'Asia/Shanghai', jp: 'Asia/Tokyo', au: 'Australia/Sydney', ma: 'Africa/Casablanca',
};
export function fuseauDuDestinataire(email: string): string {
  const d = mail(email).split('@')[1] || '';
  if (d.endsWith('.co.uk')) return 'Europe/London';
  return FUSEAUX_TLD[d.split('.').pop() || ''] || 'Europe/Luxembourg';
}
function partiesFuseau(ms: number, tz: string): { y: number; mo: number; d: number; h: number; mi: number; wd: number } {
  const f = new Intl.DateTimeFormat('en-US', { timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', weekday: 'short', hour12: false });
  const ps = f.formatToParts(new Date(ms));
  const g = (t: string) => ps.find(p => p.type === t)?.value || '';
  const wd = ({ Sun: 0, Mon: 1, Tue: 2, Wed: 3, Thu: 4, Fri: 5, Sat: 6 } as Record<string, number>)[g('weekday')] ?? 0;
  return { y: +g('year'), mo: +g('month'), d: +g('day'), h: (+g('hour')) % 24, mi: +g('minute'), wd };
}
function murVersMs(y: number, mo: number, d: number, h: number, mi: number, tz: string): number {
  const guess = Date.UTC(y, mo - 1, d, h, mi);
  const off = (t: number) => { const p = partiesFuseau(t, tz); return Date.UTC(p.y, p.mo - 1, p.d, p.h, p.mi) - t; };
  let t = guess - off(guess);
  t = guess - off(t);
  return t;
}
const p2 = (n: number) => String(n).padStart(2, '0');

/**
 * Moment d'envoi : à l'heure du destinataire, hors 8 h - 18 h ou le week-end → conseil : prochain
 * matin ouvré à 8 h 30 chez lui (`conseil` = instant ISO, `libelle` lisible, heure de Luxembourg).
 */
export function momentEnvoi(now: number, email: string): { fuseau: string; heureLocale: string; horsHeures: boolean; conseil?: string; libelle?: string } {
  const tz = fuseauDuDestinataire(email);
  const p = partiesFuseau(now, tz);
  const heureLocale = `${p.h} h ${p2(p.mi)}`;
  const weekend = p.wd === 0 || p.wd === 6;
  if (!weekend && p.h >= 8 && p.h < 18) return { fuseau: tz, heureLocale, horsHeures: false };
  let t = Date.UTC(p.y, p.mo - 1, p.d);
  if (p.h >= 8 || weekend) t += 86_400_000;
  while ([0, 6].includes(new Date(t).getUTCDay())) t += 86_400_000;
  const j = new Date(t);
  const conseil = murVersMs(j.getUTCFullYear(), j.getUTCMonth() + 1, j.getUTCDate(), 8, 30, tz);
  const lux = partiesFuseau(conseil, 'Europe/Luxembourg');
  const jour = `${p2(j.getUTCDate())}/${p2(j.getUTCMonth() + 1)}`;
  const ici = tz === 'Europe/Luxembourg';
  return { fuseau: tz, heureLocale, horsHeures: true, conseil: new Date(conseil).toISOString(), libelle: `${jour} à 8 h 30${ici ? '' : ` chez ton correspondant (${lux.h} h ${p2(lux.mi)} à Luxembourg)`}` };
}

// ── Contrôle complet ─────────────────────────────────────────────────────────

export function checkAvantEnvoi(input: SendCheckInput): SendCheckProblem[] {
  const html = input.html || (input.text ? input.text.replace(/\n/g, '<br>') : '');
  // Tous les contrôles portent sur le texte RÉDIGÉ au-dessus du fil cité, signature exclue (08/10/2026).
  const ownText = texteRedige(html);
  const threadText = quotedText(html);
  const destinataires = tousDestinataires(input);
  const estReponse = input.estReponse ?? (/^\s*(re|aw|sv|antw|r)\s*:/i.test(input.subject || '') || !!threadText);
  return [
    ...checkPieceJointe(ownText, input.piecesJointes),
    ...checkHomonymes(destinataires, input.contacts || []),
    ...checkClientAutreProjet(destinataires, input.projetDuFil, input.contacts || [], input.clients || [], input.domainesClients || {}),
    ...checkRepondreATous({ estReponse, destinataires, ownText, seuil: input.seuilRepondreATous }),
    ...(() => {
      const registre = checkRegistre({ ownText, threadText, tutoiementConnu: input.tutoiementConnu, nbDestinataires: destinataires.length });
      const clientsDest = destinatairesClients(input, destinataires);
      return [
        ...registre,
        // Tutoiement d'un client : pas en double avec l'alerte « le fil vouvoie ».
        ...(registre.some(p => p.code === 'tutoiement') ? [] : checkTutoiementClient(ownText, clientsDest, input.tutoiementConnu)),
        ...checkFournisseurEnCopie(input, clientsDest),
      ];
    })(),
    ...checkLangue(ownText, input.langueHabituelle),
    ...checkDateProjet(ownText, input.datesProjet, input.projetDuFil?.nom),
    ...checkCordialement(ownText),
  ];
}

/** true si au moins un problème retient l'envoi (Smart Alerts : « soft block »). */
export function hasBlocking(problems: SendCheckProblem[]): boolean {
  return problems.some(p => p.gravite === 'bloquant');
}

/** Texte court pour la boîte de dialogue Smart Alerts (limitée à ~500 caractères par Outlook). */
export function smartAlertMessage(problems: SendCheckProblem[]): string {
  const lignes = problems.slice(0, 4).map(p => `• ${p.message}`);
  const txt = `ATLAS : à vérifier avant d'envoyer.\n${lignes.join('\n')}`;
  return txt.length > 480 ? `${txt.slice(0, 477)}…` : txt;
}

/** Cas CRITIQUES : l'envoi est retenu tant que ce n'est pas corrigé (fournisseur visible d'un client). */
export const CODES_CRITIQUES: ReadonlySet<SendCheckProblem['code']> = new Set(['fournisseur_en_copie']);

export interface DecisionEnvoi {
  allowEvent: boolean;
  /** Boîte Smart Alerts avec « Envoyer quand même » (Mailbox 1.14 : sendModeOverride PromptUser). */
  promptUser: boolean;
  /** Texte de la boîte (allowEvent false) ou de l'avertissement montré après coup (allowEvent true). */
  message: string;
}

/**
 * Décision Smart Alerts (08/10/2026). Le manifeste déclare SoftBlock : la boîte n'offre QUE « Ne pas
 * envoyer » (la personne doit corriger puis renvoyer), ce qui empêchait d'envoyer sur un simple
 * avertissement. Désormais :
 *  - cas critique (fournisseur visible d'un client) : retenu (SoftBlock) ;
 *  - autre problème, Mailbox 1.14 disponible (nouvel Outlook Mac, Windows, web) : retenu avec
 *    « Envoyer quand même » (sendModeOverride PromptUser) ;
 *  - autre problème sans 1.14 (Outlook classique ancien) : envoi AUTORISÉ, avertissement affiché à part.
 */
export function decisionEnvoi(problems: SendCheckProblem[], supporte114: boolean): DecisionEnvoi {
  if (!problems.length) return { allowEvent: true, promptUser: false, message: '' };
  const message = smartAlertMessage(problems);
  if (problems.some(p => CODES_CRITIQUES.has(p.code))) return { allowEvent: false, promptUser: false, message };
  if (supporte114) return { allowEvent: false, promptUser: true, message };
  return { allowEvent: true, promptUser: false, message };
}
