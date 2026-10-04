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
 *  - tutoiement / vouvoiement incohérent avec le fil cité (français).
 *
 * Les fonctions de texte (ownPart, visibleText, detectLang, tutoiementMarks) sont une COPIE de
 * `src/utils/client-send-check.ts` (racine du dépôt, contrôle avant envoi client d'ATLAS) : le
 * complément est publié depuis son propre dossier et n'importe rien hors de `outlook-addin/`.
 * Toute correction de ces fonctions là-bas doit être recopiée ici.
 *
 * Gravité : « bloquant » = l'envoi est retenu (Smart Alerts « soft block » : l'utilisateur peut
 * envoyer quand même) ; « avertissement » = signalé seulement. Rien n'est jamais envoyé ni modifié.
 */

export type SendCheckGravite = 'bloquant' | 'avertissement';

export interface SendCheckProblem {
  code: 'piece_jointe' | 'homonyme' | 'client_autre_projet' | 'repondre_a_tous' | 'tutoiement' | 'vouvoiement' | 'registre_mixte';
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
  const cuts = [
    /<blockquote[\s\S]*$/i,
    /<div[^>]*id=["']?(divRplyFwdMsg|appendonsend)["']?[\s\S]*$/i,
    /<hr[^>]*>[\s\S]*?(De|From|Von)\s*:[\s\S]*$/i,
  ];
  for (const re of cuts) h = h.replace(re, '');
  return h;
}

const QUOTE_TEXT_RE = /\n\s*(-{2,}\s*(Original|Message d'origine|Ursprüngliche)|(De|From|Von)\s*:.*\n\s*(Envoyé|Sent|Gesendet)\s*:|On .{4,80} wrote:|Le .{4,80} a écrit\s*:|Am .{4,80} schrieb)/i;

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
  const m = t.search(QUOTE_TEXT_RE);
  return (m > 0 ? t.slice(0, m) : t).trim();
}

/** Texte du fil CITÉ (ce qui suit la partie rédigée), vide pour un nouveau message. */
export function quotedText(html: string): string {
  const h = html || '';
  const own = ownPart(h);
  const quotedHtml = h.slice(own.length);
  // Fil cité en texte brut, resté dans la partie « rédigée » (Outlook en mode texte).
  const t = htmlToText(own);
  const m = t.search(QUOTE_TEXT_RE);
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

// ── Contrôle complet ─────────────────────────────────────────────────────────

export function checkAvantEnvoi(input: SendCheckInput): SendCheckProblem[] {
  const html = input.html || (input.text ? input.text.replace(/\n/g, '<br>') : '');
  const ownText = visibleText(ownPart(html));
  const threadText = quotedText(html);
  const destinataires = tousDestinataires(input);
  const estReponse = input.estReponse ?? (/^\s*(re|aw|sv|antw|r)\s*:/i.test(input.subject || '') || !!threadText);
  return [
    ...checkPieceJointe(ownText, input.piecesJointes),
    ...checkHomonymes(destinataires, input.contacts || []),
    ...checkClientAutreProjet(destinataires, input.projetDuFil, input.contacts || [], input.clients || [], input.domainesClients || {}),
    ...checkRepondreATous({ estReponse, destinataires, ownText, seuil: input.seuilRepondreATous }),
    ...checkRegistre({ ownText, threadText, tutoiementConnu: input.tutoiementConnu, nbDestinataires: destinataires.length }),
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
