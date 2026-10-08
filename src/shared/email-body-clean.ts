// COPIE de src/utils/email-body-clean.ts (le complément est publié depuis son propre dossier et
// n'importe rien hors de outlook-addin/). Toute correction là-bas doit être recopiée ici :
// scripts/test-outlook-send-check.mjs vérifie que cette copie reste identique à l'original.

/**
 * email-body-clean.ts · Nettoyage du corps d'un email avant analyse IA ou affichage.
 *
 * Objectifs :
 *  - Retirer le HTML résiduel (si on reçoit du HTML brut)
 *  - Retirer les blocs de réponse cités (> at the start, "On X wrote:", "De: / Le ... a écrit :")
 *  - Retirer la signature Exclaimer ajoutée serveur côté Good Vibes (apparaît parfois dans les
 *    threads quotés). On retire aussi les signatures classiques détectées via marqueurs.
 *  - Préserver le contenu utile : ce que l'expéditeur a vraiment écrit dans CE message.
 *
 * Usage :
 *   import { cleanEmailBody } from './email-body-clean';
 *   const clean = cleanEmailBody(rawText);
 *
 * Heuristique testée sur Outlook FR/EN/DE. Pas de dépendance externe.
 */

// ── HTML → texte simple ─────────────────────────────────────────────────────

export function stripHtml(html: string): string {
  if (!html) return '';
  // Strip <style> et <script> avec leur contenu
  let s = html.replace(/<style[\s\S]*?<\/style>/gi, '');
  s = s.replace(/<script[\s\S]*?<\/script>/gi, '');
  // <br> et </p> → newlines
  s = s.replace(/<\/?(br|p|div|li)[^>]*>/gi, '\n');
  // Strip toutes les balises restantes
  s = s.replace(/<[^>]+>/g, '');
  // Décode entités HTML basiques
  s = s.replace(/&nbsp;/g, ' ')
       .replace(/&amp;/g, '&')
       .replace(/&lt;/g, '<')
       .replace(/&gt;/g, '>')
       .replace(/&quot;/g, '"')
       .replace(/&#39;/g, "'")
       .replace(/&apos;/g, "'");
  // Normalise les newlines multiples
  s = s.replace(/\r\n/g, '\n').replace(/\n{3,}/g, '\n\n');
  return s.trim();
}

// ── Historique cité en HTML ────────────────────────────────────────────────

/**
 * Début de l'historique cité dans un corps HTML (Outlook, Gmail, Apple Mail, Yahoo…), ou -1.
 * Coupé AVANT la conversion en texte : ces conteneurs disparaissent au stripHtml et les en-têtes
 * textuels (« De : … ») ne suffisent pas toujours (tableaux, lignes vides, adresses longues).
 */
const HTML_QUOTE_MARKERS_RX: RegExp[] = [
  /<div[^>]*\bid=["']?(?:divRplyFwdMsg|appendonsend|mail-editor-reference-message-container)["']?[^>]*>/i,
  /<div[^>]*\bclass=["'][^"']*\b(?:gmail_quote|yahoo_quoted|protonmail_quote|moz-cite-prefix|zmail_extra|x_divRplyFwdMsg)\b[^"']*["'][^>]*>/i,
  /<blockquote\b[^>]*(?:type=["']?cite|class=["']?(?:gmail_quote|cite))[^>]*>/i,
  /<blockquote\b[^>]*>/i,
  // Séparateur Outlook (ligne pleine) suivi de l'en-tête « De :/From: »
  /<hr[^>]*>\s*(?:<[^>]+>\s*)*(?:<b>)?\s*(?:De|From|Von)\s*:/i,
];

export function indexHtmlQuote(html: string): number {
  if (!html) return -1;
  let cutAt = -1;
  for (const rx of HTML_QUOTE_MARKERS_RX) {
    const m = rx.exec(html);
    if (m && (cutAt < 0 || m.index < cutAt)) cutAt = m.index;
  }
  return cutAt;
}

/** HTML sans l'historique cité (le contenu après le premier marqueur est ignoré). */
export function stripHtmlQuote(html: string): string {
  const i = indexHtmlQuote(html);
  return i < 0 ? html : html.slice(0, i);
}

// ── Quoted-reply detection (texte) ─────────────────────────────────────────

/**
 * Marqueurs de début de bloc cité (le contenu après est ignoré). Testés sur Outlook (Windows, nouveau
 * Mac, web), Gmail, Apple Mail, FR / EN / DE / IT / NL.
 */
const QUOTE_MARKERS_RX: RegExp[] = [
  // En-tête « De : X » suivi (dans les 4 lignes) d'un autre champ d'en-tête (Envoyé / Date / À / Objet…)
  /^[ \t]*\**(?:De|From|Von|Da|Van)\**[ \t]*:[^\n]*\n(?:[^\n]*\n){0,3}?[ \t]*\**(?:Envoy[ée]|Sent|Gesendet|Inviato|Verzonden|Date|Datum|À|A|To|An|Aan|Objet|Subject|Betreff|Oggetto|Onderwerp)\**[ \t]*:/im,
  // « Le X a écrit : » / « On X wrote: » / « Am X schrieb X: » / « Il X ha scritto: » / « Op X schreef X: »
  /^[ \t]*(?:Le|On|Am|Il|Op)\s[^\n]{1,200}?\s(?:a\s+[ée]crit|wrote|schrieb|ha\s+scritto|schreef)\s*:/im,
  // Même marqueur coupé sur 2 lignes (Apple Mail replie l'adresse) : « Le 7 oct. 2026 à 10:12, X\n<x@y> a écrit : »
  /^[ \t]*(?:Le|On|Am)\s[^\n]{1,160}\n[^\n]{0,160}?(?:a\s+[ée]crit|wrote|schrieb)\s*:/im,
  // « -----Original Message----- » / « -----Message d'origine----- » / « -----Ursprüngliche Nachricht----- »
  /^[ \t]*-{2,}\s*(?:Original\s+(?:Message|Nachricht|Appointment)|Message\s+d['’]origine|Urspr[üu]ngliche\s+Nachricht|Messaggio\s+originale|Oorspronkelijk\s+bericht|Forwarded\s+message|Message\s+transf[ée]r[ée])\s*-{2,}/im,
  // Séparateur Outlook (ligne de 20+ underscores) ou « ________ » avant l'en-tête
  /^[ \t]*_{8,}[ \t]*$/m,
  // Lignes citées « > » : dès la première ligne commençant par « > » suivie d'une autre
  /(?:^|\n)>[ \t]?[^\n]*\n>[ \t]?/,
  // Message cité de Teams / Outlook mobile : « Obtenir Outlook pour iOS/Android » suivi d'un en-tête
  /^[ \t]*(?:Get|Obtenir|Holen Sie sich)\s+Outlook\s+(?:for|pour|für)\s+(?:iOS|Android)[^\n]*\n(?:[^\n]*\n){0,3}?[ \t]*\**(?:De|From|Von)\**[ \t]*:/im,
];

/** Position du début de l'historique cité dans un texte, ou -1. */
export function indexQuotedReply(text: string): number {
  if (!text) return -1;
  let cutAt = -1;
  for (const rx of QUOTE_MARKERS_RX) {
    const m = rx.exec(text);
    if (m && (cutAt < 0 || m.index < cutAt)) cutAt = m.index;
  }
  return cutAt;
}

export function stripQuotedReply(text: string): string {
  if (!text) return '';
  const normalised = text.replace(/\r\n?/g, '\n');
  const cutAt = indexQuotedReply(normalised);
  const kept = cutAt < 0 ? normalised : normalised.slice(0, cutAt);
  // Lignes « > » isolées (citation partielle en tête de mail) : retirées, le reste est conservé.
  return kept.split('\n').filter((l) => !/^>/.test(l)).join('\n').trim();
}

// ── Disclaimers ────────────────────────────────────────────────────────────

/** Début des mentions légales de fin de mail (confidentialité, RGPD, « pensez à l'environnement »). */
const DISCLAIMER_RX: RegExp[] = [
  /^[ \t]*(?:\*{2,}|_{4,}|={4,}|-{4,})?[ \t]*(?:Ce (?:message|courriel|mail|e-mail)|Cet e-mail|Le contenu de ce (?:message|mail)|Les informations contenues)[^\n]{0,80}(?:confidenti|destin|usage exclusif|réserv)/im,
  /^[ \t]*(?:\*{2,}|_{4,}|={4,}|-{4,})?[ \t]*(?:This (?:e-?mail|message)|The (?:information|contents?) (?:contained|in this))[^\n]{0,80}(?:confidential|intended|privileged)/im,
  /^[ \t]*(?:\*{2,}|_{4,}|={4,}|-{4,})?[ \t]*(?:Diese (?:E-?Mail|Nachricht)|Der Inhalt dieser)[^\n]{0,80}(?:vertraulich|bestimmt)/im,
  /^[ \t]*(?:Pensez à l['’]environnement|Please consider the environment|Think before you print|Avant d['’]imprimer)/im,
  /^[ \t]*(?:DISCLAIMER|AVERTISSEMENT|CONFIDENTIALIT[ÉE])\s*:?\s*$/im,
];

export function stripDisclaimer(text: string): string {
  if (!text) return '';
  let cutAt = text.length;
  for (const rx of DISCLAIMER_RX) {
    const m = rx.exec(text);
    // Un disclaimer vit en fin de mail : on ne coupe pas dans la première moitié d'un mail court.
    if (m && m.index < cutAt && (m.index > 120 || text.length < 240)) cutAt = m.index;
  }
  return text.slice(0, cutAt).trim();
}

// ── Signature detection ────────────────────────────────────────────────────

/** Délimiteurs classiques de signature. */
const SIGNATURE_DELIM_RX: RegExp[] = [
  /^-- ?$/m,               // RFC standard
  /^_{4,}$/m,              // ligne de underscores
  /^={4,}$/m,              // ligne de equals
];

/** Mots-clés Exclaimer Good Vibes (signature serveur). À adapter si la signature évolue. */
const EXCLAIMER_HINTS = [
  'good vibes',
  'managing director',
  'vibes.lu',
  '+352',           // numéros Lux dans la signature
  'event management',
  'we make events',  // tagline éventuelle
];

/** Formule de politesse finale (FR / EN / DE / LU) : ce qui suit est une signature. */
const CLOSING_LINE_RX = /^[ \t]*(?:bien (?:à|a) (?:vous|toi)|(?:bien |très )?cordialement|(?:meilleures|bonnes) salutations|salutations distinguées|belle (?:journée|semaine|fin de journée)|bonne (?:journée|soirée|semaine|réception|continuation)|à (?:très )?(?:bientôt|vite|tout à l['’]heure|lundi|mardi|mercredi|jeudi|vendredi|demain)|(?:kind|best|warm) regards|regards|cheers|sincerely|mit (?:freundlichen|besten) grüßen|(?:viele|liebe|beste|schöne) grüße|freundliche grüße|mat beschte gréiss|bis geschwënn|schéinen dag|léif gréiss)\s*[,!.]?[ \t]*$/im;

/**
 * Retire la signature : formule de politesse finale (gardée, tout ce qui suit est coupé),
 * sinon délimiteur "--", sinon bloc en fin de mail avec ≥2 hints Exclaimer dans les dernières lignes.
 */
export function stripSignature(text: string): string {
  if (!text) return '';
  // 0) Formule de politesse finale : la signature (Exclaimer ou tapée) vient toujours après.
  {
    const lines = text.split('\n');
    for (let i = lines.length - 1; i >= 0; i--) {
      if (CLOSING_LINE_RX.test(lines[i])) {
        text = lines.slice(0, i + 1).join('\n').trim();
        break;
      }
    }
  }
  // 1) Délimiteur explicite → tout après est signature
  for (const rx of SIGNATURE_DELIM_RX) {
    const m = rx.exec(text);
    if (m) return text.slice(0, m.index).trim();
  }
  // 2) Heuristique Exclaimer : si les 6 dernières lignes contiennent ≥2 hints, on coupe avant
  const lines = text.split('\n');
  const tailStart = Math.max(0, lines.length - 8);
  const tail = lines.slice(tailStart).join('\n').toLowerCase();
  const hintsFound = EXCLAIMER_HINTS.filter((h) => tail.includes(h)).length;
  if (hintsFound >= 2) {
    // Cherche la première ligne du tail qui contient un hint → on coupe avant
    for (let i = tailStart; i < lines.length; i++) {
      const lc = lines[i].toLowerCase();
      if (EXCLAIMER_HINTS.some((h) => lc.includes(h))) {
        return lines.slice(0, i).join('\n').trim();
      }
    }
  }
  return text.trim();
}

// ── Pipeline complète ──────────────────────────────────────────────────────

/**
 * Nettoyage complet : HTML → texte → strip quoted → strip signature.
 * Si `maxLen` fourni, tronque la fin (préserve le début, qui est le plus important).
 */
export function cleanEmailBody(raw: string, opts: { isHtml?: boolean; maxLen?: number } = {}): string {
  if (!raw) return '';
  // HTML : l'historique cité est coupé AVANT la conversion (blockquote, gmail_quote, divRplyFwdMsg…).
  let s = opts.isHtml ? stripHtml(stripHtmlQuote(raw)) : raw;
  s = stripQuotedReply(s);
  s = stripDisclaimer(s);
  s = stripSignature(s);
  // Normalise espaces multiples
  s = s.replace(/[ \t]+/g, ' ').replace(/\n{3,}/g, '\n\n').trim();
  if (opts.maxLen && s.length > opts.maxLen) {
    s = s.slice(0, opts.maxLen) + '…[tronqué]';
  }
  return s;
}

// ── Détection du rôle destinataire (To / Cc / Bcc) ────────────────────────

export type RecipientRole = 'to' | 'cc' | 'bcc' | 'unknown';

/**
 * Détecte le rôle de `userEmail` dans la liste des destinataires.
 * - 'to'  : action attendue
 * - 'cc'  : pour information
 * - 'bcc' : passif (newsletter, blast)
 * - 'unknown' : pas trouvé (mail forwardé ou liste de diffusion)
 *
 * Note : Microsoft Graph n'expose pas les Bcc des mails reçus (raison RFC).
 * Si l'utilisateur n'est ni en To: ni en Cc:, on retourne 'bcc' par déduction.
 */
export function detectRecipientRole(
  userEmail: string,
  to: Array<{ email: string }> = [],
  cc: Array<{ email: string }> = [],
): RecipientRole {
  if (!userEmail) return 'unknown';
  const me = userEmail.toLowerCase().trim();
  if (to.some((r) => (r.email || '').toLowerCase() === me)) return 'to';
  if (cc.some((r) => (r.email || '').toLowerCase() === me)) return 'cc';
  // Ni To ni Cc → mail envoyé via Bcc ou liste de diffusion
  if (to.length > 0 || cc.length > 0) return 'bcc';
  return 'unknown';
}
