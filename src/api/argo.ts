/**
 * ARGO pour le complément Outlook ATLAS : analyse de mail et adaptation de ton.
 *
 * Depuis le 03/10/2026, toute génération IA passe par le worker (`/api/plugin/atlas/ai/*`) et
 * l'Agency Brain (journal des coûts, plafond, cache) : le complément ne détient plus de clé
 * Anthropic et n'appelle plus l'API en direct. Les prompts sont construits côté worker
 * (worker/portal-api/routes/plugin-atlas.ts) ; ici on n'envoie que des données structurées.
 */

import type { ArgoProfile } from '../types';
import { callAtlasWorker } from './worker';

/**
 * Suggère un chemin de dossier cohérent avec la structure existante de l'utilisateur
 * (ex. « Clients/Vossloh »). Chaîne vide si l'IA est indisponible.
 */
export async function suggestFolderPath(args: {
  existingFolders: string[]; // ex: ['Administration', 'Administration/Bureau', 'Markcom', 'Clients/Niessen', ...]
  iaCategory?: string;        // ex: 'demande_devis'
  iaCategoryLabel?: string;   // ex: '💼 Demande devis'
  senderName: string;         // ex: 'Aline Fontana'
  senderEmail: string;        // ex: 'aline.fontana@emotion.lu'
  subject: string;            // ex: 'RE: Devis Vossloh BBQ'
  summary?: string;           // ex: 'Vossloh demande un devis...'
}): Promise<string> {
  try {
    const r = await callAtlasWorker<{ path?: string }>('ai/folder-path', {
      ...args,
      existingFolders: args.existingFolders.slice(0, 150),
    });
    return String(r.path || '').slice(0, 200);
  } catch (e) {
    console.warn('[argo] suggestFolderPath failed:', e);
    return '';
  }
}

// ── Email Analysis ──

export interface EmailAnalysis {
  denomination: string;
  client?: string;
  typeEvenement?: string;
  debut?: string;
  lieu?: string;
  descriptif?: string;
}

export async function analyzeEmailForProjet(
  subject: string, fromName: string, fromEmail: string, bodyPreview: string
): Promise<EmailAnalysis> {
  try {
    const r = await callAtlasWorker<{ analysis?: Partial<EmailAnalysis> }>('ai/projet-extract', {
      subject, fromName, fromEmail, bodyPreview: (bodyPreview || '').slice(0, 500),
    });
    const a = r.analysis || {};
    return { ...a, denomination: a.denomination || subject, client: a.client || fromName };
  } catch {
    return { denomination: subject, client: fromName };
  }
}

// ── Analyze received email for response tone ──

export async function analyzeReceivedEmail(
  subject: string, body: string, senderName: string
): Promise<{ sentiment: string; urgence: string; tonUtilise: string; suggestions: string[] }> {
  const fallback = { sentiment: 'neutre', urgence: 'normal', tonUtilise: 'professionnel', suggestions: [] as string[] };
  try {
    const r = await callAtlasWorker<{ analysis?: typeof fallback }>('ai/tone', {
      subject, body: (body || '').slice(0, 800), senderName,
    });
    return r.analysis || fallback;
  } catch {
    return fallback;
  }
}

// ── Email Summary ──

export async function summarizeEmail(subject: string, body: string, senderName: string): Promise<string> {
  try {
    const r = await callAtlasWorker<{ summary?: string }>('ai/summary', {
      subject, body: (body || '').slice(0, 1500), senderName,
    });
    return r.summary || '';
  } catch { return ''; }
}

// ── Quick Reply Suggestions ──

export interface QuickReplySuggestion {
  label: string;
  tone: string;
  body: string;
}

export async function generateQuickReplies(
  subject: string, body: string, senderName: string, profile: ArgoProfile | null, userName: string
): Promise<QuickReplySuggestion[]> {
  try {
    const r = await callAtlasWorker<{ replies?: QuickReplySuggestion[] }>('ai/quick-replies', {
      subject, body: (body || '').slice(0, 800), senderName, profile, userName,
    });
    return Array.isArray(r.replies) ? r.replies.slice(0, 3) : [];
  } catch { return []; }
}

// ── Free-form AI Reply ──

export async function generateFreeReply(
  subject: string, body: string, senderName: string,
  instruction: string, profile: ArgoProfile | null, userName: string
): Promise<string> {
  const r = await callAtlasWorker<{ html?: string }>('ai/free-reply', {
    subject, body: (body || '').slice(0, 1500), senderName, instruction, profile, userName,
  });
  return String(r.html || '').trim();
}

// ── Tone Adaptation ──

export function getSalutation(profile: ArgoProfile | null, senderName?: string): string {
  if (!profile) return 'Bonjour,';

  const lang = profile.languePreferee || 'FR';
  const isTu = profile.tonPrefere === 'Amical' ||
    (senderName && profile.tutoiementAvec.some(n => n.toLowerCase().includes(senderName.toLowerCase())));
  const prenom = profile.prenom || '';

  switch (lang) {
    case 'EN': return prenom ? `Dear ${prenom},` : 'Dear Sir/Madam,';
    case 'DE': return isTu ? `Hallo ${prenom},` : `Sehr geehrte Damen und Herren,`;
    case 'LU': return prenom ? `Gudde Moien ${prenom},` : 'Gudde Moien,';
    default:
      if (isTu && prenom) return `Salut ${prenom},`;
      if (prenom) return `Bonjour ${prenom},`;
      return 'Bonjour,';
  }
}

export function getClosing(profile: ArgoProfile | null, senderName?: string): string {
  if (!profile) return 'Cordialement,';

  const lang = profile.languePreferee || 'FR';
  const isTu = profile.tonPrefere === 'Amical' ||
    (senderName && profile.tutoiementAvec.some(n => n.toLowerCase().includes(senderName.toLowerCase())));

  switch (lang) {
    case 'EN': return isTu ? 'Best regards,' : 'Kind regards,';
    case 'DE': return isTu ? 'Viele Grüße,' : 'Mit freundlichen Grüßen,';
    case 'LU': return 'Mat beschte Gréiss,';
    default: return isTu ? 'A bientôt,' : 'Cordialement,';
  }
}

export async function adaptEmailBody(
  html: string, profile: ArgoProfile | null, senderName?: string
): Promise<string> {
  if (!profile || !html) return html;
  try {
    const r = await callAtlasWorker<{ html?: string }>('ai/adapt', { html, profile, senderName: senderName || '' });
    const adapted = String(r.html || '').trim();
    if (adapted && adapted.includes('<')) return adapted;
    return simpleAdapt(html, profile, senderName);
  } catch {
    return simpleAdapt(html, profile, senderName);
  }
}

function simpleAdapt(html: string, profile: ArgoProfile, senderName?: string): string {
  const salutation = getSalutation(profile, senderName);
  const closing = getClosing(profile, senderName);
  const isTu = profile.tonPrefere === 'Amical' ||
    (senderName && profile.tutoiementAvec.some(n => n.toLowerCase().includes((senderName || '').toLowerCase())));

  let adapted = html;
  adapted = adapted.replace(/<p>Bonjour,<\/p>/i, `<p>${salutation}</p>`);
  adapted = adapted.replace(/^Bonjour,/im, salutation);
  adapted = adapted.replace(/<p>Cordialement,<\/p>/i, `<p>${closing}</p>`);
  adapted = adapted.replace(/Cordialement,/i, closing);

  if (isTu) {
    adapted = adapted.replace(/Veuillez trouver/gi, 'Tu trouveras');
    adapted = adapted.replace(/Veuillez cliquer/gi, 'Clique');
    adapted = adapted.replace(/Veuillez/gi, "N'hésite pas à");
    adapted = adapted.replace(/votre espace/gi, 'ton espace');
    adapted = adapted.replace(/votre demande/gi, 'ta demande');
    adapted = adapted.replace(/votre email/gi, 'ton email');
    adapted = adapted.replace(/Vous trouverez/gi, 'Tu trouveras');
    adapted = adapted.replace(/vous trouverez/gi, 'tu trouveras');
    adapted = adapted.replace(/N'hésitez pas/gi, "N'hésite pas");
    adapted = adapted.replace(/n'hésitez pas/gi, "n'hésite pas");
    adapted = adapted.replace(/je vous répondrai/gi, 'je te répondrai');
  }

  return adapted;
}
