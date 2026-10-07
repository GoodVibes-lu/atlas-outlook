/**
 * send-check-office.ts — Lecture du message en cours de rédaction (Office.js) et exécution du
 * contrôle avant envoi, SANS dépendance d'interface (ni DOM, ni MSAL, ni appels au worker) : ce
 * module est partagé par l'onglet « Relire avant envoi » (send-check-panel.ts), le gestionnaire
 * Smart Alerts du runtime navigateur (commands.ts : nouvel Outlook pour Mac, web, nouvel Outlook
 * Windows) et le runtime JavaScript seul d'Outlook classique Windows (launch-event.ts).
 *
 * Les contrôles eux-mêmes sont purs (utils/send-check.ts) : les trois surfaces appliquent EXACTEMENT
 * les mêmes.
 */
import { checkAvantEnvoi, smartAlertMessage, type SendCheckInput, type SendCheckProblem, type SendCheckRecipient } from './send-check';
import { supportsMailbox } from '../api/platform';

export function getAsyncValue<T>(getter: ((cb: (r: Office.AsyncResult<T>) => void) => void) | undefined, fallback: T): Promise<T> {
  return new Promise((resolve) => {
    if (!getter) { resolve(fallback); return; }
    try {
      getter((r) => resolve(r.status === Office.AsyncResultStatus.Succeeded ? (r.value ?? fallback) : fallback));
    } catch { resolve(fallback); }
  });
}

const toRecipients = (list: Office.EmailAddressDetails[] | undefined): SendCheckRecipient[] =>
  (list || []).map(r => ({ email: String(r?.emailAddress || ''), name: String(r?.displayName || '') })).filter(r => r.email);

export type ComposeItemInput = SendCheckInput & { conversationId: string };

/** Message en cours de rédaction → entrée des contrôles (sans les données ATLAS). */
export async function readComposeItem(): Promise<ComposeItemInput> {
  const item = Office.context.mailbox?.item as any;
  if (!item) throw new Error('Aucun message en cours de rédaction.');
  const [subject, to, cc, bcc, html] = await Promise.all([
    getAsyncValue<string>(item.subject?.getAsync?.bind(item.subject), ''),
    getAsyncValue<Office.EmailAddressDetails[]>(item.to?.getAsync?.bind(item.to), []),
    getAsyncValue<Office.EmailAddressDetails[]>(item.cc?.getAsync?.bind(item.cc), []),
    getAsyncValue<Office.EmailAddressDetails[]>(item.bcc?.getAsync?.bind(item.bcc), []),
    getAsyncValue<string>(item.body?.getAsync ? (cb: any) => item.body.getAsync(Office.CoercionType.Html, cb) : undefined, ''),
  ]);
  // Pièces jointes : Mailbox 1.8 en rédaction ; sinon inconnu (contrôle sauté plutôt que faux positif).
  let piecesJointes: number | null = null;
  if (supportsMailbox('1.8') && item.getAttachmentsAsync) {
    const list = await getAsyncValue<Office.AttachmentDetailsCompose[] | null>(item.getAttachmentsAsync.bind(item), null);
    if (list) piecesJointes = list.filter(a => !a.isInline).length;
  }
  return {
    subject,
    html,
    to: toRecipients(to),
    cc: toRecipients(cc),
    bcc: toRecipients(bcc),
    moi: String(Office.context.mailbox?.userProfile?.emailAddress || ''),
    piecesJointes,
    conversationId: String(item.conversationId || ''),
  };
}

/** Budget total du gestionnaire d'envoi (Outlook en accorde ~300 s ; on ne retient jamais l'envoi plus de 5 s). */
export const BUDGET_ENVOI_MS = 4500;

const attendre = <T>(ms: number, valeur: T): Promise<T> => new Promise(r => setTimeout(() => r(valeur), ms));

export interface ControleEnvoiOptions {
  /** Complète l'entrée avec les données ATLAS (réseau) ; ignoré au-delà du budget ou en cas d'erreur. */
  enrichir?: (input: ComposeItemInput) => Promise<SendCheckInput>;
  budgetMs?: number;
}

/**
 * Contrôle avant envoi dans le budget : lecture du message, enrichissement ATLAS facultatif (au plus
 * ~2/3 du budget, sinon contrôles locaux seuls), puis les mêmes contrôles que le bouton manuel.
 * `null` = budget dépassé ou lecture impossible : l'appelant AUTORISE l'envoi.
 */
export async function controleAvantEnvoi(opts: ControleEnvoiOptions = {}): Promise<SendCheckProblem[] | null> {
  const budget = opts.budgetMs ?? BUDGET_ENVOI_MS;
  const debut = Date.now();
  const travail = (async (): Promise<SendCheckProblem[]> => {
    const local = await readComposeItem();
    let input: SendCheckInput = local;
    if (opts.enrichir) {
      const reste = Math.max(0, Math.round(budget * 0.66) - (Date.now() - debut));
      input = await Promise.race([opts.enrichir(local).catch(() => local), attendre(reste, local)]);
    }
    return checkAvantEnvoi(input);
  })();
  try {
    return await Promise.race([travail, attendre(budget, null)]);
  } catch {
    return null;
  }
}

/**
 * Gestionnaire Smart Alerts (`OnMessageSend`, SendMode « SoftBlock ») : si quelque chose est à
 * signaler, l'envoi est RETENU avec le message et la personne peut corriger ou envoyer quand même.
 * Jamais bloqué par ATLAS : API absente, erreur, réseau en panne ou budget dépassé = envoi autorisé.
 */
export async function gererOnMessageSend(event: any, opts: ControleEnvoiOptions = {}): Promise<void> {
  let termine = false;
  const terminer = (o: { allowEvent: boolean; errorMessage?: string }) => {
    if (termine) return;
    termine = true;
    try { event.completed(o); } catch { /* déjà terminé */ }
  };
  // Garde absolue : quoi qu'il arrive, l'envoi part dans le budget.
  const garde = setTimeout(() => terminer({ allowEvent: true }), (opts.budgetMs ?? BUDGET_ENVOI_MS) + 300);
  try {
    if (!supportsMailbox('1.12')) { terminer({ allowEvent: true }); return; }
    const problems = await controleAvantEnvoi(opts);
    if (!problems || !problems.length) { terminer({ allowEvent: true }); return; }
    terminer({ allowEvent: false, errorMessage: smartAlertMessage(problems) });
  } catch (e) {
    console.warn('[ATLAS] contrôle avant envoi impossible, envoi autorisé :', e);
    terminer({ allowEvent: true });
  } finally {
    clearTimeout(garde);
  }
}
