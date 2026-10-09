/**
 * actions-mail.ts · « Traité », « Demain 8 h », « Archiver » pour le mail ouvert, par le WORKER
 * (lot 1 « vitesse », 09/10/2026).
 *
 * Les boutons du ruban (commands.ts) et l'onglet « Classer » (ia-panel.ts) passaient par le jeton de
 * boîte du complément (graph.ts › getApiContext : catégories Outlook, déplacement dans le dossier),
 * que Microsoft ne donne plus : ils échouaient, souvent après une longue attente. Tout passe
 * maintenant par le worker, dans la boîte de la personne (règle rangementClicAutorise : clic de la
 * propriétaire dans SA boîte personnelle) :
 *   - Traité  : fiche de tri ATLAS « Traité » (si elle existe) + mail marqué lu (agent/assistant/groupe) ;
 *   - Demain  : « répondre plus tard » de l'agent (agent/plus-tard, demain 8 h, jours ouvrés) + fiche « Reporté » ;
 *   - Archiver : rangé dans « Archives » par le worker (annulable) + fiche « Archivé ».
 * Chaque action réussit si au moins une des deux voies a marché ; sinon une phrase lisible.
 */
import { archiveTag, getEmailTagByConversationId, getEmailTagByEmailId, markTagDone, snoozeTag, type EmailTag } from './airtable';
import { workerRequest } from './worker';
import { invalidateJournee } from './agent';
import { humanError } from './net';

export interface MailCourant {
  /** internetMessageId (clé de l'agent). */
  messageId: string;
  /** Id REST (Graph) du mail, clé des fiches de tri ATLAS. */
  restId: string;
  conversationId: string;
}

export interface ResultatAction {
  /** Phrase courte pour la personne. */
  message: string;
  /** Action annulable par agent/annuler (archivage). */
  actionId?: string;
}

/** Fiche de tri ATLAS du mail (par id REST, sinon par fil) ; null si aucune. */
export async function ficheDeTri(m: Pick<MailCourant, 'restId' | 'conversationId'>): Promise<EmailTag | null> {
  let tag = m.restId ? await getEmailTagByEmailId(m.restId).catch(() => null) : null;
  if (!tag && m.conversationId) tag = await getEmailTagByConversationId(m.conversationId).catch(() => null);
  return tag;
}

const INCONNU = 'ATLAS ne connaît pas encore ce mail : ouvre le panneau ATLAS, puis réessaie dans un instant.';

async function groupe(action: 'archiver' | 'marquer-lu', messageId: string): Promise<{ actionIds?: string[]; dossier?: { chemin: string } } | null> {
  if (!messageId) return null;
  try {
    return await workerRequest<{ actionIds?: string[]; dossier?: { chemin: string } }>('POST', 'agent/assistant/groupe', { action, messageIds: [messageId] }, { timeoutMs: 30_000 });
  } catch (e) {
    console.warn(`[actions-mail] ${action} :`, humanError(e));
    return null;
  }
}

export async function marquerTraite(m: MailCourant, tag?: EmailTag | null): Promise<ResultatAction> {
  const fiche = tag === undefined ? await ficheDeTri(m) : tag;
  const [okFiche, lu] = await Promise.all([fiche ? markTagDone(fiche.id) : Promise.resolve(false), groupe('marquer-lu', m.messageId)]);
  if (!okFiche && !lu) throw new Error(INCONNU);
  invalidateJournee();
  return { message: 'Traité' };
}

export async function reporterDemain(m: MailCourant, tag?: EmailTag | null): Promise<ResultatAction> {
  const fiche = tag === undefined ? await ficheDeTri(m) : tag;
  const plusTard = m.messageId
    ? workerRequest<{ jusqua?: string }>('POST', 'agent/plus-tard', { messageId: m.messageId, quand: 'demain' }, { timeoutMs: 30_000 })
      .then(() => true, (e) => { console.warn('[actions-mail] plus tard :', humanError(e)); return false; })
    : Promise.resolve(false);
  const [okFiche, okAgent] = await Promise.all([fiche ? snoozeTag(fiche.id) : Promise.resolve(false), plusTard]);
  if (!okFiche && !okAgent) throw new Error(INCONNU);
  invalidateJournee();
  return { message: 'Reporté à demain 8 h' };
}

export async function archiverMail(m: MailCourant, tag?: EmailTag | null): Promise<ResultatAction> {
  const fiche = tag === undefined ? await ficheDeTri(m) : tag;
  const [okFiche, r] = await Promise.all([fiche ? archiveTag(fiche.id) : Promise.resolve(false), groupe('archiver', m.messageId)]);
  if (!okFiche && !r) throw new Error(INCONNU);
  invalidateJournee();
  const actionId = r?.actionIds?.[0];
  return { message: r ? `Archivé dans ${r.dossier?.chemin || 'Archives'}` : 'Archivé dans ATLAS (mail laissé en place)', ...(actionId ? { actionId } : {}) };
}
