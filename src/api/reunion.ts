/**
 * reunion.ts · ordre du jour collaboratif depuis une invitation Outlook (04/10/2026).
 *
 * Le complément ne détient aucun secret : tout passe par `/api/plugin/atlas/reunion/*` du worker
 * avec le jeton Microsoft de l'utilisateur (identité = jeton). Lien invitation ↔ réunion ATLAS :
 * le worker retrouve l'événement dans la boîte de l'appelant (identifiant REST) et le reconnaît
 * par son iCalUId + son début ; « Proposer un point » n'apparaît que si ATLAS suit cette réunion.
 * Règles (invité seulement, jusqu'à la veille 17 h, arbitrage réservé à l'organisateur dans ATLAS) :
 * worker/utils/reunion-oj-ops.ts. Office.js n'est utilisé que pour lire l'identifiant de l'élément.
 */
import { callAtlasWorker } from './worker';

export interface ReunionPoint {
  id: string; titre: string; dureeMin: number; projetNo?: string; statut: 'proposee' | 'acceptee' | 'refusee' | 'reportee'; motif?: string;
  piece?: { nom: string; url: string };
  delegation?: { mode: 'collegue'; email: string; nom?: string } | { mode: 'note'; texte: string } | { mode: 'audio'; dureeS?: number };
}
export interface ReunionVue {
  key: string; titre: string; debutISO: string; organisateur: string; estOrganisateur: boolean;
  limiteISO: string; ouverte: boolean; mesPoints: ReunionPoint[];
}

/** Identifiant REST + nature (message d'invitation ou événement du calendrier) de l'élément ouvert ; null hors réunion. */
export function lireElementReunion(): { restId: string; kind: 'message' | 'event' } | null {
  try {
    const item: any = Office.context.mailbox?.item;
    if (!item?.itemId) return null;
    const estRdv = item.itemType === Office.MailboxEnums.ItemType.Appointment;
    const estInvitation = item.itemType === Office.MailboxEnums.ItemType.Message && /^IPM\.Schedule\.Meeting\.(Request|Resp|Canceled)?/i.test(String(item.itemClass || ''));
    if (!estRdv && !estInvitation) return null;
    const restId = Office.context.mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0);
    return restId ? { restId, kind: estRdv ? 'event' : 'message' } : null;
  } catch { return null; }
}

export async function resoudreReunion(el: { restId: string; kind: 'message' | 'event' }): Promise<ReunionVue | null> {
  const r = await callAtlasWorker<{ enregistree: boolean; reunion?: ReunionVue }>('reunion/resoudre', el);
  return r.enregistree && r.reunion ? r.reunion : null;
}
export async function proposerPoint(key: string, p: { titre: string; dureeMin: number; projetNo?: string; pieceNom?: string; pieceUrl?: string }): Promise<ReunionVue> {
  return (await callAtlasWorker<{ reunion: ReunionVue }>('reunion/proposer', { key, ...p })).reunion;
}
export async function retirerPoint(id: string): Promise<void> { await callAtlasWorker('reunion/retirer', { id }); }
export async function deleguerPoint(id: string, delegation: unknown): Promise<void> { await callAtlasWorker('reunion/deleguer', { id, delegation }); }
export async function deposerAudio(id: string, base64: string, dureeS: number): Promise<void> { await callAtlasWorker('reunion/audio', { id, base64, dureeS }); }
export async function chargerReunion(key: string): Promise<ReunionVue> { return (await callAtlasWorker<{ reunion: ReunionVue }>('reunion/reunion', { key })).reunion; }
