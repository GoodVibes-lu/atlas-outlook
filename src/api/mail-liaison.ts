/**
 * mail-liaison.ts · Lecture du mail ouvert pour le lier à un projet, un client ou un contact (07/10/2026).
 *
 * Le complément lisait le mail avec un jeton de boîte donné par Outlook (graph.ts › getApiContext).
 * Microsoft a coupé les jetons Exchange historiques, et le nouvel Outlook pour Mac n'en donne plus :
 * « Liaison impossible : ta session Outlook doit être rouverte », même après avoir relancé Outlook.
 * Repli : le worker lit le mail avec ses droits d'application (GET /api/plugin/agent/message), dans
 * une des boîtes de la personne seulement. `viaAtlas` dit au panneau de ranger aussi par le worker.
 */
import type { MailMessageFull } from '../types';
import { AtlasError } from './net';
import { getGraphToken, getMessageForLinking, convertToRestId, messageVersLiaison } from './graph';
import { workerRequest } from './worker';

export interface MailPourLiaison {
  message: MailMessageFull;
  /** Jeton de boîte du complément ; vide quand le mail a été lu par le worker. */
  token: string;
  /** Id REST du mail (rangement avec le jeton de boîte). */
  restId: string;
  viaAtlas: boolean;
}

/** Accès direct à la boîte impossible (jeton absent ou refusé, API Outlook muette) : le worker prend le relais. */
function refusBoite(e: unknown): boolean {
  return e instanceof AtlasError && e.service === 'outlook' && e.kind !== 'hors-ligne';
}

export async function lireMailPourLiaison(itemId: string, internetMessageId?: string): Promise<MailPourLiaison> {
  const restId = convertToRestId(itemId);
  try {
    const token = await getGraphToken();
    return { message: await getMessageForLinking(token, restId), token, restId, viaAtlas: false };
  } catch (e) {
    if (!refusBoite(e)) throw e;
    console.info('[liaison] boîte refusée au complément, lecture par ATLAS :', (e as AtlasError).detail || (e as Error).message);
  }
  const q = new URLSearchParams();
  if (internetMessageId) q.set('messageId', internetMessageId);
  if (restId) q.set('graphId', restId);
  const r = await workerRequest<any>('GET', `agent/message?${q.toString()}`, undefined, { retry: true, timeoutMs: 30_000 });
  const m = r?.message;
  if (!m?.id) throw new AtlasError('introuvable', 'Ce mail est introuvable dans ta boîte.', { service: 'atlas', route: 'agent/message' });
  return { message: messageVersLiaison(m, Array.isArray(m.pieces) ? m.pieces : []), token: '', restId, viaAtlas: true };
}
