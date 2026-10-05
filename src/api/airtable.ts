/**
 * Données ATLAS pour le complément Outlook (projets, clients, contacts, mails liés, tags IA…).
 *
 * Depuis le 03/10/2026, plus AUCUN accès direct à Airtable ni jeton dans le complément : chaque
 * fonction appelle une route de la liste blanche du worker (`/api/plugin/atlas/*`, cf.
 * worker/portal-api/routes/plugin-atlas.ts) avec le jeton Microsoft de l'utilisateur.
 * Les signatures exportées sont inchangées pour les panneaux.
 */

import type { Projet, Tier, Contact, MailMessageFull, EmailTemplate, ArgoProfile } from '../types';
import { callAtlasWorker } from './worker';
import { AtlasError } from './net';

/** ATLAS injoignable (réseau, délai, session, serveur) : à montrer, pas à confondre avec « rien trouvé ». */
function isUnreachable(e: unknown): boolean {
  return e instanceof AtlasError && ['hors-ligne', 'reseau', 'delai', 'session', 'serveur', 'debit'].includes(e.kind);
}

// ── Cache ──

const cache = new Map<string, { data: unknown; ts: number }>();
const CACHE_TTL = 5 * 60 * 1000; // 5 min

function getCached<T>(key: string): T | null {
  const c = cache.get(key);
  if (c && Date.now() - c.ts < CACHE_TTL) return c.data as T;
  return null;
}

function setCache<T>(key: string, data: T): void {
  cache.set(key, { data, ts: Date.now() });
}

async function cachedList<T>(key: string, route: string, field: string): Promise<T[]> {
  const c = getCached<T[]>(key);
  if (c) return c;
  const data = await callAtlasWorker<Record<string, unknown>>(route);
  const list = (Array.isArray(data[field]) ? data[field] : []) as T[];
  setCache(key, list);
  return list;
}

// ── Projets ──

export async function getAllProjets(): Promise<Projet[]> {
  return cachedList<Projet>('projets', 'projets/list', 'projets');
}

/** Champs complémentaires d'une fiche projet (Type, Budget, Descriptif). */
export async function getProjetExtraFields(recordId: string): Promise<{ type?: string; budget?: string; descriptif?: string }> {
  try {
    const r = await callAtlasWorker<{ type?: string; budget?: string; descriptif?: string }>('projets/extra', { projetId: recordId });
    return { type: r.type || undefined, budget: r.budget || undefined, descriptif: r.descriptif || undefined };
  } catch { return {}; }
}

// ── Tiers (clients de la base Projets) ──

export async function getAllTiers(): Promise<Tier[]> {
  return cachedList<Tier>('tiers', 'tiers/list', 'tiers');
}

// ── Contacts (Contacts clients de la base Projets) ──

export async function getAllContacts(): Promise<Contact[]> {
  return cachedList<Contact>('contacts', 'contacts/list', 'contacts');
}

// ── Conversations liées ──

export async function getLinkedConversationIds(): Promise<Map<string, { projetId: string; projetName: string }>> {
  const cached = getCached<Map<string, { projetId: string; projetName: string }>>('convIds');
  if (cached) return cached;
  const r = await callAtlasWorker<{ entries?: Array<[string, { projetId: string; projetName: string }]> }>('emails/linked-conversations');
  const map = new Map(r.entries || []);
  setCache('convIds', map);
  return map;
}

// ── Mail déjà lié ? ──

export async function getAllLinkedEmailIds(): Promise<{ graphIds: Set<string>; internetIds: Set<string> }> {
  const cached = getCached<{ graphIds: Set<string>; internetIds: Set<string> }>('linkedIds');
  if (cached) return cached;
  const r = await callAtlasWorker<{ graphIds?: string[]; internetIds?: string[] }>('emails/linked-ids');
  const result = { graphIds: new Set(r.graphIds || []), internetIds: new Set(r.internetIds || []) };
  setCache('linkedIds', result);
  return result;
}

// ── Nom de tiers → id client (base Projets) ──

export async function resolveClientIdInProjetsBase(tiersName: string): Promise<string | null> {
  try {
    const r = await callAtlasWorker<{ id?: string | null }>('tiers/resolve', { name: tiersName });
    return r.id || null;
  } catch { return null; }
}

// ── Liaison mail → projet / contact ──

/** Corps du mail transmis au worker (pièces jointes : métadonnées seulement). */
function mailPayload(email: MailMessageFull) {
  return {
    id: email.id,
    subject: email.subject,
    from: email.from,
    receivedAt: email.receivedAt,
    isRead: email.isRead,
    hasAttachments: email.hasAttachments,
    bodyPreview: email.bodyPreview,
    conversationId: email.conversationId || '',
    internetMessageId: email.internetMessageId || '',
    toRecipients: email.toRecipients || [],
    ccRecipients: email.ccRecipients || [],
    bodyHtml: email.bodyHtml,
    bodyText: email.bodyText,
    attachments: (email.attachments || []).map(a => ({ id: a.id, name: a.name, size: a.size, contentType: a.contentType, isInline: a.isInline })),
  };
}

export async function linkEmailToProject(
  email: MailMessageFull,
  projetRecordId: string,
  linkedByName: string,
  direction: 'reçu' | 'envoyé' = 'reçu',
  tiersRecordId?: string,
  options?: { prive?: boolean; privePar?: string },
): Promise<string> {
  const r = await callAtlasWorker<{ id: string }>('emails/link-projet', {
    email: mailPayload(email),
    projetId: projetRecordId,
    linkedByName,
    direction,
    tiersRecordId,
    prive: !!options?.prive,
  });
  cache.delete('linkedIds');
  cache.delete('convIds');
  return r.id;
}

export async function linkEmailToContact(
  email: MailMessageFull,
  contactName: string,
  linkedByName: string,
  direction: 'reçu' | 'envoyé' = 'reçu',
  tiersName?: string,
  options?: { prive?: boolean; privePar?: string },
): Promise<string> {
  const r = await callAtlasWorker<{ id: string }>('emails/link-contact', {
    email: mailPayload(email),
    contactName,
    linkedByName,
    direction,
    tiersName,
    prive: !!options?.prive,
  });
  cache.delete('linkedIds');
  cache.delete('convIds');
  return r.id;
}

// ── Profil ARGO d'un contact ──

export async function fetchContactArgoProfile(contactEmail: string): Promise<ArgoProfile | null> {
  if (!contactEmail) return null;
  try {
    const r = await callAtlasWorker<{ profile?: ArgoProfile | null }>('contacts/argo-profile', { email: contactEmail });
    return r.profile || null;
  } catch { return null; }
}

// ── Modèles de communication ──

export async function getEmailTemplates(): Promise<EmailTemplate[]> {
  return cachedList<EmailTemplate>('templates', 'templates/list', 'templates');
}

// ── Projets d'un client ──

export async function getProjetsByClient(tiersName: string): Promise<Projet[]> {
  const all = await getAllProjets();
  return all.filter(p => (p.client || '').toLowerCase().includes(tiersName.toLowerCase()));
}

// ── Correspondances dossier Outlook (apprises par utilisateur) ──
// L'utilisateur est celui du jeton Microsoft (le worker ignore `userEmail`, gardé pour compat).

export interface FolderMapping {
  id: string;
  cle: string;
  userEmail: string;
  folderPath: string;
  folderId: string;
  scope: 'client' | 'projet';
}

export async function getFolderMapping(
  _userEmail: string, scope: 'client' | 'projet', entityId: string,
): Promise<FolderMapping | null> {
  try {
    const r = await callAtlasWorker<{ mapping?: FolderMapping | null }>('folder-mapping/get', { scope, entityId });
    return r.mapping || null;
  } catch { return null; }
}

export async function saveFolderMapping(
  _userEmail: string, scope: 'client' | 'projet', entityId: string,
  folderPath: string, folderId: string,
): Promise<void> {
  await callAtlasWorker('folder-mapping/save', { scope, entityId, folderPath, folderId });
}

// ── Nouveau tiers (client) ──

export async function createTiers(
  relation: string,
  email?: string,
): Promise<{ id: string; relation: string }> {
  const r = await callAtlasWorker<{ id: string; relation: string }>('tiers/create', { relation, email });
  cache.delete('tiers');
  return { id: r.id, relation: r.relation || relation };
}

// ── Nouveau projet ──

export interface CreateProjetInput {
  denomination: string;
  clientRecordId: string;
  types?: string[];
  budget?: string;
  mois?: string;
  annee?: string;
  enChargeRecordId?: string;
  contactClientRecordId?: string;
  dateDebut?: string;
  descriptif?: string;
}

export async function createProjet(input: CreateProjetInput): Promise<{ id: string; noProjet: number }> {
  const r = await callAtlasWorker<{ id: string; noProjet: number }>('projets/create', input);
  cache.delete('projets');
  return { id: r.id, noProjet: Number(r.noProjet) || 0 };
}

// ── Employés (liste « En charge ») ──

export interface Employe {
  id: string;
  name: string;
  email: string;
}

export async function getAllEmployes(): Promise<Employe[]> {
  return cachedList<Employe>('employes', 'employes/list', 'employes');
}

// ── Nombre de mails liés à un projet ──

export async function countLinkedEmails(projetRecordId: string): Promise<number> {
  try {
    const r = await callAtlasWorker<{ count?: number }>('emails/count', { projetId: projetRecordId });
    return Number(r.count) || 0;
  } catch { return 0; }
}

// ── Tags IA (Inbound Scanner) ──

export interface EmailTag {
  id: string;
  emailId: string;
  category: string;        // EmailTagCategory
  urgencyScore: number;    // 1-5
  summary: string;
  inboxStatus: 'inbox' | 'done' | 'snoozed' | 'archived';
  linkedProjetId?: string; // Projet auquel le mail est rattaché (pour folder mapping)
}

/** Tag IA d'un email (par EmailId Outlook). Null si pas encore taggé. */
export async function getEmailTagByEmailId(emailId: string): Promise<EmailTag | null> {
  if (!emailId) return null;
  try {
    const r = await callAtlasWorker<{ tag?: EmailTag | null }>('tags/by-email', { emailId });
    return r.tag || null;
  } catch (e) { if (isUnreachable(e)) throw e; return null; }
}

/**
 * Tag IA de la conversation (repli quand l'EmailId ne correspond pas — différences EWS / Graph
 * REST ID entre le complément et le scanner). Le conversationId est stable entre clients.
 */
export async function getEmailTagByConversationId(conversationId: string): Promise<EmailTag | null> {
  if (!conversationId) return null;
  try {
    const r = await callAtlasWorker<{ tag?: EmailTag | null }>('tags/by-conversation', { conversationId });
    return r.tag || null;
  } catch (e) { if (isUnreachable(e)) throw e; return null; }
}

async function setTagStatus(tagRecordId: string, status: 'done' | 'snoozed' | 'archived', until?: Date): Promise<boolean> {
  if (!tagRecordId) return false;
  try {
    await callAtlasWorker('tags/status', { tagId: tagRecordId, status, until: until?.toISOString() });
    return true;
  } catch (e) { console.warn(`[addin] tag ${status} failed:`, e); return false; }
}

/** Marque un tag comme « Traité ». */
export function markTagDone(tagRecordId: string): Promise<boolean> {
  return setTagStatus(tagRecordId, 'done');
}

/** Reporte un tag jusqu'à `until` (par défaut : demain 8h, calculé par le worker). */
export function snoozeTag(tagRecordId: string, until?: Date): Promise<boolean> {
  return setTagStatus(tagRecordId, 'snoozed', until);
}

/** Archive un tag. */
export function archiveTag(tagRecordId: string): Promise<boolean> {
  return setTagStatus(tagRecordId, 'archived');
}

/**
 * Crée ou remplace le tag d'un email (bouton « Re-analyser »). L'ancien tag n'est supprimé par
 * le worker que s'il porte le même mail ; le tag est attribué à l'utilisateur authentifié.
 */
export async function upsertEmailTag(input: {
  oldTagId?: string;
  emailId: string;
  conversationId: string;
  subject: string;
  fromEmail: string;
  fromName: string;
  receivedAt: string;
  category: string;
  urgencyScore: number;
  summary: string;
  detectedLanguage: string;
  userEmail: string;
}): Promise<{ id: string }> {
  const { userEmail: _ignored, ...payload } = input;
  const r = await callAtlasWorker<{ id: string }>('tags/replace', payload);
  return { id: r.id };
}

/** Corrige la catégorie d'un tag (apprentissage). */
export async function correctTagCategory(tagRecordId: string, newCategory: string): Promise<boolean> {
  if (!tagRecordId || !newCategory) return false;
  try {
    await callAtlasWorker('tags/category', { tagId: tagRecordId, category: newCategory });
    return true;
  } catch (e) { console.warn('[addin] correctTagCategory failed:', e); return false; }
}
