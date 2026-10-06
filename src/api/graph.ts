/**
 * Microsoft Graph API client for ATLAS Outlook Add-in
 * Gets the Graph token via Office.js SSO or from localStorage fallback
 */

import type { MailMessageFull, MailAttachment } from '../types';
import { readGraphToken } from './roaming-storage';
import { AtlasError, outlookFetch, outlookHttpError } from './net';
import { parcourirArborescence, estBoiteDeReception, normNom, type DossierArbre, type PageDossiers } from '../shared/mail-folder-tree';

const GRAPH_URL = 'https://graph.microsoft.com/v1.0';

// ── Token + API Base Management ──
//
// L'addin a 2 sources d'authentification possibles :
//   A. CALLBACK TOKEN Office.js — toujours dispo, ZÉRO config, mais ne marche
//      QUE sur l'endpoint Outlook REST (outlook.office.com/api/v2.0). Limité à
//      la mailbox courante de l'utilisateur. Parfait pour nos besoins :
//      list folders, scan messages d'un sender, move, create folder.
//   B. GRAPH TOKEN custom — stocké en localStorage (SSO ou token desktop).
//      Plus puissant (toute la Graph API) mais nécessite config utilisateur.
//
// On essaie d'abord A (gratuit, immédiat). Si échec → fallback B. Le base URL
// retourné dépend de la source.
//
// Outlook REST v2.0 a une shape quasi-identique à Graph v1.0 pour
// /me/messages et /me/mailFolders, donc nos helpers marchent sur les 2.

let cachedToken: string | null = null;
let cachedBase: string | null = null;
let tokenExpiry = 0;

interface ApiContext {
  token: string;
  base: string; // ex: "https://graph.microsoft.com/v1.0" ou "https://outlook.office.com/api/v2.0"
}

/**
 * Récupère un token + base URL utilisables pour les appels mailbox.
 * Préfère le callback token Office.js (gratuit, immédiat).
 */
export async function getApiContext(): Promise<ApiContext> {
  if (cachedToken && cachedBase && Date.now() < tokenExpiry) {
    return { token: cachedToken, base: cachedBase };
  }

  const errors: string[] = [];

  // ── A. Callback token Office.js avec isRest:true (préféré) ──
  try {
    if (typeof Office !== 'undefined' && Office.context?.mailbox?.getCallbackTokenAsync) {
      const result = await new Promise<Office.AsyncResult<string>>((resolve) => {
        try {
          Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, (res) => resolve(res));
        } catch (e) {
          errors.push(`isRest exception: ${(e as Error).message?.slice(0, 80)}`);
          resolve({ status: Office.AsyncResultStatus.Failed, value: '' } as Office.AsyncResult<string>);
        }
      });
      if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
        const restUrl = (Office.context.mailbox as any).restUrl || 'https://outlook.office.com/api';
        const base = `${String(restUrl).replace(/\/$/, '')}/v2.0`;
        console.info('[Mailbox API] using callback REST token, base =', base);
        cachedToken = result.value;
        cachedBase = base;
        tokenExpiry = Date.now() + 50 * 60 * 1000;
        return { token: cachedToken, base };
      }
      const err = (result as any).error;
      errors.push(`isRest status=${result.status}${err ? ` (${err.code}: ${err.message?.slice(0, 60)})` : ''}`);
    } else {
      errors.push('Office.context.mailbox.getCallbackTokenAsync indispo');
    }
  } catch (err) {
    errors.push(`isRest catch: ${(err as Error).message?.slice(0, 80)}`);
  }

  // ── B. SSO Office.auth.getAccessToken (Graph) ──
  try {
    if (typeof Office !== 'undefined' && Office.auth) {
      const ssoToken = await Office.auth.getAccessToken({ allowSignInPrompt: true });
      if (ssoToken) {
        console.info('[Mailbox API] using Office SSO token (Graph)');
        cachedToken = ssoToken;
        cachedBase = GRAPH_URL;
        tokenExpiry = Date.now() + 50 * 60 * 1000;
        return { token: ssoToken, base: GRAPH_URL };
      }
      errors.push('SSO retour vide');
    } else {
      errors.push('Office.auth indispo');
    }
  } catch (err) {
    errors.push(`SSO: ${(err as Error).message?.slice(0, 80)}`);
  }

  // ── C. Token Graph collé manuellement (session seulement, jamais persistant) ──
  const stored = readGraphToken();
  if (stored) {
    console.info('[Mailbox API] using localStorage Graph token');
    cachedToken = stored;
    cachedBase = GRAPH_URL;
    tokenExpiry = Date.now() + 30 * 60 * 1000;
    return { token: stored, base: GRAPH_URL };
  }
  errors.push('aucun jeton collé');

  throw new AtlasError('session', 'Outlook ne donne pas à ATLAS l\'accès à ta boîte sur ce poste : rouvre Outlook puis réessaie.', {
    service: 'outlook', route: 'jeton boîte', detail: errors.join(' / '),
  });
}

/**
 * @deprecated Utilise getApiContext() — retourne juste le token pour compat.
 * Conservé pour code legacy. Le base URL est implicite (Graph si token Graph,
 * sinon doit utiliser le base retourné par getApiContext).
 */
export async function getGraphToken(): Promise<string> {
  const ctx = await getApiContext();
  return ctx.token;
}

/** Retourne le base URL associé au token courant (Graph ou Outlook REST). */
export async function getApiBase(): Promise<string> {
  const ctx = await getApiContext();
  return ctx.base;
}

// ── Graph API Helpers ──


async function graphFetch<T>(path: string, token: string, baseOverride?: string): Promise<T> {
  // Si on a un token mais pas de base override → résout via getApiBase
  const base = baseOverride || await getApiBase();
  const res = await outlookFetch(`${base}${path}`, {
    headers: { 'Authorization': `Bearer ${token}` },
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw outlookHttpError(res.status, path.split("?")[0], JSON.stringify(err));
  }
  return res.json();
}

/** Page Graph / Outlook REST : chemin relatif (base de l'API ajoutée) ou URL absolue (nextLink) de la même API. */
async function lirePageAbsolue<T>(url: string, token: string): Promise<T> {
  if (!/^https:\/\//i.test(url)) return graphFetch<T>(url, token);
  const base = await getApiBase();
  const host = (u: string) => { try { return new URL(u).host.toLowerCase(); } catch { return ''; } };
  if (host(url) !== host(base)) throw new AtlasError('requete', 'lien de page inattendu', { service: 'outlook', route: 'mailFolders' });
  const res = await outlookFetch(url, { headers: { 'Authorization': `Bearer ${token}` } });
  if (!res.ok) throw outlookHttpError(res.status, 'mailFolders');
  return res.json();
}

// ── Get Message for Linking ──

export async function getMessageForLinking(token: string, messageId: string): Promise<MailMessageFull> {
  const params = new URLSearchParams({
    $select: 'id,subject,from,receivedDateTime,isRead,hasAttachments,bodyPreview,webLink,conversationId,internetMessageId,toRecipients,ccRecipients,body',
  });

  const m = await graphFetch<any>(`/me/messages/${messageId}?${params}`, token);

  const content: string = m?.body?.content ?? '';
  const contentType: string = m?.body?.contentType ?? 'text';

  let bodyHtml: string | null = null;
  let bodyText = content.trim();

  if (contentType === 'html') {
    bodyHtml = content
      .replace(/<script[\s\S]*?<\/script>/gi, '')
      .replace(/on\w+="[^"]*"/gi, '')
      .replace(/on\w+='[^']*'/gi, '')
      .trim();
    bodyText = content
      .replace(/<style[\s\S]*?<\/style>/gi, '')
      .replace(/<script[\s\S]*?<\/script>/gi, '')
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<\/p>/gi, '\n\n')
      .replace(/<\/div>/gi, '\n')
      .replace(/<[^>]+>/g, '')
      .replace(/&nbsp;/g, ' ')
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&#39;/g, "'")
      .replace(/\n{3,}/g, '\n\n')
      .trim();
  }

  // Fetch attachments
  let attachments: MailAttachment[] = [];
  if (m.hasAttachments) {
    try {
      const attData = await graphFetch<{ value: any[] }>(`/me/messages/${messageId}/attachments?$select=id,name,size,contentType,isInline`, token);
      attachments = (attData.value || []).map((a: any) => ({
        id: a.id,
        name: a.name || '',
        size: a.size || 0,
        contentType: a.contentType || '',
        isInline: a.isInline || false,
      }));
    } catch { /* non-blocking */ }
  }

  return {
    id: m.id,
    subject: m.subject ?? '(sans objet)',
    from: { name: m.from?.emailAddress?.name ?? '', email: m.from?.emailAddress?.address ?? '' },
    receivedAt: m.receivedDateTime ?? '',
    isRead: m.isRead ?? false,
    hasAttachments: m.hasAttachments ?? false,
    bodyPreview: m.bodyPreview ?? '',
    webLink: m.webLink ?? '',
    conversationId: m.conversationId ?? undefined,
    internetMessageId: m.internetMessageId ?? undefined,
    toRecipients: (m.toRecipients ?? []).map((r: any) => ({
      name: r.emailAddress?.name ?? '', email: r.emailAddress?.address ?? '',
    })),
    ccRecipients: (m.ccRecipients ?? []).map((r: any) => ({
      name: r.emailAddress?.name ?? '', email: r.emailAddress?.address ?? '',
    })),
    bodyHtml,
    bodyText,
    attachments,
  };
}

/**
 * Convert Office.js EWS item ID to REST format for use with Graph API.
 */
export function convertToRestId(ewsId: string): string {
  try {
    if (typeof Office !== 'undefined' && Office.context?.mailbox) {
      return Office.context.mailbox.convertToRestId(
        ewsId,
        Office.MailboxEnums.RestVersion.v2_0
      );
    }
  } catch { /* fallback */ }
  return ewsId;
}

/**
 * Get current user info from Graph API.
 */
export async function getCurrentUser(token: string): Promise<{ displayName: string; mail: string }> {
  return graphFetch<{ displayName: string; mail: string }>('/me?$select=displayName,mail', token);
}

// ── Mail Folder Management ──

/**
 * List mail folders (optionally under a parent folder). 07/10/2026 : TOUTES les pages suivies
 * (@odata.nextLink / Outlook REST « @odata.nextLink »), plus de coupure à 200 dossiers par niveau.
 */
export async function listMailFolders(
  token: string, parentFolderId?: string
): Promise<Array<{ id: string; displayName: string }>> {
  const base = parentFolderId
    ? `/me/mailFolders/${parentFolderId}/childFolders`
    : '/me/mailFolders';
  const out: Array<{ id: string; displayName: string }> = [];
  let url: string | undefined = `${base}?$top=250&$select=id,displayName`;
  const vues = new Set<string>();
  while (url && !vues.has(url)) {
    vues.add(url);
    const data: any = await lirePageAbsolue<any>(url, token);
    // Outlook REST v2.0 peut renvoyer PascalCase (Id/DisplayName) selon le mode.
    for (const f of (data.value ?? [])) out.push({ id: f.id || f.Id || '', displayName: f.displayName || f.DisplayName || '' });
    url = data['@odata.nextLink'] || undefined;
  }
  return out;
}

/**
 * ARBORESCENCE COMPLÈTE de la boîte (tous niveaux, pages suivies ; plafond de sûreté 5 000 signalé
 * en console). Même parcours que le worker et ATLAS (src/utils/mail-folder-tree.ts).
 */
export async function listAllMailFolders(token: string): Promise<DossierArbre[]> {
  const sel = '$top=250&$select=id,displayName,childFolderCount';
  const r = await parcourirArborescence({
    racine: `/me/mailFolders?${sel}`,
    enfants: id => `/me/mailFolders/${encodeURIComponent(id)}/childFolders?${sel}`,
    lirePage: url => lirePageAbsolue<PageDossiers>(url, token),
    surPlafond: n => console.warn(`[graph] arborescence : plafond de sûreté de ${n} dossiers atteint`),
    surErreur: (chemin, e) => console.warn('[graph] sous-dossiers illisibles :', chemin, e),
  });
  return r.dossiers;
}

/** Navigate a folder path like "Clients/LUNEX/#530 Project" and return the leaf folder ID */
export async function resolveFolderPath(token: string, folderPath: string): Promise<string | null> {
  const parts = folderPath.split('/').filter(Boolean);
  let parentId: string | undefined;

  for (const part of parts) {
    const children = await listMailFolders(token, parentId);
    const match = children.find(f => f.displayName.toLowerCase() === part.toLowerCase());
    if (!match) return null;
    parentId = match.id;
  }

  return parentId ?? null;
}

/** Create a mail folder under Inbox or a parent folder */
export async function createMailFolder(
  token: string, displayName: string, parentFolderId?: string
): Promise<{ id: string; displayName: string }> {
  const sub = parentFolderId
    ? `/me/mailFolders/${parentFolderId}/childFolders`
    : '/me/mailFolders';
  const base = await getApiBase();

  const res = await outlookFetch(`${base}${sub}`, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ displayName }),
  });

  if (!res.ok) throw outlookHttpError(res.status, "création de dossier");
  return res.json();
}

/** Create a full folder path (e.g. "Clients/LUNEX/#530") — creates missing segments */
export async function ensureFolderPath(
  token: string, folderPath: string
): Promise<string> {
  const parts = folderPath.split('/').filter(Boolean);
  let parentId: string | undefined;

  for (let i = 0; i < parts.length; i++) {
    const part = parts[i];
    const children = await listMailFolders(token, parentId);
    // 07/10/2026 : casse et accents ignorés ; 1er niveau « Inbox » = boîte de réception dans toutes les
    // langues (jamais un dossier « Inbox » créé à côté de « Boîte de réception »).
    const existing = children.find(f => normNom(f.displayName) === normNom(part))
      || (i === 0 && estBoiteDeReception(part) ? children.find(f => estBoiteDeReception(f.displayName)) : undefined);
    if (existing) {
      parentId = existing.id;
    } else {
      const created = await createMailFolder(token, part, parentId);
      parentId = created.id;
    }
  }

  return parentId!;
}

// ── Categories (tags colorés Outlook, visibles dans la liste inbox) ──

/**
 * Définition d'une catégorie ATLAS : nom affiché + couleur preset Outlook.
 * Les couleurs preset sont fixes côté Outlook : preset0=red, 1=orange,
 * 2=peach (gold), 3=yellow, 4=green, 5=teal, 6=olive, 7=blue, 8=purple,
 * 9=maroon, 10-24 = autres tons. On choisit les plus contrastés pour
 * que chaque état soit immédiatement reconnaissable.
 */
// État du mail (couleurs vives pour signaler les actions)
export const ATLAS_CATEGORIES = {
  URGENCE_5: { name: '🚨 Urgence 5', color: 'preset0' },     // red
  URGENCE_4: { name: '🔥 Urgence 4', color: 'preset1' },     // orange
  URGENCE_3: { name: '⚡ Urgence 3', color: 'preset3' },     // yellow
  SNOOZED:   { name: '⏰ Reporté',   color: 'preset7' },     // blue
  DONE:      { name: '✅ Traité',    color: 'preset4' },     // green
  ARCHIVED:  { name: '📦 Archivé',  color: 'preset8' },     // purple
};

/**
 * Catégories IA — mêmes labels que CATEGORY_LABELS dans ia-panel.ts.
 * Une couleur preset différente par catégorie pour distinguer les types
 * de mails dans la vue liste Outlook.
 */
export const ATLAS_IA_CATEGORIES: Record<string, { name: string; color: string }> = {
  demande_devis:          { name: '💼 Demande devis',          color: 'preset11' },  // steel
  validation_client:      { name: '✅ Validation client',      color: 'preset4' },   // green
  refus_client:           { name: '❌ Refus client',           color: 'preset0' },   // red
  question_staff:         { name: '👥 Question staff',         color: 'preset12' },  // dark steel
  facture_fournisseur:    { name: '🧾 Facture fournisseur',    color: 'preset9' },   // maroon
  prospection_entrante:   { name: '📞 Prospection entrante',   color: 'preset2' },   // peach
  prospection_sortante:   { name: '📤 Prospection sortante',   color: 'preset3' },   // yellow
  rdv_planning:           { name: '📅 RDV / Planning',         color: 'preset6' },   // olive
  newsletter:             { name: '📰 Newsletter',             color: 'preset14' },  // dark green
  notification_systeme:   { name: '🤖 Notification système',   color: 'preset13' },  // dark olive
  spam:                   { name: '🚫 Spam',                   color: 'preset16' },  // dark red
  autre:                  { name: '📩 Autre',                  color: 'preset15' },  // dark teal
  federation_association: { name: '🏛 Fédération / Association', color: 'preset8' }, // purple
  demande_interne_staff:  { name: '🏠 Interne staff',          color: 'preset5' },   // teal
  fournisseur:            { name: '🚚 Fournisseur',            color: 'preset10' },  // dark maroon
};

/** Liste tous les noms de catégories ATLAS pour cleanup (states + IA). */
const ALL_ATLAS_NAMES = [
  ...Object.values(ATLAS_CATEGORIES).map((c) => c.name),
  ...Object.values(ATLAS_IA_CATEGORIES).map((c) => c.name),
];

/**
 * S'assure que les catégories ATLAS existent dans la mailbox (sinon les
 * appliquer sur un mail ne donne pas la couleur). Idempotent.
 * Appelé lazy lors du premier setMessageCategories.
 */
let _categoriesEnsured = false;
async function ensureAtlasCategories(token: string, base: string): Promise<void> {
  if (_categoriesEnsured) return;
  const result = await createAtlasCategoriesVerbose(token, base);
  if (result.created > 0 || result.existed > 0) {
    _categoriesEnsured = true;
  }
  // Si tout a échoué, on laisse _categoriesEnsured=false pour retry au prochain appel
}

/**
 * Force la création de toutes les catégories ATLAS dans la mailbox.
 * Retourne un rapport détaillé (créées / existantes / échouées + raisons)
 * pour affichage UI. Appelé par le bouton "Créer les catégories ATLAS"
 * dans Settings.
 */
export async function createAtlasCategoriesVerbose(
  token?: string, base?: string,
): Promise<{ created: number; existed: number; failed: number; errors: string[] }> {
  if (!token || !base) {
    const ctx = await getApiContext();
    token = ctx.token;
    base = ctx.base;
  }
  const errors: string[] = [];
  let created = 0, existed = 0, failed = 0;

  // 1. Liste les catégories existantes
  const url = `${base}/me/outlook/masterCategories`;
  let existing: Set<string>;
  try {
    const res = await outlookFetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (!res.ok) {
      const body = await res.text().catch(() => '');
      const err = `LIST HTTP ${res.status} : ${body.slice(0, 200)}`;
      errors.push(err);
      return { created, existed, failed: 99, errors };
    }
    const data: any = await res.json();
    existing = new Set((data.value || []).map((c: any) => c.displayName || c.DisplayName));
  } catch (e) {
    errors.push(`LIST exception: ${(e as Error).message?.slice(0, 100)}`);
    return { created, existed, failed: 99, errors };
  }

  // 2. POST chaque catégorie manquante
  const allDefs = [...Object.values(ATLAS_CATEGORIES), ...Object.values(ATLAS_IA_CATEGORIES)];
  for (const def of allDefs) {
    if (existing.has(def.name)) { existed++; continue; }
    try {
      const r = await outlookFetch(url, {
        method: 'POST',
        headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ displayName: def.name, color: def.color }),
      });
      if (r.ok) {
        created++;
      } else {
        failed++;
        const body = await r.text().catch(() => '');
        errors.push(`"${def.name}" HTTP ${r.status} : ${body.slice(0, 150)}`);
      }
    } catch (e) {
      failed++;
      errors.push(`"${def.name}" exception: ${(e as Error).message?.slice(0, 100)}`);
    }
  }

  return { created, existed, failed, errors };
}

/**
 * Applique un set de catégories sur un message. Remplace toutes les
 * catégories ATLAS existantes (préserve les non-ATLAS de l'utilisateur).
 * @param messageId REST ID du message (convertToRestId)
 * @param categories array de noms (ex: [ATLAS_CATEGORIES.SNOOZED.name])
 */
export async function setMessageCategories(
  messageId: string,
  categories: string[],
): Promise<void> {
  const { token, base } = await getApiContext();
  await ensureAtlasCategories(token, base);

  // Récupère les catégories actuelles pour préserver les non-ATLAS
  let current: string[] = [];
  try {
    const res = await outlookFetch(`${base}/me/messages/${messageId}?$select=categories`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (res.ok) {
      const data: any = await res.json();
      current = data.categories || data.Categories || [];
    }
  } catch { /* ignore */ }

  // Remplace les catégories ATLAS, garde les autres
  const nonAtlas = current.filter((c) => !ALL_ATLAS_NAMES.includes(c));
  const merged = Array.from(new Set([...nonAtlas, ...categories]));

  const patchRes = await outlookFetch(`${base}/me/messages/${messageId}`, {
    method: 'PATCH',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ categories: merged }),
  });
  if (!patchRes.ok) {
    const err = await patchRes.text().catch(() => '');
    throw outlookHttpError(patchRes.status, "catégories", err);
  }
}

/**
 * Retire toutes les catégories ATLAS d'un message (préserve les autres).
 */
export async function clearAtlasCategories(messageId: string): Promise<void> {
  await setMessageCategories(messageId, []);
}

/** Move a message to a specific folder */
export async function moveMessageToFolder(
  token: string, messageId: string, folderId: string
): Promise<void> {
  const base = await getApiBase();
  const res = await outlookFetch(`${base}/me/messages/${messageId}/move`, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ destinationId: folderId }),
  });
  if (!res.ok) throw outlookHttpError(res.status, "déplacement");
}

/** Copy a message to a specific folder (keeps original in place) */
export async function copyMessageToFolder(
  token: string, messageId: string, folderId: string
): Promise<void> {
  const base = await getApiBase();
  const res = await outlookFetch(`${base}/me/messages/${messageId}/copy`, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ destinationId: folderId }),
  });
  if (!res.ok) throw outlookHttpError(res.status, "copie");
}
