/**
 * Microsoft Graph API client for ATLAS Outlook Add-in
 * Gets the Graph token via Office.js SSO or from localStorage fallback
 */

import type { MailMessageFull, MailAttachment } from '../types';

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

  // ── C. Token Graph stocké en localStorage (collé manuellement par user) ──
  const stored = localStorage.getItem('atlas_addin_graph_token');
  if (stored) {
    console.info('[Mailbox API] using localStorage Graph token');
    cachedToken = stored;
    cachedBase = GRAPH_URL;
    tokenExpiry = Date.now() + 30 * 60 * 1000;
    return { token: stored, base: GRAPH_URL };
  }
  errors.push('localStorage vide');

  throw new Error(`Aucun token mailbox dispo. Tentatives : ${errors.join(' / ')}`);
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
  const res = await fetch(`${base}${path}`, {
    headers: { 'Authorization': `Bearer ${token}` },
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(`Mailbox API ${res.status}: ${JSON.stringify(err).slice(0, 200)}`);
  }
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

/** List mail folders (optionally under a parent folder) */
export async function listMailFolders(
  token: string, parentFolderId?: string
): Promise<Array<{ id: string; displayName: string }>> {
  const base = parentFolderId
    ? `/me/mailFolders/${parentFolderId}/childFolders`
    : '/me/mailFolders';
  const data = await graphFetch<{ value: any[] }>(`${base}?$top=200&$select=id,displayName`, token);
  // Outlook REST v2.0 peut renvoyer PascalCase (Id/DisplayName) selon le mode.
  return (data.value ?? []).map((f: any) => ({
    id: f.id || f.Id || '',
    displayName: f.displayName || f.DisplayName || '',
  }));
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

  const res = await fetch(`${base}${sub}`, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ displayName }),
  });

  if (!res.ok) throw new Error(`Cannot create folder: ${res.status}`);
  return res.json();
}

/** Create a full folder path (e.g. "Clients/LUNEX/#530") — creates missing segments */
export async function ensureFolderPath(
  token: string, folderPath: string
): Promise<string> {
  const parts = folderPath.split('/').filter(Boolean);
  let parentId: string | undefined;

  for (const part of parts) {
    const children = await listMailFolders(token, parentId);
    const existing = children.find(f => f.displayName.toLowerCase() === part.toLowerCase());
    if (existing) {
      parentId = existing.id;
    } else {
      const created = await createMailFolder(token, part, parentId);
      parentId = created.id;
    }
  }

  return parentId!;
}

/** Move a message to a specific folder */
export async function moveMessageToFolder(
  token: string, messageId: string, folderId: string
): Promise<void> {
  const base = await getApiBase();
  const res = await fetch(`${base}/me/messages/${messageId}/move`, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ destinationId: folderId }),
  });
  if (!res.ok) throw new Error(`Cannot move message: ${res.status}`);
}

/** Copy a message to a specific folder (keeps original in place) */
export async function copyMessageToFolder(
  token: string, messageId: string, folderId: string
): Promise<void> {
  const base = await getApiBase();
  const res = await fetch(`${base}/me/messages/${messageId}/copy`, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ destinationId: folderId }),
  });
  if (!res.ok) throw new Error(`Cannot copy message: ${res.status}`);
}
