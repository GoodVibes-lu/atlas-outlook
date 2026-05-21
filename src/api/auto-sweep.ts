/**
 * auto-sweep.ts — Auto-archivage déclenché depuis l'addin Outlook (sans
 * dépendre de l'app Electron desktop).
 *
 * Quand l'addin se charge (Office.onReady) et à chaque ItemChanged, on
 * lance un sweep en background :
 *   1. Fetch les 100 derniers mails non rangés (encore en Inbox)
 *   2. Pour chacun, applique la GOLDEN RULE :
 *        • Doit être LU
 *        • Doit être lu depuis > 10 min (firstSeenReadAt en localStorage)
 *        • Doit être en CC/BCC (utilisateur pas destinataire principal)
 *   3. Si pass : cherche le dossier habituel via lookupSenderFolder()
 *   4. Si dossier trouvé : déplace + enrichit l'index
 *
 * Le sweep est throttlé à 1 fois toutes les 2 min (évite spam API si
 * Charles change de mail rapidement).
 *
 * Stockage du timestamp first-seen-read :
 *   localStorage 'atlas.inbox.firstSeenRead.<msgId>' → ISO timestamp
 *   Cleanup auto des entrées > 7 jours.
 */

import { getApiContext, moveMessageToFolder } from './graph';
import { lookupSenderFolder, recordSenderFolder } from './sender-folder-index';

const STORAGE_PREFIX = 'atlas.inbox.firstSeenRead.';
const ARCHIVE_DELAY_MS = 10 * 60 * 1000;          // 10 min
const STORAGE_RETENTION_MS = 7 * 24 * 3600_000;   // 7j
const THROTTLE_MS = 2 * 60 * 1000;                // 2 min entre 2 sweeps

let lastSweepAt = 0;
let purgedThisSession = false;

function getFirstSeenReadAt(id: string): string | null {
  try { return localStorage.getItem(STORAGE_PREFIX + id); } catch { return null; }
}

function setFirstSeenReadAt(id: string): void {
  try {
    if (localStorage.getItem(STORAGE_PREFIX + id)) return;
    localStorage.setItem(STORAGE_PREFIX + id, new Date().toISOString());
  } catch { /* quota */ }
}

function purgeOldEntries(): void {
  try {
    const cutoff = Date.now() - STORAGE_RETENTION_MS;
    const toRemove: string[] = [];
    for (let i = 0; i < localStorage.length; i++) {
      const k = localStorage.key(i);
      if (!k || !k.startsWith(STORAGE_PREFIX)) continue;
      const v = localStorage.getItem(k);
      if (!v) { toRemove.push(k); continue; }
      const ts = new Date(v).getTime();
      if (isNaN(ts) || ts < cutoff) toRemove.push(k);
    }
    for (const k of toRemove) localStorage.removeItem(k);
  } catch { /* noop */ }
}

export interface SweepResult {
  scanned: number;
  archived: number;
  skipped: { unread: number; tooFresh: number; notCcBcc: number; noFolder: number };
  errors: number;
}

/**
 * Force un sweep maintenant, ignore le throttle.
 * Utile pour le bouton manuel "🧹 Coup de balai".
 */
export async function forceAutoSweep(): Promise<SweepResult> {
  return runSweep(true);
}

/**
 * Sweep avec throttle : ne fait rien si < 2 min depuis le dernier appel.
 * À appeler depuis Office.onReady et ItemChanged.
 */
export async function maybeAutoSweep(): Promise<SweepResult | null> {
  if (Date.now() - lastSweepAt < THROTTLE_MS) return null;
  return runSweep(false);
}

async function runSweep(forced: boolean): Promise<SweepResult> {
  lastSweepAt = Date.now();
  if (!purgedThisSession) { purgeOldEntries(); purgedThisSession = true; }

  const result: SweepResult = {
    scanned: 0,
    archived: 0,
    skipped: { unread: 0, tooFresh: 0, notCcBcc: 0, noFolder: 0 },
    errors: 0,
  };

  try {
    const userEmail = (Office.context.mailbox?.userProfile?.emailAddress || '').toLowerCase();
    if (!userEmail) return result;

    const { token, base } = await getApiContext();

    // Récupère les 100 derniers mails de l'Inbox (lus + non lus)
    const url = `${base}/me/mailFolders/inbox/messages?$top=100&$select=id,isRead,from,toRecipients,ccRecipients,subject,receivedDateTime&$orderby=receivedDateTime desc`;
    const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (!res.ok) {
      result.errors++;
      return result;
    }
    const data: any = await res.json();
    const msgs = (data.value || []) as any[];
    result.scanned = msgs.length;

    for (const m of msgs) {
      // GOLDEN RULE 1 : jamais de non-lu
      if (!m.isRead && !m.IsRead) { result.skipped.unread++; continue; }

      // Détermine CC/BCC vs TO
      const senderEmail = (m.from?.emailAddress?.address || m.From?.EmailAddress?.Address || '').toLowerCase();
      if (senderEmail === userEmail) continue; // notre propre mail envoyé en boucle

      const to = (m.toRecipients || m.ToRecipients || []).map((r: any) =>
        (r.emailAddress?.address || r.EmailAddress?.Address || '').toLowerCase()
      );
      const cc = (m.ccRecipients || m.CcRecipients || []).map((r: any) =>
        (r.emailAddress?.address || r.EmailAddress?.Address || '').toLowerCase()
      );
      const isInTo = to.includes(userEmail);
      const isInCc = cc.includes(userEmail);
      const isCcOrBcc = isInCc || (!isInTo && !isInCc); // BCC = ni TO ni CC explicite

      if (!isCcOrBcc) {
        result.skipped.notCcBcc++;
        if (!forced) continue;
        // En mode forcé on autorise quand même les TO (Charles veut nettoyer)
      }

      // GOLDEN RULE 2 : lu depuis > 10 min
      const firstRaw = getFirstSeenReadAt(m.id);
      if (!firstRaw) {
        setFirstSeenReadAt(m.id);
        result.skipped.tooFresh++;
        continue;
      }
      const ts = new Date(firstRaw).getTime();
      if (isNaN(ts)) continue;
      if (Date.now() - ts < ARCHIVE_DELAY_MS && !forced) {
        result.skipped.tooFresh++;
        continue;
      }

      // Cherche le dossier habituel via index local
      if (!senderEmail) { result.skipped.noFolder++; continue; }
      const hit = lookupSenderFolder(senderEmail);
      if (!hit) { result.skipped.noFolder++; continue; }

      // Move
      try {
        await moveMessageToFolder(token, m.id, hit.folderId);
        recordSenderFolder(senderEmail, hit.folderId, hit.folderPath, 1);
        try { localStorage.removeItem(STORAGE_PREFIX + m.id); } catch { /* noop */ }
        result.archived++;
        console.info(`[auto-sweep] ${m.subject?.slice(0, 50)} → ${hit.folderPath}`);
      } catch (e) {
        console.warn('[auto-sweep] move failed:', e);
        result.errors++;
      }
    }
  } catch (e) {
    console.warn('[auto-sweep] sweep failed:', e);
    result.errors++;
  }

  return result;
}
