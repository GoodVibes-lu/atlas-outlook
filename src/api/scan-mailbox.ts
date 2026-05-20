/**
 * scan-mailbox.ts — Scan complet de la boîte Outlook pour construire l'index
 * sender → dossier.
 *
 * Lancé à la demande depuis la UI (bouton "🔍 Scanner ma boîte"). Pour chaque
 * dossier hors Inbox/Sent/Drafts/Junk/Deleted, récupère les N derniers mails
 * et indexe leurs expéditeurs. Plus le user a de mails dans un dossier, plus
 * le mapping est solide.
 *
 * On limite à $top=200 par dossier pour éviter d'exploser les appels API.
 * Sur 50 dossiers × 200 = 10k mails scannés en ~30 sec.
 */

import { getApiContext, listMailFolders } from './graph';
import { bulkRecord } from './sender-folder-index';

const SKIP_FOLDERS = new Set([
  'inbox', 'sent items', 'drafts', 'deleted items', 'junk email',
  'outbox', 'archive', 'rss feeds', 'conversation history', 'sync issues',
  'clutter', 'notes',
]);

export interface ScanProgress {
  foldersScanned: number;
  foldersTotal: number;
  mailsIndexed: number;
  currentFolder: string;
}

export async function scanMailboxBuildIndex(
  onProgress?: (p: ScanProgress) => void,
): Promise<{ foldersScanned: number; mailsIndexed: number }> {
  const ctx = await getApiContext();
  const { token, base } = ctx;

  // 1. Liste tous les dossiers de premier niveau, puis enfants (2 niveaux max)
  const roots = await listMailFolders(token);
  const flat: Array<{ id: string; path: string }> = [];
  for (const f of roots) {
    if (SKIP_FOLDERS.has((f.displayName || '').toLowerCase())) continue;
    if (!f.displayName) continue;
    flat.push({ id: f.id, path: f.displayName });
    try {
      const children = await listMailFolders(token, f.id);
      for (const c of children) {
        if (!c.displayName) continue;
        flat.push({ id: c.id, path: `${f.displayName}/${c.displayName}` });
        // 3e niveau
        try {
          const grand = await listMailFolders(token, c.id);
          for (const g of grand) {
            if (!g.displayName) continue;
            flat.push({ id: g.id, path: `${f.displayName}/${c.displayName}/${g.displayName}` });
          }
        } catch { /* skip */ }
      }
    } catch { /* skip */ }
  }

  let mailsIndexed = 0;
  let foldersScanned = 0;

  // 2. Pour chaque dossier, fetch les mails et indexe les senders
  for (const folder of flat) {
    onProgress?.({
      foldersScanned,
      foldersTotal: flat.length,
      mailsIndexed,
      currentFolder: folder.path,
    });
    try {
      const url = `${base}/me/mailFolders/${folder.id}/messages?$top=200&$select=id,from`;
      const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
      if (!res.ok) {
        foldersScanned++;
        continue;
      }
      const data: any = await res.json();
      const msgs = (data.value || []) as any[];
      const pairs: Array<{ sender: string; folderId: string; folderPath: string }> = [];
      for (const m of msgs) {
        // Outlook REST peut renvoyer PascalCase
        const fromObj = m.from || m.From;
        const addr = fromObj?.emailAddress?.address || fromObj?.EmailAddress?.Address || '';
        if (!addr) continue;
        pairs.push({ sender: addr, folderId: folder.id, folderPath: folder.path });
      }
      if (pairs.length > 0) {
        bulkRecord(pairs);
        mailsIndexed += pairs.length;
      }
    } catch (e) {
      console.warn('[scan-mailbox] folder failed:', folder.path, e);
    }
    foldersScanned++;
  }

  onProgress?.({
    foldersScanned,
    foldersTotal: flat.length,
    mailsIndexed,
    currentFolder: '✓ Terminé',
  });

  return { foldersScanned, mailsIndexed };
}
