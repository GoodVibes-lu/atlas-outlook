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

import { getApiContext, listAllMailFolders } from './graph';
import { dossiersDeRangement } from '../shared/mail-folder-tree';
import { outlookFetch } from './net';
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

  // 1. ARBORESCENCE COMPLÈTE (tous niveaux, pages suivies ; 07/10/2026 : plus de limite à 3 niveaux),
  //    sans les dossiers système ni leurs sous-arbres fermés (corbeille, indésirables, brouillons…).
  const flat: Array<{ id: string; path: string }> = dossiersDeRangement(await listAllMailFolders(token))
    .filter(d => d.nom && !(d.profondeur === 0 && SKIP_FOLDERS.has(d.nom.toLowerCase())))
    .map(d => ({ id: d.id, path: d.chemin }));

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
      const res = await outlookFetch(url, { headers: { Authorization: `Bearer ${token}` } });
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
