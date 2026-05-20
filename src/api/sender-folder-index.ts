/**
 * sender-folder-index.ts — Index persistant local "sender → dossier".
 *
 * Objectif : éviter de re-scanner Outlook à chaque ouverture de mail.
 * On construit une fois (scan initial) une table de correspondance entre
 * chaque expéditeur (email + domaine) et le dossier majoritaire où Charles
 * range historiquement ses mails. Puis chaque nouvelle ouverture lookup
 * en O(1) dans localStorage.
 *
 * L'index est mis à jour automatiquement à chaque action (Traité/Archiver/
 * manual-override) pour rester à jour sans re-scan complet.
 *
 * Storage : localStorage clé "atlas_addin_sender_folder_index" — JSON.
 * Limite : ~5MB total. Pour 5000 entrées (~200 bytes chacune) = 1MB. OK.
 */

const STORAGE_KEY = 'atlas_addin_sender_folder_index';

export interface FolderHit {
  folderId: string;
  folderPath: string;
  count: number;
  lastSeen: string; // ISO datetime
}

// Index par sender : on stocke pour chaque sender plusieurs dossiers possibles
// (au cas où Charles range parfois ailleurs), triés par count desc.
export interface SenderEntry {
  /** Tous les dossiers où ce sender a été vu, triés par count desc. */
  folders: FolderHit[];
}

export type SenderFolderIndex = Record<string /* email lowercase */, SenderEntry>;

let cache: SenderFolderIndex | null = null;

function loadIndex(): SenderFolderIndex {
  if (cache) return cache;
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    cache = raw ? JSON.parse(raw) : {};
  } catch {
    cache = {};
  }
  return cache!;
}

function saveIndex(): void {
  if (!cache) return;
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(cache));
  } catch (e) {
    console.warn('[sender-folder-index] save failed (quota?):', e);
  }
}

/** Cherche le top dossier pour un sender (exact match d'abord, puis domaine). */
export function lookupSenderFolder(senderEmail: string): FolderHit | null {
  const idx = loadIndex();
  const email = senderEmail.toLowerCase().trim();
  if (!email) return null;

  // 1. Match exact sur l'email
  const exact = idx[email];
  if (exact && exact.folders.length > 0) return exact.folders[0];

  // 2. Fallback : match sur le domaine (*@emotion.lu) → on agrège tous les
  // senders du même domaine et on retourne le dossier majoritaire.
  const domain = email.split('@')[1] || '';
  if (!domain) return null;
  const domainKey = `*@${domain}`;
  const domainEntry = idx[domainKey];
  if (domainEntry && domainEntry.folders.length > 0) return domainEntry.folders[0];

  return null;
}

/** Augmente le compteur (sender, folder) — appelé à chaque move/archive. */
export function recordSenderFolder(
  senderEmail: string,
  folderId: string,
  folderPath: string,
  increment = 1,
): void {
  const idx = loadIndex();
  const email = senderEmail.toLowerCase().trim();
  if (!email || !folderId) return;

  for (const key of [email, `*@${email.split('@')[1] || ''}`]) {
    if (!key || key === '*@') continue;
    if (!idx[key]) idx[key] = { folders: [] };
    const entry = idx[key];
    const existing = entry.folders.find((f) => f.folderId === folderId);
    if (existing) {
      existing.count += increment;
      existing.lastSeen = new Date().toISOString();
    } else {
      entry.folders.push({
        folderId,
        folderPath,
        count: increment,
        lastSeen: new Date().toISOString(),
      });
    }
    // Tri par count desc
    entry.folders.sort((a, b) => b.count - a.count);
    // Limite à 5 dossiers max par sender (anti-bloat)
    if (entry.folders.length > 5) entry.folders.length = 5;
  }
  saveIndex();
}

/** Vide tout l'index (debug / reset). */
export function clearIndex(): void {
  cache = {};
  localStorage.removeItem(STORAGE_KEY);
}

/** Stats : nombre de senders dans l'index. */
export function indexStats(): { senders: number; domains: number; totalMails: number; lastUpdate: string } {
  const idx = loadIndex();
  let senders = 0;
  let domains = 0;
  let totalMails = 0;
  let lastUpdate = '';
  for (const [key, entry] of Object.entries(idx)) {
    if (key.startsWith('*@')) domains++;
    else senders++;
    for (const f of entry.folders) {
      totalMails += f.count;
      if (f.lastSeen > lastUpdate) lastUpdate = f.lastSeen;
    }
  }
  return { senders, domains, totalMails, lastUpdate };
}

/**
 * Bulk import : prend des paires (senderEmail, folderId, folderPath) et
 * incrémente l'index. Utilisé par le scan initial de la mailbox.
 */
export function bulkRecord(
  pairs: Array<{ sender: string; folderId: string; folderPath: string }>,
): void {
  const idx = loadIndex();
  for (const p of pairs) {
    const email = p.sender.toLowerCase().trim();
    if (!email || !p.folderId) continue;
    for (const key of [email, `*@${email.split('@')[1] || ''}`]) {
      if (!key || key === '*@') continue;
      if (!idx[key]) idx[key] = { folders: [] };
      const entry = idx[key];
      const existing = entry.folders.find((f) => f.folderId === p.folderId);
      if (existing) {
        existing.count++;
        existing.lastSeen = new Date().toISOString();
      } else {
        entry.folders.push({
          folderId: p.folderId,
          folderPath: p.folderPath,
          count: 1,
          lastSeen: new Date().toISOString(),
        });
      }
    }
  }
  // Re-sort + truncate après tout l'import
  for (const entry of Object.values(idx)) {
    entry.folders.sort((a, b) => b.count - a.count);
    if (entry.folders.length > 5) entry.folders.length = 5;
  }
  saveIndex();
}
