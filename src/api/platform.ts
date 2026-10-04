/**
 * platform.ts — Détection de l'application Outlook et garde-fous d'API.
 *
 * Sur Outlook iPhone / Android, seul Mailbox 1.5 (+ quelques ajouts) est garanti : toute API plus
 * récente doit passer par `supportsMailbox()` / `supportsSet()` (Office.context.requirements.
 * isSetSupported) avant d'être appelée. Le panneau mobile est en lecture seule (aucune rédaction).
 */

/** Ensemble d'API minimal visé sur mobile (cf. manifeste : Mailbox 1.5). */
export const MOBILE_MAILBOX_MIN = '1.5';

/** Vrai sur Outlook iPhone / Android (paramètre du manifeste mobile, sinon détection Office.js). */
export function isMobile(): boolean {
  try {
    if (new URLSearchParams(window.location.search).get('surface') === 'mobile') return true;
  } catch { /* hors navigateur */ }
  try {
    const platform = String((Office.context as any)?.diagnostics?.platform || '');
    if (platform === 'iOS' || platform === 'Android') return true;
    const host = String(Office.context?.mailbox?.diagnostics?.hostName || '');
    if (host === 'OutlookIOS' || host === 'OutlookAndroid') return true;
  } catch { /* hors Outlook */ }
  return false;
}

/** Vrai si l'ensemble d'API demandé est disponible dans ce client Outlook. */
export function supportsSet(name: string, version: string): boolean {
  try {
    return !!Office.context?.requirements?.isSetSupported?.(name, version);
  } catch {
    return false;
  }
}

/** Raccourci pour Mailbox x.y. */
export function supportsMailbox(version: string): boolean {
  return supportsSet('Mailbox', version);
}

/**
 * Ouvre une adresse dans le navigateur de l'appareil (ATLAS, fiche…). Sur mobile et dans le
 * nouvel Outlook, `openBrowserWindow` (OpenBrowserWindowApi 1.1) est la seule voie fiable ;
 * ailleurs, repli sur window.open.
 */
export function openExternal(url: string): void {
  // Seuls https et le lien profond d'ATLAS sont ouverts (jamais javascript:, data:, file:…).
  if (!/^(https:\/\/|atlas-app:\/\/)/i.test(String(url || '').trim())) return;
  try {
    if (supportsSet('OpenBrowserWindowApi', '1.1') && (Office.context as any)?.ui?.openBrowserWindow) {
      (Office.context as any).ui.openBrowserWindow(url);
      return;
    }
  } catch { /* repli ci-dessous */ }
  window.open(url, '_blank', 'noopener');
}

/** Adresse de l'application ATLAS (liens « Ouvrir dans ATLAS »). */
export const ATLAS_BASE = 'https://atlas.vibes.lu';
