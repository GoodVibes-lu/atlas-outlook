/**
 * Échappement et assainissement du HTML injecté dans le panneau (audit M21).
 * Toute donnée qui vient d'ailleurs (expéditeur, Airtable, modèles, IA, messages d'erreur)
 * passe par escapeHtml (texte) ou sanitizeHtml (HTML riche) avant d'être injectée.
 */
import DOMPurify from 'dompurify';

export function escapeHtml(str: unknown): string {
  if (str === null || str === undefined || str === '') return '';
  return String(str)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

/** Message d'une erreur quelconque, prêt à être injecté. */
export function escapeError(err: unknown): string {
  return escapeHtml(err instanceof Error ? err.message : String(err ?? ''));
}

/**
 * HTML riche (corps de modèle, texte rendu par l'IA) : retire scripts, gestionnaires d'événements,
 * iframes, formulaires et URL javascript:. Les balises de mise en forme restent.
 */
export function sanitizeHtml(html: string | undefined | null): string {
  if (!html) return '';
  return DOMPurify.sanitize(String(html), {
    USE_PROFILES: { html: true },
    FORBID_TAGS: ['style', 'form', 'input', 'button', 'textarea', 'select', 'iframe', 'object', 'embed', 'link', 'meta'],
    FORBID_ATTR: ['srcset'],
  });
}
