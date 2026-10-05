/**
 * states.ts · États partagés du panneau (05/10/2026) : chargement, vide, erreur avec « Réessayer ».
 *
 * Règle : jamais de texte technique brut à l'écran. Toute erreur passe par `humanError` (net.ts),
 * le détail technique reste dans Réglages › Diagnostic. Un bloc en échec n'empêche pas le reste du
 * panneau de s'afficher : chaque section gère son propre état.
 */

import { icon } from './icons';
import { humanError, AtlasError } from '../api/net';
import { resetWorkerToken } from '../api/worker';
import { escapeHtml } from '../utils/html';

/** Chargement : lignes squelettes + libellé lisible par un lecteur d'écran. */
export function loadingHtml(label = 'Chargement…', lines = 3): string {
  const widths = [92, 76, 58, 84, 40];
  return `<div class="state-loading" role="status" aria-live="polite">
    <span class="sr-only">${escapeHtml(label)}</span>
    ${Array.from({ length: lines }, (_, i) => `<div class="skeleton-line" style="width:${widths[i % widths.length]}%"></div>`).join('')}
    <p class="state-loading-label" aria-hidden="true">${escapeHtml(label)}</p>
  </div>`;
}

/** Chargement court en ligne (bouton, petite section). */
export function inlineLoadingHtml(label = 'Chargement…'): string {
  return `<span class="state-inline-loading" role="status"><span class="spinner spinner-sm" aria-hidden="true"></span>${escapeHtml(label)}</span>`;
}

export interface EmptyOptions {
  icon?: string;
  title: string;
  text?: string;
  /** Bouton d'action : identifiant `data-action` lu par l'appelant. */
  action?: { id: string; label: string; icon?: string };
}

/** État vide utile : ce qui se passe, et quoi faire. */
export function emptyHtml(o: EmptyOptions): string {
  return `<div class="state-empty">
    <div class="state-empty-icon">${icon(o.icon || 'info', 20)}</div>
    <p class="state-empty-title">${escapeHtml(o.title)}</p>
    ${o.text ? `<p class="state-empty-text">${escapeHtml(o.text)}</p>` : ''}
    ${o.action ? `<button type="button" class="btn btn-secondary btn-sm" data-action="${escapeHtml(o.action.id)}">${o.action.icon ? icon(o.action.icon, 14) : ''}${escapeHtml(o.action.label)}</button>` : ''}
  </div>`;
}

function errorIcon(e: unknown): string {
  if (e instanceof AtlasError) {
    if (e.kind === 'hors-ligne' || e.kind === 'reseau') return 'wifi-off';
    if (e.kind === 'session' || e.kind === 'acces') return 'lock';
    if (e.kind === 'delai') return 'clock';
  }
  return 'alert';
}

/** HTML d'une erreur lisible ; le bouton porte `data-action="retry"` si `retry` est vrai. */
export function errorHtml(e: unknown, opts: { title?: string; retry?: boolean; compact?: boolean } = {}): string {
  const msg = humanError(e);
  const title = opts.title || 'Impossible de charger';
  return `<div class="state-error${opts.compact ? ' is-compact' : ''}" role="alert">
    <div class="state-error-icon">${icon(errorIcon(e), opts.compact ? 16 : 20)}</div>
    <div class="state-error-body">
      <p class="state-error-title">${escapeHtml(title)}</p>
      <p class="state-error-text">${escapeHtml(msg)}</p>
      ${opts.retry === false ? '' : `<button type="button" class="btn btn-secondary btn-sm" data-action="retry">${icon('refresh', 14)}Réessayer</button>`}
    </div>
  </div>`;
}

/**
 * Affiche une erreur dans `host` avec un bouton « Réessayer » qui rappelle `onRetry`.
 * Sur une erreur de session, le jeton en cache est oublié avant la nouvelle tentative.
 */
export function renderError(host: HTMLElement, e: unknown, onRetry?: () => void, opts: { title?: string; compact?: boolean } = {}): void {
  host.innerHTML = errorHtml(e, { ...opts, retry: !!onRetry });
  if (!onRetry) return;
  host.querySelector<HTMLButtonElement>('[data-action="retry"]')?.addEventListener('click', () => {
    if (e instanceof AtlasError && e.kind === 'session') resetWorkerToken();
    onRetry();
  });
}
