/**
 * ui.ts · petites briques d'interface du tableau de bord : échappement, notification discrète
 * (avec « Annuler »), fenêtre modale, choix d'une date, chiffres animés, mouvement réduit.
 */
import { escapeHtml } from '../utils/html';
import { valeurAnimee, presetsMoments, depuisChampDate, versChampDate, type Preset } from './logique';

export const h = (s: unknown) => escapeHtml(String(s ?? ''));
export const mouvementReduit = () => window.matchMedia?.('(prefers-reduced-motion: reduce)').matches === true;

let toastTimer: number | null = null;
/** Notification discrète en bas à droite ; `annuler` ajoute un bouton (5 s). */
export function toast(message: string, type: 'info' | 'success' | 'error' = 'info', annuler?: () => void, dureeMs?: number): void {
  document.querySelectorAll('.tb-toast').forEach(el => el.remove());
  if (toastTimer) window.clearTimeout(toastTimer);
  const el = document.createElement('div');
  el.className = `tb-toast${type === 'error' ? ' is-error' : ''}`;
  el.setAttribute('role', type === 'error' ? 'alert' : 'status');
  el.innerHTML = `<span>${h(message)}</span>${annuler ? '<button type="button">Annuler</button>' : ''}`;
  el.querySelector('button')?.addEventListener('click', () => { el.remove(); annuler?.(); });
  document.body.appendChild(el);
  toastTimer = window.setTimeout(() => el.remove(), dureeMs ?? (annuler ? 6000 : 3500));
}

/** Ouvre une couche (palette, modale) ; Échap ou clic dehors la ferme. Renvoie la fonction de fermeture. */
export function ouvrirCouche(contenu: HTMLElement, onClose?: () => void): () => void {
  const overlay = document.createElement('div');
  overlay.className = 'tb-overlay';
  overlay.appendChild(contenu);
  const precedent = document.activeElement as HTMLElement | null;
  let ferme = false;
  const fermer = () => {
    if (ferme) return;
    ferme = true;
    overlay.classList.remove('is-on');
    document.removeEventListener('keydown', onKey, true);
    window.setTimeout(() => overlay.remove(), mouvementReduit() ? 0 : 150);
    onClose?.();
    precedent?.focus?.();
  };
  const onKey = (e: KeyboardEvent) => { if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); fermer(); } };
  overlay.addEventListener('mousedown', e => { if (e.target === overlay) fermer(); });
  document.addEventListener('keydown', onKey, true);
  document.body.appendChild(overlay);
  requestAnimationFrame(() => overlay.classList.add('is-on'));
  return fermer;
}

/**
 * Choix d'un moment (mettre de côté, envoyer plus tard) : préréglages + date libre. Résout l'ISO
 * choisi, ou null si la personne ferme. `extra` : bloc HTML ajouté (ex. relance), lu par `lireExtra`.
 */
export function choisirMoment<T = undefined>(titre: string, opts: { envoi?: boolean; extra?: string; lireExtra?: (el: HTMLElement) => T } = {}): Promise<{ iso: string; extra?: T } | null> {
  return new Promise(resolve => {
    const now = Date.now();
    const presets: Preset[] = presetsMoments(now, { envoi: opts.envoi });
    const el = document.createElement('div');
    el.className = 'tb-modal';
    el.setAttribute('role', 'dialog');
    el.setAttribute('aria-modal', 'true');
    el.setAttribute('aria-label', titre);
    el.innerHTML = `
      <h3>${h(titre)}</h3>
      <div class="tb-presets">${presets.map((p, i) => `<button type="button" class="tb-btn" data-iso="${p.iso}"><span class="tb-kbd">${i + 1}</span>${h(p.libelle)}</button>`).join('')}</div>
      <label class="tb-note" for="tb-quand">Ou une date précise</label>
      <div class="tb-actions"><input id="tb-quand" class="tb-input" type="datetime-local" value="${versChampDate(presets[0]?.iso || new Date(now + 3_600_000).toISOString())}"><button type="button" class="tb-btn is-primary" id="tb-quand-ok">Valider</button></div>
      ${opts.extra || ''}`;
    let fini = false;
    const finir = (iso: string | null) => {
      if (fini) return;
      fini = true;
      const extra = iso && opts.lireExtra ? opts.lireExtra(el) : undefined;
      fermer();
      resolve(iso ? { iso, ...(extra !== undefined ? { extra } : {}) } : null);
    };
    const fermer = ouvrirCouche(el, () => { if (!fini) { fini = true; resolve(null); } });
    el.querySelectorAll<HTMLButtonElement>('[data-iso]').forEach(b => b.addEventListener('click', () => finir(b.dataset.iso!)));
    el.querySelector('#tb-quand-ok')?.addEventListener('click', () => {
      const iso = depuisChampDate((el.querySelector('#tb-quand') as HTMLInputElement).value);
      if (!iso || Date.parse(iso) < Date.now() + 60_000) { toast('Choisis un moment dans le futur', 'error'); return; }
      finir(iso);
    });
    el.addEventListener('keydown', e => {
      const n = Number(e.key);
      if (n >= 1 && n <= presets.length && !(e.target instanceof HTMLInputElement)) { e.preventDefault(); finir(presets[n - 1].iso); }
    });
    (el.querySelector('[data-iso]') as HTMLButtonElement | null)?.focus();
  });
}

/** Chiffre animé (compte jusqu'à la nouvelle valeur), sans animation si le mouvement est réduit. */
export function animerChiffre(el: HTMLElement, vers: number): void {
  const de = Number(el.dataset.v ?? el.textContent ?? 0) || 0;
  el.dataset.v = String(vers);
  if (de === vers || mouvementReduit()) { el.textContent = String(vers); return; }
  const debut = performance.now(), duree = 520;
  const pas = (t: number) => {
    const x = Math.min(1, (t - debut) / duree);
    el.textContent = String(valeurAnimee(de, vers, x));
    if (x < 1 && el.dataset.v === String(vers)) requestAnimationFrame(pas);
  };
  requestAnimationFrame(pas);
}

/** Confirmation courte (fenêtre ATLAS, jamais window.confirm). */
export function confirmer(titre: string, texte: string, ok = 'Confirmer'): Promise<boolean> {
  return new Promise(resolve => {
    const el = document.createElement('div');
    el.className = 'tb-modal';
    el.setAttribute('role', 'alertdialog');
    el.innerHTML = `<h3>${h(titre)}</h3><p>${h(texte)}</p><div class="tb-actions"><button type="button" class="tb-btn is-primary" data-ok>${h(ok)}</button><button type="button" class="tb-btn is-ghost" data-non>Annuler</button></div>`;
    let fini = false;
    const fermer = ouvrirCouche(el, () => { if (!fini) { fini = true; resolve(false); } });
    el.querySelector('[data-ok]')?.addEventListener('click', () => { fini = true; fermer(); resolve(true); });
    el.querySelector('[data-non]')?.addEventListener('click', () => { fini = true; fermer(); resolve(false); });
    (el.querySelector('[data-ok]') as HTMLButtonElement).focus();
  });
}
