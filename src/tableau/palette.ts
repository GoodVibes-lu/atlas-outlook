/**
 * palette.ts · palette ⌘K du tableau de bord : une seule entrée pour chercher un mail, changer de
 * section ou agir sur le mail choisi (ARGO, modèle, de côté, projet, je prends). Flèches, Entrée,
 * Échap ; recherche sans accents (scoreRecherche, logique.ts).
 */
import { filtrerCommandes, type Commande } from './logique';
import { h, ouvrirCouche } from './ui';

export interface ActionPalette extends Commande { faire: () => unknown }

export function ouvrirPalette(cmds: ActionPalette[]): void {
  const el = document.createElement('div');
  el.className = 'tb-palette';
  el.setAttribute('role', 'dialog');
  el.setAttribute('aria-label', 'Palette de commandes');
  el.innerHTML = '<input type="text" placeholder="Chercher un mail, une section, une action…" aria-label="Commande" aria-controls="tb-pal-list"><ul id="tb-pal-list" role="listbox"></ul>';
  const input = el.querySelector('input')!;
  const ul = el.querySelector('ul')!;
  let vis: ActionPalette[] = [];
  let idx = 0;
  const dessiner = () => {
    vis = filtrerCommandes(cmds, input.value, input.value.trim() ? 14 : 18);
    idx = Math.min(idx, Math.max(0, vis.length - 1));
    let groupe = '';
    ul.innerHTML = vis.length ? vis.map((c, i) => {
      const titre = c.groupe !== groupe ? `<li class="tb-grp" role="presentation">${h(c.groupe)}</li>` : '';
      groupe = c.groupe;
      return `${titre}<li role="option" data-i="${i}" aria-selected="${i === idx}"><span>${h(c.libelle)}</span>${c.raccourci ? `<small><span class="tb-kbd">${h(c.raccourci)}</span></small>` : ''}</li>`;
    }).join('') : '<li class="tb-grp">Rien ne correspond</li>';
    ul.querySelector('[aria-selected="true"]')?.scrollIntoView({ block: 'nearest' });
  };
  const fermer = ouvrirCouche(el);
  const valider = (i: number) => { const c = vis[i]; if (!c) return; fermer(); window.setTimeout(() => { void c.faire(); }, 0); };
  input.addEventListener('input', () => { idx = 0; dessiner(); });
  input.addEventListener('keydown', e => {
    if (e.key === 'ArrowDown') { e.preventDefault(); idx = Math.min(vis.length - 1, idx + 1); dessiner(); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); idx = Math.max(0, idx - 1); dessiner(); }
    else if (e.key === 'Enter') { e.preventDefault(); valider(idx); }
  });
  ul.addEventListener('mousemove', e => {
    const li = (e.target as HTMLElement).closest<HTMLElement>('[data-i]');
    if (li && Number(li.dataset.i) !== idx) { idx = Number(li.dataset.i); ul.querySelectorAll('[data-i]').forEach(x => x.setAttribute('aria-selected', String(Number((x as HTMLElement).dataset.i) === idx))); }
  });
  ul.addEventListener('click', e => { const li = (e.target as HTMLElement).closest<HTMLElement>('[data-i]'); if (li) valider(Number(li.dataset.i)); });
  dessiner();
  input.focus();
}
