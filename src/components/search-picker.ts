/**
 * Universal Search Picker — Search across projets, tiers, contacts
 */

import { getAllProjets, getAllTiers, getAllContacts } from '../api/airtable';
import type { SearchResult } from '../types';
import { escapeHtml } from '../utils/html';
import { humanError } from '../api/net';
import { icon } from '../ui/icons';
import { renderError } from '../ui/states';

function normalize(str: string): string {
  return str.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
}

export class SearchPicker {
  private partialFailure: unknown = null;
  private container: HTMLElement;
  private onSelect: (result: SearchResult) => void;
  private searchInput!: HTMLInputElement;
  private resultsList!: HTMLElement;
  private skeletonContainer!: HTMLElement;
  private allResults: SearchResult[] = [];
  private loaded = false;
  private searchTimeout: ReturnType<typeof setTimeout> | null = null;

  constructor(container: HTMLElement, onSelect: (result: SearchResult) => void) {
    this.container = container;
    this.onSelect = onSelect;
    this.render();
    this.loadData();
  }

  private render(): void {
    this.container.innerHTML = `
      <div class="search-wrapper">
        <span class="search-icon">${icon('search', 14)}</span>
        <input type="search" class="search-input" placeholder="Chargement des projets et contacts…" aria-label="Rechercher un projet, un client ou un contact" disabled />
      </div>
      <div class="search-skeleton">
        <div class="skeleton-line skeleton-lg"></div>
        <div class="skeleton-line"></div>
        <div class="skeleton-line skeleton-sm"></div>
      </div>
      <div class="search-results" style="display:none;"></div>
    `;

    this.searchInput = this.container.querySelector('.search-input')!;
    this.resultsList = this.container.querySelector('.search-results')!;
    this.skeletonContainer = this.container.querySelector('.search-skeleton')!;

    this.searchInput.addEventListener('input', () => this.onSearchInput());
  }

  private async loadData(): Promise<void> {
    try {
      // Chaque liste est indépendante : une liste en échec n'empêche pas de chercher dans les autres.
      const settled = await Promise.allSettled([getAllProjets(), getAllTiers(), getAllContacts()]);
      const failed = settled.find(r => r.status === 'rejected') as PromiseRejectedResult | undefined;
      if (settled.every(r => r.status === 'rejected')) throw failed!.reason;
      const val = <T,>(r: PromiseSettledResult<T[]>): T[] => (r.status === 'fulfilled' ? r.value : []);
      const projets = val(settled[0] as PromiseSettledResult<Awaited<ReturnType<typeof getAllProjets>>[number][]>);
      const tiers = val(settled[1] as PromiseSettledResult<Awaited<ReturnType<typeof getAllTiers>>[number][]>);
      const contacts = val(settled[2] as PromiseSettledResult<Awaited<ReturnType<typeof getAllContacts>>[number][]>);
      this.partialFailure = failed ? failed.reason : null;

      this.allResults = [
        ...projets.map(p => ({
          type: 'projet' as const,
          id: p.id,
          label: `#${p.noProjet} ${p.denomination}`,
          detail: p.client || '',
        })),
        ...tiers.map(t => ({
          type: 'tiers' as const,
          id: t.id,
          label: t.relation,
          detail: t.categorie || '',
        })),
        ...contacts.map(c => ({
          type: 'contact' as const,
          id: c.id,
          label: c.personneDeContact,
          detail: c.relationSociete || '',
        })),
      ];

      this.loaded = true;

      // Remove skeletons, enable input, show results area
      this.skeletonContainer.style.display = 'none';
      this.resultsList.style.display = '';
      this.searchInput.disabled = false;
      this.searchInput.placeholder = `Projet, client ou contact (${this.allResults.length})`;
      if (this.partialFailure) {
        const note = document.createElement('p');
        note.className = 'help';
        note.textContent = 'Une partie de la liste est indisponible pour le moment : certains résultats peuvent manquer.';
        this.container.querySelector('.search-wrapper')?.after(note);
      }
      this.searchInput.focus();
    } catch (err: any) {
      console.error('[SearchPicker] loadData failed:', err);
      const msg = humanError(err);
      // Replace skeleton with error
      this.skeletonContainer.style.display = 'none';
      this.resultsList.style.display = '';
      void msg;
      renderError(this.resultsList, err, () => { this.render(); this.loadData(); }, { title: 'Recherche indisponible', compact: true });
    }
  }

  private onSearchInput(): void {
    // Debounce search and show a small loading indicator
    if (this.searchTimeout) clearTimeout(this.searchTimeout);

    const query = this.searchInput.value.trim();
    if (!query || query.length < 1) {
      this.resultsList.innerHTML = '';
      return;
    }

    // Brief loading indicator for perceived responsiveness
    this.resultsList.innerHTML = '<p class="empty-state" style="opacity:0.5;">Recherche...</p>';

    this.searchTimeout = setTimeout(() => {
      this.search();
    }, 80);
  }

  private search(): void {
    const query = normalize(this.searchInput.value.trim());
    if (!query || query.length < 1) {
      this.resultsList.innerHTML = '';
      return;
    }

    const matches = this.allResults
      .filter(r => normalize(r.label).includes(query) || normalize(r.detail).includes(query))
      .slice(0, 15);

    if (matches.length === 0) {
      this.resultsList.innerHTML = '<p class="empty-state">Rien ne correspond : essaie un autre mot (nom du client, numéro de projet).</p>';
      return;
    }

    this.resultsList.innerHTML = matches.map(r => `
      <div class="suggestion-item" role="button" tabindex="0" data-id="${escapeHtml(r.id)}" data-type="${escapeHtml(r.type)}" data-label="${this.escapeAttr(r.label)}" data-detail="${this.escapeAttr(r.detail)}">
        <span class="suggestion-badge badge-${escapeHtml(r.type)}">${r.type === 'projet' ? 'Projet' : r.type === 'tiers' ? 'Client' : 'Contact'}</span>
        <span class="suggestion-name">${this.highlight(r.label, query)}</span>
        <span class="suggestion-detail">${escapeHtml(r.detail)}</span>
      </div>
    `).join('');

    // Click handlers
    this.resultsList.querySelectorAll('.suggestion-item').forEach(el => {
      const choose = () => this.onSelect({
        type: el.getAttribute('data-type') as SearchResult['type'],
        id: el.getAttribute('data-id')!,
        label: el.getAttribute('data-label')!,
        detail: el.getAttribute('data-detail')!,
      });
      el.addEventListener('click', choose);
      el.addEventListener('keydown', (ev) => {
        const k = (ev as KeyboardEvent).key;
        if (k === 'Enter' || k === ' ') { ev.preventDefault(); choose(); }
      });
    });
  }

  private highlight(text: string, query: string): string {
    const idx = normalize(text).indexOf(query);
    if (idx === -1) return this.escapeHtml(text);
    const before = text.slice(0, idx);
    const match = text.slice(idx, idx + query.length);
    const after = text.slice(idx + query.length);
    return `${this.escapeHtml(before)}<strong>${this.escapeHtml(match)}</strong>${this.escapeHtml(after)}`;
  }

  private escapeHtml(str: string): string {
    return escapeHtml(str);
  }

  private escapeAttr(str: string): string {
    return escapeHtml(str);
  }

  destroy(): void {
    if (this.searchTimeout) clearTimeout(this.searchTimeout);
    this.container.innerHTML = '';
  }
}
