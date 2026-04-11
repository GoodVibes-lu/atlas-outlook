/**
 * Universal Search Picker — Search across projets, tiers, contacts
 */

import { getAllProjets, getAllTiers, getAllContacts } from '../api/airtable';
import type { SearchResult } from '../types';

function normalize(str: string): string {
  return str.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
}

export class SearchPicker {
  private container: HTMLElement;
  private onSelect: (result: SearchResult) => void;
  private searchInput!: HTMLInputElement;
  private resultsList!: HTMLElement;
  private allResults: SearchResult[] = [];
  private loaded = false;

  constructor(container: HTMLElement, onSelect: (result: SearchResult) => void) {
    this.container = container;
    this.onSelect = onSelect;
    this.render();
    this.loadData();
  }

  private render(): void {
    this.container.innerHTML = `
      <div class="search-wrapper">
        <span class="search-icon">🔍</span>
        <input type="text" class="search-input" placeholder="Rechercher projet, tiers, contact..." />
      </div>
      <div class="search-results"></div>
    `;

    this.searchInput = this.container.querySelector('.search-input')!;
    this.resultsList = this.container.querySelector('.search-results')!;

    this.searchInput.addEventListener('input', () => this.search());
    this.searchInput.focus();
  }

  private async loadData(): Promise<void> {
    try {
      const [projets, tiers, contacts] = await Promise.all([
        getAllProjets(),
        getAllTiers(),
        getAllContacts(),
      ]);

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
      this.searchInput.placeholder = `Rechercher parmi ${this.allResults.length} éléments...`;
    } catch (err) {
      this.resultsList.innerHTML = '<p class="empty-state">Erreur de chargement. Vérifiez le token Airtable.</p>';
    }
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
      this.resultsList.innerHTML = '<p class="empty-state">Aucun résultat</p>';
      return;
    }

    this.resultsList.innerHTML = matches.map(r => `
      <div class="suggestion-item" data-id="${r.id}" data-type="${r.type}" data-label="${this.escapeAttr(r.label)}" data-detail="${this.escapeAttr(r.detail)}">
        <span class="suggestion-badge badge-${r.type}">${r.type === 'projet' ? '📁' : r.type === 'tiers' ? '🏢' : '👤'}</span>
        <span class="suggestion-name">${this.highlight(r.label, query)}</span>
        <span class="suggestion-detail">${r.detail}</span>
      </div>
    `).join('');

    // Click handlers
    this.resultsList.querySelectorAll('.suggestion-item').forEach(el => {
      el.addEventListener('click', () => {
        this.onSelect({
          type: el.getAttribute('data-type') as SearchResult['type'],
          id: el.getAttribute('data-id')!,
          label: el.getAttribute('data-label')!,
          detail: el.getAttribute('data-detail')!,
        });
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
    return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  private escapeAttr(str: string): string {
    return str.replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  destroy(): void {
    this.container.innerHTML = '';
  }
}
