/**
 * Create Project Panel — Full project creation form with ARGO pre-fill,
 * searchable client dropdown, inline new-tiers creation, and Airtable submission.
 */

import {
  createProjet,
  createTiers,
  getAllTiers,
  getAllEmployes,
  linkEmailToProject,
  type CreateProjetInput,
  type Employe,
} from '../api/airtable';
import { getGraphToken, getMessageForLinking, convertToRestId } from '../api/graph';
import { analyzeEmailForProjet } from '../api/argo';
import { showToast } from '../taskpane';
import type { Tier } from '../types';
import { escapeHtml } from '../utils/html';
import { humanError } from '../api/net';
import { icon } from '../ui/icons';
import { loadingHtml, emptyHtml, errorHtml, renderError } from '../ui/states';

// ── Constants ──

const PROJECT_TYPES = [
  'Evenementiel',
  'Communication',
  'Staffing',
  'Print',
  'Digital',
  'Social',
  'Activation',
  'Retail',
  'Logistique',
] as const;

const BUDGET_RANGES = [
  "< 1'000\u20ac",
  "1'000 - 5'000\u20ac",
  "5'000 - 10'000\u20ac",
  "10'000 - 25'000\u20ac",
  "25'000 - 50'000\u20ac",
  "50'000 - 100'000\u20ac",
  "100'000 - 250'000\u20ac",
  "250'000 - 500'000\u20ac",
  "> 500'000\u20ac",
] as const;

const MONTHS = [
  'Janvier', 'Fevrier', 'Mars', 'Avril', 'Mai', 'Juin',
  'Juillet', 'Aout', 'Septembre', 'Octobre', 'Novembre', 'Decembre',
] as const;

function normalize(str: string): string {
  return str.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
}

// ── Panel Class ──

export class CreateProjectPanel {
  private container: HTMLElement;
  private userName: string;

  // Data
  private tiers: Tier[] = [];
  private employes: Employe[] = [];
  private selectedClientId: string | null = null;
  private selectedClientName = '';
  private selectedEnChargeId: string | null = null;
  private selectedTypes: Set<string> = new Set();

  // DOM refs
  private clientInput: HTMLInputElement | null = null;
  private clientDropdown: HTMLElement | null = null;
  private enChargeInput: HTMLInputElement | null = null;
  private enChargeDropdown: HTMLElement | null = null;
  private newTiersSection: HTMLElement | null = null;
  private createBtn: HTMLButtonElement | null = null;

  constructor(container: HTMLElement, userName: string) {
    this.container = container;
    this.userName = userName;
    this.render();
    this.init();
  }

  // ── Initial render (loading state) ──

  private render(): void {
    this.container.innerHTML = `
      <div class="panel-scroll">
        <h2 class="section-heading">Créer un projet à partir de ce mail</h2>
        <div id="create-notice"></div>
        <div id="create-form">
          <div class="argo-progress" role="status"><span class="argo-progress-bar" aria-hidden="true"></span><span>ARGO prépare la fiche d'après le mail…</span></div>
          ${loadingHtml('Préparation de la fiche…', 4)}
        </div>
      </div>
    `;
  }

  // ── Init: load data + ARGO analysis in parallel, then render form ──

  private async init(): Promise<void> {
    const formEl = document.getElementById('create-form')!;

    try {
      const item = Office.context.mailbox.item;
      if (!item) {
        formEl.innerHTML = emptyHtml({ icon: 'mail', title: 'Aucun mail sélectionné', text: 'Ouvre un mail pour créer un projet à partir de lui.' });
        return;
      }

      const subject = (item as any).subject || '';
      const from = (item as any).from;
      const fromName = from?.displayName || '';
      const fromEmail = from?.emailAddress || '';
      const bodyPreview = (item as any).bodyPreview || '';

      // Tout en parallèle ; un appel en échec n'empêche pas de remplir la fiche à la main.
      const [rAnalysis, rTiers, rEmployes] = await Promise.allSettled([
        analyzeEmailForProjet(subject, fromName, fromEmail, bodyPreview),
        getAllTiers(),
        getAllEmployes(),
      ]);
      const analysis: { denomination?: string; client?: string; descriptif?: string } =
        rAnalysis.status === 'fulfilled' ? rAnalysis.value : {};
      this.tiers = rTiers.status === 'fulfilled' ? rTiers.value : [];
      this.employes = rEmployes.status === 'fulfilled' ? rEmployes.value : [];
      const notice = document.getElementById('create-notice');
      const listsFailed = rTiers.status === 'rejected' ? rTiers.reason : rEmployes.status === 'rejected' ? rEmployes.reason : null;
      if (notice && listsFailed) {
        renderError(notice, listsFailed, () => { notice.innerHTML = ''; this.init(); }, { title: 'Clients ou équipe indisponibles : la recherche sera vide', compact: true });
      } else if (notice && rAnalysis.status === 'rejected') {
        notice.innerHTML = `<p class="help" style="margin-bottom:10px;">${icon('info', 12)}ARGO n'a pas pu préremplir la fiche : complète-la à la main.</p>`;
      }

      // Try to auto-match client from ARGO analysis
      if (analysis.client) {
        const matchedClient = this.findBestTiersMatch(analysis.client);
        if (matchedClient) {
          this.selectedClientId = matchedClient.id;
          this.selectedClientName = matchedClient.relation;
        }
      }

      this.renderForm(analysis, subject);
    } catch (err) {
      renderError(formEl, err, () => { formEl.innerHTML = loadingHtml('Préparation de la fiche…', 4); this.init(); }, { title: 'Fiche impossible à préparer' });
    }
  }

  // ── Render the full form ──

  private renderForm(analysis: { denomination?: string; client?: string; descriptif?: string }, fallbackDenomination: string): void {
    const formEl = document.getElementById('create-form')!;
    const currentYear = new Date().getFullYear();
    const currentMonth = new Date().getMonth(); // 0-indexed
    const years = Array.from({ length: 5 }, (_, i) => currentYear + i);

    formEl.innerHTML = `
      <!-- Denomination -->
      <div class="form-group">
        <label class="form-label" for="proj-denomination">Dénomination <span aria-hidden="true">*</span></label>
        <input type="text" class="form-input" id="proj-denomination" required
               value="${this.escapeAttr(analysis.denomination || fallbackDenomination)}"
               placeholder="Nom du projet" />
      </div>

      <!-- Client (searchable dropdown) -->
      <div class="form-group">
        <label class="form-label" for="proj-client-input">Client <span aria-hidden="true">*</span></label>
        <div style="position:relative;">
          <input type="text" class="form-input" id="proj-client-input"
                 value="${this.escapeAttr(this.selectedClientName)}"
                 placeholder="Rechercher un client"
                 autocomplete="off" />
          <div id="proj-client-dropdown" class="search-dropdown" style="display:none;"></div>
        </div>
        <button type="button" class="btn btn-ghost btn-sm" id="new-tiers-toggle" style="margin-top:6px;">${icon('plus', 14)}Client absent de la liste</button>
        <div id="new-tiers-section" class="card stack-sm" style="display:none;margin-top:8px;">
          <div class="form-group" style="margin-bottom:8px;">
            <label class="form-label" for="new-tiers-nom">Nom de la société <span aria-hidden="true">*</span></label>
            <input type="text" class="form-input" id="new-tiers-nom" placeholder="Luxair S.A." />
          </div>
          <div class="form-group" style="margin-bottom:8px;">
            <label class="form-label" for="new-tiers-email">E-mail (facultatif)</label>
            <input type="email" class="form-input" id="new-tiers-email" placeholder="contact@societe.lu" />
          </div>
          <div class="btn-row">
            <button type="button" class="btn btn-primary btn-sm" id="new-tiers-create">${icon('check', 14)}Créer le client</button>
            <button type="button" class="btn btn-ghost btn-sm" id="new-tiers-cancel">Annuler</button>
          </div>
          <div id="new-tiers-status" style="margin-top:6px;"></div>
        </div>
      </div>

      <!-- Type (multi-select checkboxes) -->
      <div class="form-group">
        <label class="form-label">Type</label>
        <div id="proj-types" style="display:flex;flex-wrap:wrap;gap:6px;">
          ${PROJECT_TYPES.map(t => `
            <label class="checkbox-pill" style="display:inline-flex;align-items:center;gap:5px;padding:4px 9px;border:1px solid var(--border-mid);font-size:12px;cursor:pointer;user-select:none;">
              <input type="checkbox" value="${escapeHtml(t)}" style="margin:0;width:14px;height:14px;accent-color:var(--accent-ink);" />
              ${escapeHtml(t)}
            </label>
          `).join('')}
        </div>
      </div>

      <!-- Budget -->
      <div class="form-group">
        <label class="form-label">Budget</label>
        <select class="dropdown-select" id="proj-budget">
          <option value="">Choisir</option>
          ${BUDGET_RANGES.map(b => `<option value="${b}">${b}</option>`).join('')}
        </select>
      </div>

      <!-- Mois / Annee (side by side) -->
      <div class="form-group">
        <label class="form-label">Mois et année de réalisation</label>
        <div style="display:flex;gap:8px;">
          <select class="dropdown-select" id="proj-mois" style="flex:1;">
            <option value="">Mois</option>
            ${MONTHS.map((m, i) => `<option value="${m}" ${i === currentMonth ? 'selected' : ''}>${m}</option>`).join('')}
          </select>
          <select class="dropdown-select" id="proj-annee" style="flex:1;">
            <option value="">Année</option>
            ${years.map(y => `<option value="${y}" ${y === currentYear ? 'selected' : ''}>${y}</option>`).join('')}
          </select>
        </div>
      </div>

      <!-- En charge (searchable dropdown) -->
      <div class="form-group">
        <label class="form-label">En charge</label>
        <div style="position:relative;">
          <input type="text" class="form-input" id="proj-encharge-input"
                 placeholder="Rechercher une personne de l'équipe"
                 autocomplete="off" />
          <div id="proj-encharge-dropdown" class="search-dropdown" style="display:none;"></div>
        </div>
      </div>

      <!-- Date souhaitee -->
      <div class="form-group">
        <label class="form-label" for="proj-date">Date souhaitée</label>
        <input type="date" class="form-input" id="proj-date" />
      </div>

      <!-- Descriptif -->
      <div class="form-group">
        <label class="form-label" for="proj-descriptif">Descriptif</label>
        <textarea class="form-input" id="proj-descriptif" rows="3"
                  style="resize:vertical;"
                  placeholder="Notes, détails, contexte">${this.escapeHtml(analysis.descriptif || '')}</textarea>
      </div>

      <!-- Submit -->
      <button type="button" class="btn btn-primary btn-block" id="create-btn" style="margin-top:4px;">${icon('plus', 14)}Créer le projet</button>
      <p class="help" style="margin-top:8px;">Le projet est créé dans ATLAS et ce mail y est lié automatiquement.</p>
    `;

    // Cache DOM refs
    this.clientInput = document.getElementById('proj-client-input') as HTMLInputElement;
    this.clientDropdown = document.getElementById('proj-client-dropdown')!;
    this.enChargeInput = document.getElementById('proj-encharge-input') as HTMLInputElement;
    this.enChargeDropdown = document.getElementById('proj-encharge-dropdown')!;
    this.newTiersSection = document.getElementById('new-tiers-section')!;
    this.createBtn = document.getElementById('create-btn') as HTMLButtonElement;

    this.bindEvents();
  }

  // ── Bind all event listeners ──

  private bindEvents(): void {
    // Client searchable dropdown
    this.clientInput!.addEventListener('input', () => this.filterClients());
    this.clientInput!.addEventListener('focus', () => this.filterClients());
    this.clientInput!.addEventListener('blur', () => {
      // Delay to allow click on dropdown item
      setTimeout(() => this.hideDropdown(this.clientDropdown!), 200);
    });

    // En charge searchable dropdown
    this.enChargeInput!.addEventListener('input', () => this.filterEmployes());
    this.enChargeInput!.addEventListener('focus', () => this.filterEmployes());
    this.enChargeInput!.addEventListener('blur', () => {
      setTimeout(() => this.hideDropdown(this.enChargeDropdown!), 200);
    });

    // New tiers toggle
    document.getElementById('new-tiers-toggle')?.addEventListener('click', () => {
      const section = this.newTiersSection!;
      const isVisible = section.style.display !== 'none';
      section.style.display = isVisible ? 'none' : 'block';
    });

    // New tiers create
    document.getElementById('new-tiers-create')?.addEventListener('click', () => this.handleCreateTiers());

    // New tiers cancel
    document.getElementById('new-tiers-cancel')?.addEventListener('click', () => {
      this.newTiersSection!.style.display = 'none';
    });

    // Type checkboxes
    document.querySelectorAll('#proj-types input[type="checkbox"]').forEach(cb => {
      cb.addEventListener('change', (e) => {
        const input = e.target as HTMLInputElement;
        const pill = input.closest('.checkbox-pill') as HTMLElement;
        if (input.checked) {
          this.selectedTypes.add(input.value);
          pill.style.background = 'var(--accent)';
          pill.style.color = 'var(--accent-ink)';
          pill.style.borderColor = 'var(--accent)';
        } else {
          this.selectedTypes.delete(input.value);
          pill.style.background = '';
          pill.style.color = '';
          pill.style.borderColor = '';
        }
      });
    });

    // Create button
    this.createBtn!.addEventListener('click', () => this.handleCreateProject());

    // Close dropdowns on outside click
    document.addEventListener('click', (e) => {
      const target = e.target as HTMLElement;
      if (!target.closest('#proj-client-input') && !target.closest('#proj-client-dropdown')) {
        this.hideDropdown(this.clientDropdown!);
      }
      if (!target.closest('#proj-encharge-input') && !target.closest('#proj-encharge-dropdown')) {
        this.hideDropdown(this.enChargeDropdown!);
      }
    });
  }

  // ── Client dropdown filtering ──

  private filterClients(): void {
    const query = normalize(this.clientInput!.value.trim());
    this.selectedClientId = null;
    this.selectedClientName = '';

    const filtered = query.length === 0
      ? this.tiers.slice(0, 20)
      : this.tiers.filter(t => normalize(t.relation).includes(query)).slice(0, 20);

    if (filtered.length === 0) {
      this.clientDropdown!.innerHTML = '<div class="dropdown-empty">Aucun resultat</div>';
      this.showDropdown(this.clientDropdown!);
      return;
    }

    this.clientDropdown!.innerHTML = filtered.map(t => `
      <div class="dropdown-item" data-id="${escapeHtml(t.id)}" data-name="${this.escapeAttr(t.relation)}">
        <span style="font-weight:500;">${this.highlight(t.relation, query)}</span>
        ${t.email ? `<span style="font-size:11px;color:var(--atlas-text-muted);margin-left:4px;">${this.escapeHtml(t.email)}</span>` : ''}
      </div>
    `).join('');

    this.showDropdown(this.clientDropdown!);

    this.clientDropdown!.querySelectorAll('.dropdown-item').forEach(el => {
      el.addEventListener('mousedown', (e) => {
        e.preventDefault(); // Prevent blur
        this.selectedClientId = el.getAttribute('data-id')!;
        this.selectedClientName = el.getAttribute('data-name')!;
        this.clientInput!.value = this.selectedClientName;
        this.hideDropdown(this.clientDropdown!);
      });
    });
  }

  // ── Employe dropdown filtering ──

  private filterEmployes(): void {
    const query = normalize(this.enChargeInput!.value.trim());
    this.selectedEnChargeId = null;

    const filtered = query.length === 0
      ? this.employes.slice(0, 20)
      : this.employes.filter(e => normalize(e.name).includes(query)).slice(0, 20);

    if (filtered.length === 0) {
      this.enChargeDropdown!.innerHTML = '<div class="dropdown-empty">Aucun resultat</div>';
      this.showDropdown(this.enChargeDropdown!);
      return;
    }

    this.enChargeDropdown!.innerHTML = filtered.map(e => `
      <div class="dropdown-item" data-id="${escapeHtml(e.id)}" data-name="${this.escapeAttr(e.name)}">
        ${this.highlight(e.name, query)}
      </div>
    `).join('');

    this.showDropdown(this.enChargeDropdown!);

    this.enChargeDropdown!.querySelectorAll('.dropdown-item').forEach(el => {
      el.addEventListener('mousedown', (e) => {
        e.preventDefault();
        this.selectedEnChargeId = el.getAttribute('data-id')!;
        this.enChargeInput!.value = el.getAttribute('data-name')!;
        this.hideDropdown(this.enChargeDropdown!);
      });
    });
  }

  // ── Dropdown show/hide helpers ──

  private showDropdown(el: HTMLElement): void {
    el.style.display = 'block';
  }

  private hideDropdown(el: HTMLElement): void {
    el.style.display = 'none';
  }

  // ── Create new Tiers inline ──

  private async handleCreateTiers(): Promise<void> {
    const nomInput = document.getElementById('new-tiers-nom') as HTMLInputElement;
    const emailInput = document.getElementById('new-tiers-email') as HTMLInputElement;
    const statusEl = document.getElementById('new-tiers-status')!;
    const createBtn = document.getElementById('new-tiers-create') as HTMLButtonElement;

    const nom = nomInput.value.trim();
    if (!nom) {
      statusEl.innerHTML = '<p class="help" role="alert" style="color:var(--danger);">Le nom est obligatoire.</p>';
      return;
    }

    createBtn.disabled = true;
    createBtn.textContent = 'Création…';
    statusEl.innerHTML = '';

    try {
      const email = emailInput.value.trim() || undefined;
      const result = await createTiers(nom, email);

      // Add to local tiers list
      const newTier: Tier = { id: result.id, relation: result.relation, categorie: '', email: email || '', telephone: '' };
      this.tiers.unshift(newTier);

      // Auto-select as client
      this.selectedClientId = result.id;
      this.selectedClientName = result.relation;
      this.clientInput!.value = result.relation;

      // Close the new tiers section
      this.newTiersSection!.style.display = 'none';
      nomInput.value = '';
      emailInput.value = '';

      showToast(`Client « ${result.relation} » créé`, 'success');
    } catch (err) {
      statusEl.innerHTML = errorHtml(err, { title: 'Client non créé', retry: false, compact: true });
    } finally {
      createBtn.disabled = false;
      createBtn.innerHTML = `${icon('check', 14)}Créer le client`;
    }
  }

  // ── Create Project ──

  private async handleCreateProject(): Promise<void> {
    const btn = this.createBtn!;
    const denomination = (document.getElementById('proj-denomination') as HTMLInputElement).value.trim();
    const budget = (document.getElementById('proj-budget') as HTMLSelectElement).value;
    const mois = (document.getElementById('proj-mois') as HTMLSelectElement).value;
    const annee = (document.getElementById('proj-annee') as HTMLSelectElement).value;
    const dateDebut = (document.getElementById('proj-date') as HTMLInputElement).value;
    const descriptif = (document.getElementById('proj-descriptif') as HTMLTextAreaElement).value.trim();

    // Validation
    if (!denomination) {
      showToast('La dénomination est obligatoire.', 'error');
      (document.getElementById('proj-denomination') as HTMLInputElement).focus();
      return;
    }

    if (!this.selectedClientId) {
      showToast('Choisis un client dans la liste.', 'error');
      this.clientInput!.focus();
      return;
    }

    btn.disabled = true;
    btn.textContent = 'Création en cours…';

    try {
      const input: CreateProjetInput = {
        denomination,
        clientRecordId: this.selectedClientId,
        types: this.selectedTypes.size > 0 ? Array.from(this.selectedTypes) : undefined,
        budget: budget || undefined,
        mois: mois || undefined,
        annee: annee || undefined,
        enChargeRecordId: this.selectedEnChargeId || undefined,
        dateDebut: dateDebut || undefined,
        descriptif: descriptif || undefined,
      };

      const result = await createProjet(input);

      showToast(`Projet #${escapeHtml(result.noProjet)} cree avec succes !`, 'success');

      // Link the current email to the new project
      await this.linkCurrentEmail(result.id, result.noProjet);

      // Show success state
      const formEl = document.getElementById('create-form')!;
      formEl.innerHTML = `
        <article class="card project-card">
          <span class="eyebrow">${icon('check-circle', 12)}Projet créé</span>
          <h3 class="card-title">#${escapeHtml(result.noProjet)} ${this.escapeHtml(denomination)}</h3>
          <dl class="facts"><dt>Client</dt><dd>${this.escapeHtml(this.selectedClientName)}</dd></dl>
          <p class="status-linked" style="margin-top:10px;">${icon('link', 14)}Ce mail est lié au projet</p>
        </article>
      `;
    } catch (err) {
      showToast(`${humanError(err)}`, 'error');
      btn.disabled = false;
      btn.innerHTML = `${icon('plus', 14)}Créer le projet`;
    }
  }

  // ── Link current email to the newly created project ──

  private async linkCurrentEmail(projetRecordId: string, noProjet: number): Promise<void> {
    try {
      const item = Office.context.mailbox.item;
      if (!item || !item.itemId) return;

      const token = await getGraphToken();
      const restId = convertToRestId(item.itemId);
      const fullMessage = await getMessageForLinking(token, restId);

      // Determine direction
      const userEmail = localStorage.getItem('atlas_addin_user_email') || Office.context?.mailbox?.userProfile?.emailAddress || '';
      const isSent = fullMessage.from.email.toLowerCase() === userEmail.toLowerCase();
      const direction: 'reçu' | 'envoyé' = isSent ? 'envoyé' : 'reçu';

      await linkEmailToProject(
        fullMessage,
        projetRecordId,
        this.userName + ' (Outlook)',
        direction,
      );
    } catch (err) {
      console.warn('[CreateProject] Failed to link email:', err);
      // Non-blocking: project was created, email link is best-effort
    }
  }

  // ── Helpers ──

  private findBestTiersMatch(name: string): Tier | null {
    const q = normalize(name);
    // Exact match first
    const exact = this.tiers.find(t => normalize(t.relation) === q);
    if (exact) return exact;
    // Partial match
    const partial = this.tiers.find(t => normalize(t.relation).includes(q) || q.includes(normalize(t.relation)));
    return partial || null;
  }

  private highlight(text: string, query: string): string {
    if (!query) return this.escapeHtml(text);
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
    this.container.innerHTML = '';
  }
}
