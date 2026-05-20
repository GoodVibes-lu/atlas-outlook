/**
 * Project Info Panel — Show project details when email is linked
 * Detects projects by #NNN in subject or by conversation ID match.
 * Shows extended info, ARGO contact card, email count, and ATLAS deep link.
 */

import {
  getAllProjets,
  getLinkedConversationIds,
  countLinkedEmails,
  fetchContactArgoProfile,
  getProjetsByClient,
} from '../api/airtable';
import type { Projet, ArgoProfile } from '../types';

// ── Status color mapping ──

const STATUS_COLORS: Record<string, { bg: string; text: string }> = {
  'Demande':      { bg: '#FFF3CD', text: '#856404' },
  'Devis envoyé': { bg: '#CCE5FF', text: '#004085' },
  'Confirmé':     { bg: '#D4EDDA', text: '#155724' },
  'En cours':     { bg: '#D1ECF1', text: '#0C5460' },
  'Terminé':      { bg: '#E2E3E5', text: '#383D41' },
  'Facturé':      { bg: '#D6D8DB', text: '#1B1E21' },
  'Annulé':       { bg: '#F8D7DA', text: '#721C24' },
  'Archivé':      { bg: '#E2E3E5', text: '#6C757D' },
};

function getStatusBadge(statut: string): string {
  const colors = STATUS_COLORS[statut] || { bg: '#E9ECEF', text: '#495057' };
  return `<span style="
    display:inline-block;
    padding:2px 8px;
    border-radius:10px;
    font-size:11px;
    font-weight:600;
    background:${colors.bg};
    color:${colors.text};
    letter-spacing:0.3px;
  ">${escapeHtml(statut || '—')}</span>`;
}

function escapeHtml(str: string | undefined | null): string {
  if (!str) return '';
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

export class ProjectInfoPanel {
  private container: HTMLElement;
  private projet: Projet | null = null;
  private projetExtra: { type?: string; budget?: string; descriptif?: string } = {};
  private emailCount = 0;
  private argoProfile: ArgoProfile | null = null;
  private senderEmail = '';

  constructor(container: HTMLElement) {
    this.container = container;
    this.render();
    this.detectProject();
  }

  private render(): void {
    this.container.innerHTML = `
      <div class="panel-scroll">
        <div class="section-heading">Projet detect&eacute;</div>
        <div id="project-info-content">
          <div class="spinner" style="margin:12px auto;"></div>
        </div>
        <div id="argo-contact-card" style="display:none;"></div>
      </div>
    `;
  }

  private async detectProject(): Promise<void> {
    const content = document.getElementById('project-info-content')!;

    try {
      const item = Office.context.mailbox.item;
      if (!item) {
        content.innerHTML = '<p class="empty-state">Aucun email s&eacute;lectionn&eacute;</p>';
        return;
      }

      const subject = (item as any).subject || '';
      this.senderEmail = (item as any).from?.emailAddress || '';

      // 1) Try to detect #NNN in subject
      const match = subject.match(/#\s*(\d{2,4})/);
      if (match) {
        const projets = await getAllProjets();
        this.projet = projets.find(p => String(p.noProjet) === match[1]) || null;
      }

      // 2) Fallback: conversation ID match
      if (!this.projet) {
        const conversationId = (item as any).conversationId || '';
        if (conversationId) {
          const convMap = await getLinkedConversationIds();
          const linked = convMap.get(conversationId);
          if (linked) {
            const projets = await getAllProjets();
            this.projet = projets.find(p => p.id === linked.projetId) || null;
          }
        }
      }

      if (this.projet) {
        // Fetch extra fields, email count, and ARGO profile in parallel
        const [emailCount, argoProfile, extraFields] = await Promise.all([
          countLinkedEmails(this.projet.id),
          this.senderEmail ? fetchContactArgoProfile(this.senderEmail) : Promise.resolve(null),
          this.fetchExtraFields(this.projet.id),
        ]);

        this.emailCount = emailCount;
        this.argoProfile = argoProfile;
        this.projetExtra = extraFields;

        this.showProjectInfo(content);

        if (this.argoProfile) {
          this.showArgoCard();
        }
      } else {
        content.innerHTML = `
          <div class="empty-state">
            <div class="empty-state-icon">&#128193;</div>
            <p>Aucun projet d&eacute;tect&eacute; dans le sujet ou la conversation.<br/>
            Utilisez l'onglet "Lier" pour associer manuellement.</p>
          </div>
        `;
      }
    } catch (err) {
      content.innerHTML = `<p style="color:var(--atlas-danger);font-size:12px;">${escapeHtml((err as Error).message)}</p>`;
    }
  }

  /**
   * Fetch extra fields (Type, Budget, Descriptif) for a specific project record.
   */
  private async fetchExtraFields(recordId: string): Promise<{ type?: string; budget?: string; descriptif?: string }> {
    try {
      const API_URL = 'https://api.airtable.com/v0';
      const PROJETS_BASE = 'appKiJY0qjI4UTrWU';
      const PROJETS_TABLE = 'tblKBSumqrxAQFt2u';
      const TYPE_FIELD = 'fldzcy1D0UhFDnusK';
      const BUDGET_FIELD = 'fldB7LExyJIQySlG6';
      const DESCRIPTIF_FIELD = 'fldo2A4Ja7UaiQ3QO';

      const token = localStorage.getItem('atlas_addin_airtable_token') || '';
      const fields = [TYPE_FIELD, BUDGET_FIELD, DESCRIPTIF_FIELD].map(f => `fields%5B%5D=${f}`).join('&');
      const url = `${API_URL}/${PROJETS_BASE}/${PROJETS_TABLE}/${recordId}?returnFieldsByFieldId=true&${fields}`;

      const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
      if (!res.ok) return {};

      const data = await res.json();
      const f = data.fields || {};

      // Type can be multipleSelects
      let typeVal = '';
      const typeRaw = f[TYPE_FIELD];
      if (Array.isArray(typeRaw)) {
        typeVal = typeRaw.map((v: any) => typeof v === 'string' ? v : v?.name || '').filter(Boolean).join(', ');
      } else if (typeRaw) {
        typeVal = typeof typeRaw === 'string' ? typeRaw : typeRaw?.name || '';
      }

      return {
        type: typeVal || undefined,
        budget: f[BUDGET_FIELD] ? String(f[BUDGET_FIELD]) : undefined,
        descriptif: f[DESCRIPTIF_FIELD] || undefined,
      };
    } catch {
      return {};
    }
  }

  private showProjectInfo(container: HTMLElement): void {
    if (!this.projet) return;

    const p = this.projet;
    const extra = this.projetExtra;

    // Truncate descriptif to 200 chars
    let descriptifDisplay = '';
    if (extra.descriptif) {
      descriptifDisplay = extra.descriptif.length > 200
        ? extra.descriptif.slice(0, 200) + '...'
        : extra.descriptif;
    }

    container.innerHTML = `
      <div class="project-info" style="position:relative;">
        <div style="display:flex;align-items:center;gap:8px;margin-bottom:8px;">
          <h4 style="margin:0;flex:1;font-size:14px;">${p.refProjet ? escapeHtml(p.refProjet) : '#' + escapeHtml(String(p.noProjet))} ${escapeHtml(p.denomination)}</h4>
          ${getStatusBadge(p.statut)}
        </div>

        ${p.client ? `
        <div class="project-info-row">
          <span class="project-info-label">Client</span>
          <span class="project-info-value">${escapeHtml(p.client)}</span>
        </div>
        ` : ''}

        ${extra.type ? `
        <div class="project-info-row">
          <span class="project-info-label">Type</span>
          <span class="project-info-value">${escapeHtml(extra.type)}</span>
        </div>
        ` : ''}

        ${extra.budget ? `
        <div class="project-info-row">
          <span class="project-info-label">Budget</span>
          <span class="project-info-value">${escapeHtml(extra.budget)}</span>
        </div>
        ` : ''}

        ${p.enCharge ? `
        <div class="project-info-row">
          <span class="project-info-label">En charge</span>
          <span class="project-info-value">${escapeHtml(p.enCharge)}</span>
        </div>
        ` : ''}

        ${p.dateDebut ? `
        <div class="project-info-row">
          <span class="project-info-label">D&eacute;but</span>
          <span class="project-info-value">${new Date(p.dateDebut).toLocaleDateString('fr-LU', { day: 'numeric', month: 'long', year: 'numeric' })}</span>
        </div>
        ` : ''}

        ${p.dateFin ? `
        <div class="project-info-row">
          <span class="project-info-label">Fin</span>
          <span class="project-info-value">${new Date(p.dateFin).toLocaleDateString('fr-LU', { day: 'numeric', month: 'long', year: 'numeric' })}</span>
        </div>
        ` : ''}

        ${descriptifDisplay ? `
        <div style="margin-top:8px;padding:6px 8px;background:var(--atlas-bg-hover,#f5f5f5);border-radius:6px;font-size:11px;color:var(--atlas-text-secondary,#666);line-height:1.4;">
          ${escapeHtml(descriptifDisplay)}
        </div>
        ` : ''}

        <div style="display:flex;align-items:center;gap:12px;margin-top:10px;padding-top:8px;border-top:1px solid var(--atlas-border,#e0e0e0);">
          <span style="font-size:11px;color:var(--atlas-text-secondary,#888);">
            &#128231; ${this.emailCount} email${this.emailCount !== 1 ? 's' : ''} li&eacute;${this.emailCount !== 1 ? 's' : ''}
          </span>
          <span style="flex:1;"></span>
          <button id="btn-open-atlas" class="btn btn-sm btn-primary" style="font-size:11px;padding:4px 10px;">
            Ouvrir dans ATLAS
          </button>
        </div>
      </div>
    `;

    // Bind ATLAS deep link button
    const btn = document.getElementById('btn-open-atlas');
    btn?.addEventListener('click', () => {
      const url = `atlas-app://open?entity=projet&id=${encodeURIComponent(p.id)}`;
      window.open(url, '_blank');
    });
  }

  private showArgoCard(): void {
    if (!this.argoProfile) return;

    const card = document.getElementById('argo-contact-card');
    if (!card) return;

    const a = this.argoProfile;
    const fullName = [a.prenom, a.nom].filter(Boolean).join(' ') || this.senderEmail;

    const tonLabel = a.tonPrefere === 'Amical' ? 'Tu' : a.tonPrefere === 'Professionnel' ? 'Vous' : '';
    const tonBadge = tonLabel
      ? `<span style="
          display:inline-block;
          padding:1px 6px;
          border-radius:8px;
          font-size:10px;
          font-weight:600;
          background:${tonLabel === 'Tu' ? '#D4EDDA' : '#CCE5FF'};
          color:${tonLabel === 'Tu' ? '#155724' : '#004085'};
          margin-left:6px;
        ">${tonLabel}</span>`
      : '';

    const langBadge = a.languePreferee
      ? `<span style="
          display:inline-block;
          padding:1px 6px;
          border-radius:8px;
          font-size:10px;
          font-weight:500;
          background:#E9ECEF;
          color:#495057;
          margin-left:4px;
        ">${escapeHtml(a.languePreferee)}</span>`
      : '';

    card.style.display = 'block';
    card.innerHTML = `
      <div class="section-heading" style="margin-top:12px;">Profil ARGO &mdash; Exp&eacute;diteur</div>
      <div style="
        background:var(--atlas-bg-hover,#f8f9fa);
        border:1px solid var(--atlas-border,#e0e0e0);
        border-radius:8px;
        padding:10px 12px;
        font-size:12px;
      ">
        <div style="display:flex;align-items:center;gap:6px;margin-bottom:4px;">
          <span style="font-weight:600;font-size:13px;">${escapeHtml(fullName)}</span>
          ${tonBadge}
          ${langBadge}
        </div>
        <div style="color:var(--atlas-text-secondary,#888);font-size:11px;">
          ${escapeHtml(this.senderEmail)}
        </div>
        ${a.tutoiementAvec.length > 0 ? `
        <div style="margin-top:6px;font-size:11px;color:var(--atlas-text-secondary,#666);">
          Tutoiement avec : ${a.tutoiementAvec.map(t => escapeHtml(t)).join(', ')}
        </div>
        ` : ''}
      </div>
    `;
  }

  destroy(): void {
    this.container.innerHTML = '';
  }
}
