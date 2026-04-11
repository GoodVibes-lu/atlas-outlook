/**
 * Project Info Panel — Show project details when email is linked
 */

import { getAllProjets, getProjetsByClient } from '../api/airtable';
import type { Projet } from '../types';

export class ProjectInfoPanel {
  private container: HTMLElement;
  private projet: Projet | null = null;

  constructor(container: HTMLElement) {
    this.container = container;
    this.render();
    this.detectProject();
  }

  private render(): void {
    this.container.innerHTML = `
      <div class="panel-scroll">
        <div class="section-heading">Projet détecté</div>
        <div id="project-info-content">
          <div class="spinner" style="margin:12px auto;"></div>
        </div>
        <div id="related-emails" style="display:none;">
          <div class="section-heading">Derniers emails liés</div>
          <div id="related-emails-list"></div>
        </div>
      </div>
    `;
  }

  private async detectProject(): Promise<void> {
    const content = document.getElementById('project-info-content')!;

    try {
      const item = Office.context.mailbox.item;
      if (!item) {
        content.innerHTML = '<p class="empty-state">Aucun email sélectionné</p>';
        return;
      }

      const subject = (item as any).subject || '';

      // Try to detect #NNN in subject
      const match = subject.match(/#\s*(\d{2,4})/);
      if (match) {
        const projets = await getAllProjets();
        this.projet = projets.find(p => p.noProjet === match[1]) || null;
      }

      if (this.projet) {
        this.showProjectInfo(content);
      } else {
        content.innerHTML = `
          <div class="empty-state">
            <div class="empty-state-icon">📁</div>
            <p>Aucun projet détecté dans le sujet de cet email.<br/>
            Utilisez l'onglet "Lier" pour associer manuellement.</p>
          </div>
        `;
      }
    } catch (err) {
      content.innerHTML = `<p style="color:var(--atlas-danger);font-size:12px;">${(err as Error).message}</p>`;
    }
  }

  private showProjectInfo(container: HTMLElement): void {
    if (!this.projet) return;

    container.innerHTML = `
      <div class="project-info">
        <h4>#${this.projet.noProjet} ${this.escapeHtml(this.projet.denomination)}</h4>
        <div class="project-info-row">
          <span class="project-info-label">Client</span>
          <span class="project-info-value">${this.escapeHtml(this.projet.client || '—')}</span>
        </div>
        <div class="project-info-row">
          <span class="project-info-label">Statut</span>
          <span class="project-info-value">${this.escapeHtml(this.projet.statut || '—')}</span>
        </div>
        <div class="project-info-row">
          <span class="project-info-label">Commercial</span>
          <span class="project-info-value">${this.escapeHtml(this.projet.commercial || '—')}</span>
        </div>
        <div class="project-info-row">
          <span class="project-info-label">Chef de projet</span>
          <span class="project-info-value">${this.escapeHtml(this.projet.chefDeProjet || '—')}</span>
        </div>
        <div class="project-info-row">
          <span class="project-info-label">Début</span>
          <span class="project-info-value">${this.projet.dateDebut ? new Date(this.projet.dateDebut).toLocaleDateString('fr-LU') : '—'}</span>
        </div>
        <div class="project-info-row">
          <span class="project-info-label">Fin</span>
          <span class="project-info-value">${this.projet.dateFin ? new Date(this.projet.dateFin).toLocaleDateString('fr-LU') : '—'}</span>
        </div>
      </div>
    `;
  }

  private escapeHtml(str: string): string {
    return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  destroy(): void {
    this.container.innerHTML = '';
  }
}
