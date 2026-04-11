/**
 * Create Project Panel — Create a new project from an email via ARGO analysis
 */

import { analyzeEmailForProjet } from '../api/argo';
import { showToast } from '../taskpane';

export class CreateProjectPanel {
  private container: HTMLElement;

  constructor(container: HTMLElement) {
    this.container = container;
    this.render();
    this.analyzeEmail();
  }

  private render(): void {
    this.container.innerHTML = `
      <div class="panel-scroll">
        <div class="section-heading">Créer un projet depuis cet email</div>
        <div id="create-form">
          <div class="spinner" style="margin:12px auto;"></div>
          <p style="text-align:center;font-size:12px;color:var(--atlas-text-secondary);">Analyse ARGO en cours...</p>
        </div>
      </div>
    `;
  }

  private async analyzeEmail(): Promise<void> {
    const formEl = document.getElementById('create-form')!;

    try {
      const item = Office.context.mailbox.item;
      if (!item) {
        formEl.innerHTML = '<p class="empty-state">Aucun email sélectionné</p>';
        return;
      }

      const subject = (item as any).subject || '';
      const from = (item as any).from;
      const fromName = from?.displayName || '';
      const fromEmail = from?.emailAddress || '';
      const bodyPreview = (item as any).bodyPreview || '';

      // Analyze with ARGO
      const analysis = await analyzeEmailForProjet(subject, fromName, fromEmail, bodyPreview);

      formEl.innerHTML = `
        <div class="form-group">
          <label class="form-label">Dénomination</label>
          <input type="text" class="form-input" id="proj-denomination" value="${this.escapeAttr(analysis.denomination || subject)}" />
        </div>
        <div class="form-group">
          <label class="form-label">Client</label>
          <input type="text" class="form-input" id="proj-client" value="${this.escapeAttr(analysis.client || fromName)}" />
        </div>
        <div class="form-group">
          <label class="form-label">Type d'événement</label>
          <select class="dropdown-select" id="proj-type">
            <option value="">— Sélectionner —</option>
            ${['Cocktail', 'Conférence', 'Gala', 'Séminaire', 'Workshop', 'Lancement', 'Team Building', 'Autre']
              .map(t => `<option value="${t}" ${analysis.typeEvenement === t ? 'selected' : ''}>${t}</option>`)
              .join('')}
          </select>
        </div>
        <div class="form-group">
          <label class="form-label">Date début</label>
          <input type="date" class="form-input" id="proj-date" value="${analysis.debut || ''}" />
        </div>
        <div class="form-group">
          <label class="form-label">Lieu</label>
          <input type="text" class="form-input" id="proj-lieu" value="${this.escapeAttr(analysis.lieu || '')}" />
        </div>
        <div class="form-group">
          <label class="form-label">Descriptif</label>
          <textarea class="form-input" id="proj-descriptif" rows="3" style="resize:vertical;">${this.escapeHtml(analysis.descriptif || '')}</textarea>
        </div>
        <button class="btn btn-primary btn-block" id="create-btn">Créer le projet</button>
        <p style="font-size:11px;color:var(--atlas-text-muted);margin-top:8px;text-align:center;">
          Le projet sera créé dans ATLAS et l'email sera automatiquement lié.
        </p>
      `;

      document.getElementById('create-btn')?.addEventListener('click', () => this.createProject());

    } catch (err) {
      formEl.innerHTML = `
        <div class="empty-state">
          <p>Impossible d'analyser l'email. Vérifiez la clé API Anthropic dans les paramètres.</p>
          <p style="color:var(--atlas-danger);font-size:11px;margin-top:8px;">${(err as Error).message}</p>
        </div>
      `;
    }
  }

  private async createProject(): Promise<void> {
    const btn = document.getElementById('create-btn') as HTMLButtonElement;
    btn.disabled = true;
    btn.textContent = 'Création...';

    try {
      // For now, show a message that this feature is coming
      // Full implementation requires creating a record in the Projets table
      // with the correct structure (linked fields, auto-increment N° Projet, etc.)
      showToast('Fonctionnalité en cours de développement — créez le projet dans ATLAS Desktop', 'info');
    } catch (err) {
      showToast(`Erreur : ${(err as Error).message}`, 'error');
    } finally {
      btn.disabled = false;
      btn.textContent = 'Créer le projet';
    }
  }

  private escapeHtml(str: string): string {
    return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  private escapeAttr(str: string): string {
    return str.replace(/"/g, '&quot;').replace(/'/g, '&#39;').replace(/</g, '&lt;');
  }

  destroy(): void {
    this.container.innerHTML = '';
  }
}
