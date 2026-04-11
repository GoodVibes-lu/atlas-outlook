/**
 * Compose Panel — Templates + ARGO tone adaptation for email composition
 */

import { getEmailTemplates, fetchContactArgoProfile } from '../api/airtable';
import { getSalutation, getClosing, adaptEmailBody, analyzeReceivedEmail } from '../api/argo';
import type { EmailTemplate, ArgoProfile } from '../types';
import { showToast } from '../taskpane';

export class ComposePanel {
  private container: HTMLElement;
  private userName: string;
  private templates: EmailTemplate[] = [];
  private selectedTemplate: EmailTemplate | null = null;
  private argoProfile: ArgoProfile | null = null;
  private recipientEmail = '';
  private isReply = false;

  constructor(container: HTMLElement, userName: string, options?: { isReply?: boolean }) {
    this.container = container;
    this.userName = userName;
    this.isReply = options?.isReply ?? false;
    this.render();
    this.loadContext();
  }

  private render(): void {
    this.container.innerHTML = `
      <div class="panel-scroll">
        <div class="section-heading">Destinataire</div>
        <div id="recipient-info" class="email-card">
          <div class="spinner" style="margin:8px auto;"></div>
        </div>

        <div id="argo-profile-section" style="display:none;">
          <div class="section-heading">Profil ARGO</div>
          <div id="argo-profile-info" class="email-card"></div>
        </div>

        <div id="received-analysis" style="display:none;">
          <div class="section-heading">Analyse du mail reçu</div>
          <div id="received-analysis-info" class="email-card"></div>
        </div>

        <div class="section-heading">Template</div>
        <div class="form-group">
          <select class="dropdown-select" id="template-select">
            <option value="">Chargement des templates...</option>
          </select>
        </div>

        <div id="template-variables" style="display:none;">
          <div class="section-heading">Variables</div>
          <div id="variables-container"></div>
        </div>

        <div id="template-preview-section" style="display:none;">
          <div class="section-heading">Aperçu</div>
          <div id="template-preview" class="template-preview"></div>
        </div>

        <div style="margin-top:16px; display:flex; gap:8px;">
          <button class="btn btn-primary btn-block" id="insert-btn" disabled>
            Insérer dans le mail
          </button>
        </div>
      </div>
    `;

    document.getElementById('template-select')?.addEventListener('change', (e) => {
      const id = (e.target as HTMLSelectElement).value;
      this.selectTemplate(id);
    });

    document.getElementById('insert-btn')?.addEventListener('click', () => this.insertIntoEmail());
  }

  private async loadContext(): Promise<void> {
    const recipientInfo = document.getElementById('recipient-info')!;

    try {
      const item = Office.context.mailbox.item;
      if (!item) {
        recipientInfo.innerHTML = '<p class="empty-state">Aucun email en cours de rédaction</p>';
        return;
      }

      // Get recipients from compose item
      if ((item as any).to?.getAsync) {
        (item as any).to.getAsync((result: any) => {
          if (result.status === Office.AsyncResultStatus.Succeeded && result.value.length > 0) {
            const to = result.value;
            this.recipientEmail = to[0]?.emailAddress || '';
            recipientInfo.innerHTML = `
              <div class="email-card-subject">${to.map((r: any) => `${r.displayName || ''} &lt;${r.emailAddress}&gt;`).join(', ')}</div>
            `;
            this.loadArgoProfile();
          } else {
            recipientInfo.innerHTML = '<p style="font-size:12px;color:var(--atlas-text-secondary);">Ajoutez un destinataire pour activer l\'adaptation ARGO</p>';
          }
        });
      } else if ((item as any).to) {
        // Read mode - get "to" directly
        const to = (item as any).to;
        if (Array.isArray(to) && to.length > 0) {
          this.recipientEmail = to[0]?.emailAddress || '';
          recipientInfo.innerHTML = `
            <div class="email-card-subject">${to.map((r: any) => `${r.displayName || ''} &lt;${r.emailAddress}&gt;`).join(', ')}</div>
          `;
          this.loadArgoProfile();
        }
      }

      // Check if this is a reply
      const subject = typeof (item as any).subject === 'string'
        ? (item as any).subject
        : await new Promise<string>((resolve) => {
            (item as any).subject?.getAsync?.((r: any) => resolve(r.value || ''));
            setTimeout(() => resolve(''), 1000);
          });

      this.isReply = subject.startsWith('Re:') || subject.startsWith('RE:');

      // If reply, analyze the received email
      if (this.isReply) {
        this.analyzeReceivedContent();
      }

      // Load templates
      await this.loadTemplates();

    } catch (err) {
      recipientInfo.innerHTML = `<p style="color:var(--atlas-danger);font-size:12px;">${(err as Error).message}</p>`;
    }
  }

  private async loadArgoProfile(): Promise<void> {
    if (!this.recipientEmail) return;

    try {
      this.argoProfile = await fetchContactArgoProfile(this.recipientEmail);
      if (this.argoProfile) {
        const profileSection = document.getElementById('argo-profile-section')!;
        const profileInfo = document.getElementById('argo-profile-info')!;
        profileSection.style.display = 'block';

        const isTu = this.argoProfile.tonPrefere === 'Amical' ||
          this.argoProfile.tutoiementAvec.some(n => n.toLowerCase().includes(this.userName.toLowerCase()));

        profileInfo.innerHTML = `
          <div style="font-size:12px;">
            <strong>${this.argoProfile.prenom} ${this.argoProfile.nom}</strong><br/>
            Ton : <span style="color:${isTu ? 'var(--atlas-success)' : 'var(--atlas-primary)'}">${isTu ? '👋 Tutoiement' : '🤝 Vouvoiement'}</span><br/>
            Langue : ${this.argoProfile.languePreferee || 'FR'}
            ${this.argoProfile.tutoiementAvec.length > 0 ? `<br/>Tutoiement avec : ${this.argoProfile.tutoiementAvec.join(', ')}` : ''}
          </div>
        `;
      }
    } catch { /* non-blocking */ }
  }

  private async analyzeReceivedContent(): Promise<void> {
    try {
      const item = Office.context.mailbox.item;
      if (!item) return;

      // Get the body of the email being replied to
      (item as any).body?.getAsync?.(Office.CoercionType.Text, async (result: any) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) return;

        const body = result.value || '';
        const subject = typeof (item as any).subject === 'string' ? (item as any).subject : '';
        const from = (item as any).from?.displayName || '';

        const analysis = await analyzeReceivedEmail(subject, body, from);

        const section = document.getElementById('received-analysis')!;
        const info = document.getElementById('received-analysis-info')!;
        section.style.display = 'block';

        info.innerHTML = `
          <div style="font-size:12px;">
            Sentiment : <strong>${analysis.sentiment}</strong><br/>
            Urgence : <strong>${analysis.urgence}</strong><br/>
            Ton détecté : <strong>${analysis.tonUtilise}</strong>
            ${analysis.suggestions.length > 0 ? `<br/><br/>Suggestions :<ul style="margin:4px 0 0 16px;">${analysis.suggestions.map(s => `<li>${s}</li>`).join('')}</ul>` : ''}
          </div>
        `;
      });
    } catch { /* non-blocking */ }
  }

  private async loadTemplates(): Promise<void> {
    try {
      this.templates = await getEmailTemplates();
      const select = document.getElementById('template-select') as HTMLSelectElement;

      select.innerHTML = `
        <option value="">— Sélectionner un template —</option>
        ${this.templates.map(t => `
          <option value="${t.id}">${t.nom} ${t.marque ? `(${t.marque})` : ''}</option>
        `).join('')}
      `;
    } catch (err) {
      showToast('Erreur chargement templates', 'error');
    }
  }

  private selectTemplate(templateId: string): void {
    this.selectedTemplate = this.templates.find(t => t.id === templateId) || null;
    const previewSection = document.getElementById('template-preview-section')!;
    const preview = document.getElementById('template-preview')!;
    const insertBtn = document.getElementById('insert-btn') as HTMLButtonElement;
    const varsSection = document.getElementById('template-variables')!;

    if (!this.selectedTemplate) {
      previewSection.style.display = 'none';
      varsSection.style.display = 'none';
      insertBtn.disabled = true;
      return;
    }

    // Show variables if any
    const vars = this.selectedTemplate.variables
      .split(',')
      .map(v => v.trim())
      .filter(Boolean);

    if (vars.length > 0) {
      varsSection.style.display = 'block';
      const varsContainer = document.getElementById('variables-container')!;
      varsContainer.innerHTML = vars.map(v => `
        <div class="form-group">
          <label class="form-label">${v}</label>
          <input type="text" class="form-input var-input" data-var="${v}" placeholder="${v}" />
        </div>
      `).join('');

      // Update preview on variable change
      varsContainer.querySelectorAll('.var-input').forEach(input => {
        input.addEventListener('input', () => this.updatePreview());
      });
    } else {
      varsSection.style.display = 'none';
    }

    previewSection.style.display = 'block';
    insertBtn.disabled = false;
    this.updatePreview();
  }

  private updatePreview(): void {
    if (!this.selectedTemplate) return;

    const preview = document.getElementById('template-preview')!;
    let body = this.selectedTemplate.corpsFR || '';
    let subject = this.selectedTemplate.sujetFR || '';

    // Replace variables
    document.querySelectorAll('.var-input').forEach((input) => {
      const varName = (input as HTMLInputElement).getAttribute('data-var')!;
      const value = (input as HTMLInputElement).value || `{{${varName}}}`;
      body = body.replaceAll(`{{${varName}}}`, value);
      subject = subject.replaceAll(`{{${varName}}}`, value);
    });

    // Apply ARGO salutation/closing
    const salutation = getSalutation(this.argoProfile, this.userName);
    const closing = getClosing(this.argoProfile, this.userName);

    preview.innerHTML = `
      <div style="margin-bottom:8px;font-weight:600;font-size:12px;">Objet: ${this.escapeHtml(subject)}</div>
      <hr style="border:none;border-top:1px solid var(--atlas-border);margin:8px 0;"/>
      <p>${salutation}</p>
      <div>${body}</div>
      <p>${closing}</p>
      <p>${this.userName}<br/>GOOD VIBES events & communications</p>
    `;
  }

  private async insertIntoEmail(): Promise<void> {
    if (!this.selectedTemplate) return;

    const insertBtn = document.getElementById('insert-btn') as HTMLButtonElement;
    insertBtn.disabled = true;
    insertBtn.textContent = 'Insertion...';

    try {
      let body = this.selectedTemplate.corpsFR || '';
      let subject = this.selectedTemplate.sujetFR || '';

      // Replace variables
      document.querySelectorAll('.var-input').forEach((input) => {
        const varName = (input as HTMLInputElement).getAttribute('data-var')!;
        const value = (input as HTMLInputElement).value || '';
        body = body.replaceAll(`{{${varName}}}`, value);
        body = body.replaceAll(`{{${varName.replace(/_/g, '\\_')}}}`, value);
        subject = subject.replaceAll(`{{${varName}}}`, value);
      });

      // Apply ARGO adaptation
      const salutation = getSalutation(this.argoProfile, this.userName);
      const closing = getClosing(this.argoProfile, this.userName);

      let fullHtml = `<p>${salutation}</p>${body}<p>${closing}</p><p>${this.userName}<br/>GOOD VIBES events &amp; communications</p>`;

      // Full AI adaptation if Anthropic key available
      if (this.argoProfile && localStorage.getItem('atlas_addin_anthropic_key')) {
        try {
          fullHtml = await adaptEmailBody(fullHtml, this.argoProfile, this.userName);
        } catch { /* fallback to simple version */ }
      }

      // Insert into Outlook compose window
      const item = Office.context.mailbox.item;
      if (item) {
        // Set subject
        if (subject) {
          (item as any).subject?.setAsync?.(subject);
        }

        // Set body
        (item as any).body?.setAsync?.(fullHtml, { coercionType: Office.CoercionType.Html }, (result: any) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            showToast('Template inséré avec succès', 'success');
          } else {
            showToast('Erreur d\'insertion', 'error');
          }
        });
      }
    } catch (err) {
      showToast(`Erreur : ${(err as Error).message}`, 'error');
    } finally {
      insertBtn.disabled = false;
      insertBtn.textContent = 'Insérer dans le mail';
    }
  }

  private escapeHtml(str: string): string {
    return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  destroy(): void {
    this.container.innerHTML = '';
  }
}
