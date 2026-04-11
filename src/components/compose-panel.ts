/**
 * Compose Panel — Templates + ARGO tone adaptation for email composition
 */

import { getEmailTemplates, fetchContactArgoProfile } from '../api/airtable';
import { getSalutation, getClosing, adaptEmailBody, analyzeReceivedEmail, generateQuickReplies, generateFreeReply } from '../api/argo';
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
          <div class="section-heading">Analyse du mail recu</div>
          <div id="received-analysis-info" class="email-card"></div>
        </div>

        ${this.isReply ? `
        <div id="quick-replies-section" style="display:none;">
          <div style="display:flex; align-items:center; gap:6px; margin-bottom:8px;">
            <span style="font-size:14px;">⚡</span>
            <span class="section-heading" style="margin:0;">Reponses rapides IA</span>
          </div>
          <div id="quick-replies-container" style="display:flex; flex-direction:column; gap:6px;"></div>
        </div>

        <div id="free-reply-section">
          <div style="display:flex; align-items:center; gap:6px; margin-bottom:8px;">
            <span style="font-size:14px;">✨</span>
            <span class="section-heading" style="margin:0;">Reponse libre IA</span>
          </div>
          <div class="form-group" style="margin-bottom:8px;">
            <textarea class="form-input" id="free-reply-instruction" rows="2" placeholder="Ex: Confirmer la date, demander un devis, proposer un rdv mardi..."></textarea>
          </div>
          <button class="btn btn-primary btn-block" id="free-reply-btn" style="font-size:13px;">
            ✨ Generer la reponse
          </button>
          <div id="free-reply-preview" style="display:none; margin-top:10px;">
            <div class="section-heading" style="font-size:10px;">APERCU</div>
            <div id="free-reply-content" class="template-preview"></div>
            <button class="btn btn-primary btn-block" id="free-reply-send-btn" style="margin-top:8px; font-size:14px;">
              Repondre avec ce texte
            </button>
          </div>
        </div>

        <div style="border-top:1px solid var(--atlas-border); margin:16px 0 8px; padding-top:12px;">
          <div class="section-heading" style="font-size:10px; color:var(--atlas-text-muted);">OU UTILISER UN TEMPLATE</div>
        </div>
        ` : ''}

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

        <div style="margin-top:16px; display:flex; flex-direction:column; gap:8px;">
          ${this.isReply ? `
            <button class="btn btn-primary btn-block" id="reply-btn" disabled style="font-size:14px;">
              Repondre avec ce template
            </button>
          ` : `
            <button class="btn btn-primary btn-block" id="insert-btn" disabled>
              Inserer au curseur
            </button>
            <button class="btn btn-secondary btn-block" id="replace-btn" disabled style="font-size:11px;">
              Remplacer tout le contenu
            </button>
          `}
        </div>
      </div>
    `;

    document.getElementById('template-select')?.addEventListener('change', (e) => {
      const id = (e.target as HTMLSelectElement).value;
      this.selectTemplate(id);
    });

    document.getElementById('insert-btn')?.addEventListener('click', () => this.insertIntoEmail('cursor'));
    document.getElementById('replace-btn')?.addEventListener('click', () => this.insertIntoEmail('replace'));
    document.getElementById('reply-btn')?.addEventListener('click', () => this.replyWithTemplate());
    document.getElementById('free-reply-btn')?.addEventListener('click', () => this.generateFreeReplyContent());
    document.getElementById('free-reply-send-btn')?.addEventListener('click', () => this.sendFreeReply());
  }

  private async loadContext(): Promise<void> {
    const recipientInfo = document.getElementById('recipient-info')!;

    try {
      const item = Office.context.mailbox.item;
      if (!item) {
        recipientInfo.innerHTML = '<p class="empty-state">Aucun email selectionne</p>';
        return;
      }

      if (this.isReply) {
        // ── Reply mode (read mode in Inbox) ──
        // The "recipient" of our reply is the SENDER of the current email
        const from = item.from;
        if (from) {
          this.recipientEmail = from.emailAddress || '';
          recipientInfo.innerHTML = `
            <div style="font-size:11px;color:var(--atlas-text-secondary);text-transform:uppercase;letter-spacing:0.5px;margin-bottom:4px;">Repondre a</div>
            <div class="email-card-subject">${this.escapeHtml(from.displayName || '')} &lt;${this.escapeHtml(from.emailAddress || '')}&gt;</div>
            <div class="email-card-meta" style="margin-top:4px;">${this.escapeHtml(typeof item.subject === 'string' ? item.subject : '')}</div>
          `;
          this.loadArgoProfile();
        }
        // Analyze the received email for tone + generate quick replies
        this.analyzeReceivedContent();
        this.loadQuickReplies();
      } else {
        // ── Compose mode (new email or forward) ──
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
          const to = (item as any).to;
          if (Array.isArray(to) && to.length > 0) {
            this.recipientEmail = to[0]?.emailAddress || '';
            recipientInfo.innerHTML = `
              <div class="email-card-subject">${to.map((r: any) => `${r.displayName || ''} &lt;${r.emailAddress}&gt;`).join(', ')}</div>
            `;
            this.loadArgoProfile();
          }
        }
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
    const actionBtn = (document.getElementById('reply-btn') || document.getElementById('insert-btn')) as HTMLButtonElement;
    const varsSection = document.getElementById('template-variables')!;

    if (!this.selectedTemplate) {
      previewSection.style.display = 'none';
      varsSection.style.display = 'none';
      if (actionBtn) actionBtn.disabled = true;
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
    if (actionBtn) actionBtn.disabled = false;
    this.updatePreview();
  }

  /** Pick the right language version of the template based on ARGO profile */
  private getTemplateContent(): { body: string; subject: string; lang: string } {
    if (!this.selectedTemplate) return { body: '', subject: '', lang: 'FR' };
    const lang = this.argoProfile?.languePreferee || 'FR';
    const t = this.selectedTemplate;
    switch (lang) {
      case 'EN': return { body: t.corpsEN || t.corpsFR || '', subject: t.sujetEN || t.sujetFR || '', lang: 'EN' };
      case 'DE': return { body: (t as any).corpsDE || t.corpsFR || '', subject: (t as any).sujetDE || t.sujetFR || '', lang: 'DE' };
      case 'LU': return { body: (t as any).corpsLU || t.corpsFR || '', subject: (t as any).sujetLU || t.sujetFR || '', lang: 'LU' };
      default: return { body: t.corpsFR || '', subject: t.sujetFR || '', lang: 'FR' };
    }
  }

  private updatePreview(): void {
    if (!this.selectedTemplate) return;

    const preview = document.getElementById('template-preview')!;
    const content = this.getTemplateContent();
    let body = content.body;
    let subject = content.subject;

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

    const hasAI = !!localStorage.getItem('atlas_addin_anthropic_key') && !!this.argoProfile;
    const langBadge = content.lang !== 'FR' ? `<span style="display:inline-block;background:var(--atlas-primary);color:#fff;border-radius:3px;padding:1px 6px;font-size:9px;margin-left:6px;">${content.lang}</span>` : '';

    preview.innerHTML = `
      ${hasAI ? `<div style="display:flex;align-items:center;gap:4px;margin-bottom:6px;font-size:10px;color:var(--atlas-primary);"><span>✨</span> Sera adapte par ARGO (${this.argoProfile!.tonPrefere === 'Amical' ? 'tu' : 'vous'}, ${content.lang})${langBadge}</div>` : ''}
      <div style="margin-bottom:8px;font-weight:600;font-size:12px;">Objet: ${this.escapeHtml(subject)}</div>
      <hr style="border:none;border-top:1px solid var(--atlas-border);margin:8px 0;"/>
      <p>${salutation}</p>
      <div>${body}</div>
      <p>${closing}</p>
      <p>${this.userName}<br/>GOOD VIBES events & communications</p>
    `;
  }

  private async insertIntoEmail(mode: 'cursor' | 'replace' = 'cursor'): Promise<void> {
    if (!this.selectedTemplate) return;

    const insertBtn = document.getElementById('insert-btn') as HTMLButtonElement;
    const replaceBtn = document.getElementById('replace-btn') as HTMLButtonElement;
    insertBtn.disabled = true;
    replaceBtn.disabled = true;
    insertBtn.textContent = 'Personnalisation IA...';

    try {
      const content = this.getTemplateContent();
      let body = content.body;
      let subject = content.subject;

      // Replace variables
      document.querySelectorAll('.var-input').forEach((input) => {
        const varName = (input as HTMLInputElement).getAttribute('data-var')!;
        const value = (input as HTMLInputElement).value || '';
        body = body.replaceAll(`{{${varName}}}`, value);
        subject = subject.replaceAll(`{{${varName}}}`, value);
      });

      // Apply ARGO salutation + closing
      const salutation = getSalutation(this.argoProfile, this.userName);
      const closing = getClosing(this.argoProfile, this.userName);

      let fullHtml = `<p>${salutation}</p>${body}<p>${closing}</p><p>${this.userName}<br/>GOOD VIBES events &amp; communications</p>`;

      // Full AI adaptation if Anthropic key available — Claude rewrites naturally
      if (this.argoProfile && localStorage.getItem('atlas_addin_anthropic_key')) {
        try {
          fullHtml = await adaptEmailBody(fullHtml, this.argoProfile, this.userName);
        } catch { /* fallback to assembled version */ }
      }

      // Insert into Outlook compose window
      const item = Office.context.mailbox.item;
      if (!item) { showToast('Aucun email ouvert', 'error'); return; }

      // Set subject (only if template has one and it's not a reply)
      if (subject && !this.isReply) {
        (item as any).subject?.setAsync?.(subject);
      }

      if (mode === 'replace') {
        // Replace entire body
        (item as any).body?.setAsync?.(fullHtml, { coercionType: Office.CoercionType.Html }, (result: any) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            showToast('Template applique — contenu remplace', 'success');
          } else {
            showToast('Erreur d\'insertion', 'error');
          }
        });
      } else {
        // Insert at cursor position (keeps existing content)
        (item as any).body?.setSelectedDataAsync?.(fullHtml, { coercionType: Office.CoercionType.Html }, (result: any) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            showToast('Template insere au curseur', 'success');
          } else {
            // Fallback: prepend if setSelectedDataAsync not supported
            (item as any).body?.prependAsync?.(fullHtml, { coercionType: Office.CoercionType.Html }, (r2: any) => {
              if (r2.status === Office.AsyncResultStatus.Succeeded) {
                showToast('Template insere en debut de mail', 'success');
              } else {
                showToast('Erreur d\'insertion', 'error');
              }
            });
          }
        });
      }
    } catch (err) {
      showToast(`Erreur : ${(err as Error).message}`, 'error');
    } finally {
      insertBtn.disabled = false;
      replaceBtn.disabled = false;
      insertBtn.textContent = 'Inserer au curseur';
    }
  }

  private freeReplyHtml = '';

  /** Load 3 quick reply suggestions from AI */
  private async loadQuickReplies(): Promise<void> {
    if (!localStorage.getItem('atlas_addin_anthropic_key')) return;
    try {
      const item = Office.context.mailbox.item;
      if (!item) return;
      const subject = typeof item.subject === 'string' ? item.subject : '';
      const bodyText = await new Promise<string>((resolve) => {
        (item as any).body?.getAsync?.(Office.CoercionType.Text, (r: any) => {
          resolve(r?.status === Office.AsyncResultStatus.Succeeded ? r.value || '' : '');
        });
        setTimeout(() => resolve(''), 3000);
      });
      if (!bodyText) return;

      const senderName = item.from?.displayName || '';
      const replies = await generateQuickReplies(subject, bodyText, senderName, this.argoProfile, this.userName);
      if (replies.length === 0) return;

      const section = document.getElementById('quick-replies-section')!;
      const container = document.getElementById('quick-replies-container')!;
      section.style.display = 'block';

      container.innerHTML = replies.map((r, i) => `
        <button class="btn btn-secondary btn-block quick-reply-btn" data-idx="${i}" style="text-align:left; padding:10px 12px; font-size:12px; line-height:1.4;">
          <strong>${this.escapeHtml(r.label)}</strong>
        </button>
      `).join('');

      container.querySelectorAll('.quick-reply-btn').forEach(btn => {
        btn.addEventListener('click', () => {
          const idx = parseInt(btn.getAttribute('data-idx')!);
          const reply = replies[idx];
          if (reply) {
            const mailItem = Office.context.mailbox.item;
            (mailItem as any)?.displayReplyForm?.({ htmlBody: reply.body });
            showToast('Reponse rapide ouverte', 'success');
          }
        });
      });
    } catch { /* non-blocking */ }
  }

  /** Generate a free-form reply from user instruction */
  private async generateFreeReplyContent(): Promise<void> {
    const instruction = (document.getElementById('free-reply-instruction') as HTMLTextAreaElement)?.value?.trim();
    if (!instruction) { showToast('Decrivez ce que vous voulez repondre', 'error'); return; }

    const btn = document.getElementById('free-reply-btn') as HTMLButtonElement;
    btn.disabled = true;
    btn.textContent = '✨ Generation en cours...';

    try {
      const item = Office.context.mailbox.item;
      if (!item) return;
      const subject = typeof item.subject === 'string' ? item.subject : '';
      const bodyText = await new Promise<string>((resolve) => {
        (item as any).body?.getAsync?.(Office.CoercionType.Text, (r: any) => {
          resolve(r?.status === Office.AsyncResultStatus.Succeeded ? r.value || '' : '');
        });
        setTimeout(() => resolve(''), 3000);
      });

      const senderName = item.from?.displayName || '';
      this.freeReplyHtml = await generateFreeReply(subject, bodyText, senderName, instruction, this.argoProfile, this.userName);

      // Show preview
      const previewSection = document.getElementById('free-reply-preview')!;
      const previewContent = document.getElementById('free-reply-content')!;
      previewSection.style.display = 'block';
      previewContent.innerHTML = this.freeReplyHtml;
    } catch (err) {
      showToast(`Erreur IA : ${(err as Error).message}`, 'error');
    } finally {
      btn.disabled = false;
      btn.textContent = '✨ Generer la reponse';
    }
  }

  /** Send the free-form generated reply */
  private async sendFreeReply(): Promise<void> {
    if (!this.freeReplyHtml) return;
    const item = Office.context.mailbox.item;
    (item as any)?.displayReplyForm?.({ htmlBody: this.freeReplyHtml });
    showToast('Reponse ouverte', 'success');
  }

  /** Reply mode: open Outlook reply form with adapted template content */
  private async replyWithTemplate(): Promise<void> {
    if (!this.selectedTemplate) return;

    const replyBtn = document.getElementById('reply-btn') as HTMLButtonElement;
    replyBtn.disabled = true;
    replyBtn.textContent = 'Personnalisation IA...';

    try {
      const content = this.getTemplateContent();
      let body = content.body;

      // Replace variables
      document.querySelectorAll('.var-input').forEach((input) => {
        const varName = (input as HTMLInputElement).getAttribute('data-var')!;
        const value = (input as HTMLInputElement).value || '';
        body = body.replaceAll(`{{${varName}}}`, value);
      });

      // Apply ARGO salutation + closing
      const salutation = getSalutation(this.argoProfile, this.userName);
      const closing = getClosing(this.argoProfile, this.userName);
      let fullHtml = `<p>${salutation}</p>${body}<p>${closing}</p><p>${this.userName}<br/>GOOD VIBES events &amp; communications</p>`;

      // Full AI adaptation if Anthropic key available
      if (this.argoProfile && localStorage.getItem('atlas_addin_anthropic_key')) {
        try {
          fullHtml = await adaptEmailBody(fullHtml, this.argoProfile, this.userName);
        } catch { /* fallback */ }
      }

      // Open Outlook reply form with the adapted content
      const item = Office.context.mailbox.item;
      if (item) {
        (item as any).displayReplyForm?.({
          htmlBody: fullHtml,
        });
        showToast('Reponse ouverte avec le template adapte', 'success');
      }
    } catch (err) {
      showToast(`Erreur : ${(err as Error).message}`, 'error');
    } finally {
      replyBtn.disabled = false;
      replyBtn.textContent = 'Repondre avec ce template';
    }
  }

  private escapeHtml(str: string): string {
    return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  destroy(): void {
    this.container.innerHTML = '';
  }
}
