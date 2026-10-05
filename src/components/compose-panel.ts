/**
 * Compose Panel — Templates + ARGO tone adaptation for email composition
 *
 * Phase 6 de l'agent d'inbox (à la demande) : « Modèles de réponse » de l'agent (fiches
 * Communications remplies pour le dernier mail reçu du fil, `conversationId`), insérés au curseur
 * si Mailbox 1.2 est disponible (`body.setSelectedDataAsync`, testé par isSetSupported), sinon
 * texte à copier ; « Pièces jointes lourdes ? » (Mailbox 1.8, `getAttachmentsAsync`, testé) →
 * « Déposer sur le cloud et insérer les liens » : brouillon enregistré (saveAsync), le worker lit
 * les pièces par Graph, les dépose dans le transit du CLOUD (30 jours, jamais le NAS) et renvoie
 * les liens courts vibes.lu/p/…, insérés au curseur (bouton compatible Outlook) ; pièces déposées
 * retirées du brouillon si la case reste cochée (removeAttachmentAsync). Lien court d'un lien de
 * partage collé : toujours possible. Aucun envoi, jamais de signature.
 */

import { getEmailTemplates, fetchContactArgoProfile } from '../api/airtable';
import { getSalutation, getClosing, adaptEmailBody, analyzeReceivedEmail, generateQuickReplies, generateFreeReply } from '../api/argo';
import type { EmailTemplate, ArgoProfile } from '../types';
import { showToast } from '../taskpane';
import { supportsMailbox } from '../api/platform';
import { renderModeles, renderPiecesLourdes, type DepotCtx } from './agent-outils';
import { convertToRestId } from '../api/graph';
import { escapeHtml, escapeError, sanitizeHtml } from '../utils/html';
import { humanError } from '../api/net';
import { icon } from '../ui/icons';
import { inlineLoadingHtml, errorHtml } from '../ui/states';

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
    this.renderOutilsAgent();
  }

  private render(): void {
    this.container.innerHTML = `
      <div class="panel-scroll stack">
        <section class="section" aria-labelledby="cp-h-dest">
          <h2 class="section-heading" id="cp-h-dest">${this.isReply ? 'Tu réponds à' : 'Destinataire'}</h2>
          <div id="recipient-info" class="email-card">${inlineLoadingHtml('Lecture du destinataire…')}</div>
        </section>

        <section id="argo-profile-section" class="section" style="display:none;" aria-labelledby="cp-h-prof">
          <h2 class="section-heading" id="cp-h-prof">Comment lui écrire, selon ARGO</h2>
          <div id="argo-profile-info" class="card"></div>
        </section>

        <section id="received-analysis" class="section" style="display:none;" aria-labelledby="cp-h-ana">
          <h2 class="section-heading" id="cp-h-ana">Le mail reçu, lu par ARGO</h2>
          <div id="received-analysis-info" class="card"></div>
        </section>

        ${this.isReply ? `
        <section id="quick-replies-section" class="section" style="display:none;" aria-labelledby="cp-h-quick">
          <h2 class="section-heading" id="cp-h-quick">Réponses proposées par ARGO</h2>
          <div id="quick-replies-container" class="stack-sm"></div>
        </section>

        <section id="free-reply-section" class="section" aria-labelledby="cp-h-free">
          <h2 class="section-heading" id="cp-h-free">Dis à ARGO quoi répondre</h2>
          <textarea class="form-input" id="free-reply-instruction" rows="3" aria-labelledby="cp-h-free" placeholder="En quelques mots : « décline, trop cher, on verra l'an prochain » ou « confirme et propose mardi 10 h ». Vide : réponse polie d'attente."></textarea>
          <p class="help">Salutation, langue et tutoiement repris du mail reçu. La signature est ajoutée par Exclaimer.</p>
          <button type="button" class="btn btn-argo btn-block" id="free-reply-btn">${icon('bolt', 14)}<span>Rédiger avec ARGO</span><span class="argo-tag" aria-hidden="true">IA</span></button>
          <div id="free-reply-preview" class="stack-sm" style="display:none;">
            <p class="eyebrow">Aperçu</p>
            <div id="free-reply-content" class="template-preview"></div>
            <button type="button" class="btn btn-primary btn-block" id="free-reply-send-btn">${icon('reply', 14)}Répondre avec ce texte</button>
          </div>
        </section>
        ` : ''}

        <div id="agent-modeles-compose" class="agent-section" hidden></div>
        <div id="agent-pj-compose-wrap" class="agent-section" hidden>
          <button type="button" class="btn btn-secondary btn-block agent-btn" id="agent-pj-compose-btn">${icon('paperclip', 14)}Pièces jointes lourdes ?</button>
          <div id="agent-pj-compose" hidden></div>
        </div>

        <section class="section" aria-labelledby="cp-h-tpl">
          <h2 class="section-heading" id="cp-h-tpl">${this.isReply ? 'Ou partir d\'un modèle' : 'Modèle de mail'}</h2>
          <label class="sr-only" for="template-select">Choisir un modèle</label>
          <select class="dropdown-select" id="template-select">
            <option value="">Chargement des modèles…</option>
          </select>
          <div id="template-variables" class="stack-sm" style="display:none;">
            <p class="eyebrow">À compléter</p>
            <div id="variables-container"></div>
          </div>
          <div id="template-preview-section" class="stack-sm" style="display:none;">
            <p class="eyebrow">Aperçu</p>
            <div id="template-preview" class="template-preview"></div>
          </div>
          <div class="stack-sm">
            ${this.isReply ? `
              <button type="button" class="btn btn-primary btn-block" id="reply-btn" disabled>${icon('reply', 14)}Répondre avec ce modèle</button>
            ` : `
              <button type="button" class="btn btn-primary btn-block" id="insert-btn" disabled>${icon('template', 14)}Insérer au curseur</button>
              <button type="button" class="btn btn-ghost btn-block" id="replace-btn" disabled>Remplacer tout le contenu</button>
            `}
          </div>
        </section>
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

  // ── Agent d'inbox, phase 6 : modèles de réponse et pièces jointes lourdes ──

  private renderOutilsAgent(): void {
    let conversationId = '';
    try { conversationId = String((Office.context.mailbox?.item as any)?.conversationId || ''); } catch { /* hors Outlook */ }
    const modeles = this.container.querySelector<HTMLElement>('#agent-modeles-compose');
    // Réponse (fil connu) seulement : le modèle est rempli pour le dernier mail reçu du fil.
    if (modeles && conversationId) {
      renderModeles(modeles, {
        conversationId, onInfo: showToast,
        ...(supportsMailbox('1.2') ? { inserer: (texte: string) => this.insererTexte(texte) } : {}),
      });
    }
    // getAttachmentsAsync (compose) : Mailbox 1.8, absent sur mobile → bouton masqué.
    const wrap = this.container.querySelector<HTMLElement>('#agent-pj-compose-wrap');
    if (!wrap || !supportsMailbox('1.8')) return;
    wrap.hidden = false;
    const btn = this.container.querySelector<HTMLButtonElement>('#agent-pj-compose-btn');
    const host = this.container.querySelector<HTMLElement>('#agent-pj-compose');
    btn?.addEventListener('click', () => {
      if (!host) return;
      btn.disabled = true;
      try {
        (Office.context.mailbox.item as any).getAttachmentsAsync((r: any) => {
          btn.disabled = false;
          if (r?.status !== Office.AsyncResultStatus.Succeeded) { showToast('Pièces jointes illisibles', 'error'); return; }
          const pieces = (Array.isArray(r.value) ? r.value : []).map((a: any) => ({ id: a?.id ? String(a.id) : undefined, nom: String(a?.name || ''), octets: Number(a?.size) || 0, isInline: !!a?.isInline }));
          void renderPiecesLourdes(host, { pieces, onInfo: showToast, depot: this.depotBrouillon() }).then(() => {
            if (host.hidden) showToast('Pièces jointes sous le seuil : rien à faire', 'info');
          });
        });
      } catch {
        btn.disabled = false;
        showToast('Pièces jointes illisibles', 'error');
      }
    });
  }

  /**
   * Dépôt des pièces du brouillon : enregistrement (saveAsync, Mailbox 1.3) pour que le worker lise
   * les pièces par Graph, insertion des liens au curseur (Mailbox 1.2), retrait des pièces (1.1).
   */
  private depotBrouillon(): DepotCtx | undefined {
    if (!supportsMailbox('1.3')) return undefined;
    const item = Office.context.mailbox.item as any;
    return {
      source: () => new Promise((resolve, reject) => {
        try {
          item.saveAsync((r: any) => {
            if (r?.status !== Office.AsyncResultStatus.Succeeded || !r.value) { reject(new Error('Brouillon non enregistré')); return; }
            let mailbox = '';
            try { mailbox = String(Office.context.mailbox.userProfile?.emailAddress || ''); } catch { /* boîte par défaut côté worker */ }
            resolve({ messageId: convertToRestId(String(r.value)), ...(mailbox ? { mailbox } : {}) });
          });
        } catch (e) { reject(e instanceof Error ? e : new Error('Brouillon non enregistré')); }
      }),
      idRest: (id: string) => convertToRestId(id),
      inserer: (html: string) => new Promise<boolean>(resolve => {
        if (!supportsMailbox('1.2')) { resolve(false); return; }
        try {
          item.body.setSelectedDataAsync(sanitizeHtml(html), { coercionType: Office.CoercionType.Html }, (r: any) => resolve(r?.status === Office.AsyncResultStatus.Succeeded));
        } catch { resolve(false); }
      }),
      retirer: async (ids: string[]) => {
        let n = 0;
        for (const id of ids) {
          const ok = await new Promise<boolean>(resolve => {
            try { item.removeAttachmentAsync(id, (r: any) => resolve(r?.status === Office.AsyncResultStatus.Succeeded)); } catch { resolve(false); }
          });
          if (ok) n++;
        }
        return n;
      },
    };
  }

  /** Insère un texte (modèle rempli) au curseur : body.setSelectedDataAsync, Mailbox 1.2 (testé par l'appelant). */
  private insererTexte(texte: string): Promise<boolean> {
    const esc = (t: string) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const html = texte.split(/\n{2,}/).map(p => `<p>${esc(p).replace(/\n/g, '<br/>')}</p>`).join('');
    return new Promise(resolve => {
      try {
        (Office.context.mailbox.item as any).body.setSelectedDataAsync(html, { coercionType: Office.CoercionType.Html }, (r: any) => {
          resolve(r?.status === Office.AsyncResultStatus.Succeeded);
        });
      } catch { resolve(false); }
    });
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
            <div class="email-card-subject">${this.escapeHtml(from.displayName || from.emailAddress || '')}</div>
            <div class="email-card-meta">${this.escapeHtml(from.emailAddress || '')}</div>
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
                <div class="email-card-subject">${to.map((r: any) => `${escapeHtml(r.displayName || '')} &lt;${escapeHtml(r.emailAddress)}&gt;`).join(', ')}</div>
              `;
              this.loadArgoProfile();
            } else {
              recipientInfo.innerHTML = '<p class="help">Ajoute un destinataire : ARGO adaptera le ton à cette personne.</p>';
            }
          });
        } else if ((item as any).to) {
          const to = (item as any).to;
          if (Array.isArray(to) && to.length > 0) {
            this.recipientEmail = to[0]?.emailAddress || '';
            recipientInfo.innerHTML = `
              <div class="email-card-subject">${to.map((r: any) => `${escapeHtml(r.displayName || '')} &lt;${escapeHtml(r.emailAddress)}&gt;`).join(', ')}</div>
            `;
            this.loadArgoProfile();
          }
        }
      }

      // Load templates
      await this.loadTemplates();

    } catch (err) {
      recipientInfo.innerHTML = errorHtml(err, { title: 'Destinataire illisible', retry: false, compact: true });
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
          <p class="person-name">${escapeHtml(this.argoProfile.prenom)} ${escapeHtml(this.argoProfile.nom)}</p>
          <div class="chips"><span class="chip">${isTu ? 'Tutoiement' : 'Vouvoiement'}</span><span class="chip">${escapeHtml(this.argoProfile.languePreferee || 'FR')}</span></div>
          ${this.argoProfile.tutoiementAvec.length > 0 ? `<p class="help" style="margin-top:6px;">Tutoie : ${escapeHtml(this.argoProfile.tutoiementAvec.join(', '))}</p>` : ''}
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
          <dl class="facts" style="margin-top:0;">
            <dt>Ressenti</dt><dd>${escapeHtml(analysis.sentiment)}</dd>
            <dt>Urgence</dt><dd>${escapeHtml(analysis.urgence)}</dd>
            <dt>Ton</dt><dd>${escapeHtml(analysis.tonUtilise)}</dd>
          </dl>
          ${analysis.suggestions.length > 0 ? `<p class="eyebrow" style="margin-top:10px;">Pistes de réponse</p><ul class="agent-facts" style="margin-top:4px;">${analysis.suggestions.map(s => `<li>${escapeHtml(s)}</li>`).join('')}</ul>` : ''}
        `;
      });
    } catch { /* non-blocking */ }
  }

  private async loadTemplates(): Promise<void> {
    try {
      this.templates = await getEmailTemplates();
      const select = document.getElementById('template-select') as HTMLSelectElement;

      select.innerHTML = `
        <option value="">Choisir un modèle</option>
        ${this.templates.map(t => `
          <option value="${escapeHtml(t.id)}">${escapeHtml(t.nom)} ${t.marque ? `(${escapeHtml(t.marque)})` : ''}</option>
        `).join('')}
      `;
    } catch (err) {
      showToast('Modèles indisponibles pour le moment : réessaie dans un instant.', 'error');
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
          <label class="form-label">${escapeHtml(v)}</label>
          <input type="text" class="form-input var-input" data-var="${escapeHtml(v)}" placeholder="${escapeHtml(v)}" />
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
      body = body.replaceAll(`{{${varName}}}`, escapeHtml(value));
      subject = subject.replaceAll(`{{${varName}}}`, value);
    });

    // Apply ARGO salutation/closing
    const salutation = getSalutation(this.argoProfile, this.userName);
    const closing = getClosing(this.argoProfile, this.userName);

    // IA servie par le worker (aucune clé dans le complément) : adaptation dès qu'un profil ARGO existe.
    const hasAI = !!this.argoProfile;
    const langBadge = content.lang !== 'FR' ? ` <span class="badge">${escapeHtml(content.lang)}</span>` : '';

    preview.innerHTML = `
      ${hasAI ? `<p class="help" style="margin-bottom:6px;">${icon('bolt', 12)}ARGO l'adaptera à la personne (${this.argoProfile!.tonPrefere === 'Amical' ? 'tu' : 'vous'}, ${escapeHtml(content.lang)})${langBadge}</p>` : ''}
      <div style="margin-bottom:8px;font-weight:650;font-size:12px;">Objet : ${this.escapeHtml(subject)}</div>
      <hr style="border:none;border-top:1px solid var(--atlas-border);margin:8px 0;"/>
      <p>${escapeHtml(salutation)}</p>
      <div>${sanitizeHtml(body)}</div>
      <p>${escapeHtml(closing)}</p>
      <p>${escapeHtml(this.userName)}<br/>GOOD VIBES events & communications</p>
    `;
  }

  private async insertIntoEmail(mode: 'cursor' | 'replace' = 'cursor'): Promise<void> {
    if (!this.selectedTemplate) return;

    const insertBtn = document.getElementById('insert-btn') as HTMLButtonElement;
    const replaceBtn = document.getElementById('replace-btn') as HTMLButtonElement;
    insertBtn.disabled = true;
    replaceBtn.disabled = true;
    insertBtn.textContent = 'ARGO adapte le modèle…';

    try {
      const content = this.getTemplateContent();
      let body = content.body;
      let subject = content.subject;

      // Replace variables
      document.querySelectorAll('.var-input').forEach((input) => {
        const varName = (input as HTMLInputElement).getAttribute('data-var')!;
        const value = (input as HTMLInputElement).value || '';
        body = body.replaceAll(`{{${varName}}}`, escapeHtml(value));
        subject = subject.replaceAll(`{{${varName}}}`, value);
      });

      // Apply ARGO salutation + closing
      const salutation = getSalutation(this.argoProfile, this.userName);
      const closing = getClosing(this.argoProfile, this.userName);

      let fullHtml = `<p>${escapeHtml(salutation)}</p>${sanitizeHtml(body)}<p>${escapeHtml(closing)}</p><p>${escapeHtml(this.userName)}<br/>GOOD VIBES events &amp; communications</p>`;

      // Adaptation IA complète (worker + Agency Brain) — Claude réécrit naturellement
      if (this.argoProfile) {
        try {
          fullHtml = sanitizeHtml(await adaptEmailBody(fullHtml, this.argoProfile, this.userName));
        } catch { /* fallback to assembled version */ }
      }

      // Insert into Outlook compose window
      const item = Office.context.mailbox.item;
      if (!item) { showToast('Aucun mail ouvert', 'error'); return; }

      // Set subject (only if template has one and it's not a reply)
      if (subject && !this.isReply) {
        (item as any).subject?.setAsync?.(subject);
      }

      if (mode === 'replace') {
        // Replace entire body
        (item as any).body?.setAsync?.(fullHtml, { coercionType: Office.CoercionType.Html }, (result: any) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            showToast('Modèle appliqué : le contenu a été remplacé', 'success');
          } else {
            showToast('Insertion impossible : réessaie.', 'error');
          }
        });
      } else {
        // Insert at cursor position (keeps existing content)
        (item as any).body?.setSelectedDataAsync?.(fullHtml, { coercionType: Office.CoercionType.Html }, (result: any) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            showToast('Modèle inséré au curseur', 'success');
          } else {
            // Fallback: prepend if setSelectedDataAsync not supported
            (item as any).body?.prependAsync?.(fullHtml, { coercionType: Office.CoercionType.Html }, (r2: any) => {
              if (r2.status === Office.AsyncResultStatus.Succeeded) {
                showToast('Modèle inséré en début de mail', 'success');
              } else {
                showToast('Insertion impossible : réessaie.', 'error');
              }
            });
          }
        });
      }
    } catch (err) {
      showToast(`${humanError(err)}`, 'error');
    } finally {
      insertBtn.disabled = false;
      replaceBtn.disabled = false;
      insertBtn.innerHTML = `${icon('template', 14)}Insérer au curseur`;
    }
  }

  private freeReplyHtml = '';

  /** Load 3 quick reply suggestions from AI */
  private async loadQuickReplies(): Promise<void> {
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
        <button type="button" class="btn btn-secondary btn-block quick-reply-btn" data-idx="${i}" style="white-space:normal; text-align:left;">${icon('reply', 14)}
          <strong>${this.escapeHtml(r.label)}</strong>
        </button>
      `).join('');

      container.querySelectorAll('.quick-reply-btn').forEach(btn => {
        btn.addEventListener('click', () => {
          const idx = parseInt(btn.getAttribute('data-idx')!);
          const reply = replies[idx];
          if (reply) {
            const mailItem = Office.context.mailbox.item;
            (mailItem as any)?.displayReplyForm?.({ htmlBody: sanitizeHtml(reply.body) });
            showToast('Réponse ouverte', 'success');
          }
        });
      });
    } catch { /* non-blocking */ }
  }

  /** Generate a free-form reply from user instruction */
  private async generateFreeReplyContent(): Promise<void> {
    const instruction = (document.getElementById('free-reply-instruction') as HTMLTextAreaElement)?.value?.trim();
    if (!instruction) { showToast('Dis en quelques mots ce que tu veux répondre.', 'error'); return; }

    const btn = document.getElementById('free-reply-btn') as HTMLButtonElement;
    btn.disabled = true;
    btn.innerHTML = `${icon('bolt', 14)}<span>ARGO rédige…</span>`;

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
      this.freeReplyHtml = sanitizeHtml(await generateFreeReply(subject, bodyText, senderName, instruction, this.argoProfile, this.userName));

      // Show preview
      const previewSection = document.getElementById('free-reply-preview')!;
      const previewContent = document.getElementById('free-reply-content')!;
      previewSection.style.display = 'block';
      previewContent.innerHTML = this.freeReplyHtml;
    } catch (err) {
      showToast(`${humanError(err)}`, 'error');
    } finally {
      btn.disabled = false;
      btn.innerHTML = `${icon('bolt', 14)}<span>Rédiger avec ARGO</span><span class="argo-tag" aria-hidden="true">IA</span>`;
    }
  }

  /** Send the free-form generated reply */
  private async sendFreeReply(): Promise<void> {
    if (!this.freeReplyHtml) return;
    const item = Office.context.mailbox.item;
    (item as any)?.displayReplyForm?.({ htmlBody: this.freeReplyHtml });
    showToast('Réponse ouverte', 'success');
  }

  /** Reply mode: open Outlook reply form with adapted template content */
  private async replyWithTemplate(): Promise<void> {
    if (!this.selectedTemplate) return;

    const replyBtn = document.getElementById('reply-btn') as HTMLButtonElement;
    replyBtn.disabled = true;
    replyBtn.textContent = 'ARGO adapte le modèle…';

    try {
      const content = this.getTemplateContent();
      let body = content.body;

      // Replace variables
      document.querySelectorAll('.var-input').forEach((input) => {
        const varName = (input as HTMLInputElement).getAttribute('data-var')!;
        const value = (input as HTMLInputElement).value || '';
        body = body.replaceAll(`{{${varName}}}`, escapeHtml(value));
      });

      // Apply ARGO salutation + closing
      const salutation = getSalutation(this.argoProfile, this.userName);
      const closing = getClosing(this.argoProfile, this.userName);
      let fullHtml = `<p>${escapeHtml(salutation)}</p>${sanitizeHtml(body)}<p>${escapeHtml(closing)}</p><p>${escapeHtml(this.userName)}<br/>GOOD VIBES events &amp; communications</p>`;

      // Adaptation IA complète (worker + Agency Brain)
      if (this.argoProfile) {
        try {
          fullHtml = sanitizeHtml(await adaptEmailBody(fullHtml, this.argoProfile, this.userName));
        } catch { /* fallback */ }
      }

      // Open Outlook reply form with the adapted content
      const item = Office.context.mailbox.item;
      if (item) {
        (item as any).displayReplyForm?.({
          htmlBody: fullHtml,
        });
        showToast('Réponse ouverte avec le modèle adapté', 'success');
      }
    } catch (err) {
      showToast(`${humanError(err)}`, 'error');
    } finally {
      replyBtn.disabled = false;
      replyBtn.innerHTML = `${icon('reply', 14)}Répondre avec ce modèle`;
    }
  }

  private escapeHtml(str: string): string {
    return escapeHtml(str);
  }

  destroy(): void {
    this.container.innerHTML = '';
  }
}
