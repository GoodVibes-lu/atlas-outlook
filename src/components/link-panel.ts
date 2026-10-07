/**
 * Link Panel — Link current email to a projet/tiers/contact
 */

import { getLinkedConversationIds, getAllLinkedEmailIds, linkEmailToProject, linkEmailToContact, getAllProjets, resolveClientIdInProjetsBase } from '../api/airtable';
import { lireMailPourLiaison } from '../api/mail-liaison';
import { renderClasserOutlook } from './agent-dossiers';
import { fetchProjetDuMail } from '../api/parite';
import { convertToRestId } from '../api/graph';
import { summarizeEmail } from '../api/argo';
import { SearchPicker } from './search-picker';
import type { SearchResult } from '../types';
import { showToast } from '../taskpane';
import { escapeHtml } from '../utils/html';
import { loadingHtml, inlineLoadingHtml, emptyHtml, errorHtml, renderError } from '../ui/states';
import { icon } from '../ui/icons';

interface EmailInfo {
  subject: string;
  from: string;
  fromEmail: string;
  to: string;
  date: string;
  conversationId?: string;
  internetMessageId?: string;
  itemId: string;
  isAlreadyLinked: boolean;
}

export class LinkPanel {
  private container: HTMLElement;
  private userName: string;
  private emailInfo: EmailInfo | null = null;
  /** Projet auquel le mail est lié dans ATLAS (fil lié ou appris), affiché en tête du panneau. */
  private projetLie: { id: string; libelle: string } | null = null;
  private isPrive = false;
  private searchPicker: SearchPicker | null = null;
  private searchExpanded = false;

  constructor(container: HTMLElement, userName: string) {
    this.container = container;
    this.userName = userName;
    this.render();
    this.loadEmailInfo();
  }

  private render(): void {
    this.container.innerHTML = `
      <div class="panel-scroll stack">
        <section class="section" aria-labelledby="lk-h-mail">
          <h2 class="section-heading" id="lk-h-mail">Ce mail</h2>
          <div id="email-info-card" class="email-card">${loadingHtml('Lecture du mail…', 2)}</div>
        </section>

        <section id="ai-summary" class="section" style="display:none;" aria-labelledby="lk-h-sum">
          <h2 class="section-heading" id="lk-h-sum">Résumé ARGO</h2>
          <p id="ai-summary-content" class="card card-note argo-summary"></p>
        </section>

        <div id="link-status" style="display:none;"></div>

        <div id="auto-suggestions" style="display:none;">
          <div id="auto-suggestion-content"></div>
        </div>

        <section id="link-search-section" class="section" aria-labelledby="lk-h-search">
          <h2 class="section-heading" id="lk-h-search">Lier à un projet, un client ou un contact</h2>
          <label class="toggle-row">
            <span class="toggle-switch" id="prive-toggle" role="switch" aria-checked="false" tabindex="0" aria-label="Marquer comme privé"></span>
            <span>${icon('lock', 14)} Privé : visible par toi seul dans ATLAS</span>
          </label>
          <div id="search-toggle-section" style="display:none;">
            <button type="button" class="agent-link" id="search-toggle-link" aria-expanded="false">${icon('search', 14)}<span>Chercher un autre projet, client ou contact</span></button>
            <div id="search-container" style="display:none;"></div>
          </div>
          <div id="search-direct-container"></div>
        </section>
      </div>
    `;

    // Toggle privé
    document.getElementById('prive-toggle')?.addEventListener('click', (e) => {
      const el = e.currentTarget as HTMLElement;
      this.isPrive = !this.isPrive;
      el.classList.toggle('active', this.isPrive);
      el.setAttribute('aria-checked', String(this.isPrive));
    });
    document.getElementById('prive-toggle')?.addEventListener('keydown', (e) => {
      if (e.key === ' ' || e.key === 'Enter') { e.preventDefault(); (e.currentTarget as HTMLElement).click(); }
    });

    // Search toggle link
    document.getElementById('search-toggle-link')?.addEventListener('click', (e) => {
      e.preventDefault();
      this.searchExpanded = !this.searchExpanded;
      const link = e.currentTarget as HTMLElement;
      const searchContainer = document.getElementById('search-container')!;
      link.setAttribute('aria-expanded', String(this.searchExpanded));
      searchContainer.style.display = this.searchExpanded ? 'block' : 'none';
    });

    // Init search picker in the collapsible container
    const searchContainer = document.getElementById('search-container')!;
    this.searchPicker = new SearchPicker(searchContainer, (result) => this.handleLink(result));

    // Also init a second search picker for the direct (no auto-detect) case
    const directContainer = document.getElementById('search-direct-container')!;
    this.searchPicker = new SearchPicker(directContainer, (result) => this.handleLink(result));
  }

  private async loadEmailInfo(): Promise<void> {
    const card = document.getElementById('email-info-card')!;

    try {
      const item = Office.context.mailbox.item;
      if (!item) {
        card.outerHTML = emptyHtml({ icon: 'mail', title: 'Aucun mail sélectionné', text: 'Ouvre un mail pour le lier à un projet.' });
        return;
      }

      // Read email properties via Office.js callbacks
      const subject = await this.getAsync<string>(item, 'subject');
      const from = item.from;
      const to = await this.getRecipientsAsync(item);
      const conversationId = item.conversationId;
      const internetMessageId = (item as any).internetMessageId;
      const itemId = item.itemId;

      this.emailInfo = {
        subject: subject || '(sans objet)',
        from: from?.displayName || '',
        fromEmail: from?.emailAddress || '',
        to: to,
        date: '',
        conversationId: conversationId || undefined,
        internetMessageId: internetMessageId || undefined,
        itemId: itemId || '',
        isAlreadyLinked: false,
      };

      // Check if already linked
      const { graphIds, internetIds } = await getAllLinkedEmailIds();
      const restId = convertToRestId(itemId);
      if (graphIds.has(restId) || (internetMessageId && internetIds.has(internetMessageId))) {
        this.emailInfo.isAlreadyLinked = true;
      }
      // Projet lié (07/10/2026, retour de Charles : « on ne voit nulle part à quel projet il est lié ») :
      // même source que l'Inbox ATLAS (fil lié dans ATLAS, sinon fil appris par l'agent).
      this.projetLie = null;
      if (internetMessageId) {
        this.projetLie = await fetchProjetDuMail(internetMessageId).catch(() => null);
        if (this.projetLie) this.emailInfo.isAlreadyLinked = true;
      }

      // Render email info
      card.innerHTML = `
        <div class="email-card-subject">${this.escapeHtml(this.emailInfo.subject)}</div>
        <dl class="facts">
          <dt>De</dt><dd>${this.escapeHtml(this.emailInfo.from || this.emailInfo.fromEmail)}${this.emailInfo.from ? `<br/><span class="meta">${this.escapeHtml(this.emailInfo.fromEmail)}</span>` : ''}</dd>
          <dt>À</dt><dd>${this.escapeHtml(this.emailInfo.to)}</dd>
        </dl>
        ${this.projetLie
          ? `<p class="status-linked" style="margin-top:10px;">${icon('check-circle', 14)}Lié au projet ${this.escapeHtml(this.projetLie.libelle)}</p>`
          : this.emailInfo.isAlreadyLinked ? `<p class="status-linked" style="margin-top:10px;">${icon('check-circle', 14)}Déjà lié dans ATLAS</p>` : ''}
      `;
      // Mail déjà lié à un projet : « Classer » dans le dossier Outlook du projet, comme l'Inbox ATLAS.
      if (this.projetLie && internetMessageId) this.afficherClassement(internetMessageId, this.projetLie);

      // Auto-detect project from subject (#NNN)
      await this.autoDetect();

      // Generate AI summary (non-blocking)
      this.loadAiSummary();

    } catch (err) {
      renderError(card, err, () => { card.innerHTML = loadingHtml('Lecture du mail…', 2); this.loadEmailInfo(); }, { title: 'Mail illisible pour le moment', compact: true });
    }
  }

  private async loadAiSummary(): Promise<void> {
    if (!this.emailInfo) return;
    try {
      // Get email body via Office.js for summary
      const item = Office.context.mailbox.item;
      if (!item) return;
      const bodyText = await new Promise<string>((resolve) => {
        (item as any).body?.getAsync?.(Office.CoercionType.Text, (r: any) => {
          resolve(r?.status === Office.AsyncResultStatus.Succeeded ? r.value || '' : '');
        });
        setTimeout(() => resolve(''), 3000);
      });
      if (!bodyText || bodyText.length < 20) return;

      const summary = await summarizeEmail(
        this.emailInfo.subject,
        bodyText,
        this.emailInfo.from
      );
      if (summary) {
        const el = document.getElementById('ai-summary')!;
        const content = document.getElementById('ai-summary-content')!;
        el.style.display = 'block';
        content.textContent = summary;
      }
    } catch { /* non-blocking */ }
  }

  private async autoDetect(): Promise<void> {
    if (!this.emailInfo) return;

    const autoSection = document.getElementById('auto-suggestions')!;
    const autoContent = document.getElementById('auto-suggestion-content')!;
    const searchToggleSection = document.getElementById('search-toggle-section')!;
    const directContainer = document.getElementById('search-direct-container')!;

    // 1. Check #NNN in subject
    const projetMatch = this.emailInfo.subject.match(/#\s*(\d{2,4})/);
    if (projetMatch) {
      const noProjet = projetMatch[1];
      const projets = await getAllProjets();
      const found = projets.find(p => String(p.noProjet) === noProjet);
      if (found) {
        autoSection.style.display = 'block';
        // Hide direct search, show collapsible toggle instead
        directContainer.style.display = 'none';
        searchToggleSection.style.display = 'block';

        const isLinked = this.projetLie ? this.projetLie.id === found.id : this.emailInfo.isAlreadyLinked;

        autoContent.innerHTML = `
          <section class="section" aria-labelledby="lk-h-det">
            <h2 class="section-heading" id="lk-h-det">Projet repéré dans l'objet</h2>
            <article class="card detected-project-card">
              <span class="eyebrow">${icon('folder', 12)}Projet #${this.escapeHtml(found.noProjet)}</span>
              <h3 class="card-title">${this.escapeHtml(found.denomination)}</h3>
              ${found.client || found.enCharge ? `<dl class="facts">${found.client ? `<dt>Client</dt><dd>${this.escapeHtml(found.client)}</dd>` : ''}${found.enCharge ? `<dt>En charge</dt><dd>${this.escapeHtml(found.enCharge)}</dd>` : ''}</dl>` : ''}
              <div class="card-foot">
                ${isLinked
                  ? `<p class="status-linked">${icon('check-circle', 14)}Ce mail est déjà lié à ce projet</p>`
                  : `<button type="button" class="btn btn-primary btn-block" id="auto-link-btn">${icon('link', 14)}Lier ce mail au projet #${this.escapeHtml(noProjet)}</button>`
                }
              </div>
            </article>
          </section>
        `;

        if (!isLinked) {
          document.getElementById('auto-link-btn')?.addEventListener('click', () => {
            this.handleLink({ type: 'projet', id: found.id, label: found.denomination, detail: found.client });
          });
        }
        return;
      }
    }

    // 2. Check conversation ID
    if (this.emailInfo.conversationId) {
      const convMap = await getLinkedConversationIds();
      const match = convMap.get(this.emailInfo.conversationId);
      if (match) {
        autoSection.style.display = 'block';
        // Hide direct search, show collapsible toggle
        directContainer.style.display = 'none';
        searchToggleSection.style.display = 'block';

        autoContent.innerHTML = `
          <section class="section" aria-labelledby="lk-h-conv">
            <h2 class="section-heading" id="lk-h-conv">Fil déjà lié à un projet</h2>
            <article class="card detected-project-card">
              <span class="eyebrow">${icon('mail', 12)}Même conversation</span>
              <h3 class="card-title">${this.escapeHtml(match.projetName)}</h3>
              <p class="help">Un mail précédent de ce fil est déjà rattaché à ce projet.</p>
              <div class="card-foot">
                <button type="button" class="btn btn-primary btn-block" id="auto-conv-btn">${icon('link', 14)}Lier ce mail au projet</button>
              </div>
            </article>
          </section>
        `;
        document.getElementById('auto-conv-btn')?.addEventListener('click', () => {
          this.handleLink({ type: 'projet', id: match.projetId, label: match.projetName, detail: '' });
        });
        return;
      }
    }

    // No auto-detection — show direct search, hide collapsible toggle
    directContainer.style.display = 'block';
    searchToggleSection.style.display = 'none';
  }

  private async handleLink(result: SearchResult): Promise<void> {
    if (!this.emailInfo) return;

    const statusEl = document.getElementById('link-status')!;
    statusEl.style.display = 'block';
    statusEl.innerHTML = inlineLoadingHtml('Liaison en cours…');

    try {
      const lu = await lireMailPourLiaison(this.emailInfo.itemId, this.emailInfo.internetMessageId);
      const fullMessage = lu.message;

      // Determine direction
      const userEmail = localStorage.getItem('atlas_addin_user_email') || Office.context?.mailbox?.userProfile?.emailAddress || '';
      const isSent = fullMessage.from.email.toLowerCase() === userEmail.toLowerCase();
      const direction = isSent ? 'envoyé' as const : 'reçu' as const;

      const priveOpts = this.isPrive
        ? { prive: true, privePar: userEmail }
        : undefined;

      if (result.type === 'projet') {
        await linkEmailToProject(fullMessage, result.id, this.userName + ' (Outlook)', direction, undefined, priveOpts);
      } else if (result.type === 'tiers') {
        const tiersId = await resolveClientIdInProjetsBase(result.label);
        if (tiersId) {
          await linkEmailToProject(fullMessage, '', this.userName + ' (Outlook)', direction, tiersId, priveOpts);
        }
      } else if (result.type === 'contact') {
        await linkEmailToContact(fullMessage, result.label, this.userName + ' (Outlook)', direction, result.detail, priveOpts);
      }

      statusEl.innerHTML = `<p class="status-linked">${icon('check-circle', 14)}Mail lié à ${this.escapeHtml(result.label)}</p>`;
      showToast(`Email lié à ${result.label}`, 'success');

      // Hide search section
      document.getElementById('link-search-section')!.style.display = 'none';
      document.getElementById('auto-suggestions')!.style.display = 'none';

      // Offer to file in Outlook folder if mapping exists
      // Classement dans Outlook, comme l'Inbox ATLAS après une liaison.
      if (result.type === 'projet') this.projetLie = { id: result.id, libelle: result.label };
      this.afficherClassement(fullMessage.internetMessageId || this.emailInfo.internetMessageId || '', this.projetLie || undefined);

    } catch (err) {
      statusEl.innerHTML = errorHtml(err, { title: 'Liaison impossible', retry: false, compact: true });
      showToast('Erreur de liaison', 'error');
    }
  }

  // ── Folder Filing ──

  /**
   * Classement dans Outlook (07/10/2026) : la même carte que l'onglet « Ce mail » (agent-dossiers.ts ›
   * renderClasserOutlook) : dossier du projet en un clic, sinon création au chemin d'ATLAS, ou un autre
   * dossier par recherche. Le dossier du projet est retenu pour ATLAS (rien à ressaisir dans l'Inbox).
   */
  private afficherClassement(internetMessageId: string, projet?: { id: string; libelle: string }): void {
    const statusEl = document.getElementById('link-status');
    if (!statusEl || !internetMessageId) return;
    statusEl.style.display = 'block';
    let zone = document.getElementById('link-classer');
    if (!zone) {
      zone = document.createElement('div');
      zone.id = 'link-classer';
      zone.style.marginTop = '10px';
      statusEl.appendChild(zone);
    }
    renderClasserOutlook(zone, { messageId: internetMessageId, mailbox: '', onInfo: showToast, ...(projet ? { projetId: projet.id, projetLibelle: projet.libelle } : {}) });
  }

  // ── Office.js helpers ──

  private getAsync<T>(item: Office.MessageRead, prop: string): Promise<T> {
    return new Promise((resolve) => {
      // In Office.js, some properties are direct, some need getAsync
      const val = (item as any)[prop];
      resolve(val as T);
    });
  }

  private getRecipientsAsync(item: Office.MessageRead): Promise<string> {
    return new Promise((resolve) => {
      try {
        const to = (item as any).to;
        if (Array.isArray(to)) {
          resolve(to.map((r: any) => r.displayName || r.emailAddress || '').join(', '));
        } else {
          resolve('');
        }
      } catch { resolve(''); }
    });
  }

  private escapeHtml(str: string | undefined | null): string {
    return escapeHtml(str);
  }

  destroy(): void {
    this.searchPicker?.destroy();
    this.container.innerHTML = '';
  }
}
