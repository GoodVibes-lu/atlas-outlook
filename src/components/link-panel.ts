/**
 * Link Panel — Link current email to a projet/tiers/contact
 */

import { getLinkedConversationIds, getAllLinkedEmailIds, linkEmailToProject, linkEmailToContact, getAllProjets, resolveClientIdInProjetsBase, getFolderMapping, saveFolderMapping } from '../api/airtable';
import { getGraphToken, getMessageForLinking, convertToRestId, moveMessageToFolder, ensureFolderPath, resolveFolderPath } from '../api/graph';
import { summarizeEmail } from '../api/argo';
import { SearchPicker } from './search-picker';
import type { SearchResult, MailMessageFull, Projet } from '../types';
import { showToast } from '../taskpane';

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
      <div class="panel-scroll">
        <div class="section-heading">Email courant</div>
        <div id="email-info-card" class="email-card">
          <div class="spinner" style="margin: 12px auto;"></div>
        </div>

        <div id="ai-summary" style="display:none;">
          <div style="display:flex; align-items:center; gap:6px; margin-bottom:6px;">
            <span style="font-size:12px;">✨</span>
            <span class="section-heading" style="margin:0; font-size:10px;">RESUME IA</span>
          </div>
          <div id="ai-summary-content" style="font-size:12px; color:var(--atlas-text-secondary); line-height:1.5; padding:10px 12px; background:linear-gradient(135deg, rgba(99,102,241,0.05), rgba(139,92,246,0.05)); border:1px solid rgba(99,102,241,0.15); border-radius:var(--atlas-radius);"></div>
        </div>

        <div id="link-status" style="display:none;"></div>

        <div id="auto-suggestions" style="display:none;">
          <div id="auto-suggestion-content"></div>
        </div>

        <div id="link-search-section">
          <div class="toggle-row">
            <div class="toggle-switch" id="prive-toggle"></div>
            <span>🔒 Marquer comme privé</span>
          </div>
          <div id="search-toggle-section" style="display:none;">
            <a href="#" id="search-toggle-link" style="font-size:12px; color:var(--atlas-primary); text-decoration:none; display:inline-block; margin-bottom:8px;">Rechercher un autre projet, tiers ou contact ▼</a>
            <div id="search-container" style="display:none;"></div>
          </div>
          <div id="search-direct-container"></div>
        </div>
      </div>
    `;

    // Toggle privé
    document.getElementById('prive-toggle')?.addEventListener('click', (e) => {
      const el = e.currentTarget as HTMLElement;
      this.isPrive = !this.isPrive;
      el.classList.toggle('active', this.isPrive);
    });

    // Search toggle link
    document.getElementById('search-toggle-link')?.addEventListener('click', (e) => {
      e.preventDefault();
      this.searchExpanded = !this.searchExpanded;
      const link = e.currentTarget as HTMLElement;
      const searchContainer = document.getElementById('search-container')!;
      if (this.searchExpanded) {
        link.textContent = 'Rechercher un autre projet, tiers ou contact ▲';
        searchContainer.style.display = 'block';
      } else {
        link.textContent = 'Rechercher un autre projet, tiers ou contact ▼';
        searchContainer.style.display = 'none';
      }
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
        card.innerHTML = '<p class="empty-state">Aucun email sélectionné</p>';
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
      const { graphIds } = await getAllLinkedEmailIds();
      const restId = convertToRestId(itemId);
      if (graphIds.has(restId)) {
        this.emailInfo.isAlreadyLinked = true;
      }

      // Render email info
      card.innerHTML = `
        <div class="email-card-subject">${this.escapeHtml(this.emailInfo.subject)}</div>
        <div class="email-card-meta">
          De: ${this.escapeHtml(this.emailInfo.from)} &lt;${this.escapeHtml(this.emailInfo.fromEmail)}&gt;<br/>
          À: ${this.escapeHtml(this.emailInfo.to)}
        </div>
        ${this.emailInfo.isAlreadyLinked ? '<div class="status-linked" style="margin-top:8px;">✓ Déjà lié dans ATLAS</div>' : ''}
      `;

      // Auto-detect project from subject (#NNN)
      await this.autoDetect();

      // Generate AI summary (non-blocking)
      this.loadAiSummary();

    } catch (err) {
      card.innerHTML = `<p class="empty-state">Erreur : ${(err as Error).message}</p>`;
    }
  }

  private async loadAiSummary(): Promise<void> {
    if (!this.emailInfo || !localStorage.getItem('atlas_addin_anthropic_key')) return;
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

        const isLinked = this.emailInfo.isAlreadyLinked;

        autoContent.innerHTML = `
          <div class="detected-project-card" style="
            border-left: 4px solid var(--atlas-primary);
            background: var(--atlas-bg-secondary);
            border-radius: var(--atlas-radius);
            padding: 14px 16px;
            margin-bottom: 12px;
          ">
            <div style="font-size:11px; color:var(--atlas-text-secondary); text-transform:uppercase; letter-spacing:0.5px; margin-bottom:4px;">Projet détecté</div>
            <div style="font-size:16px; font-weight:700; color:var(--atlas-text-primary); margin-bottom:2px;">
              #${this.escapeHtml(found.noProjet)} ${this.escapeHtml(found.denomination)}
            </div>
            <div style="font-size:12px; color:var(--atlas-text-secondary); margin-bottom:4px;">
              ${this.escapeHtml(found.client || '')}
            </div>
            ${found.enCharge ? `<div style="font-size:11px; color:var(--atlas-text-secondary);">En charge : ${this.escapeHtml(found.enCharge)}</div>` : ''}
            <div style="margin-top:12px;">
              ${isLinked
                ? `<div style="
                    display:inline-flex; align-items:center; gap:6px;
                    background:var(--atlas-success-bg, rgba(40,167,69,0.1));
                    color:var(--atlas-success, #28a745);
                    border:1px solid var(--atlas-success, #28a745);
                    border-radius:var(--atlas-radius);
                    padding:6px 12px;
                    font-size:13px; font-weight:600;
                  ">✓ Email déjà lié à ce projet</div>`
                : `<button class="btn btn-primary" id="auto-link-btn" style="width:100%; padding:10px; font-size:14px; font-weight:600;">
                    Lier cet email au projet #${this.escapeHtml(noProjet)}
                  </button>`
              }
            </div>
          </div>
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
          <div class="detected-project-card" style="
            border-left: 4px solid var(--atlas-primary);
            background: var(--atlas-bg-secondary);
            border-radius: var(--atlas-radius);
            padding: 14px 16px;
            margin-bottom: 12px;
          ">
            <div style="font-size:11px; color:var(--atlas-text-secondary); text-transform:uppercase; letter-spacing:0.5px; margin-bottom:4px;">Conversation liée</div>
            <div style="font-size:16px; font-weight:700; color:var(--atlas-text-primary); margin-bottom:2px;">
              ${this.escapeHtml(match.projetName)}
            </div>
            <div style="font-size:12px; color:var(--atlas-text-secondary); margin-bottom:4px;">Même fil de conversation</div>
            <div style="margin-top:12px;">
              <button class="btn btn-primary" id="auto-conv-btn" style="width:100%; padding:10px; font-size:14px; font-weight:600;">
                Lier cet email au projet
              </button>
            </div>
          </div>
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
    statusEl.innerHTML = '<div class="spinner" style="margin:8px auto;"></div><p style="text-align:center;font-size:12px;">Liaison en cours...</p>';

    try {
      const token = await getGraphToken();
      const restId = convertToRestId(this.emailInfo.itemId);
      const fullMessage = await getMessageForLinking(token, restId);

      // Determine direction
      const userEmail = localStorage.getItem('atlas_addin_user_email') || '';
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

      statusEl.innerHTML = `<div class="status-linked">✓ Email lié à : ${this.escapeHtml(result.label)}</div>`;
      showToast(`Email lié à ${result.label}`, 'success');

      // Hide search section
      document.getElementById('link-search-section')!.style.display = 'none';
      document.getElementById('auto-suggestions')!.style.display = 'none';

      // Offer to file in Outlook folder if mapping exists
      await this.offerFolderFiling(result, restId, token);

    } catch (err) {
      statusEl.innerHTML = `<p style="color:var(--atlas-danger);font-size:12px;">Erreur : ${(err as Error).message}</p>`;
      showToast('Erreur de liaison', 'error');
    }
  }

  // ── Folder Filing ──

  private async offerFolderFiling(result: SearchResult, messageId: string, token: string): Promise<void> {
    const userEmail = localStorage.getItem('atlas_addin_user_email') || '';
    if (!userEmail) return;

    const statusEl = document.getElementById('link-status')!;

    try {
      // Check if user has a folder mapping for this entity
      const mapping = await getFolderMapping(userEmail, result.type === 'projet' ? 'projet' : 'client', result.id);

      if (mapping && mapping.folderId) {
        // User has an existing folder — offer to move
        statusEl.innerHTML += `
          <div style="margin-top:12px; padding:10px; background:var(--atlas-bg-secondary); border:1px solid var(--atlas-border); border-radius:var(--atlas-radius);">
            <p style="font-size:12px; margin-bottom:8px;">📂 Dossier Outlook trouvé : <strong>${this.escapeHtml(mapping.folderPath)}</strong></p>
            <div style="display:flex; gap:6px;">
              <button class="btn btn-primary btn-sm" id="move-to-folder-btn">Déplacer dans le dossier</button>
              <button class="btn btn-secondary btn-sm" id="skip-folder-btn">Ignorer</button>
            </div>
          </div>
        `;

        document.getElementById('move-to-folder-btn')?.addEventListener('click', async () => {
          try {
            // Verify folder still exists
            let folderId = mapping.folderId;
            const resolved = await resolveFolderPath(token, mapping.folderPath);
            if (resolved) {
              folderId = resolved;
            } else {
              // Folder was deleted — recreate it
              folderId = await ensureFolderPath(token, mapping.folderPath);
              await saveFolderMapping(userEmail, mapping.scope, result.id, mapping.folderPath, folderId);
            }

            await moveMessageToFolder(token, messageId, folderId);
            showToast(`Email déplacé dans ${mapping.folderPath}`, 'success');
            document.getElementById('move-to-folder-btn')!.closest('div')!.parentElement!.innerHTML =
              '<p style="font-size:12px;color:var(--atlas-success);">✓ Email déplacé dans le dossier</p>';
          } catch (err) {
            showToast(`Erreur déplacement : ${(err as Error).message}`, 'error');
          }
        });

        document.getElementById('skip-folder-btn')?.addEventListener('click', () => {
          document.getElementById('move-to-folder-btn')!.closest('div')!.parentElement!.remove();
        });
      } else {
        // No mapping — offer to create a folder
        const suggestedPath = result.type === 'projet'
          ? `Clients/${result.detail || 'Client'}/${result.label}`
          : `Clients/${result.label}`;

        statusEl.innerHTML += `
          <div style="margin-top:12px; padding:10px; background:var(--atlas-bg-secondary); border:1px solid var(--atlas-border); border-radius:var(--atlas-radius);">
            <p style="font-size:12px; margin-bottom:8px;">📂 Créer un dossier Outlook et y déplacer l'email ?</p>
            <div class="form-group" style="margin-bottom:8px;">
              <input type="text" class="form-input" id="folder-path-input" value="${this.escapeAttr(suggestedPath)}" style="font-size:12px;" />
            </div>
            <div style="display:flex; gap:6px;">
              <button class="btn btn-primary btn-sm" id="create-folder-btn">Créer et déplacer</button>
              <button class="btn btn-secondary btn-sm" id="skip-create-btn">Ignorer</button>
            </div>
          </div>
        `;

        document.getElementById('create-folder-btn')?.addEventListener('click', async () => {
          const pathInput = document.getElementById('folder-path-input') as HTMLInputElement;
          const folderPath = pathInput.value.trim();
          if (!folderPath) return;

          try {
            const btn = document.getElementById('create-folder-btn') as HTMLButtonElement;
            btn.disabled = true;
            btn.textContent = 'Création...';

            const folderId = await ensureFolderPath(token, folderPath);
            await moveMessageToFolder(token, messageId, folderId);
            await saveFolderMapping(userEmail, result.type === 'projet' ? 'projet' : 'client', result.id, folderPath, folderId);

            showToast(`Dossier créé et email déplacé`, 'success');
            btn.closest('div')!.parentElement!.innerHTML =
              `<p style="font-size:12px;color:var(--atlas-success);">✓ Email déplacé dans <strong>${this.escapeHtml(folderPath)}</strong></p>`;
          } catch (err) {
            showToast(`Erreur : ${(err as Error).message}`, 'error');
          }
        });

        document.getElementById('skip-create-btn')?.addEventListener('click', () => {
          document.getElementById('create-folder-btn')!.closest('div')!.parentElement!.remove();
        });
      }
    } catch {
      // Non-blocking — folder filing is optional
    }
  }

  private escapeAttr(str: string): string {
    return str.replace(/"/g, '&quot;').replace(/'/g, '&#39;');
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
    if (!str) return '';
    return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  destroy(): void {
    this.searchPicker?.destroy();
    this.container.innerHTML = '';
  }
}
