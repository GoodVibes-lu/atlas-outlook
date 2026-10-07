/**
 * Link Panel — Link current email to a projet/tiers/contact
 */

import { getLinkedConversationIds, getAllLinkedEmailIds, linkEmailToProject, linkEmailToContact, getAllProjets, resolveClientIdInProjetsBase, getFolderMapping, saveFolderMapping } from '../api/airtable';
import { lireMailPourLiaison } from '../api/mail-liaison';
import { fetchDossierProjet, rangerDansDossier } from '../api/agent';
import { convertToRestId, moveMessageToFolder, ensureFolderPath, resolveFolderPath, listAllMailFolders } from '../api/graph';
import { dossiersDeRangement, trouverDossierProjet, cheminDossierProjetPropose } from '../shared/mail-folder-tree';
import { summarizeEmail } from '../api/argo';
import { SearchPicker } from './search-picker';
import type { SearchResult, MailMessageFull, Projet } from '../types';
import { showToast } from '../taskpane';
import { escapeHtml } from '../utils/html';
import { loadingHtml, inlineLoadingHtml, emptyHtml, errorHtml, renderError } from '../ui/states';
import { humanError } from '../api/net';
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
      const { graphIds } = await getAllLinkedEmailIds();
      const restId = convertToRestId(itemId);
      if (graphIds.has(restId)) {
        this.emailInfo.isAlreadyLinked = true;
      }

      // Render email info
      card.innerHTML = `
        <div class="email-card-subject">${this.escapeHtml(this.emailInfo.subject)}</div>
        <dl class="facts">
          <dt>De</dt><dd>${this.escapeHtml(this.emailInfo.from || this.emailInfo.fromEmail)}${this.emailInfo.from ? `<br/><span class="meta">${this.escapeHtml(this.emailInfo.fromEmail)}</span>` : ''}</dd>
          <dt>À</dt><dd>${this.escapeHtml(this.emailInfo.to)}</dd>
        </dl>
        ${this.emailInfo.isAlreadyLinked ? `<p class="status-linked" style="margin-top:10px;">${icon('check-circle', 14)}Déjà lié dans ATLAS</p>` : ''}
      `;

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

        const isLinked = this.emailInfo.isAlreadyLinked;

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
      const { token, restId } = lu;
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
      if (lu.viaAtlas) await this.offerFolderFilingViaAtlas(result, fullMessage.internetMessageId || this.emailInfo.internetMessageId || '');
      else await this.offerFolderFiling(result, restId, token);

    } catch (err) {
      statusEl.innerHTML = errorHtml(err, { title: 'Liaison impossible', retry: false, compact: true });
      showToast('Erreur de liaison', 'error');
    }
  }

  // ── Folder Filing ──

  /**
   * Rangement par le worker (07/10/2026) : quand Outlook ne donne pas de jeton de boîte au complément,
   * le dossier du projet est cherché et le mail rangé par ATLAS (mêmes routes que le tableau de bord :
   * dossier trouvé n'importe où dans l'arbre, appris, création des niveaux manquants, annulable).
   */
  private async offerFolderFilingViaAtlas(result: SearchResult, internetMessageId: string): Promise<void> {
    if (!internetMessageId) return;
    const statusEl = document.getElementById('link-status')!;
    try {
      const projetId = result.type === 'projet' ? result.id : '';
      const d = projetId ? await fetchDossierProjet(projetId) : { existant: null, propose: `Clients/${result.label}`, verrou: false };
      if (d.existant) {
        statusEl.innerHTML += `
          <div class="card folder-card is-known" style="margin-top:10px;">
            <div class="folder-card-icon">${icon('folder', 18)}</div>
            <div class="folder-card-body"><span class="eyebrow">Dossier du projet</span><p class="folder-card-title">${this.escapeHtml(d.existant.chemin)}</p>
            <div class="btn-row">
              <button class="btn btn-primary btn-sm" id="move-atlas-btn">Déplacer dans le dossier</button>
            </div></div>
          </div>`;
        document.getElementById('move-atlas-btn')?.addEventListener('click', async () => {
          try {
            await rangerDansDossier({ messageId: internetMessageId, dossierId: d.existant!.id, ...(projetId ? { projetId } : {}) });
            showToast(`Email déplacé dans ${d.existant!.chemin}`, 'success');
            document.getElementById('move-atlas-btn')!.closest('div')!.parentElement!.innerHTML =
              `<p class="status-linked">${icon('check-circle', 14)}Mail rangé dans le dossier</p>`;
          } catch (err) { showToast(`${humanError(err)}`, 'error'); }
        });
        return;
      }
      if (!d.propose) return;
      statusEl.innerHTML += `
        <div class="card stack-sm" style="margin-top:10px;">
          <label class="form-label" for="folder-path-atlas">Ranger dans un nouveau dossier Outlook ?</label>
          <input type="text" class="form-input" id="folder-path-atlas" value="${this.escapeAttr(d.propose)}" />
          <div class="btn-row">
            <button class="btn btn-primary btn-sm" id="create-atlas-btn">Créer et déplacer</button>
            <button class="btn btn-ghost btn-sm" id="skip-atlas-btn">Ignorer</button>
          </div>
        </div>`;
      document.getElementById('create-atlas-btn')?.addEventListener('click', async () => {
        const chemin = (document.getElementById('folder-path-atlas') as HTMLInputElement).value.trim();
        if (!chemin) return;
        const btn = document.getElementById('create-atlas-btn') as HTMLButtonElement;
        btn.disabled = true;
        btn.textContent = 'Création...';
        try {
          const r = await rangerDansDossier({ messageId: internetMessageId, creer: { chemin, ...(projetId ? { projetId } : {}) } });
          showToast('Dossier créé et email déplacé', 'success');
          btn.closest('div')!.parentElement!.innerHTML =
            `<p class="status-linked">${icon('check-circle', 14)}Mail rangé dans ${this.escapeHtml(r.dossier.chemin || chemin)}</p>`;
        } catch (err) {
          btn.disabled = false;
          btn.textContent = 'Créer et déplacer';
          showToast(`${humanError(err)}`, 'error');
        }
      });
      document.getElementById('skip-atlas-btn')?.addEventListener('click', () => {
        document.getElementById('create-atlas-btn')!.closest('div')!.parentElement!.remove();
      });
    } catch {
      // Non bloquant : le rangement reste facultatif
    }
  }

  private async offerFolderFiling(result: SearchResult, messageId: string, token: string): Promise<void> {
    const userEmail = localStorage.getItem('atlas_addin_user_email') || Office.context?.mailbox?.userProfile?.emailAddress || '';
    if (!userEmail) return;

    const statusEl = document.getElementById('link-status')!;

    try {
      // Check if user has a folder mapping for this entity
      const mapping = await getFolderMapping(userEmail, result.type === 'projet' ? 'projet' : 'client', result.id);

      if (mapping && mapping.folderId) {
        // User has an existing folder — offer to move
        statusEl.innerHTML += `
          <div class="card folder-card is-known" style="margin-top:10px;">
            <div class="folder-card-icon">${icon('folder', 18)}</div>
            <div class="folder-card-body"><span class="eyebrow">Dossier habituel</span><p class="folder-card-title">${this.escapeHtml(mapping.folderPath)}</p>
            <div class="btn-row">
              <button class="btn btn-primary btn-sm" id="move-to-folder-btn">Déplacer dans le dossier</button>
              <button class="btn btn-ghost btn-sm" id="skip-folder-btn">Ignorer</button>
            </div></div>
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
              `<p class="status-linked">${icon('check-circle', 14)}Mail rangé dans le dossier</p>`;
          } catch (err) {
            showToast(`${humanError(err)}`, 'error');
          }
        });

        document.getElementById('skip-folder-btn')?.addEventListener('click', () => {
          document.getElementById('move-to-folder-btn')!.closest('div')!.parentElement!.remove();
        });
      } else {
        // Pas de dossier appris (07/10/2026, parité Inbox ATLAS) : dossier du projet cherché dans TOUTE
        // l'arborescence (n° de projet d'abord, puis nom) ; sinon chemin proposé à la même place et avec
        // le même nom que l'Inbox ATLAS (« Clients/<Client>/#755 Nom »).
        let suggestedPath = `Clients/${result.label}`;
        if (result.type === 'projet') {
          const arbre = dossiersDeRangement(await listAllMailFolders(token));
          const p = (await getAllProjets().catch(() => [] as Projet[])).find(x => x.id === result.id);
          const projet = { numero: p?.noProjet, nom: p?.denomination || result.label, clientNom: p?.client || result.detail || '' };
          const trouve = trouverDossierProjet(arbre, projet);
          if (trouve) {
            statusEl.innerHTML += `
              <div class="card folder-card is-known" style="margin-top:10px;">
                <div class="folder-card-icon">${icon('folder', 18)}</div>
                <div class="folder-card-body"><span class="eyebrow">Dossier du projet</span><p class="folder-card-title">${this.escapeHtml(trouve.dossier.chemin)}</p>
                <div class="btn-row">
                  <button class="btn btn-primary btn-sm" id="move-found-btn">Déplacer dans le dossier</button>
                </div></div>
              </div>`;
            document.getElementById('move-found-btn')?.addEventListener('click', async () => {
              try {
                await moveMessageToFolder(token, messageId, trouve.dossier.id);
                await saveFolderMapping(userEmail, 'projet', result.id, trouve.dossier.chemin, trouve.dossier.id).catch(() => undefined);
                showToast(`Email déplacé dans ${trouve.dossier.chemin}`, 'success');
                document.getElementById('move-found-btn')!.closest('div')!.parentElement!.innerHTML =
                  `<p class="status-linked">${icon('check-circle', 14)}Mail rangé dans le dossier</p>`;
              } catch (err) { showToast(`${humanError(err)}`, 'error'); }
            });
            return;
          }
          suggestedPath = cheminDossierProjetPropose(projet, arbre);
        }

        statusEl.innerHTML += `
          <div class="card stack-sm" style="margin-top:10px;">
            <label class="form-label" for="folder-path-input">Ranger dans un nouveau dossier Outlook ?</label>
            <input type="text" class="form-input" id="folder-path-input" value="${this.escapeAttr(suggestedPath)}" />
            <div class="btn-row">
              <button class="btn btn-primary btn-sm" id="create-folder-btn">Créer et déplacer</button>
              <button class="btn btn-ghost btn-sm" id="skip-create-btn">Ignorer</button>
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
              `<p class="status-linked">${icon('check-circle', 14)}Mail rangé dans ${this.escapeHtml(folderPath)}</p>`;
          } catch (err) {
            showToast(`${humanError(err)}`, 'error');
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
    return escapeHtml(str);
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
