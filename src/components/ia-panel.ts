/**
 * ia-panel.ts — Panneau "IA" du task-pane ATLAS dans Outlook.
 *
 * Reflète les actions IA disponibles dans l'app desktop ATLAS :
 *   • Affiche la catégorie + l'urgence détectées par l'inbound scanner
 *   • Permet de : marquer Traité / Reporter / Archiver / Corriger la catégorie
 *
 * Le tagging IA lui-même est fait côté worker/app (inbound-scanner.service.ts) —
 * ici on consomme le résultat depuis Airtable et on déclenche les actions.
 */

import { showToast } from '../taskpane';
import { outlookFetch, humanError, noteDiag } from '../api/net';
import { escapeHtml } from '../utils/html';
import { icon } from '../ui/icons';
import { loadingHtml, inlineLoadingHtml, emptyHtml, errorHtml, renderError } from '../ui/states';
import {
  getEmailTagByEmailId,
  getEmailTagByConversationId,
  getFolderMapping,
  saveFolderMapping,
  getAllProjets,
  markTagDone,
  snoozeTag,
  archiveTag,
  correctTagCategory,
  upsertEmailTag,
  type EmailTag,
} from '../api/airtable';
import {
  analyzeEmailWithClaude,
  isAiAvailable,
} from '../api/claude';
import { lookupSenderFolder, recordSenderFolder } from '../api/sender-folder-index';
import { getMessageForLinking } from '../api/graph';
import type { Projet } from '../types';
import {
  getGraphToken,
  moveMessageToFolder,
  convertToRestId,
  listMailFolders,
  ensureFolderPath,
  setMessageCategories,
  clearAtlasCategories,
  ATLAS_CATEGORIES,
  ATLAS_IA_CATEGORIES,
} from '../api/graph';

const CATEGORIES = [
  'demande_devis', 'validation_client', 'refus_client', 'question_staff',
  'facture_fournisseur', 'prospection_entrante', 'prospection_sortante', 'rdv_planning',
  'newsletter', 'notification_systeme', 'spam', 'autre',
  'federation_association', 'demande_interne_staff', 'fournisseur',
];

const CATEGORY_LABELS: Record<string, string> = {
  demande_devis: 'Demande de devis',
  validation_client: 'Validation client',
  refus_client: 'Refus client',
  question_staff: 'Question staff',
  facture_fournisseur: 'Facture fournisseur',
  prospection_entrante: 'Prospection entrante',
  prospection_sortante: 'Prospection sortante',
  rdv_planning: 'Rendez-vous, planning',
  newsletter: 'Newsletter',
  notification_systeme: 'Notification automatique',
  spam: 'Indésirable',
  autre: 'Autre',
  federation_association: 'Fédération, association',
  demande_interne_staff: 'Demande interne',
  fournisseur: 'Fournisseur',
};

/** Icône (forme, jamais couleur) de chaque thème. */
const CATEGORY_ICONS: Record<string, string> = {
  demande_devis: 'template', validation_client: 'check-circle', refus_client: 'x', question_staff: 'team',
  facture_fournisseur: 'paperclip', prospection_entrante: 'inbox', prospection_sortante: 'external',
  rdv_planning: 'calendar-plus', newsletter: 'list', notification_systeme: 'activity', spam: 'alert',
  autre: 'mail', federation_association: 'building', demande_interne_staff: 'user', fournisseur: 'archive',
};

/**
 * Mapping catégorie IA → mots-clés qu'on cherche dans les noms de dossiers
 * Outlook pour suggérer un dossier de classement quand on n'a aucun autre signal.
 * Le matching est case-insensitive et substring (cf. findFolderForCategory).
 */
const CATEGORY_FOLDER_ALIASES: Record<string, string[]> = {
  demande_devis: ['devis', 'demande', 'offre'],
  validation_client: ['client', 'validation'],
  refus_client: ['refus', 'client'],
  question_staff: ['staff', 'équipe', 'equipe', 'interne'],
  facture_fournisseur: ['facture', 'comptabilité', 'compta', 'fournisseur'],
  prospection_entrante: ['prospect', 'prospection', 'lead', 'entrant'],
  prospection_sortante: ['prospection', 'sortant', 'outbound', 'commercial'],
  rdv_planning: ['rdv', 'rendez', 'planning', 'agenda'],
  newsletter: ['newsletter', 'newsletters', 'news'],
  notification_systeme: ['notification', 'système', 'systeme', 'auto'],
  spam: ['spam', 'junk', 'courrier indésirable'],
  autre: [],
  federation_association: ['fédération', 'federation', 'fédérations', 'federations', 'association', 'associations'],
  demande_interne_staff: ['interne', 'staff', 'équipe', 'equipe'],
  fournisseur: ['fournisseur', 'fournisseurs', 'suppliers'],
};

const URGENCY_LABELS: Record<number, string> = {
  1: 'Peut attendre', 2: 'Faible', 3: 'À traiter', 4: 'Urgent', 5: 'Très urgent',
};

/** Jauge d'urgence : 5 segments, le rouge seulement à partir de « Urgent » (statut réel). */
function urgencyHtml(score: number): string {
  const n = Math.max(0, Math.min(5, Math.round(Number(score) || 0)));
  const label = URGENCY_LABELS[n] || 'Non évaluée';
  return `<span class="urgency${n >= 4 ? ' is-high' : ''}" role="img" aria-label="Urgence ${n} sur 5 : ${label}">
    <span class="urgency-bars" aria-hidden="true">${[1, 2, 3, 4, 5].map(i => `<i class="${i <= n ? 'on' : ''}"></i>`).join('')}</span>
    <span class="urgency-label">${label}</span>
  </span>`;
}

const STATUS_LABELS: Record<string, string> = {
  done: 'Traité', snoozed: 'Reporté', archived: 'Archivé', inbox: 'À traiter',
};

// Suggestion de dossier de classement résolue dynamiquement.
interface FolderSuggestion {
  /**
   * Origine de la suggestion :
   *   - mapped : folder mapping Airtable appris (explicite)
   *   - sender-pattern : pattern observé sur les mails de ce sender dans Outlook
   *     (continue l'existant — Charles range déjà ces mails dans ce dossier)
   *   - match : top match fuzzy sur les noms de dossiers existants
   *   - create : aucun match → on propose un nouveau dossier à créer
   *   - none : aucun signal exploitable
   */
  source: 'index-sender' | 'mapped' | 'sender-pattern' | 'category-match' | 'match' | 'create' | 'manual-override' | 'none';
  /** ID Outlook si dossier existant. */
  folderId?: string;
  /** Chemin lisible (ex: "Markcom/Creativity Camp"). */
  folderPath: string;
  /** Projet lié (utilisé pour la sauvegarde du mapping). */
  projetId?: string;
  /** Pour 'sender-pattern' : nombre de mails du sender dans ce dossier (preuve). */
  patternMailCount?: number;
  /** Diagnostic visible UI quand rien ne match (debug sans DevTools). */
  debug?: string;
}

export class IAPanel {
  private root: HTMLElement;
  private tag: EmailTag | null = null;
  private emailId = '';
  private linkedProjet: Projet | null = null;
  private folderSuggestion: FolderSuggestion | null = null;
  // Picker manuel : état + cache de tous les dossiers Outlook (chargé lazy)
  private folderPickerOpen = false;
  private allFoldersCache: Array<{ id: string; path: string }> | null = null;

  constructor(root: HTMLElement) {
    this.root = root;
    this.renderLoading();
    this.load();
  }

  destroy(): void { /* no async work to cancel */ }

  private renderLoading(): void {
    this.root.innerHTML = `<div class="panel-scroll">${loadingHtml('Lecture du classement ARGO…', 4)}</div>`;
  }

  private async load(): Promise<void> {
    try {
      const item = Office.context.mailbox?.item;
      if (!item) {
        this.renderEmpty('Aucun mail sélectionné.');
        return;
      }
      // Outlook fournit itemId au format EWS. On utilise convertToRestId pour avoir l'ID Graph.
      const ewsId = (item as any).itemId;
      const token = (item as any).restUrl ? '' : '';
      // Note : ici on n'a pas besoin du token Graph — l'inbound-scanner stocke
      // l'ID Graph dans Airtable (champ EmailId). On peut donc retrouver le tag
      // soit via Graph-restId (si déjà converti côté app), soit via le format
      // EWS si stocké comme tel. Pour V1 on essaie les deux variantes.
      this.emailId = ewsId || '';

      const tag = await this.tryFetchTag(this.emailId);
      this.tag = tag;
      if (tag) {
        // Résolution du dossier de classement (en parallèle du rendu initial)
        this.render();
        this.resolveFolderSuggestion().then(() => this.render()).catch(() => {});
      } else {
        // Plus d'auto-analyse à l'ouverture (phase 1.5, 03/10/2026) : l'agent serveur lit chaque
        // mail une seule fois (onglet Agent) ; ici l'analyse IA reste une action volontaire
        // (bouton « Analyser maintenant » / « Re-analyser »).
        this.renderNotTagged();
      }
    } catch (e) {
      console.warn('[IAPanel] load failed:', e);
      this.root.innerHTML = '<div class="panel-scroll"><div id="ia-err"></div></div>';
      renderError(this.root.querySelector('#ia-err')!, e, () => { this.renderLoading(); this.load(); }, { title: 'Classement indisponible' });
    }
  }

  /**
   * Affiche pendant l'auto-analyse au chargement (entre "pas de tag trouvé"
   * et "résultat de Claude"). Évite que Charles voie le toast "Pas encore
   * taggé" puis doit cliquer.
   */
  private renderAutoAnalyzing(): void {
    this.root.innerHTML = `<div class="panel-scroll">
      <div class="argo-progress" role="status"><span class="argo-progress-bar" aria-hidden="true"></span><span>ARGO classe ce mail, quelques secondes…</span></div>
    </div>`;
  }

  /**
   * Essaie de récupérer le tag pour ce message via plusieurs stratégies :
   *   1. EWS itemId direct
   *   2. EWS → REST conversion (Graph format)
   *   3. **Fallback conversationId** — le scanner backend stocke le Graph REST
   *      ID dans `EmailId`, qui peut ne pas matcher l'EWS ID de l'addin selon
   *      le client (Outlook web vs desktop). Mais le `conversationId` est
   *      stable cross-client : on lookup le tag du thread.
   */
  private async tryFetchTag(rawId: string): Promise<EmailTag | null> {
    if (!rawId) return null;
    // 1. Direct (cas où l'ID est déjà au format REST)
    let tag = await getEmailTagByEmailId(rawId);
    if (tag) return tag;
    // 2. EWS → REST conversion
    try {
      const restId = Office.context.mailbox?.convertToRestId?.(
        rawId,
        // @ts-ignore — la constante est exposée même si non typée
        Office.MailboxEnums?.RestVersion?.v2_0 ?? 'v2.0',
      );
      if (restId && restId !== rawId) {
        tag = await getEmailTagByEmailId(restId);
        if (tag) return tag;
      }
    } catch { /* ignore */ }
    // 3. Fallback conversationId — le plus fiable cross-client (Outlook web,
    //    desktop, mobile). Le scanner stocke le conversationId pour CHAQUE mail,
    //    donc on retombe sur le tag du même thread (catégorie identique 99% du temps).
    try {
      const convId = (Office.context.mailbox?.item as any)?.conversationId;
      if (convId) {
        tag = await getEmailTagByConversationId(convId);
        if (tag) return tag;
      }
    } catch { /* ignore */ }
    return null;
  }

  private renderEmpty(msg: string): void {
    this.root.innerHTML = `<div class="panel-scroll">${emptyHtml({ icon: 'mail', title: msg, text: 'Sélectionne un mail dans ta boîte pour voir son classement.' })}</div>`;
  }

  private renderNotTagged(): void {
    const hasKey = isAiAvailable();
    this.root.innerHTML = `
      <div class="panel-scroll stack">
        ${emptyHtml({ icon: 'tag', title: 'Pas encore classé', text: 'ARGO n\'a pas encore lu ce mail. Lance le classement : thème, urgence et dossier conseillé en quelques secondes.' })}
        <button type="button" data-action="analyze-now" class="btn btn-argo btn-block" ${hasKey ? '' : 'disabled'}>
          ${icon('bolt', 14)}<span>Classer avec ARGO</span><span class="argo-tag" aria-hidden="true">IA</span>
        </button>
        <div id="ia-progress" hidden></div>
        ${hasKey ? '' : '<p class="help">ARGO est indisponible pour le moment.</p>'}
      </div>
    `;
    const progress = this.root.querySelector<HTMLElement>('#ia-progress');
    this.root.querySelector<HTMLButtonElement>('button[data-action="analyze-now"]')
      ?.addEventListener('click', async (ev) => {
        const btn = ev.currentTarget as HTMLButtonElement;
        btn.disabled = true;
        if (progress) { progress.hidden = false; progress.innerHTML = '<div class="argo-progress" role="status"><span class="argo-progress-bar" aria-hidden="true"></span><span>ARGO lit le mail…</span></div>'; }
        try {
          const ok = await this.reanalyzeNow();
          if (ok) this.render(); // bascule sur la vue classée
          else { btn.disabled = false; if (progress) progress.hidden = true; }
        } catch {
          btn.disabled = false;
          if (progress) progress.hidden = true;
        }
      });
  }

  private render(): void {
    const tag = this.tag!;
    const catLabel = CATEGORY_LABELS[tag.category] || tag.category;
    const status = tag.inboxStatus;

    this.root.innerHTML = `
      <div class="panel-scroll stack">
        <section class="section" aria-labelledby="ia-h-class">
          <h2 class="section-heading" id="ia-h-class">Classement ARGO</h2>
          <div class="card ia-class">
            <div class="ia-class-row">
              <span class="ia-cat">${icon(CATEGORY_ICONS[tag.category] || 'tag', 16)}<span>${escapeHtml(catLabel)}</span></span>
              ${status && status !== 'inbox' ? `<span class="badge badge-neutral">${escapeHtml(STATUS_LABELS[status] || status)}</span>` : ''}
            </div>
            ${urgencyHtml(tag.urgencyScore)}
            ${tag.summary ? `<p class="card-note">${escapeHtml(tag.summary)}</p>` : ''}
          </div>
        </section>

        <section class="section" aria-labelledby="ia-h-folder">
          <h2 class="section-heading" id="ia-h-folder">Où ranger ce mail</h2>
          ${this.renderFolderSection()}
        </section>

        <section class="section" aria-labelledby="ia-h-actions">
          <h2 class="section-heading" id="ia-h-actions">Que faire</h2>
          <div class="action-grid">
            <button type="button" data-action="done" class="btn btn-primary">${icon('check', 14)}Traité</button>
            <button type="button" data-action="snooze" class="btn btn-secondary">${icon('clock', 14)}Demain 8 h</button>
            <button type="button" data-action="archive" class="btn btn-secondary">${icon('archive', 14)}Archiver</button>
            <button type="button" data-action="broom" class="btn btn-secondary" title="Range maintenant les mails lus depuis plus de 10 minutes dans leurs dossiers habituels">${icon('broom', 14)}Ranger la boîte</button>
          </div>
          <p class="help">Traité et Archiver rangent aussi le mail dans le dossier ci-dessus.</p>
        </section>

        <details class="disclosure">
          <summary>${icon('chevron-right', 14)}Corriger le classement</summary>
          <div class="disclosure-body stack-sm">
            <label class="form-label" for="ia-cat-select">Le bon thème</label>
            <select id="ia-cat-select" class="form-input">
              ${CATEGORIES.map(cat => `<option value="${cat}" ${cat === tag.category ? 'selected' : ''}>${escapeHtml(CATEGORY_LABELS[cat] || cat)}</option>`).join('')}
            </select>
            <button type="button" data-action="correct" class="btn btn-secondary btn-sm">${icon('check', 14)}Enregistrer, ARGO apprend</button>
            <button type="button" data-action="reanalyze" class="btn btn-argo btn-sm" ${isAiAvailable() ? '' : 'disabled'}>${icon('bolt', 14)}<span>Reclasser avec ARGO</span><span class="argo-tag" aria-hidden="true">IA</span></button>
          </div>
        </details>
      </div>
    `;

    // Click handlers
    this.root.querySelectorAll<HTMLButtonElement>('button[data-action]').forEach(btn => {
      btn.addEventListener('click', () => this.onAction(btn.dataset.action!));
    });

    // Picker dossier : search live + click sur ligne pour sélectionner
    const searchInput = this.root.querySelector<HTMLInputElement>('#folder-picker-search');
    if (searchInput) {
      searchInput.addEventListener('input', (e) => {
        this.folderPickerQuery = (e.target as HTMLInputElement).value;
        // Re-render uniquement la section picker (préserve focus)
        const oldFocus = searchInput.selectionStart;
        this.render();
        const newInput = this.root.querySelector<HTMLInputElement>('#folder-picker-search');
        if (newInput) {
          newInput.focus();
          if (oldFocus !== null) newInput.setSelectionRange(oldFocus, oldFocus);
        }
      });
    }
    this.root.querySelectorAll<HTMLLIElement>('li[data-folder-id]').forEach((li) => {
      const pick = () => {
        this.folderPickerSelectedId = li.dataset.folderId || '';
        this.folderPickerSelectedPath = li.dataset.folderPath || '';
        this.render();
        this.root.querySelector<HTMLElement>('li.is-selected')?.focus();
      };
      li.addEventListener('click', pick);
      li.addEventListener('keydown', (ev) => { if (ev.key === 'Enter' || ev.key === ' ') { ev.preventDefault(); pick(); } });
    });
    // Champ création de dossier — bind le change pour mémoriser
    const newPathInput = this.root.querySelector<HTMLInputElement>('#folder-picker-new-path');
    if (newPathInput) {
      newPathInput.addEventListener('input', (e) => {
        this.folderPickerNewPath = (e.target as HTMLInputElement).value;
      });
      newPathInput.focus();
    }
  }

  private async onAction(action: string): Promise<void> {
    // Picker dossier : ne nécessite pas de tag (peut être déclenché avant que
    // la classification soit faite, juste pour préparer le mapping).
    if (action === 'broom') {
      showToast('Rangement de la boîte en cours…', 'info');
      try {
        const { forceAutoSweep } = await import('../api/auto-sweep');
        const r = await forceAutoSweep();
        if (r.archived > 0) {
          showToast(`${r.archived} mail${r.archived > 1 ? 's' : ''} rangé${r.archived > 1 ? 's' : ''} dans leurs dossiers habituels.`, 'success');
        } else {
          showToast(`Rien à ranger : ${r.scanned} mails vus, ${r.skipped.noFolder} sans dossier connu.`, 'info');
        }
      } catch (e) {
        showToast(`${humanError(e)}`, 'error');
      }
      return;
    }
    if (action === 'retry' && this.folderPickerOpen) {
      this.foldersLoadError = null;
      this.allFoldersCache = null;
      this.loadAllFolders().then(() => this.render()).catch(() => this.render());
      this.render();
      return;
    }
    if (action === 'change-folder') {
      this.folderPickerOpen = true;
      this.folderPickerQuery = '';
      this.folderPickerSelectedId = '';
      this.folderPickerSelectedPath = '';
      this.folderPickerCreateMode = false;
      this.folderPickerNewPath = '';
      // Charge la liste des dossiers en background
      this.loadAllFolders().then(() => this.render()).catch(() => this.render());
      this.render();
      return;
    }
    if (action === 'cancel-folder') {
      this.folderPickerOpen = false;
      this.folderPickerCreateMode = false;
      this.render();
      return;
    }
    if (action === 'confirm-folder') {
      const folderId = this.folderPickerSelectedId;
      const folderPath = this.folderPickerSelectedPath;
      if (!folderId || !folderPath) {
        showToast('Sélectionne un dossier dans la liste', 'error');
        return;
      }
      this.folderSuggestion = {
        source: 'manual-override',
        folderId,
        folderPath,
        projetId: this.tag?.linkedProjetId,
      };
      this.folderPickerOpen = false;
      showToast(`Dossier choisi : ${folderPath}. Clique sur Traité ou Archiver pour y ranger le mail.`, 'info');
      this.render();
      return;
    }
    if (action === 'open-create-folder') {
      this.folderPickerCreateMode = true;
      this.folderPickerNewPath = this.folderPickerQuery || ''; // reuse search query as default
      this.folderPickerSuggestingAI = true;
      this.render();
      // Lance la suggestion IA en background
      this.fetchAIFolderSuggestion().catch(() => {});
      return;
    }
    if (action === 'create-folder-cancel') {
      this.folderPickerCreateMode = false;
      this.render();
      return;
    }
    if (action === 'create-folder-confirm') {
      const path = this.folderPickerNewPath.trim();
      if (!path) {
        showToast('Indique le nom du dossier à créer', 'error');
        return;
      }
      try {
        showToast(`Création du dossier "${path}"…`, 'info');
        const { getApiContext } = await import('../api/graph');
        const ctx = await getApiContext();
        const folderId = await ensureFolderPath(ctx.token, path);
        // Ajoute au cache local
        if (!this.allFoldersCache) this.allFoldersCache = [];
        this.allFoldersCache.push({ id: folderId, path });
        this.folderSuggestion = {
          source: 'manual-override',
          folderId,
          folderPath: path,
          projetId: this.tag?.linkedProjetId,
        };
        this.folderPickerOpen = false;
        this.folderPickerCreateMode = false;
        showToast(`Dossier « ${path} » créé : clique sur Traité ou Archiver pour y ranger le mail.`, 'success');
        this.render();
      } catch (e) {
        showToast(`${humanError(e)}`, 'error');
      }
      return;
    }

    if (!this.tag) return;
    const buttons = this.root.querySelectorAll<HTMLButtonElement>('button');
    buttons.forEach(b => b.disabled = true);
    try {
      let ok = false;
      switch (action) {
        case 'done':
          ok = await markTagDone(this.tag.id);
          if (ok) {
            this.tag.inboxStatus = 'done';
            this.applyCategoryForState('done').catch(() => {});
            const r = await this.tryMoveToHabitualFolder('done');
            if (r.folderPath) {
              showToast(`Traité et rangé dans ${r.folderPath}`, 'success');
            } else if (r.error && r.error !== 'aucune suggestion') {
              showToast(`Traité, mais pas rangé : ${r.error}`, 'error');
            } else {
              showToast('Marqué comme traité', 'success');
            }
          }
          break;
        case 'snooze':
          ok = await snoozeTag(this.tag.id);
          if (ok) {
            this.tag.inboxStatus = 'snoozed';
            // Applique catégorie ⏰ ATLAS · Reporté — visible dans la liste inbox
            // sans avoir besoin d'ouvrir le task-pane
            this.applyCategoryForState('snoozed').catch(() => {});
            showToast('Reporté à demain 8 h', 'success');
          }
          break;
        case 'archive':
          ok = await archiveTag(this.tag.id);
          if (ok) {
            this.tag.inboxStatus = 'archived';
            this.applyCategoryForState('archived').catch(() => {});
            const r = await this.tryMoveToHabitualFolder('archive');
            if (r.folderPath) {
              showToast(`Archivé dans ${r.folderPath}`, 'success');
            } else if (r.error && r.error !== 'aucune suggestion') {
              showToast(`Archivé, mais pas rangé : ${r.error}`, 'error');
            } else {
              showToast('Archivé (aucun dossier habituel : range-le à la main)', 'success');
            }
          }
          break;
        case 'correct': {
          const sel = this.root.querySelector<HTMLSelectElement>('#ia-cat-select');
          const newCat = sel?.value || '';
          if (newCat && newCat !== this.tag.category) {
            ok = await correctTagCategory(this.tag.id, newCat);
            if (ok) { this.tag.category = newCat; showToast(`Catégorie corrigée : ${CATEGORY_LABELS[newCat] || newCat}`, 'success'); }
          } else {
            showToast('Aucune modification', 'info');
            ok = true;
          }
          break;
        }
        case 'reanalyze': {
          // reanalyzeNow gère son propre toast d'erreur, on skip le générique
          ok = await this.reanalyzeNow();
          if (!ok) {
            buttons.forEach(b => b.disabled = false);
            return; // évite le "Échec de l'action" générique qui écrase le vrai message
          }
          break;
        }
      }
      if (!ok) showToast('Échec de l\'action', 'error');
      this.render();
    } catch (e) {
      console.warn('[IAPanel] action failed:', e);
      showToast('Action impossible pour le moment : réessaie.', 'error');
    } finally {
      buttons.forEach(b => b.disabled = false);
    }
  }

  /**
   * Force une re-analyse Claude pour le mail courant. Purge l'ancien tag +
   * écrit le nouveau résultat. Résultat immédiat — pas d'attente du scan
   * périodique côté app desktop. L'IA est servie par le worker (Agency Brain),
   * aucune clé dans le complément ; `auto` = lancée à l'ouverture (tâche de fond).
   */
  private async reanalyzeNow(opts: { auto?: boolean } = {}): Promise<boolean> {
    if (!isAiAvailable()) {
      showToast('ARGO est indisponible pour le moment.', 'error');
      return false;
    }
    try {
      const item = Office.context.mailbox?.item as any;
      if (!item) { showToast('Aucun mail sélectionné', 'error'); return false; }
      const userEmail = Office.context.mailbox?.userProfile?.emailAddress || '';
      if (!userEmail) { showToast('Adresse de ta boîte introuvable : rouvre le panneau', 'error'); return false; }

      // 1. Récupère les infos du mail via Office.js (pas besoin de token Graph)
      const ewsId: string = item.itemId || '';
      const subject: string = item.subject || '';
      const from = {
        name: item.from?.displayName || '',
        email: item.from?.emailAddress || '',
      };
      const toRecipients = (item.to || []).map((r: any) => ({
        name: r.displayName || '',
        email: r.emailAddress || '',
      }));
      const ccRecipients = (item.cc || []).map((r: any) => ({
        name: r.displayName || '',
        email: r.emailAddress || '',
      }));
      const receivedAt = item.dateTimeCreated
        ? new Date(item.dateTimeCreated).toISOString()
        : new Date().toISOString();
      const conversationId: string = item.conversationId || '';

      // Convertit l'EWS itemId en Graph REST ID pour cohérence avec le scanner backend
      let restEmailId = ewsId;
      try {
        const { convertToRestId } = await import('../api/graph');
        restEmailId = convertToRestId(ewsId);
      } catch { /* fallback ewsId brut */ }

      // Récupère le body via Office.js (text plain — propre pour Claude)
      const body: string = await new Promise((resolve) => {
        try {
          item.body.getAsync(Office.CoercionType.Text, (res: any) => {
            resolve(res?.status === Office.AsyncResultStatus.Succeeded ? (res.value || '') : '');
          });
        } catch { resolve(''); }
      });

      showToast('ARGO lit le mail…', 'info');

      // 2. Appelle Claude
      const analysis = await analyzeEmailWithClaude({
        subject,
        from,
        toRecipients,
        ccRecipients,
        body,
        receivedAt,
        userEmail,
        auto: !!opts.auto,
      });

      // 3. Upsert le tag dans Airtable (DELETE ancien + CREATE nouveau)
      const upserted = await upsertEmailTag({
        oldTagId: this.tag?.id,
        emailId: restEmailId,
        conversationId,
        subject,
        fromEmail: from.email,
        fromName: from.name,
        receivedAt,
        category: analysis.category,
        urgencyScore: analysis.urgencyScore,
        summary: analysis.summary,
        detectedLanguage: analysis.detectedLanguage,
        userEmail,
      });

      // 4. Refresh local
      this.tag = {
        id: upserted.id,
        emailId: restEmailId,
        category: analysis.category,
        urgencyScore: analysis.urgencyScore,
        summary: analysis.summary,
        inboxStatus: 'inbox',
        linkedProjetId: this.tag?.linkedProjetId,
      };
      showToast(`Classé : ${CATEGORY_LABELS[analysis.category] || analysis.category}`, 'success');
      // Applique la catégorie urgence (visible dans la liste inbox)
      this.applyCategoryForState('tagged').catch(() => {});
      // Re-résout aussi le dossier de classement (la catégorie a peut-être changé)
      this.resolveFolderSuggestion().then(() => this.render()).catch(() => {});
      return true;
    } catch (e) {
      console.warn('[IAPanel] reanalyze failed:', e);
      showToast(`${humanError(e)}`, 'error');
      return false;
    }
  }

  /** Section « Où ranger ce mail » affichée dans le rendu du classement. */
  private renderFolderSection(): string {
    const s = this.folderSuggestion;
    if (!s) {
      return `<div class="card folder-card is-pending">${inlineLoadingHtml('Recherche du dossier habituel…')}</div>`;
    }
    const picker = this.folderPickerOpen ? this.renderFolderPicker() : '';
    if (s.source === 'none') {
      // Le détail technique va dans Réglages › Diagnostic, jamais à l'écran.
      if (s.debug && !this.folderDebugNoted) {
        this.folderDebugNoted = true;
        noteDiag({ at: Date.now(), service: 'outlook', label: 'dossier conseillé : aucun signal', status: 0, ms: 0, detail: s.debug.slice(0, 200) });
      }
      return `
        <div class="card folder-card is-empty">
          <div class="folder-card-icon">${icon('folder', 18)}</div>
          <div class="folder-card-body">
            <p class="folder-card-title">Aucun dossier connu pour cet expéditeur</p>
            <p class="help">Choisis-en un : ATLAS le retiendra pour les prochains mails.</p>
            ${this.folderPickerOpen ? '' : `<button type="button" data-action="change-folder" class="btn btn-secondary btn-sm">${icon('folder-move', 14)}Choisir un dossier</button>`}
          </div>
        </div>
        ${picker}
      `;
    }
    const n = s.patternMailCount || 0;
    const plural = n > 1 ? 's' : '';
    const META: Record<string, { label: string; why: string; tone: string }> = {
      'index-sender': { label: 'Ton habitude', why: `${n} mail${plural} de cet expéditeur déjà rangé${plural} ici.`, tone: 'is-known' },
      mapped: { label: 'Dossier du projet', why: 'Appris lors d\'un rangement précédent.', tone: 'is-known' },
      'manual-override': { label: 'Ton choix', why: 'Le mail ira ici et ATLAS retiendra ce choix.', tone: 'is-chosen' },
      'sender-pattern': { label: 'Ton habitude', why: `${n} mail${plural} de cet expéditeur déjà dans ce dossier.`, tone: 'is-known' },
      'category-match': { label: 'Suggestion', why: `Dossier proche du thème « ${CATEGORY_LABELS[this.tag?.category || ''] || 'du mail'} ».`, tone: 'is-suggested' },
      match: { label: 'Suggestion', why: 'Nom de dossier proche du client ou du projet.', tone: 'is-suggested' },
      create: { label: 'Nouveau dossier', why: 'Aucun dossier existant ne correspond : ATLAS le créera au rangement.', tone: 'is-suggested' },
    };
    const m = META[s.source] || META.match;
    return `
      <div class="card folder-card ${m.tone}">
        <div class="folder-card-icon">${icon(s.source === 'create' ? 'plus' : 'folder', 18)}</div>
        <div class="folder-card-body">
          <span class="eyebrow">${m.tone === 'is-known' ? icon('check', 12) : ''}${escapeHtml(m.label)}</span>
          <p class="folder-card-title" title="${escapeHtml(s.folderPath)}">${escapeHtml(s.folderPath)}</p>
          <p class="help">${escapeHtml(m.why)}</p>
        </div>
        ${this.folderPickerOpen ? '' : `<button type="button" data-action="change-folder" class="btn btn-ghost btn-sm folder-card-change">Changer</button>`}
      </div>
      ${picker}
    `;
  }

  // État du picker : query de recherche + valeur sélectionnée
  private folderPickerQuery = '';
  private folderPickerSelectedId = '';
  private folderPickerSelectedPath = '';
  private folderPickerCreateMode = false;
  private folderPickerNewPath = '';
  private folderPickerSuggestingAI = false; // ARGO propose un nom de dossier
  private folderDebugNoted = false;
  private foldersLoadError: unknown = null;

  /** Sélecteur de dossier (recherche + liste) ou création d'un dossier. */
  private renderFolderPicker(): string {
    const folders = this.allFoldersCache || [];
    if (this.foldersLoadError) {
      return `<div class="picker">${errorHtml(this.foldersLoadError, { title: 'Liste des dossiers indisponible', compact: true })}</div>`;
    }
    if (folders.length === 0) {
      return `<div class="picker">${inlineLoadingHtml('Lecture de tes dossiers Outlook (quelques secondes sur une grosse boîte)…')}</div>`;
    }

    const q = this.folderPickerQuery.toLowerCase().trim();
    const all = folders
      .slice()
      .sort((a, b) => a.path.localeCompare(b.path, 'fr'))
      .filter((f) => !q || f.path.toLowerCase().includes(q));
    const filtered = all.slice(0, 100); // 100 lignes au plus

    if (this.folderPickerCreateMode) {
      const busy = this.folderPickerSuggestingAI;
      return `
        <div class="picker">
          <p class="picker-title">${icon('plus', 14)}Nouveau dossier</p>
          ${busy
            ? '<div class="argo-progress" role="status"><span class="argo-progress-bar" aria-hidden="true"></span><span>ARGO propose un nom d\'après tes dossiers…</span></div>'
            : this.folderPickerNewPath ? `<p class="help">${icon('bolt', 12)}Proposé par ARGO d'après tes dossiers, modifiable.</p>` : ''}
          <label class="form-label" for="folder-picker-new-path">Nom du dossier</label>
          <input type="text" class="form-input" id="folder-picker-new-path" placeholder="Clients/Vossloh ou Vossloh"
            value="${escapeHtml(this.folderPickerNewPath)}" ${busy ? 'disabled' : ''} />
          <p class="help">Utilise « / » pour un sous-dossier (Clients/Vossloh/2026) : les niveaux manquants sont créés.</p>
          <div class="btn-row">
            <button type="button" data-action="create-folder-confirm" class="btn btn-primary btn-sm" ${busy ? 'disabled' : ''}>${icon('check', 14)}Créer et utiliser</button>
            <button type="button" data-action="create-folder-cancel" class="btn btn-ghost btn-sm">Annuler</button>
          </div>
        </div>
      `;
    }

    const list = filtered.length === 0
      ? `<li class="picker-empty">Aucun dossier ne contient « ${escapeHtml(this.folderPickerQuery)} »</li>`
      : filtered.map((f) => {
          const isSel = f.id === this.folderPickerSelectedId;
          const depth = f.path.split('/').length - 1;
          const name = f.path.split('/').pop() || f.path;
          const parent = depth ? f.path.slice(0, f.path.length - name.length - 1) : '';
          return `<li role="option" tabindex="0" aria-selected="${isSel}" data-folder-id="${escapeHtml(f.id)}" data-folder-path="${escapeHtml(f.path)}" class="folder-row${isSel ? ' is-selected' : ''}" title="${escapeHtml(f.path)}">
            ${icon(isSel ? 'check' : 'folder', 14)}<span class="folder-row-name">${escapeHtml(name)}</span>${parent ? `<span class="folder-row-parent">${escapeHtml(parent)}</span>` : ''}
          </li>`;
        }).join('');

    return `
      <div class="picker">
        <div class="picker-head">
          <p class="picker-title">${icon('folder-move', 14)}Choisir un dossier <span class="picker-count">${folders.length}</span></p>
          <button type="button" data-action="open-create-folder" class="btn btn-ghost btn-sm">${icon('plus', 14)}Nouveau</button>
        </div>
        <div class="search-wrapper">
          <span class="search-icon">${icon('search', 14)}</span>
          <input type="search" class="search-input" id="folder-picker-search" placeholder="Rechercher un dossier" aria-label="Rechercher un dossier"
            value="${escapeHtml(this.folderPickerQuery)}" />
        </div>
        <ul class="picker-list" role="listbox" aria-label="Dossiers Outlook">${list}</ul>
        ${all.length > 100 ? `<p class="help">100 premiers dossiers affichés : précise ta recherche.</p>` : ''}
        <div class="btn-row">
          <button type="button" data-action="confirm-folder" class="btn btn-primary btn-sm" ${!this.folderPickerSelectedId ? 'disabled' : ''}>
            ${this.folderPickerSelectedId ? `${icon('check', 14)}Ranger ici` : 'Sélectionne un dossier'}
          </button>
          <button type="button" data-action="cancel-folder" class="btn btn-ghost btn-sm">Annuler</button>
        </div>
      </div>
    `;
  }

  /**
   * Résout la meilleure destination de classement pour ce mail :
   *   1. Folder mapping appris pour le projet lié (Airtable FolderMappings).
   *   2. Sinon : top match parmi les dossiers Outlook existants (par nom du
   *      client / dénomination du projet).
   *   3. Sinon : suggestion d'un nouveau dossier à créer
   *      ("Clients/<Client>/#<NoProjet> <Dénomination>").
   *   4. Sinon : 'none' — aucune suggestion (mail sans projet lié).
   */
  private async resolveFolderSuggestion(): Promise<void> {
    const tag = this.tag;
    const userEmail = Office.context.mailbox?.userProfile?.emailAddress || '';
    const senderEmail = (Office.context.mailbox?.item as any)?.from?.emailAddress || '';

    try {
      // ── Stratégie 0 : Index local sender → dossier ──
      //   Source de vérité primaire. Construit une fois par le scan initial
      //   (bouton "Scanner ma boîte" en Settings) et enrichi à chaque action.
      //   Lookup O(1) — pas d'appel API à chaque ouverture de mail.
      if (senderEmail) {
        const hit = lookupSenderFolder(senderEmail);
        if (hit && hit.count >= 2) {
          this.folderSuggestion = {
            source: 'index-sender',
            folderId: hit.folderId,
            folderPath: hit.folderPath,
            projetId: tag?.linkedProjetId,
            patternMailCount: hit.count,
          };
          return;
        }
      }

      // ── Stratégie 1 : Folder mapping appris (Airtable) pour le projet lié ──
      if (tag?.linkedProjetId && userEmail) {
        const mapping = await getFolderMapping(userEmail, 'projet', tag.linkedProjetId);
        if (mapping?.folderId) {
          this.folderSuggestion = {
            source: 'mapped',
            folderId: mapping.folderId,
            folderPath: mapping.folderPath || '(dossier mémorisé)',
            projetId: tag.linkedProjetId,
          };
          return;
        }
      }

      // ── Stratégie 2 : Pattern d'expéditeur — où Charles range-t-il déjà
      //    les mails de ce sender ? Continue l'existant : si les 30 derniers
      //    mails du sender sont majoritairement dans 1 dossier, c'est CE
      //    dossier qu'on suggère, peu importe le projet lié ou non.
      let patternDebug = '';
      if (senderEmail) {
        try {
          const { getApiContext } = await import('../api/graph');
          const ctx = await getApiContext();
          const apiKind = ctx.base.includes('graph.microsoft.com') ? 'Graph' : 'OutlookREST';
          const result = await this.findSenderPatternFolderWithDebug(ctx.token, ctx.base, senderEmail);
          patternDebug = `${apiKind} · ${result.debug}`;
          if (result.match) {
            this.folderSuggestion = {
              source: 'sender-pattern',
              folderId: result.match.folderId,
              folderPath: result.match.folderPath,
              projetId: tag?.linkedProjetId,
              patternMailCount: result.match.count,
            };
            return;
          }
        } catch (e) {
          patternDebug = `Erreur lookup : ${humanError(e)}`;
          console.warn('[IAPanel] sender-pattern lookup failed:', e);
        }
      }

      // ── Stratégie 2.5 : Category-match — pas de pattern, pas de projet lié,
      //    mais l'IA a classé en (ex) "Fédération / Association" et l'utilisateur
      //    a un dossier "Fédérations" quelque part dans Outlook ? Suggère-le.
      //    Évite que des mails non-projet (newsletters, fédérations, fournisseurs)
      //    restent éternellement en Inbox faute de signal.
      if (tag?.category) {
        try {
          const { getApiContext } = await import('../api/graph');
          const ctx = await getApiContext();
          const all = await this.scanAllFoldersRecursive(ctx.token);
          const catMatch = this.findFolderForCategory(all, tag.category);
          if (catMatch) {
            this.folderSuggestion = {
              source: 'category-match',
              folderId: catMatch.id,
              folderPath: catMatch.path,
              projetId: tag.linkedProjetId,
            };
            return;
          }
        } catch (e) {
          console.warn('[IAPanel] category-match lookup failed:', e);
        }
      }

      // ── Stratégies 3 + 4 : nécessitent un projet lié (client + dénomination)
      if (!tag?.linkedProjetId) {
        this.folderSuggestion = { source: 'none', folderPath: '', debug: patternDebug };
        return;
      }
      const projets = await getAllProjets();
      const projet = projets.find((p) => p.id === tag.linkedProjetId);
      this.linkedProjet = projet || null;
      if (!projet) {
        this.folderSuggestion = { source: 'none', folderPath: '', debug: patternDebug };
        return;
      }

      // 3. Match fuzzy par nom sur les dossiers existants
      try {
        const token = await getGraphToken();
        const all = await this.scanAllFoldersRecursive(token);
        const match = this.findBestFolderMatch(all, projet);
        if (match) {
          this.folderSuggestion = {
            source: 'match',
            folderId: match.id,
            folderPath: match.path,
            projetId: tag.linkedProjetId,
          };
          return;
        }
      } catch (e) {
        console.warn('[IAPanel] folder scan failed:', e);
      }

      // 4. Création — propose un nouveau dossier basé sur Client / Projet
      const client = (projet.client || 'Clients').trim();
      const refOrNo = projet.refProjet || '';
      const denomination = (projet.denomination || '').slice(0, 60).trim();
      const suggestedPath = refOrNo
        ? `${client}/${refOrNo} ${denomination}`.trim()
        : `${client}/${denomination}`.trim();
      this.folderSuggestion = {
        source: 'create',
        folderPath: suggestedPath,
        projetId: tag.linkedProjetId,
      };
    } catch (e) {
      console.warn('[IAPanel] resolveFolderSuggestion failed:', e);
      this.folderSuggestion = { source: 'none', folderPath: '' };
    }
  }

  /**
   * Cherche le dossier où Charles range déjà les mails de ce sender.
   * Continue l'existant : on observe les 30 derniers mails du sender DANS
   * Outlook (toutes localisations sauf Inbox/Sent), on compte le dossier
   * majoritaire. Si > 50% → on suggère ce dossier.
   */
  /**
   * Wrapper de findSenderPatternFolder qui retourne aussi un texte de debug
   * visible dans l'UI (pour diagnostiquer sans DevTools).
   */
  private async findSenderPatternFolderWithDebug(
    token: string,
    apiBase: string,
    senderEmail: string,
  ): Promise<{ match: { folderId: string; folderPath: string; count: number } | null; debug: string }> {
    const dbg: string[] = [];
    const orig = console.info;
    const captured: string[] = [];
    console.info = (...args: any[]) => {
      try { captured.push(args.map((a) => (typeof a === 'string' ? a : JSON.stringify(a))).join(' ')); } catch {}
      orig(...args);
    };
    try {
      const match = await this.findSenderPatternFolder(token, apiBase, senderEmail);
      dbg.push(...captured.filter((c) => c.includes('sender-pattern')).map((c) => c.replace('[IAPanel] sender-pattern', '').trim()));
      return { match, debug: dbg.join(' | ') || 'aucune réponse API' };
    } finally {
      console.info = orig;
    }
  }

  private async findSenderPatternFolder(
    token: string,
    apiBase: string,
    senderEmail: string,
  ): Promise<{ folderId: string; folderPath: string; count: number } | null> {
    // Filtre : from = senderEmail. On exclut côté JS les mails encore en Inbox.
    // On augmente $top à 100 pour catcher l'historique long (50+ mails sur un sender).
    const filter = encodeURIComponent(`from/emailAddress/address eq '${senderEmail.replace(/'/g, "''")}'`);
    const url = `${apiBase}/me/messages?$filter=${filter}&$top=100&$select=id,parentFolderId,subject&$orderby=receivedDateTime desc`;
    let res: Response;
    try {
      res = await outlookFetch(url, { headers: { Authorization: `Bearer ${token}` } });
    } catch (e) {
      console.error('[IAPanel] sender-pattern fetch failed (network):', e);
      return null;
    }
    if (!res.ok) {
      const errBody = await res.text().catch(() => '');
      console.error(`[IAPanel] sender-pattern HTTP ${res.status} on ${url}:`, errBody.slice(0, 300));
      return null;
    }
    const data: any = await res.json();
    const rawMsgs = (data.value || []) as any[];
    // Outlook REST v2.0 peut renvoyer PascalCase (ParentFolderId) ou camelCase
    // selon le mode. On normalise pour gérer les 2.
    const msgs = rawMsgs.map((m) => ({
      parentFolderId: m.parentFolderId || m.ParentFolderId || '',
      subject: m.subject || m.Subject || '',
    }));
    console.info(`[IAPanel] sender-pattern: ${msgs.length} mails de ${senderEmail} (sample: ${JSON.stringify(msgs.slice(0, 3))})`);
    if (msgs.length < 3) return null;

    // Récupère l'ID du dossier Inbox pour l'exclure
    const inboxRes = await outlookFetch(`${apiBase}/me/mailFolders/inbox?$select=id`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    const inboxData = inboxRes.ok ? await inboxRes.json() : {};
    const inboxId = inboxData.id || inboxData.Id || '';

    // Compte par dossier (hors Inbox)
    const counts = new Map<string, number>();
    let skippedInbox = 0;
    let skippedNoId = 0;
    for (const m of msgs) {
      if (!m.parentFolderId) { skippedNoId++; continue; }
      if (m.parentFolderId === inboxId) { skippedInbox++; continue; }
      counts.set(m.parentFolderId, (counts.get(m.parentFolderId) || 0) + 1);
    }
    console.info(`[IAPanel] sender-pattern counts: ${counts.size} folders, ${skippedInbox} in Inbox, ${skippedNoId} no parentFolderId`);
    if (counts.size === 0) return null;

    // Top dossier
    let bestId = '';
    let bestCount = 0;
    for (const [fid, c] of counts.entries()) {
      if (c > bestCount) { bestCount = c; bestId = fid; }
    }
    // Pattern significatif : au moins 3 mails ET >= 50% des mails classés
    const totalClassified = Array.from(counts.values()).reduce((a, b) => a + b, 0);
    if (bestCount < 3 || bestCount / totalClassified < 0.5) return null;

    // Résoudre le path lisible du folder (1 niveau parent suffit pour le contexte)
    let folderPath = bestId.slice(0, 8); // fallback ID si pas de nom
    try {
      const fres = await outlookFetch(`${apiBase}/me/mailFolders/${bestId}?$select=displayName,parentFolderId`, {
        headers: { Authorization: `Bearer ${token}` },
      });
      if (fres.ok) {
        const f: any = await fres.json();
        const dn = f.displayName || f.DisplayName || '';
        const pfId = f.parentFolderId || f.ParentFolderId || '';
        folderPath = dn || folderPath;
        // Tente de récupérer le parent pour afficher "Parent/Folder"
        if (pfId && pfId !== inboxId) {
          const pres = await outlookFetch(`${apiBase}/me/mailFolders/${pfId}?$select=displayName`, {
            headers: { Authorization: `Bearer ${token}` },
          });
          if (pres.ok) {
            const p: any = await pres.json();
            const pdn = p.displayName || p.DisplayName || '';
            if (pdn && !['Top of Information Store', 'Inbox'].includes(pdn)) {
              folderPath = `${pdn}/${dn}`;
            }
          }
        }
      }
    } catch { /* ignore — on garde l'ID en fallback */ }

    return { folderId: bestId, folderPath, count: bestCount };
  }

  /**
   * Scan tous les dossiers Outlook récursivement (jusqu'à 4 niveaux de
   * profondeur). On exclut les dossiers système.
   */
  private async scanAllFoldersRecursive(token: string): Promise<Array<{ id: string; path: string }>> {
    const SKIP = new Set(['inbox', 'sent items', 'drafts', 'deleted items', 'junk email', 'outbox', 'archive', 'rss feeds', 'conversation history', 'sync issues', 'clutter', 'notes']);
    const out: Array<{ id: string; path: string }> = [];

    const walk = async (parentId: string | undefined, parentPath: string, depth: number): Promise<void> => {
      if (depth > 4) return;
      let children: Array<{ id: string; displayName: string }> = [];
      try { children = await listMailFolders(token, parentId); } catch { return; }
      for (const f of children) {
        const name = f.displayName || '';
        if (!name) continue;
        if (depth === 0 && SKIP.has(name.toLowerCase())) continue;
        const path = parentPath ? `${parentPath}/${name}` : name;
        out.push({ id: f.id, path });
        await walk(f.id, path, depth + 1);
      }
    };

    await walk(undefined, '', 0);
    return out;
  }

  /**
   * Cherche un dossier Outlook dont le nom match les alias de la catégorie IA.
   * Ex: catégorie "federation_association" → cherche un dossier nommé
   * "Fédérations", "Federation", "Associations" etc.
   * Scoring : alias plus long = match plus précis. Min 4 chars pour éviter
   * les faux positifs sur des alias courts comme "news".
   */
  private findFolderForCategory(
    folders: Array<{ id: string; path: string }>,
    category: string,
  ): { id: string; path: string } | null {
    const aliases = (CATEGORY_FOLDER_ALIASES[category] || []).filter((a) => a.length >= 4);
    if (aliases.length === 0) return null;

    let best: { id: string; path: string; score: number } | null = null;
    for (const f of folders) {
      const p = f.path.toLowerCase();
      const leaf = (p.split('/').pop() || '').toLowerCase();
      for (const alias of aliases) {
        const a = alias.toLowerCase();
        let s = 0;
        if (leaf === a) s = 100;                  // match exact sur le nom du dossier
        else if (leaf.includes(a)) s = 60;        // substring sur le nom du dossier
        else if (p.includes(a)) s = 30;           // substring quelque part dans le path
        if (s > (best?.score || 0)) best = { id: f.id, path: f.path, score: s };
      }
    }
    if (!best || best.score < 30) return null;
    return { id: best.id, path: best.path };
  }

  /**
   * Cherche le dossier dont le nom (ou un segment du path) match le mieux le
   * client / la dénomination du projet. Scoring simple : substring case-insensitive.
   */
  private findBestFolderMatch(
    folders: Array<{ id: string; path: string }>,
    projet: Projet,
  ): { id: string; path: string } | null {
    const client = (projet.client || '').toLowerCase().trim();
    const denomination = (projet.denomination || '').toLowerCase().trim();
    const ref = (projet.refProjet || '').toLowerCase().trim();
    if (!client && !denomination && !ref) return null;

    type Scored = { id: string; path: string; score: number };
    const scored: Scored[] = folders.map((f) => {
      const p = f.path.toLowerCase();
      let score = 0;
      if (ref && p.includes(ref)) score += 100;            // match référence projet = très fort
      if (denomination && p.includes(denomination)) score += 50;
      if (client) {
        // Match client : on découpe en mots pour catcher "Markcom" même si le path est "Fédérations/Markcom/2026"
        const clientWords = client.split(/[\s,&]+/).filter((w) => w.length >= 3);
        for (const w of clientWords) {
          if (p.includes(w)) score += 20;
        }
      }
      return { id: f.id, path: f.path, score };
    });
    scored.sort((a, b) => b.score - a.score);
    const best = scored[0];
    if (!best || best.score < 20) return null;
    return { id: best.id, path: best.path };
  }

  /**
   * Déplace le mail courant dans le dossier de destination résolu (mapped /
   * match / create). Sauvegarde le folder mapping pour l'apprentissage.
   * Retourne le path du dossier si OK, '' si rien fait.
   */
  private async tryMoveToHabitualFolder(
    _actionName: string,
  ): Promise<{ folderPath: string; error?: string }> {
    const suggestion = this.folderSuggestion;
    if (!suggestion || suggestion.source === 'none') {
      return { folderPath: '', error: 'aucune suggestion' };
    }

    try {
      const userEmail = Office.context.mailbox?.userProfile?.emailAddress || '';
      const item = Office.context.mailbox?.item;
      const rawId = (item as any)?.itemId;
      if (!rawId) return { folderPath: '', error: 'pas d\'itemId' };
      const restId = convertToRestId(rawId);
      const token = await getGraphToken();

      let folderId = suggestion.folderId;
      // Cas 'create' : on crée le dossier dans Outlook (ensureFolderPath crée récursivement)
      if (suggestion.source === 'create' && !folderId) {
        folderId = await ensureFolderPath(token, suggestion.folderPath);
      }
      if (!folderId) {
        return { folderPath: '', error: `pas de folderId (source=${suggestion.source})` };
      }

      // Move — peut throw si l'API rejette
      await moveMessageToFolder(token, restId, folderId);

      // Apprentissage local : enrichit l'index sender → dossier IMMÉDIATEMENT.
      const senderEmail = (item as any)?.from?.emailAddress || '';
      if (senderEmail) {
        const weight = suggestion.source === 'manual-override' ? 2 : 1;
        recordSenderFolder(senderEmail, folderId, suggestion.folderPath, weight);
      }

      // Apprentissage : sauvegarde le mapping (Airtable) si projet lié
      if (suggestion.source !== 'mapped' && suggestion.projetId && userEmail) {
        try {
          await saveFolderMapping(userEmail, 'projet', suggestion.projetId, suggestion.folderPath, folderId);
        } catch (e) {
          console.warn('[IAPanel] saveFolderMapping failed (non-fatal):', e);
        }
      }

      return { folderPath: suggestion.folderPath };
    } catch (e) {
      const msg = humanError(e);
      console.warn('[IAPanel] tryMoveToHabitualFolder failed:', e);
      return { folderPath: '', error: msg };
    }
  }

  /**
   * Applique la catégorie Outlook qui correspond à l'état + l'urgence.
   * Visible directement dans la liste inbox d'Outlook (tag coloré à côté
   * du sujet). Permet à Charles de voir d'un coup d'œil ce qui est
   * reporté/traité/urgent sans ouvrir le task-pane.
   */
  private async applyCategoryForState(state: 'done' | 'snoozed' | 'archived' | 'tagged'): Promise<void> {
    const item = Office.context.mailbox?.item;
    const rawId = (item as any)?.itemId;
    if (!rawId) return;
    const restId = convertToRestId(rawId);

    // On applique simultanément plusieurs catégories (Outlook les affiche
    // toutes côte-à-côte dans la vue liste) :
    //   - IA (type de mail : "💼 Demande devis", "🏛 Fédération / Association"...)
    //   - Urgence si >= 4 (signaux visuels forts pour ne rien rater)
    //   - État (Reporté / Traité / Archivé) si pas en inbox
    const cats: string[] = [];

    // 1) Catégorie IA (toujours, dès qu'on a un tag)
    if (this.tag?.category) {
      const ia = ATLAS_IA_CATEGORIES[this.tag.category];
      if (ia) cats.push(ia.name);
    }

    // 2) Urgence haute (>= 4)
    if (this.tag) {
      const u = this.tag.urgencyScore || 0;
      if (u >= 5) cats.push(ATLAS_CATEGORIES.URGENCE_5.name);
      else if (u === 4) cats.push(ATLAS_CATEGORIES.URGENCE_4.name);
    }

    // 3) État
    if (state === 'snoozed') cats.push(ATLAS_CATEGORIES.SNOOZED.name);
    else if (state === 'done') cats.push(ATLAS_CATEGORIES.DONE.name);
    else if (state === 'archived') cats.push(ATLAS_CATEGORIES.ARCHIVED.name);

    try {
      await setMessageCategories(restId, cats);
    } catch (e) {
      console.warn('[IAPanel] setMessageCategories failed:', e);
    }
  }

  /**
   * Appelle Claude pour suggérer un path de dossier cohérent avec la
   * structure existante. Pre-fill du champ création.
   */
  private async fetchAIFolderSuggestion(): Promise<void> {
    try {
      const item = Office.context.mailbox?.item as any;
      if (!item) { this.folderPickerSuggestingAI = false; this.render(); return; }
      const senderName = item.from?.displayName || '';
      const senderEmail = item.from?.emailAddress || '';
      const subject = item.subject || '';

      // S'assure que le cache des dossiers est chargé
      if (!this.allFoldersCache || this.allFoldersCache.length === 0) {
        await this.loadAllFolders();
      }

      const { suggestFolderPath } = await import('../api/argo');
      const suggested = await suggestFolderPath({
        existingFolders: (this.allFoldersCache || []).map((f) => f.path),
        iaCategory: this.tag?.category,
        iaCategoryLabel: this.tag?.category ? (CATEGORY_LABELS[this.tag.category] || this.tag.category) : undefined,
        senderName,
        senderEmail,
        subject,
        summary: this.tag?.summary,
      });
      if (suggested && this.folderPickerCreateMode) {
        this.folderPickerNewPath = suggested;
      }
    } catch (e) {
      console.warn('[IAPanel] fetchAIFolderSuggestion failed:', e);
    } finally {
      this.folderPickerSuggestingAI = false;
      this.render();
    }
  }

  /**
   * Charge la liste complète des dossiers Outlook (cache en mémoire).
   * Utilisé par le picker manuel.
   */
  private async loadAllFolders(): Promise<void> {
    if (this.allFoldersCache && this.allFoldersCache.length > 0) return;
    this.foldersLoadError = null;
    try {
      const { getApiContext } = await import('../api/graph');
      const ctx = await getApiContext();
      this.allFoldersCache = await this.scanAllFoldersRecursive(ctx.token);
    } catch (e) {
      console.warn('[IAPanel] loadAllFolders failed:', e);
      this.allFoldersCache = [];
      this.foldersLoadError = e;
    }
  }
}

