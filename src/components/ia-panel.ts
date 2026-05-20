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
  hasAnthropicToken,
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
} from '../api/graph';

const CATEGORIES = [
  'demande_devis', 'validation_client', 'refus_client', 'question_staff',
  'facture_fournisseur', 'prospection_entrante', 'rdv_planning',
  'newsletter', 'notification_systeme', 'spam', 'autre',
  'federation_association', 'demande_interne_staff', 'fournisseur',
];

const CATEGORY_LABELS: Record<string, string> = {
  demande_devis: '💼 Demande devis',
  validation_client: '✅ Validation client',
  refus_client: '❌ Refus client',
  question_staff: '👥 Question staff',
  facture_fournisseur: '🧾 Facture fournisseur',
  prospection_entrante: '📞 Prospection entrante',
  rdv_planning: '📅 RDV / Planning',
  newsletter: '📰 Newsletter',
  notification_systeme: '🤖 Notification système',
  spam: '🚫 Spam',
  autre: '📩 Autre',
  federation_association: '🏛 Fédération / Association',
  demande_interne_staff: '🏠 Interne staff',
  fournisseur: '🚚 Fournisseur',
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
  prospection_entrante: ['prospect', 'prospection', 'lead'],
  rdv_planning: ['rdv', 'rendez', 'planning', 'agenda'],
  newsletter: ['newsletter', 'newsletters', 'news'],
  notification_systeme: ['notification', 'système', 'systeme', 'auto'],
  spam: ['spam', 'junk', 'courrier indésirable'],
  autre: [],
  federation_association: ['fédération', 'federation', 'fédérations', 'federations', 'association', 'associations'],
  demande_interne_staff: ['interne', 'staff', 'équipe', 'equipe'],
  fournisseur: ['fournisseur', 'fournisseurs', 'suppliers'],
};

const URGENCY_COLORS: Record<number, string> = {
  1: '#94a3b8', 2: '#60a5fa', 3: '#f59e0b', 4: '#ef4444', 5: '#dc2626',
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
    this.root.innerHTML = `
      <div style="padding: 16px; text-align: center; color: #64748b;">
        <div class="spinner"></div>
        <p style="font-size: 12px; margin-top: 8px;">Chargement IA…</p>
      </div>
    `;
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
        this.renderNotTagged();
      }
    } catch (e) {
      console.warn('[IAPanel] load failed:', e);
      this.renderEmpty('Erreur lors du chargement.');
    }
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
    this.root.innerHTML = `
      <div style="padding: 20px; text-align: center; color: #64748b; font-size: 13px;">${msg}</div>
    `;
  }

  private renderNotTagged(): void {
    const hasKey = hasAnthropicToken();
    this.root.innerHTML = `
      <div style="padding: 16px; display: flex; flex-direction: column; gap: 12px;">
        <div style="padding: 12px; background: #fff7ed; border: 1px solid #fed7aa; border-radius: 8px; color: #9a3412;">
          <strong>Pas encore taggé par l'IA</strong>
          <p style="margin: 6px 0 0; font-size: 12px;">Le scanner ATLAS ne l'a pas (encore) analysé. Tu peux forcer l'analyse maintenant — résultat en ~3 sec.</p>
        </div>
        <button data-action="analyze-now" class="ia-btn ia-btn-secondary" ${hasKey ? '' : 'disabled title="Configure Clé Anthropic dans Settings"'}
          style="padding: 10px 16px; font-size: 14px; font-weight: 600;">
          ✨ Analyser maintenant
        </button>
        ${hasKey ? '' : '<div style="font-size: 11px; color: #9a3412;">⚠️ Clé Anthropic manquante — onglet Settings ⚙️.</div>'}
      </div>
    `;
    this.root.querySelector<HTMLButtonElement>('button[data-action="analyze-now"]')
      ?.addEventListener('click', async (ev) => {
        const btn = ev.currentTarget as HTMLButtonElement;
        btn.disabled = true;
        btn.textContent = '⏳ Analyse en cours…';
        try {
          const ok = await this.reanalyzeNow();
          if (ok) this.render(); // bascule sur la vue taggée
          else {
            btn.disabled = false;
            btn.textContent = '✨ Analyser maintenant';
          }
        } catch {
          btn.disabled = false;
          btn.textContent = '✨ Analyser maintenant';
        }
      });
  }

  private render(): void {
    const tag = this.tag!;
    const catLabel = CATEGORY_LABELS[tag.category] || tag.category;
    const urgColor = URGENCY_COLORS[tag.urgencyScore] || '#94a3b8';
    const status = tag.inboxStatus;

    this.root.innerHTML = `
      <div style="padding: 14px; display: flex; flex-direction: column; gap: 14px;">
        <!-- Catégorie + urgence -->
        <div style="display: flex; flex-direction: column; gap: 6px;">
          <div style="font-size: 10px; font-weight: 700; color: #475569; text-transform: uppercase; letter-spacing: 0.5px;">Classification IA</div>
          <div style="display: flex; align-items: center; gap: 8px; flex-wrap: wrap;">
            <span style="padding: 4px 10px; background: #f1f5f9; border-radius: 12px; font-size: 13px; font-weight: 600;">
              ${catLabel}
            </span>
            <span style="padding: 4px 10px; background: ${urgColor}; color: #fff; border-radius: 12px; font-size: 12px; font-weight: 700;">
              Urgence ${tag.urgencyScore}/5
            </span>
            ${status !== 'inbox' ? `<span style="padding: 4px 10px; background: #ecfdf5; color: #065f46; border-radius: 12px; font-size: 11px;">État : ${status}</span>` : ''}
          </div>
          ${tag.summary ? `<p style="margin: 4px 0 0; font-size: 12px; color: #475569; line-height: 1.4;">${escapeHtml(tag.summary)}</p>` : ''}
        </div>

        <!-- Dossier de classement (suggestion intelligente) -->
        ${this.renderFolderSection()}

        <!-- Actions principales -->
        <div style="display: flex; flex-direction: column; gap: 8px;">
          <div style="font-size: 10px; font-weight: 700; color: #475569; text-transform: uppercase; letter-spacing: 0.5px;">Actions</div>
          <div style="display: flex; gap: 6px; flex-wrap: wrap;">
            <button data-action="done" class="ia-btn ia-btn-success">✓ Traité</button>
            <button data-action="snooze" class="ia-btn ia-btn-warning">⏰ Reporter (demain 8h)</button>
            <button data-action="archive" class="ia-btn ia-btn-muted">📦 Archiver</button>
            <button data-action="reanalyze" class="ia-btn ia-btn-secondary" ${hasAnthropicToken() ? '' : 'disabled title="Configure Clé Anthropic dans Settings"'}>🔄 Re-analyser</button>
          </div>
        </div>

        <!-- Correction -->
        <div style="display: flex; flex-direction: column; gap: 6px;">
          <label style="font-size: 10px; font-weight: 700; color: #475569; text-transform: uppercase; letter-spacing: 0.5px;">Corriger la catégorie (apprentissage)</label>
          <select id="ia-cat-select" style="padding: 8px; border: 1px solid #cbd5e1; border-radius: 6px; font-size: 13px; color: #1a1a1a; background: #fff;">
            ${CATEGORIES.map(cat => `<option value="${cat}" ${cat === tag.category ? 'selected' : ''}>${CATEGORY_LABELS[cat] || cat}</option>`).join('')}
          </select>
          <button data-action="correct" class="ia-btn ia-btn-primary" style="margin-top: 4px;">Enregistrer la correction</button>
        </div>
      </div>

      <style>
        .ia-btn { padding: 8px 12px; border: 0; border-radius: 6px; font-weight: 600; font-size: 12px; cursor: pointer; }
        .ia-btn-success { background: #10b981; color: #fff; }
        .ia-btn-warning { background: #f59e0b; color: #fff; }
        .ia-btn-muted { background: #94a3b8; color: #fff; }
        .ia-btn-primary { background: #cc2200; color: #fff; }
        .ia-btn-secondary { background: #6366f1; color: #fff; }
        .ia-btn:hover { opacity: 0.9; }
        .ia-btn:disabled { opacity: 0.5; cursor: not-allowed; }
      </style>
    `;

    // Click handlers
    this.root.querySelectorAll<HTMLButtonElement>('button[data-action]').forEach(btn => {
      btn.addEventListener('click', () => this.onAction(btn.dataset.action!));
    });
  }

  private async onAction(action: string): Promise<void> {
    // Picker dossier : ne nécessite pas de tag (peut être déclenché avant que
    // la classification soit faite, juste pour préparer le mapping).
    if (action === 'change-folder') {
      this.folderPickerOpen = true;
      // Charge la liste des dossiers en background
      this.loadAllFolders().then(() => this.render()).catch(() => this.render());
      this.render();
      return;
    }
    if (action === 'cancel-folder') {
      this.folderPickerOpen = false;
      this.render();
      return;
    }
    if (action === 'confirm-folder') {
      const sel = this.root.querySelector<HTMLSelectElement>('#folder-picker-select');
      const folderId = sel?.value || '';
      const folderPath = sel?.selectedOptions[0]?.getAttribute('data-path') || '';
      if (!folderId || !folderPath) {
        showToast('Sélectionne un dossier', 'error');
        return;
      }
      this.folderSuggestion = {
        source: 'manual-override',
        folderId,
        folderPath,
        projetId: this.tag?.linkedProjetId,
      };
      this.folderPickerOpen = false;
      showToast(`Dossier choisi : ${folderPath}. Clic Traité ou Archiver pour confirmer.`, 'info');
      this.render();
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
            const moved = await this.tryMoveToHabitualFolder('done');
            showToast(moved ? `Traité ✓ + déplacé dans ${moved} 📁` : 'Marqué comme traité ✓', 'success');
          }
          break;
        case 'snooze':
          ok = await snoozeTag(this.tag.id);
          if (ok) { this.tag.inboxStatus = 'snoozed'; showToast('Reporté à demain 8h ⏰', 'success'); }
          break;
        case 'archive':
          ok = await archiveTag(this.tag.id);
          if (ok) {
            this.tag.inboxStatus = 'archived';
            // Pour Archive : on tente de déplacer dans le dossier habituel (projet
            // lié → folder mapping). Si pas de mapping → laisse le mail en inbox
            // (l'utilisateur peut toujours le déplacer manuellement).
            const moved = await this.tryMoveToHabitualFolder('archive');
            showToast(moved ? `Archivé + déplacé dans ${moved} 📁` : 'Archivé 📦 (aucun dossier habituel — déplace manuellement si besoin)', 'success');
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
      showToast('Erreur', 'error');
    } finally {
      buttons.forEach(b => b.disabled = false);
    }
  }

  /**
   * Force une re-analyse Claude pour le mail courant. Purge l'ancien tag +
   * écrit le nouveau résultat. Résultat immédiat — pas d'attente du scan
   * périodique côté app desktop. Nécessite ANTHROPIC_API_KEY configurée
   * dans Settings de l'addin.
   */
  private async reanalyzeNow(): Promise<boolean> {
    if (!hasAnthropicToken()) {
      showToast('Configure Clé Anthropic dans Settings de l\'addin', 'error');
      return false;
    }
    try {
      const item = Office.context.mailbox?.item as any;
      if (!item) { showToast('Aucun mail sélectionné', 'error'); return false; }
      const userEmail = Office.context.mailbox?.userProfile?.emailAddress || '';
      if (!userEmail) { showToast('Pas d\'email utilisateur Office.js', 'error'); return false; }

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

      showToast('Analyse Claude en cours…', 'info');

      // 2. Appelle Claude
      const analysis = await analyzeEmailWithClaude({
        subject,
        from,
        toRecipients,
        ccRecipients,
        body,
        receivedAt,
        userEmail,
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
      showToast(`Re-analysé ✓ : ${CATEGORY_LABELS[analysis.category] || analysis.category}`, 'success');
      // Re-résout aussi le dossier de classement (la catégorie a peut-être changé)
      this.resolveFolderSuggestion().then(() => this.render()).catch(() => {});
      return true;
    } catch (e) {
      console.warn('[IAPanel] reanalyze failed:', e);
      showToast(`Erreur re-analyse : ${(e as Error).message || 'inconnue'}`, 'error');
      return false;
    }
  }

  /** Section "Dossier de classement" affichée dans le rendu IA. */
  private renderFolderSection(): string {
    const s = this.folderSuggestion;
    if (!s) {
      return `<div style="padding: 10px; background: #f8fafc; border: 1px dashed #cbd5e1; border-radius: 6px; font-size: 11px; color: #64748b;">📁 Résolution du dossier en cours…</div>`;
    }
    if (s.source === 'none') {
      // Aucune suggestion exploitable. On affiche un mini diagnostic visible pour
      // que Charles voie POURQUOI (sans DevTools).
      const dbg = s.debug ? escapeHtml(s.debug).slice(0, 400) : 'aucun signal';
      return `
        <div style="padding: 10px; background: #fafafa; border: 1px dashed #cbd5e1; border-radius: 6px;">
          <div style="font-size: 10px; font-weight: 700; color: #64748b; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 4px;">📁 Aucun dossier suggéré</div>
          <div style="font-size: 11px; color: #94a3b8; font-family: monospace;">Debug : ${dbg}</div>
        </div>
      `;
    }
    // Construit la suggestion + bouton "Changer" (chip discret en haut-droite).
    // Le bouton ouvre un dropdown <select> avec tous les dossiers Outlook.
    const changeBtn = `
      <div style="display: flex; justify-content: flex-end; margin-top: 6px;">
        <button data-action="change-folder" style="background: rgba(255,255,255,0.6); border: 1px solid currentColor; padding: 3px 8px; border-radius: 4px; font-size: 10px; cursor: pointer; opacity: 0.7; color: inherit;">
          📂 Changer
        </button>
      </div>
    `;
    const picker = this.folderPickerOpen ? this.renderFolderPicker() : '';

    if (s.source === 'index-sender') {
      return `
        <div style="padding: 10px; background: #ecfdf5; border: 1px solid #6ee7b7; border-radius: 6px; color: #047857;">
          <div style="font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 4px;">📁 Tu ranges déjà ces mails ici ✓</div>
          <div style="font-size: 13px; font-weight: 600; color: #065f46;">${escapeHtml(s.folderPath)}</div>
          <div style="font-size: 11px; margin-top: 4px;">Index local : ${s.patternMailCount} mail${(s.patternMailCount || 0) > 1 ? 's' : ''} de cet expéditeur déjà rangés ici. ATLAS suit ton habitude.</div>
          ${changeBtn}
        </div>
        ${picker}
      `;
    }
    if (s.source === 'mapped') {
      return `
        <div style="padding: 10px; background: #ecfdf5; border: 1px solid #6ee7b7; border-radius: 6px; color: #047857;">
          <div style="font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 4px;">📁 Dossier habituel (appris)</div>
          <div style="font-size: 13px; font-weight: 600; color: #065f46;">${escapeHtml(s.folderPath)} ✓</div>
          <div style="font-size: 11px; margin-top: 2px;">Sera utilisé automatiquement au clic sur Traité ou Archiver.</div>
          ${changeBtn}
        </div>
        ${picker}
      `;
    }
    if (s.source === 'manual-override') {
      return `
        <div style="padding: 10px; background: #fef3c7; border: 1px solid #fcd34d; border-radius: 6px; color: #92400e;">
          <div style="font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 4px;">📁 Dossier choisi (à apprendre)</div>
          <div style="font-size: 13px; font-weight: 600; color: #78350f;">${escapeHtml(s.folderPath)}</div>
          <div style="font-size: 11px; margin-top: 4px;">Au clic Traité/Archiver, le mail file ici et ATLAS retient ton choix pour la prochaine fois.</div>
          ${changeBtn}
        </div>
        ${picker}
      `;
    }
    if (s.source === 'sender-pattern') {
      return `
        <div style="padding: 10px; background: #f5f3ff; border: 1px solid #c4b5fd; border-radius: 6px; color: #6d28d9;">
          <div style="font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 4px;">📁 Tu ranges déjà ces mails ici</div>
          <div style="font-size: 13px; font-weight: 600; color: #5b21b6;">${escapeHtml(s.folderPath)}</div>
          <div style="font-size: 11px; margin-top: 4px;">Observé : ${s.patternMailCount} mail${(s.patternMailCount || 0) > 1 ? 's' : ''} de cet expéditeur déjà dans ce dossier.</div>
          ${changeBtn}
        </div>
        ${picker}
      `;
    }
    if (s.source === 'category-match') {
      return `
        <div style="padding: 10px; background: #eff6ff; border: 1px solid #93c5fd; border-radius: 6px; color: #1d4ed8;">
          <div style="font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 4px;">📁 Dossier suggéré (par catégorie IA)</div>
          <div style="font-size: 13px; font-weight: 600; color: #1e3a8a;">${escapeHtml(s.folderPath)}</div>
          <div style="font-size: 11px; margin-top: 4px;">Match basé sur la catégorie. Si c'est pas le bon dossier, change-le ↓</div>
          ${changeBtn}
        </div>
        ${picker}
      `;
    }
    if (s.source === 'match') {
      return `
        <div style="padding: 10px; background: #eff6ff; border: 1px solid #93c5fd; border-radius: 6px; color: #1d4ed8;">
          <div style="font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 4px;">📁 Dossier suggéré (match par nom)</div>
          <div style="font-size: 13px; font-weight: 600; color: #1e3a8a;">${escapeHtml(s.folderPath)}</div>
          <div style="font-size: 11px; margin-top: 4px;">ATLAS y range le mail au clic Traité/Archiver et s'en souviendra.</div>
          ${changeBtn}
        </div>
        ${picker}
      `;
    }
    // source === 'create'
    return `
      <div style="padding: 10px; background: #fef3c7; border: 1px solid #fcd34d; border-radius: 6px; color: #92400e;">
        <div style="font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 4px;">📁 Nouveau dossier suggéré</div>
        <div style="font-size: 13px; font-weight: 600; color: #78350f;">${escapeHtml(s.folderPath)}</div>
        <div style="font-size: 11px; margin-top: 4px;">Aucun dossier Outlook existant ne match. Au clic Traité/Archiver, ATLAS le crée.</div>
        ${changeBtn}
      </div>
      ${picker}
    `;
  }

  /**
   * Dropdown <select> avec tous les dossiers Outlook. Apparaît au clic sur
   * "Changer". Sélection → tap Confirmer → la suggestion devient
   * 'manual-override' avec ce dossier.
   */
  private renderFolderPicker(): string {
    const folders = this.allFoldersCache || [];
    if (folders.length === 0) {
      return `
        <div style="padding: 10px; background: #f8fafc; border: 1px dashed #cbd5e1; border-radius: 6px; margin-top: 6px;">
          <div style="font-size: 11px; color: #64748b;">⏳ Chargement de la liste de dossiers…</div>
        </div>
      `;
    }
    const opts = folders
      .slice()
      .sort((a, b) => a.path.localeCompare(b.path, 'fr'))
      .map((f) => `<option value="${escapeHtml(f.id)}" data-path="${escapeHtml(f.path)}">${escapeHtml(f.path)}</option>`)
      .join('');
    return `
      <div style="padding: 10px; background: #f8fafc; border: 1px solid #cbd5e1; border-radius: 6px; margin-top: 6px;">
        <div style="font-size: 10px; font-weight: 700; color: #475569; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 6px;">📂 Choisir un dossier</div>
        <select id="folder-picker-select" style="width: 100%; padding: 6px 8px; border: 1px solid #cbd5e1; border-radius: 4px; font-size: 12px; margin-bottom: 6px;">
          <option value="">— Sélectionner —</option>
          ${opts}
        </select>
        <div style="display: flex; gap: 6px;">
          <button data-action="confirm-folder" style="flex: 1; padding: 6px; background: #4f46e5; color: white; border: none; border-radius: 4px; font-size: 12px; font-weight: 600; cursor: pointer;">Confirmer</button>
          <button data-action="cancel-folder" style="padding: 6px 10px; background: #e2e8f0; color: #334155; border: none; border-radius: 4px; font-size: 12px; cursor: pointer;">Annuler</button>
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
          patternDebug = `Erreur lookup : ${(e as Error).message?.slice(0, 100)}`;
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
      res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
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
    const inboxRes = await fetch(`${apiBase}/me/mailFolders/inbox?$select=id`, {
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
      const fres = await fetch(`${apiBase}/me/mailFolders/${bestId}?$select=displayName,parentFolderId`, {
        headers: { Authorization: `Bearer ${token}` },
      });
      if (fres.ok) {
        const f: any = await fres.json();
        const dn = f.displayName || f.DisplayName || '';
        const pfId = f.parentFolderId || f.ParentFolderId || '';
        folderPath = dn || folderPath;
        // Tente de récupérer le parent pour afficher "Parent/Folder"
        if (pfId && pfId !== inboxId) {
          const pres = await fetch(`${apiBase}/me/mailFolders/${pfId}?$select=displayName`, {
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
   * Scan tous les dossiers Outlook récursivement (depuis racine), 2 niveaux
   * max pour ne pas exploser le nombre d'appels Graph.
   */
  private async scanAllFoldersRecursive(token: string): Promise<Array<{ id: string; path: string }>> {
    const out: Array<{ id: string; path: string }> = [];
    const roots = await listMailFolders(token);
    // Filtre les system folders inutiles (Inbox/Sent/Drafts/Deleted Items/Junk)
    const SKIP = new Set(['inbox', 'sent items', 'drafts', 'deleted items', 'junk email', 'outbox', 'archive', 'rss feeds', 'conversation history', 'sync issues']);
    for (const f of roots) {
      if (SKIP.has(f.displayName.toLowerCase())) continue;
      out.push({ id: f.id, path: f.displayName });
      try {
        const children = await listMailFolders(token, f.id);
        for (const c of children) {
          out.push({ id: c.id, path: `${f.displayName}/${c.displayName}` });
        }
      } catch { /* skip */ }
    }
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
  private async tryMoveToHabitualFolder(_actionName: string): Promise<string> {
    try {
      const suggestion = this.folderSuggestion;
      if (!suggestion || suggestion.source === 'none') return '';

      const userEmail = Office.context.mailbox?.userProfile?.emailAddress || '';
      const item = Office.context.mailbox?.item;
      const rawId = (item as any)?.itemId;
      if (!rawId) return '';
      const restId = convertToRestId(rawId);
      const token = await getGraphToken();

      let folderId = suggestion.folderId;
      // Cas 'create' : on crée le dossier dans Outlook (ensureFolderPath crée récursivement)
      if (suggestion.source === 'create' && !folderId) {
        folderId = await ensureFolderPath(token, suggestion.folderPath);
      }
      if (!folderId) return '';

      // Move
      await moveMessageToFolder(token, restId, folderId);

      // Apprentissage local : enrichit l'index sender → dossier IMMÉDIATEMENT.
      // Lookup O(1) à la prochaine ouverture d'un mail du même sender.
      const senderEmail = (item as any)?.from?.emailAddress || '';
      if (senderEmail) {
        // +2 si manual-override (signal fort de Charles), +1 sinon
        const weight = suggestion.source === 'manual-override' ? 2 : 1;
        recordSenderFolder(senderEmail, folderId, suggestion.folderPath, weight);
      }

      // Apprentissage : sauvegarde le mapping si c'était un match, une création
      // ou un manual-override (Charles a corrigé). Sur 'mapped' on ne sauve pas
      // (déjà fait précédemment). Le scope est 'projet' si un projet est lié.
      if (suggestion.source !== 'mapped' && suggestion.projetId && userEmail) {
        try {
          await saveFolderMapping(userEmail, 'projet', suggestion.projetId, suggestion.folderPath, folderId);
        } catch (e) {
          console.warn('[IAPanel] saveFolderMapping failed (non-fatal):', e);
        }
      }
      // Si pas de projet lié, le déplacement physique du mail vers le dossier
      // sert lui-même de signal : la prochaine fois, findSenderPatternFolder
      // détectera que ce sender a déjà des mails dans ce dossier.

      return suggestion.folderPath;
    } catch (e) {
      console.warn('[IAPanel] tryMoveToHabitualFolder failed:', e);
      return '';
    }
  }

  /**
   * Charge la liste complète des dossiers Outlook (cache en mémoire).
   * Utilisé par le picker manuel.
   */
  private async loadAllFolders(): Promise<void> {
    if (this.allFoldersCache && this.allFoldersCache.length > 0) return;
    try {
      const { getApiContext } = await import('../api/graph');
      const ctx = await getApiContext();
      this.allFoldersCache = await this.scanAllFoldersRecursive(ctx.token);
    } catch (e) {
      console.warn('[IAPanel] loadAllFolders failed:', e);
      this.allFoldersCache = [];
      showToast(`Impossible de charger les dossiers : ${(e as Error).message?.slice(0, 80)}`, 'error');
    }
  }
}

function escapeHtml(s: string): string {
  return s.replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c] as string));
}
