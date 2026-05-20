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
  source: 'mapped' | 'sender-pattern' | 'match' | 'create' | 'none';
  /** ID Outlook si dossier existant. */
  folderId?: string;
  /** Chemin lisible (ex: "Markcom/Creativity Camp"). */
  folderPath: string;
  /** Projet lié (utilisé pour la sauvegarde du mapping). */
  projetId?: string;
  /** Pour 'sender-pattern' : nombre de mails du sender dans ce dossier (preuve). */
  patternMailCount?: number;
}

export class IAPanel {
  private root: HTMLElement;
  private tag: EmailTag | null = null;
  private emailId = '';
  private linkedProjet: Projet | null = null;
  private folderSuggestion: FolderSuggestion | null = null;

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
    this.root.innerHTML = `
      <div style="padding: 16px;">
        <div style="padding: 12px; background: #fff7ed; border: 1px solid #fed7aa; border-radius: 8px; color: #9a3412;">
          <strong>Pas encore taggé par l'IA</strong>
          <p style="margin: 6px 0 0; font-size: 12px;">L'inbound scanner ATLAS (côté serveur) n'a pas encore analysé ce mail. Cela peut prendre 1-2 min après réception. Reviens dans un instant.</p>
        </div>
      </div>
    `;
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
          ok = await this.reanalyzeNow();
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
      const item = Office.context.mailbox?.item;
      if (!item) return false;
      const userEmail = Office.context.mailbox?.userProfile?.emailAddress || '';
      const ewsId = (item as any).itemId;
      if (!ewsId || !userEmail) return false;

      // 1. Fetch le mail complet via Graph (corps + recipients)
      const { getGraphToken, convertToRestId } = await import('../api/graph');
      const token = await getGraphToken();
      const restId = convertToRestId(ewsId);
      const msg = await getMessageForLinking(token, restId);

      showToast('Analyse Claude en cours…', 'info');

      // 2. Appelle Claude
      const analysis = await analyzeEmailWithClaude({
        subject: msg.subject,
        from: { name: msg.from.name, email: msg.from.email },
        toRecipients: (msg.toRecipients || []).map((r) => ({ name: r.name, email: r.email })),
        ccRecipients: (msg.ccRecipients || []).map((r) => ({ name: r.name, email: r.email })),
        body: msg.bodyText || msg.bodyPreview || '',
        receivedAt: msg.receivedAt,
        userEmail,
      });

      // 3. Upsert le tag dans Airtable (DELETE ancien + CREATE nouveau)
      const upserted = await upsertEmailTag({
        oldTagId: this.tag?.id,
        emailId: msg.id,
        conversationId: msg.conversationId || '',
        subject: msg.subject,
        fromEmail: msg.from.email,
        fromName: msg.from.name,
        receivedAt: msg.receivedAt,
        category: analysis.category,
        urgencyScore: analysis.urgencyScore,
        summary: analysis.summary,
        detectedLanguage: analysis.detectedLanguage,
        userEmail,
      });

      // 4. Refresh local
      this.tag = {
        id: upserted.id,
        emailId: msg.id,
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
    if (!s || s.source === 'none') {
      // Pas de projet lié → pas de suggestion. Section silencieuse.
      return this.tag?.linkedProjetId
        ? `<div style="padding: 10px; background: #f8fafc; border: 1px dashed #cbd5e1; border-radius: 6px; font-size: 11px; color: #64748b;">📁 Résolution du dossier en cours…</div>`
        : '';
    }
    if (s.source === 'mapped') {
      return `
        <div style="padding: 10px; background: #ecfdf5; border: 1px solid #6ee7b7; border-radius: 6px;">
          <div style="font-size: 10px; font-weight: 700; color: #047857; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 4px;">📁 Dossier habituel (appris)</div>
          <div style="font-size: 13px; font-weight: 600; color: #065f46;">${escapeHtml(s.folderPath)} ✓</div>
          <div style="font-size: 11px; color: #047857; margin-top: 2px;">Sera utilisé automatiquement au clic sur Traité ou Archiver.</div>
        </div>
      `;
    }
    if (s.source === 'sender-pattern') {
      return `
        <div style="padding: 10px; background: #f5f3ff; border: 1px solid #c4b5fd; border-radius: 6px;">
          <div style="font-size: 10px; font-weight: 700; color: #6d28d9; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 4px;">📁 Tu ranges déjà ces mails ici</div>
          <div style="font-size: 13px; font-weight: 600; color: #5b21b6;">${escapeHtml(s.folderPath)}</div>
          <div style="font-size: 11px; color: #6d28d9; margin-top: 4px;">Observé : ${s.patternMailCount} mail${(s.patternMailCount || 0) > 1 ? 's' : ''} de cet expéditeur déjà dans ce dossier. ATLAS suit ton habitude — au clic Traité/Archiver, il s'en souvient pour de bon.</div>
        </div>
      `;
    }
    if (s.source === 'match') {
      return `
        <div style="padding: 10px; background: #eff6ff; border: 1px solid #93c5fd; border-radius: 6px;">
          <div style="font-size: 10px; font-weight: 700; color: #1d4ed8; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 4px;">📁 Dossier suggéré (match par nom)</div>
          <div style="font-size: 13px; font-weight: 600; color: #1e3a8a;">${escapeHtml(s.folderPath)}</div>
          <div style="font-size: 11px; color: #1d4ed8; margin-top: 4px;">Si tu cliques Traité ou Archiver, le mail ira là et ATLAS s'en souviendra pour les prochains mails de ce projet.</div>
        </div>
      `;
    }
    // source === 'create'
    return `
      <div style="padding: 10px; background: #fef3c7; border: 1px solid #fcd34d; border-radius: 6px;">
        <div style="font-size: 10px; font-weight: 700; color: #92400e; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 4px;">📁 Nouveau dossier suggéré</div>
        <div style="font-size: 13px; font-weight: 600; color: #78350f;">${escapeHtml(s.folderPath)}</div>
        <div style="font-size: 11px; color: #92400e; margin-top: 4px;">Aucun dossier Outlook existant ne match. Au clic sur Traité ou Archiver, ATLAS le crée et mémorise la règle.</div>
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
      if (senderEmail) {
        try {
          const token = await getGraphToken();
          const pattern = await this.findSenderPatternFolder(token, senderEmail);
          if (pattern) {
            this.folderSuggestion = {
              source: 'sender-pattern',
              folderId: pattern.folderId,
              folderPath: pattern.folderPath,
              projetId: tag?.linkedProjetId,
              patternMailCount: pattern.count,
            };
            return;
          }
        } catch (e) {
          console.warn('[IAPanel] sender-pattern lookup failed:', e);
        }
      }

      // ── Stratégies 3 + 4 : nécessitent un projet lié (client + dénomination)
      if (!tag?.linkedProjetId) {
        this.folderSuggestion = { source: 'none', folderPath: '' };
        return;
      }
      const projets = await getAllProjets();
      const projet = projets.find((p) => p.id === tag.linkedProjetId);
      this.linkedProjet = projet || null;
      if (!projet) {
        this.folderSuggestion = { source: 'none', folderPath: '' };
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
  private async findSenderPatternFolder(
    token: string,
    senderEmail: string,
  ): Promise<{ folderId: string; folderPath: string; count: number } | null> {
    // Filtre Graph : from = senderEmail. On exclut côté JS les mails encore en Inbox.
    const filter = encodeURIComponent(`from/emailAddress/address eq '${senderEmail.replace(/'/g, "''")}'`);
    const url = `https://graph.microsoft.com/v1.0/me/messages?$filter=${filter}&$top=30&$select=id,parentFolderId,subject&$orderby=receivedDateTime desc`;
    const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (!res.ok) return null;
    const data: any = await res.json();
    const msgs = (data.value || []) as Array<{ parentFolderId: string }>;
    if (msgs.length < 3) return null;

    // Récupère l'ID du dossier Inbox pour l'exclure
    const inboxRes = await fetch('https://graph.microsoft.com/v1.0/me/mailFolders/inbox?$select=id', {
      headers: { Authorization: `Bearer ${token}` },
    });
    const inboxId = inboxRes.ok ? (await inboxRes.json()).id : '';

    // Compte par dossier (hors Inbox)
    const counts = new Map<string, number>();
    for (const m of msgs) {
      if (!m.parentFolderId || m.parentFolderId === inboxId) continue;
      counts.set(m.parentFolderId, (counts.get(m.parentFolderId) || 0) + 1);
    }
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
      const fres = await fetch(`https://graph.microsoft.com/v1.0/me/mailFolders/${bestId}?$select=displayName,parentFolderId`, {
        headers: { Authorization: `Bearer ${token}` },
      });
      if (fres.ok) {
        const f: any = await fres.json();
        folderPath = f.displayName || folderPath;
        // Tente de récupérer le parent pour afficher "Parent/Folder"
        if (f.parentFolderId && f.parentFolderId !== inboxId) {
          const pres = await fetch(`https://graph.microsoft.com/v1.0/me/mailFolders/${f.parentFolderId}?$select=displayName`, {
            headers: { Authorization: `Bearer ${token}` },
          });
          if (pres.ok) {
            const p: any = await pres.json();
            if (p.displayName && !['Top of Information Store', 'Inbox'].includes(p.displayName)) {
              folderPath = `${p.displayName}/${f.displayName}`;
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
      const tag = this.tag;
      if (!suggestion || suggestion.source === 'none' || !tag) return '';

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

      // Apprentissage : sauvegarde le mapping si c'était un match ou une création.
      // Sur 'mapped' on ne sauve pas (déjà fait).
      if (suggestion.source !== 'mapped' && suggestion.projetId && userEmail) {
        try {
          await saveFolderMapping(userEmail, 'projet', suggestion.projetId, suggestion.folderPath, folderId);
        } catch (e) {
          console.warn('[IAPanel] saveFolderMapping failed (non-fatal):', e);
        }
      }
      return suggestion.folderPath;
    } catch (e) {
      console.warn('[IAPanel] tryMoveToHabitualFolder failed:', e);
      return '';
    }
  }
}

function escapeHtml(s: string): string {
  return s.replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c] as string));
}
