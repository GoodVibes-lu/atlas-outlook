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
  markTagDone,
  snoozeTag,
  archiveTag,
  correctTagCategory,
  type EmailTag,
} from '../api/airtable';
import { getGraphToken, moveMessageToFolder, convertToRestId } from '../api/graph';

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

export class IAPanel {
  private root: HTMLElement;
  private tag: EmailTag | null = null;
  private emailId = '';

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
        this.render();
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

        <!-- Actions principales -->
        <div style="display: flex; flex-direction: column; gap: 8px;">
          <div style="font-size: 10px; font-weight: 700; color: #475569; text-transform: uppercase; letter-spacing: 0.5px;">Actions</div>
          <div style="display: flex; gap: 6px; flex-wrap: wrap;">
            <button data-action="done" class="ia-btn ia-btn-success">✓ Traité</button>
            <button data-action="snooze" class="ia-btn ia-btn-warning">⏰ Reporter (demain 8h)</button>
            <button data-action="archive" class="ia-btn ia-btn-muted">📦 Archiver</button>
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
   * Déplace le mail courant dans le dossier habituel appris pour le projet lié.
   * Retourne le nom du dossier si déplacement effectué, '' sinon.
   *
   * Gating : ne touche QUE si on a (a) un projet lié au tag, (b) un folder
   * mapping appris pour ce user+projet. Sinon ne déplace rien — le mail reste
   * en inbox Outlook et c'est OK (l'utilisateur peut filer manuellement).
   */
  private async tryMoveToHabitualFolder(_actionName: string): Promise<string> {
    try {
      const tag = this.tag;
      if (!tag?.linkedProjetId) return '';
      const userEmail = Office.context.mailbox?.userProfile?.emailAddress || '';
      if (!userEmail) return '';
      const mapping = await getFolderMapping(userEmail, 'projet', tag.linkedProjetId);
      if (!mapping?.folderId) return '';
      const item = Office.context.mailbox?.item;
      const rawId = (item as any)?.itemId;
      if (!rawId) return '';
      // Outlook côté addin renvoie un EWS itemId — Graph veut un REST ID.
      const restId = convertToRestId(rawId);
      const token = await getGraphToken();
      await moveMessageToFolder(token, restId, mapping.folderId);
      return mapping.folderPath || mapping.folderId.slice(0, 8);
    } catch (e) {
      console.warn('[IAPanel] tryMoveToHabitualFolder failed:', e);
      return '';
    }
  }
}

function escapeHtml(s: string): string {
  return s.replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c] as string));
}
