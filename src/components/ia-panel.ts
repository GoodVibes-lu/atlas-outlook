/**
 * ia-panel.ts · onglet « Classer » du volet ATLAS dans Outlook.
 *
 *   • thème et urgence du mail (fiche de tri ATLAS), correction (ARGO apprend), reclassement ARGO ;
 *   • « Classer ce mail » : la même carte que l'onglet « Ce mail » (renderClasserOutlook : dossier du
 *     projet en un clic, création au chemin d'ATLAS, recherche dans tout l'arbre), par le worker ;
 *   • Traité / Demain 8 h / Archiver : par le worker (api/actions-mail.ts), comme les boutons du ruban.
 *
 * Lot 1 « vitesse » (09/10/2026) : l'ancien onglet passait par le jeton de boîte du complément
 * (liste des dossiers, déplacement, catégories, « Ranger la boîte ») que Microsoft ne donne plus : il
 * attendait puis échouait. Plus aucun accès direct à la boîte ici (règle rangementClicAutorise : le
 * worker range dans la boîte de la personne, sur son clic) ; le rangement automatique est fait par
 * l'agent du worker. L'analyse IA reste une action volontaire (bouton), jamais à l'ouverture.
 */

import { showToast } from '../taskpane';
import { humanError } from '../api/net';
import { escapeHtml } from '../utils/html';
import { icon } from '../ui/icons';
import { loadingHtml, emptyHtml, renderError } from '../ui/states';
import { correctTagCategory, upsertEmailTag, type EmailTag } from '../api/airtable';
import { analyzeEmailWithClaude, isAiAvailable } from '../api/claude';
import { convertToRestId } from '../api/graph';
import { archiverMail, ficheDeTri, marquerTraite, reporterDemain, type MailCourant } from '../api/actions-mail';
import { annulerAction } from '../api/agent';
import { renderClasserOutlook } from './agent-dossiers';

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


function mailCourant(): MailCourant | null {
  const item = Office.context.mailbox?.item as any;
  if (!item?.itemId) return null;
  let restId = String(item.itemId);
  try { restId = convertToRestId(restId); } catch { /* id brut */ }
  return { restId, messageId: String(item.internetMessageId || ''), conversationId: String(item.conversationId || '') };
}

export class IAPanel {
  private root: HTMLElement;
  private tag: EmailTag | null = null;
  private mail: MailCourant | null = mailCourant();
  private destroyed = false;

  constructor(root: HTMLElement) {
    this.root = root;
    this.renderLoading();
    void this.load();
  }

  destroy(): void { this.destroyed = true; }

  private renderLoading(): void {
    this.root.innerHTML = `<div class="panel-scroll">${loadingHtml('Lecture du classement ARGO…', 4)}</div>`;
  }

  private async load(): Promise<void> {
    if (!this.mail) { this.renderEmpty('Aucun mail sélectionné.'); return; }
    try {
      this.tag = await ficheDeTri(this.mail);
      if (this.destroyed) return;
      // Pas d'analyse IA à l'ouverture : l'agent du worker lit chaque mail une fois (onglet « Ce mail »).
      if (this.tag) this.render(); else this.renderNotTagged();
    } catch (e) {
      if (this.destroyed) return;
      this.root.innerHTML = '<div class="panel-scroll"><div id="ia-err"></div></div>';
      renderError(this.root.querySelector('#ia-err')!, e, () => { this.renderLoading(); void this.load(); }, { title: 'Classement indisponible' });
    }
  }

  private renderEmpty(msg: string): void {
    this.root.innerHTML = `<div class="panel-scroll">${emptyHtml({ icon: 'mail', title: msg, text: 'Sélectionne un mail dans ta boîte pour voir son classement.' })}</div>`;
  }

  /** Bloc commun : actions sur le mail et « Classer ce mail » (avec ou sans fiche de tri). */
  private actionsHtml(): string {
    return `
      <section class="section" aria-labelledby="ia-h-actions">
        <h2 class="section-heading" id="ia-h-actions">Que faire</h2>
        <div class="action-grid">
          <button type="button" data-action="done" class="btn btn-primary">${icon('check', 14)}Traité</button>
          <button type="button" data-action="snooze" class="btn btn-secondary">${icon('clock', 14)}Demain 8 h</button>
          <button type="button" data-action="archive" class="btn btn-secondary">${icon('archive', 14)}Archiver</button>
        </div>
        <div id="ia-res" class="tool-row-result" aria-live="polite" hidden></div>
      </section>
      <section id="ia-classer" class="section agent-section"></section>`;
  }

  private brancherCommun(): void {
    this.root.querySelectorAll<HTMLButtonElement>('button[data-action]').forEach(btn => {
      btn.addEventListener('click', () => void this.onAction(btn.dataset.action!, btn));
    });
    const classer = this.root.querySelector<HTMLElement>('#ia-classer');
    const m = this.mail;
    if (classer && m?.messageId) {
      const mailbox = String(Office.context.mailbox?.userProfile?.emailAddress || '');
      renderClasserOutlook(classer, { messageId: m.messageId, mailbox, conversationId: m.conversationId, onInfo: showToast, ...(this.tag?.linkedProjetId ? { projetId: this.tag.linkedProjetId } : {}) });
    }
  }

  private renderNotTagged(): void {
    const hasKey = isAiAvailable();
    this.root.innerHTML = `
      <div class="panel-scroll stack">
        ${emptyHtml({ icon: 'tag', title: 'Pas encore classé', text: 'ARGO n\'a pas encore lu ce mail. Lance le classement : thème et urgence en quelques secondes.' })}
        <button type="button" data-ia="analyze-now" class="btn btn-argo btn-block" ${hasKey ? '' : 'disabled'}>
          ${icon('bolt', 14)}<span>Classer avec ARGO</span><span class="argo-tag" aria-hidden="true">IA</span>
        </button>
        <div id="ia-progress" hidden></div>
        ${hasKey ? '' : '<p class="help">ARGO est indisponible pour le moment.</p>'}
        ${this.actionsHtml()}
      </div>
    `;
    const progress = this.root.querySelector<HTMLElement>('#ia-progress');
    this.root.querySelector<HTMLButtonElement>('button[data-ia="analyze-now"]')?.addEventListener('click', async (ev) => {
      const btn = ev.currentTarget as HTMLButtonElement;
      btn.disabled = true;
      if (progress) { progress.hidden = false; progress.innerHTML = '<div class="argo-progress" role="status"><span class="argo-progress-bar" aria-hidden="true"></span><span>ARGO lit le mail…</span></div>'; }
      const ok = await this.reanalyzeNow();
      if (ok) this.render();
      else { btn.disabled = false; if (progress) progress.hidden = true; }
    });
    this.brancherCommun();
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
        ${this.actionsHtml()}
        <details class="disclosure">
          <summary>${icon('chevron-right', 14)}Corriger le classement</summary>
          <div class="disclosure-body stack-sm">
            <label class="form-label" for="ia-cat-select">Le bon thème</label>
            <select id="ia-cat-select" class="form-input">
              ${CATEGORIES.map(cat => `<option value="${cat}" ${cat === tag.category ? 'selected' : ''}>${escapeHtml(CATEGORY_LABELS[cat] || cat)}</option>`).join('')}
            </select>
            <button type="button" data-ia="correct" class="btn btn-secondary btn-sm">${icon('check', 14)}Enregistrer, ARGO apprend</button>
            <button type="button" data-ia="reanalyze" class="btn btn-argo btn-sm" ${isAiAvailable() ? '' : 'disabled'}>${icon('bolt', 14)}<span>Reclasser avec ARGO</span><span class="argo-tag" aria-hidden="true">IA</span></button>
          </div>
        </details>
      </div>
    `;
    this.root.querySelector<HTMLButtonElement>('button[data-ia="correct"]')?.addEventListener('click', ev => void this.corriger(ev.currentTarget as HTMLButtonElement));
    this.root.querySelector<HTMLButtonElement>('button[data-ia="reanalyze"]')?.addEventListener('click', async ev => {
      const b = ev.currentTarget as HTMLButtonElement;
      b.disabled = true;
      if (await this.reanalyzeNow()) this.render(); else b.disabled = false;
    });
    this.brancherCommun();
  }

  /** Traité / Demain 8 h / Archiver, par le worker ; Annuler pour un archivage. */
  private async onAction(action: string, btn: HTMLButtonElement): Promise<void> {
    const m = this.mail;
    if (!m) return;
    const boutons = this.root.querySelectorAll<HTMLButtonElement>('button[data-action]');
    boutons.forEach(b => { b.disabled = true; });
    const res = this.root.querySelector<HTMLElement>('#ia-res');
    const lib = btn.innerHTML;
    btn.textContent = 'Un instant…';
    try {
      const faire = action === 'done' ? marquerTraite : action === 'snooze' ? reporterDemain : archiverMail;
      const r = await faire(m, this.tag);
      if (this.tag) this.tag.inboxStatus = action === 'done' ? 'done' : action === 'snooze' ? 'snoozed' : 'archived';
      showToast(r.message, 'success');
      if (res) {
        res.hidden = false;
        res.innerHTML = `<span>${icon('check-circle', 14)} ${escapeHtml(r.message)}</span>${r.actionId ? ' <button type="button" class="agent-link" data-annuler>Annuler</button>' : ''}`;
        const a = res.querySelector<HTMLButtonElement>('[data-annuler]');
        if (a && r.actionId) {
          const id = r.actionId;
          a.addEventListener('click', async () => {
            a.disabled = true;
            try { await annulerAction(id); a.replaceWith(Object.assign(document.createElement('span'), { className: 'agent-muted', textContent: 'Annulé : mail remis à sa place' })); }
            catch (e) { a.disabled = false; showToast(`Annulation impossible : ${humanError(e)}`, 'error'); }
          });
        }
      }
    } catch (e) {
      showToast(humanError(e), 'error');
    } finally {
      btn.innerHTML = lib;
      boutons.forEach(b => { b.disabled = false; });
    }
  }

  private async corriger(btn: HTMLButtonElement): Promise<void> {
    if (!this.tag) return;
    const newCat = this.root.querySelector<HTMLSelectElement>('#ia-cat-select')?.value || '';
    if (!newCat || newCat === this.tag.category) { showToast('Aucune modification', 'info'); return; }
    btn.disabled = true;
    const ok = await correctTagCategory(this.tag.id, newCat);
    btn.disabled = false;
    if (!ok) { showToast('Correction impossible pour le moment : réessaie.', 'error'); return; }
    this.tag.category = newCat;
    showToast(`Catégorie corrigée : ${CATEGORY_LABELS[newCat] || newCat}`, 'success');
    this.render();
  }

  /**
   * Analyse ARGO du mail ouvert (sur clic) : IA servie par le worker (Agency Brain), fiche de tri
   * remplacée. Lu par Office.js, aucun accès à la boîte.
   */
  private async reanalyzeNow(): Promise<boolean> {
    if (!isAiAvailable()) { showToast('ARGO est indisponible pour le moment.', 'error'); return false; }
    try {
      const item = Office.context.mailbox?.item as any;
      if (!item || !this.mail) { showToast('Aucun mail sélectionné', 'error'); return false; }
      const userEmail = Office.context.mailbox?.userProfile?.emailAddress || '';
      if (!userEmail) { showToast('Adresse de ta boîte introuvable : rouvre le panneau', 'error'); return false; }
      const personnes = (xs: any[]) => (xs || []).map((r: any) => ({ name: r.displayName || '', email: r.emailAddress || '' }));
      const subject: string = item.subject || '';
      const from = { name: item.from?.displayName || '', email: item.from?.emailAddress || '' };
      const receivedAt = item.dateTimeCreated ? new Date(item.dateTimeCreated).toISOString() : new Date().toISOString();
      const body: string = await new Promise((resolve) => {
        try {
          item.body.getAsync(Office.CoercionType.Text, (res: any) => {
            resolve(res?.status === Office.AsyncResultStatus.Succeeded ? (res.value || '') : '');
          });
        } catch { resolve(''); }
      });
      showToast('ARGO lit le mail…', 'info');
      const analysis = await analyzeEmailWithClaude({ subject, from, toRecipients: personnes(item.to), ccRecipients: personnes(item.cc), body, receivedAt, userEmail });
      const upserted = await upsertEmailTag({
        oldTagId: this.tag?.id, emailId: this.mail.restId, conversationId: this.mail.conversationId,
        subject, fromEmail: from.email, fromName: from.name, receivedAt,
        category: analysis.category, urgencyScore: analysis.urgencyScore, summary: analysis.summary,
        detectedLanguage: analysis.detectedLanguage, userEmail,
      });
      this.tag = {
        id: upserted.id, emailId: this.mail.restId, category: analysis.category, urgencyScore: analysis.urgencyScore,
        summary: analysis.summary, inboxStatus: 'inbox', linkedProjetId: this.tag?.linkedProjetId,
      };
      showToast(`Classé : ${CATEGORY_LABELS[analysis.category] || analysis.category}`, 'success');
      return true;
    } catch (e) {
      console.warn('[IAPanel] reanalyze failed:', e);
      showToast(humanError(e), 'error');
      return false;
    }
  }
}
