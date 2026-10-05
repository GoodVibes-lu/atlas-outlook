/**
 * quick-draft.ts — Onglet « ⚡ Rapide » du panneau RÉDACTION (bureau et web uniquement).
 *
 * Reprend les fonctions de rédaction du complément « ATLAS Assistant » (outlook-plugin, remplacé) :
 *   • Réponse rapide : 4 modèles fixes × FR / EN / DE × tu / vous (route worker
 *     /api/plugin/email/quick-reply, SANS IA) ;
 *   • Brouillon automatique : réponse contextuelle rédigée par le worker (route
 *     /api/plugin/email/auto-draft, IA par l'Agency Brain), sur clic uniquement.
 * Le texte est inséré EN TÊTE du message en cours (le fil cité est conservé) ; rien n'est envoyé.
 * Jamais de signature (Exclaimer l'ajoute à l'envoi).
 */

import { fetchQuickReply, fetchAutoDraft, type QuickReplyType } from '../api/agent';
import { showToast } from '../taskpane';
import { humanError } from '../api/net';
import { icon } from '../ui/icons';

const QUICK_BUTTONS: Array<{ type: QuickReplyType; label: string; icon: string }> = [
  { type: 'acknowledge', label: 'Bien reçu', icon: 'check' },
  { type: 'quote_pending', label: 'Le devis suit', icon: 'template' },
  { type: 'polite_refusal', label: 'Refus poli', icon: 'x' },
  { type: 'reschedule', label: 'Reporter', icon: 'calendar-plus' },
];

function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** Texte brut → HTML simple (paragraphes / retours à la ligne). Laisse passer un HTML déjà formé. */
function textToHtml(text: string): string {
  if (/<(p|div|br)\b/i.test(text)) return text;
  return text.split(/\n{2,}/).map(p => `<p>${escapeHtml(p).replace(/\n/g, '<br/>')}</p>`).join('');
}

function getAsyncValue<T>(getter: ((cb: (r: Office.AsyncResult<T>) => void) => void) | undefined, fallback: T): Promise<T> {
  return new Promise((resolve) => {
    if (!getter) { resolve(fallback); return; }
    try {
      getter((r) => resolve(r.status === Office.AsyncResultStatus.Succeeded ? (r.value ?? fallback) : fallback));
    } catch { resolve(fallback); }
  });
}

export class QuickDraftPanel {
  private container: HTMLElement;
  private busy = false;

  constructor(container: HTMLElement) {
    this.container = container;
    this.render();
  }

  destroy(): void {
    this.container.innerHTML = '';
  }

  private render(): void {
    this.container.innerHTML = `
      <div class="panel-scroll stack">
        <section class="section" aria-labelledby="qd-h-quick">
          <h2 class="section-heading" id="qd-h-quick">Réponse en un clic</h2>
          <p class="help">Un texte type est inséré en tête du message. Sans ARGO, gratuit.</p>
          <div class="quick-grid">
            ${QUICK_BUTTONS.map(b => `<button type="button" class="btn btn-secondary agent-btn quick-btn" data-type="${b.type}">${icon(b.icon, 14)}${b.label}</button>`).join('')}
          </div>
          <div class="quick-options">
            <label>Langue
              <select class="dropdown-select" id="quick-lang">
                <option value="FR">Français</option>
                <option value="EN">English</option>
                <option value="DE">Deutsch</option>
              </select>
            </label>
            <label>Ton
              <select class="dropdown-select" id="quick-address">
                <option value="vous">Vouvoiement</option>
                <option value="tu">Tutoiement</option>
              </select>
            </label>
          </div>
        </section>

        <section class="section" aria-labelledby="qd-h-auto">
          <h2 class="section-heading" id="qd-h-auto">Brouillon complet par ARGO</h2>
          <p class="help">ARGO rédige une réponse d'après le fil (langue et tutoiement repris). Rien n'est envoyé : tu relis avant.</p>
          <button type="button" class="btn btn-argo btn-block agent-btn" id="auto-draft-btn">${icon('bolt', 14)}<span>Rédiger avec ARGO</span><span class="argo-tag" aria-hidden="true">IA</span></button>
        </section>
      </div>
    `;
    this.container.querySelectorAll<HTMLButtonElement>('.quick-btn').forEach(btn => {
      btn.addEventListener('click', () => this.onQuickReply(btn.dataset.type as QuickReplyType));
    });
    this.container.querySelector('#auto-draft-btn')?.addEventListener('click', () => this.onAutoDraft());
  }

  private item(): any {
    return Office.context.mailbox?.item as any;
  }

  private async firstRecipient(): Promise<{ displayName: string; emailAddress: string }> {
    const item = this.item();
    const list = await getAsyncValue<any[]>(item?.to?.getAsync?.bind(item.to), []);
    return list[0] || { displayName: '', emailAddress: '' };
  }

  /** Insère en tête du corps, au format du message (HTML ou texte). */
  private async insertAtTop(text: string): Promise<void> {
    const item = this.item();
    if (!item?.body?.prependAsync) throw new Error('Insertion impossible dans ce message.');
    const type = await getAsyncValue<string>(item.body.getTypeAsync?.bind(item.body), 'html');
    const isHtml = String(type).toLowerCase() === 'html';
    await new Promise<void>((resolve, reject) => {
      item.body.prependAsync(
        isHtml ? textToHtml(text) : `${text}\n\n`,
        { coercionType: isHtml ? Office.CoercionType.Html : Office.CoercionType.Text },
        (r: Office.AsyncResult<void>) => (r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(new Error(r.error?.message || 'insertion échouée'))),
      );
    });
  }

  private async onQuickReply(type: QuickReplyType): Promise<void> {
    if (this.busy) return;
    this.busy = true;
    try {
      const recipient = await this.firstRecipient();
      const lang = (this.container.querySelector<HTMLSelectElement>('#quick-lang')?.value || 'FR') as 'FR' | 'EN' | 'DE';
      const address = (this.container.querySelector<HTMLSelectElement>('#quick-address')?.value || 'vous') as 'tu' | 'vous';
      const firstName = (recipient.displayName || '').split(/\s+/)[0] || '';
      const body = await fetchQuickReply({ type, language: lang, address, recipientName: firstName });
      if (!body) throw new Error('modèle vide');
      await this.insertAtTop(body);
      showToast('Réponse rapide insérée', 'success');
    } catch (err) {
      showToast(`${humanError(err)}`, 'error');
    } finally {
      this.busy = false;
    }
  }

  private async onAutoDraft(): Promise<void> {
    if (this.busy) return;
    const btn = this.container.querySelector<HTMLButtonElement>('#auto-draft-btn');
    this.busy = true;
    if (btn) { btn.disabled = true; btn.innerHTML = `${icon('bolt', 14)}<span>ARGO rédige…</span>`; }
    try {
      const item = this.item();
      const subject = await getAsyncValue<string>(item?.subject?.getAsync?.bind(item.subject), '');
      const recipient = await this.firstRecipient();
      // Le corps en cours contient le fil cité : il sert de contexte au worker.
      const thread = await getAsyncValue<string>(
        item?.body?.getAsync ? (cb: any) => item.body.getAsync(Office.CoercionType.Text, cb) : undefined,
        '',
      );
      const draft = await fetchAutoDraft({
        email: {
          id: String(item?.itemId || 'compose-new'),
          subject: subject.replace(/^(re|tr|aw|fw|fwd)\s*:\s*/i, ''),
          from: { name: recipient.displayName || '', email: recipient.emailAddress || '' },
          receivedAt: new Date().toISOString(),
          bodyPreview: thread.slice(0, 500),
        },
        fullBody: thread.slice(0, 8000) || undefined,
      });
      if (!draft?.body) { showToast('Pas de brouillon généré', 'info'); return; }
      await this.insertAtTop(draft.body);
      showToast('Brouillon inséré : relis-le avant d\'envoyer', 'success');
    } catch (err) {
      showToast(`${humanError(err)}`, 'error');
    } finally {
      this.busy = false;
      if (btn) { btn.disabled = false; btn.innerHTML = `${icon('bolt', 14)}<span>Rédiger avec ARGO</span><span class="argo-tag" aria-hidden="true">IA</span>`; }
    }
  }
}
