/**
 * agent-lists.ts — Listes de l'agent d'inbox, communes au panneau (bandeau « Ma journée », onglet
 * Agent) et au tableau de bord de la barre de gauche (tableau-de-bord.ts). Phase 2.
 *
 *  • Liste des mails d'une pile (GET /api/plugin/agent/liste) : clic = ouvre le mail ;
 *    pile « À filtrer » : boutons Accepter / Refuser l'expéditeur (POST /api/plugin/agent/expediteur).
 *  • Journal « Ce que l'agent a fait » (GET /api/plugin/agent/journal) avec Annuler
 *    (POST /api/plugin/agent/annuler).
 *  • Phase 3 : pile « Plus tard » (date de retour + « Remettre maintenant », DELETE
 *    /api/plugin/agent/plus-tard), mentions « brouillon prêt » et « relance prévue le … ».
 *  • Phase 5 (membres de good@) : listes « à attribuer » et « pour toi sur good@ »
 *    (`liste?pile=equipe_a_attribuer | equipe_pour_moi`), mentions « Pour X » / « Pris par X ».
 *
 * Ouverture d'un mail :
 *  - complément bureau / web : `displayMessageForm` (Mailbox 1.0) avec l'identifiant EWS obtenu par
 *    `convertToEwsId` (Mailbox 1.3, vérifié) ;
 *  - mobile (rien au-delà de Mailbox 1.5 sans garde, `displayMessageForm` n'y existe pas) et
 *    tableau de bord (pas d'Office.js) : `webLink` (Outlook sur le web) dans le navigateur.
 * Aucun appel IA.
 */

import {
  fetchListe, fetchJournal, annulerAction, decideExpediteur, retirerPlusTard,
  type AgentListePile, type AgentListeElement, type AgentJournalAction,
} from '../api/agent';
import { isMobile, openExternal, supportsMailbox } from '../api/platform';
import { escapeHtml } from '../utils/html';
import { humanError } from '../api/net';

export const PILE_LIBELLES: Record<AgentListePile, string> = {
  a_traiter: 'À traiter par toi',
  en_attente: 'En attente d\'une réponse',
  pour_info: 'Pour info',
  bruit: 'Bruit',
  a_filtrer: 'À filtrer (nouveaux expéditeurs)',
  plus_tard: 'Plus tard (répondre plus tard)',
  equipe_a_attribuer: 'good@ : à attribuer',
  equipe_pour_moi: 'good@ : pour toi',
};

export { escapeHtml };

/** Date courte à l'heure du Luxembourg : « 14:05 » aujourd'hui, sinon « 3 oct. ». */
export function dateCourte(value: string | undefined): string {
  if (!value) return '';
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return '';
  const tz = 'Europe/Luxembourg';
  const jour = (x: Date) => x.toLocaleDateString('fr-FR', { timeZone: tz });
  if (jour(d) === jour(new Date())) {
    return d.toLocaleTimeString('fr-FR', { timeZone: tz, hour: '2-digit', minute: '2-digit' });
  }
  return d.toLocaleDateString('fr-FR', { timeZone: tz, day: 'numeric', month: 'short' });
}

/**
 * Ouvre un mail de la liste. Renvoie false si rien n'a pu l'ouvrir (ni Office.js ni webLink).
 * `openLink` : ouverture d'une adresse propre à l'hôte (tableau de bord : Teams `app.openLink`).
 */
export function openAgentMail(el: Pick<AgentListeElement, 'graphId' | 'webLink'>, openLink?: (url: string) => void): boolean {
  let mailbox: Office.Mailbox | undefined;
  try { mailbox = typeof Office !== 'undefined' ? Office.context?.mailbox : undefined; } catch { mailbox = undefined; }
  if (mailbox && el.graphId && !isMobile() && supportsMailbox('1.3')) {
    try {
      const ewsId = mailbox.convertToEwsId(el.graphId, Office.MailboxEnums.RestVersion.v2_0);
      mailbox.displayMessageForm(ewsId);
      return true;
    } catch (e) {
      console.warn('[agent-lists] displayMessageForm impossible, repli sur le lien web :', e);
    }
  }
  if (el.webLink) {
    (openLink || openExternal)(el.webLink);
    return true;
  }
  return false;
}

// ── Liste d'une pile ──

export interface MailListOptions {
  limite?: number;
  /** Ouverture d'une adresse propre à l'hôte (tableau de bord). */
  openLink?: (url: string) => void;
  /** Appelé après une décision Accepter / Refuser (pour relire les compteurs). */
  onChange?: () => void;
  /** Message affiché si un mail ne peut pas être ouvert. */
  onInfo?: (message: string, type?: 'success' | 'error' | 'info') => void;
}

/** Charge et affiche les mails de `pile` dans `host`. */
export async function renderMailList(host: HTMLElement, pile: AgentListePile, opts: MailListOptions = {}): Promise<void> {
  host.innerHTML = '<div class="agent-loading"><div class="spinner"></div><span>Chargement de la liste…</span></div>';
  let elements: AgentListeElement[];
  try {
    elements = await fetchListe(pile, opts.limite ?? 50);
  } catch (e) {
    host.innerHTML = `<p class="agent-muted">Liste indisponible (${escapeHtml(humanError(e))}).</p>`;
    return;
  }
  if (!host.isConnected) return;
  if (!elements.length) {
    host.innerHTML = `<p class="agent-muted">Rien dans « ${escapeHtml(PILE_LIBELLES[pile] || pile)} ».</p>`;
    return;
  }
  const filtrer = pile === 'a_filtrer';
  const plusTard = pile === 'plus_tard';
  const jour = (v: string) => {
    const seul = /^\d{4}-\d{2}-\d{2}$/.test(v);
    const d = new Date(seul ? `${v}T12:00:00Z` : v);
    return Number.isNaN(d.getTime()) ? v : d.toLocaleDateString('fr-FR', { timeZone: seul ? 'UTC' : 'Europe/Luxembourg', weekday: 'short', day: 'numeric', month: 'short' });
  };
  const meta = (el: AgentListeElement): string => {
    const m: string[] = [];
    if (el.plusTardJusqua) m.push(`Revient ${jour(el.plusTardJusqua)}`);
    if (el.relanceLe) m.push(`Relance prévue ${jour(el.relanceLe)}`);
    else if (el.relanceGeree) m.push(el.relanceGeree);
    if (el.brouillonPret) m.push('Brouillon prêt');
    if (el.equipe?.prisPar) m.push(`Pris par ${el.equipe.prisParNom || el.equipe.prisPar}`);
    else if (el.equipe?.assigneA) m.push(`Pour ${el.equipe.assigneNom || el.equipe.assigneA}`);
    return m.length ? `<span class="mail-meta">${escapeHtml(m.join(' · '))}</span>` : '';
  };
  host.innerHTML = `
    <ul class="mail-list" aria-label="${escapeHtml(PILE_LIBELLES[pile] || pile)}">
      ${elements.map((el, i) => `
        <li class="mail-item${el.urgence >= 3 ? ' mail-urgent' : ''}" data-i="${i}">
          <button type="button" class="mail-open" data-open="${i}" title="Ouvrir le mail">
            <span class="mail-line1">
              <span class="mail-from">${escapeHtml(el.from?.name || el.from?.email || 'Expéditeur inconnu')}</span>
              <span class="mail-date">${escapeHtml(dateCourte(el.receivedAt))}</span>
            </span>
            <span class="mail-subject">${escapeHtml(el.subject || '(sans objet)')}</span>
            ${el.resume ? `<span class="mail-resume">${escapeHtml(el.resume)}</span>` : ''}
            ${meta(el)}
          </button>
          ${plusTard ? `
            <div class="mail-filtre-btns mail-plustard">
              <button type="button" class="btn btn-secondary agent-btn" data-remettre="${i}">Remettre maintenant</button>
            </div>` : ''}
          ${filtrer && el.from?.email ? `
            <div class="mail-filtre" data-email="${escapeHtml(el.from.email)}">
              <span class="agent-muted">${escapeHtml(el.from.email)}</span>
              <span class="mail-filtre-btns">
                <button type="button" class="btn btn-primary agent-btn" data-decision="accepter">Accepter</button>
                <button type="button" class="btn btn-secondary agent-btn" data-decision="refuser">Refuser</button>
              </span>
            </div>` : ''}
        </li>`).join('')}
    </ul>
  `;
  host.querySelectorAll<HTMLButtonElement>('button[data-open]').forEach(btn => {
    btn.addEventListener('click', () => {
      const el = elements[Number(btn.dataset.open)];
      if (el && !openAgentMail(el, opts.openLink)) {
        opts.onInfo?.('Ce mail ne peut pas être ouvert d\'ici : retrouve-le dans ta boîte.', 'info');
      }
    });
  });
  host.querySelectorAll<HTMLButtonElement>('button[data-remettre]').forEach(btn => {
    btn.addEventListener('click', async () => {
      const el = elements[Number(btn.dataset.remettre)];
      if (!el) return;
      btn.disabled = true;
      try {
        await retirerPlusTard(el.messageId);
        btn.closest('li')?.remove();
        if (!host.querySelector('li')) host.innerHTML = '<p class="agent-muted">Plus rien en « plus tard ».</p>';
        opts.onInfo?.('Mail remis dans « à traiter »', 'success');
        opts.onChange?.();
      } catch (e) {
        btn.disabled = false;
        opts.onInfo?.(`${humanError(e)}`, 'error');
      }
    });
  });
  host.querySelectorAll<HTMLButtonElement>('button[data-decision]').forEach(btn => {
    btn.addEventListener('click', async () => {
      const wrap = btn.closest<HTMLElement>('.mail-filtre');
      const email = wrap?.dataset.email || '';
      const decision = btn.dataset.decision === 'accepter' ? 'accepter' : 'refuser';
      if (!email || !wrap) return;
      wrap.querySelectorAll<HTMLButtonElement>('button').forEach(b => { b.disabled = true; });
      try {
        await decideExpediteur(email, decision);
        // Tous les mails de cet expéditeur quittent la liste.
        host.querySelectorAll<HTMLElement>('.mail-filtre').forEach(w => {
          if (w.dataset.email === email) w.closest('li')?.remove();
        });
        if (!host.querySelector('li')) host.innerHTML = '<p class="agent-muted">Plus aucun expéditeur à filtrer.</p>';
        opts.onInfo?.(decision === 'accepter' ? `${email} accepté` : `${email} refusé : ses mails iront au bruit`, 'success');
        opts.onChange?.();
      } catch (e) {
        wrap.querySelectorAll<HTMLButtonElement>('button').forEach(b => { b.disabled = false; });
        opts.onInfo?.(`${humanError(e)}`, 'error');
      }
    });
  });
}

// ── Journal « Ce que l'agent a fait » + Annuler ──

function mailDe(a: AgentJournalAction): string {
  const f = a.mail?.from;
  const from = typeof f === 'string' ? f : (f?.name || f?.email || '');
  const parts = [a.mail?.subject ? `« ${a.mail.subject} »` : '', from ? `de ${from}` : ''].filter(Boolean);
  return parts.join(' ');
}

/** Charge et affiche le journal de l'agent dans `host`, avec un bouton Annuler par action annulable. */
export async function renderJournal(host: HTMLElement, opts: { limite?: number; onChange?: () => void; onInfo?: MailListOptions['onInfo'] } = {}): Promise<void> {
  host.innerHTML = '<div class="agent-loading"><div class="spinner"></div><span>Lecture du journal…</span></div>';
  let actions: AgentJournalAction[];
  try {
    actions = await fetchJournal(opts.limite ?? 30);
  } catch (e) {
    host.innerHTML = `<p class="agent-muted">Journal indisponible (${escapeHtml(humanError(e))}).</p>`;
    return;
  }
  if (!host.isConnected) return;
  if (!actions.length) {
    host.innerHTML = '<p class="agent-muted">L\'agent n\'a encore rien fait (en mode à blanc, il observe sans agir).</p>';
    return;
  }
  host.innerHTML = `
    <ul class="journal-list">
      ${actions.map((a, i) => `
        <li class="journal-item">
          <div class="journal-text">
            <span class="journal-libelle">${escapeHtml(a.libelle || a.type)}</span>
            ${mailDe(a) ? `<span class="agent-muted">${escapeHtml(mailDe(a))}</span>` : ''}
            <span class="agent-muted">${escapeHtml(dateCourte(a.date))}</span>
          </div>
          ${a.annulable ? `<button type="button" class="btn btn-secondary journal-undo" data-undo="${i}">Annuler</button>` : ''}
        </li>`).join('')}
    </ul>
  `;
  host.querySelectorAll<HTMLButtonElement>('button[data-undo]').forEach(btn => {
    btn.addEventListener('click', async () => {
      const a = actions[Number(btn.dataset.undo)];
      if (!a) return;
      btn.disabled = true;
      btn.textContent = 'Annulation…';
      try {
        await annulerAction(a.id);
        btn.replaceWith(Object.assign(document.createElement('span'), { className: 'agent-muted', textContent: 'Annulé' }));
        opts.onInfo?.('Action annulée', 'success');
        opts.onChange?.();
      } catch (e) {
        btn.disabled = false;
        btn.textContent = 'Annuler';
        opts.onInfo?.(`${humanError(e)}`, 'error');
      }
    });
  });
}
