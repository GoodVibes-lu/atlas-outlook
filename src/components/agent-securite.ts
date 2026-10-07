/**
 * agent-securite.ts · BANDEAU DE SÉCURITÉ d'un mail (backlog reczJ0zLhPXqkWF9S, 07/10/2026), en tête
 * de l'onglet « Ce mail » et du détail du tableau de bord.
 *
 * Rien n'est affiché pour un mail à risque faible. Sinon : niveau (« À vérifier » / « Risque élevé »),
 * raisons lisibles, autres boîtes qui ont reçu le même mail, surveillance renforcée, et trois boutons :
 *  - « Mettre en quarantaine » : le worker range le mail dans « ATLAS · Suspect » (créé s'il manque),
 *    sur le clic de la propriétaire de la boîte, annulable ;
 *  - « Ce mail est sûr » : plus d'alerte pour ces signaux de cet expéditeur (dans cette boîte) ;
 *  - « Prévenir l'expéditeur » : SANS passer par sa boîte (peut-être piratée) : téléphone de la fiche
 *    ATLAS + texte à copier (aucun envoi de SMS depuis ATLAS), ou mail préparé pour un collègue de la
 *    même société (fenêtre de rédaction d'Outlook : la personne relit et envoie elle-même).
 * Nouvel Outlook pour Mac : seules API utilisées : displayNewMessageForm (Mailbox 1.6, gardé par
 * supportsMailbox) ; repli : texte copié et lien mailto.
 */
import { fetchSecurite, fetchPrevenir, marquerSur, mettreEnQuarantaine, type Prevenir, type SecuriteMail } from '../api/securite';
import { annulerAction } from '../api/agent';
import { copierTexte } from './agent-outils';
import { escapeHtml } from '../utils/html';
import { humanError } from '../api/net';
import { icon } from '../ui/icons';
import { openExternal, supportsMailbox } from '../api/platform';

type InfoFn = (message: string, type?: 'success' | 'error' | 'info') => void;

export interface CtxSecurite {
  messageId: string;
  mailbox?: string;
  onInfo?: InfoFn;
  /** Prévenu quand le mail est (ou n'est plus) protégé : l'appelant peut masquer ses actions automatiques. */
  onProtege?: (protege: boolean) => void;
}

const dateCourte = (iso?: string) => {
  if (!iso) return '';
  const d = new Date(iso);
  return Number.isFinite(d.getTime()) ? new Intl.DateTimeFormat('fr-FR', { timeZone: 'Europe/Luxembourg', day: '2-digit', month: '2-digit', year: 'numeric' }).format(d) : '';
};

function peutRediger(): boolean {
  try {
    const mb = (globalThis as any).Office?.context?.mailbox;
    return !!mb && typeof mb.displayNewMessageForm === 'function' && supportsMailbox('1.6');
  } catch { return false; }
}

function bandeauHtml(s: SecuriteMail, quarantaineDisponible: boolean): string {
  const eleve = s.niveau === 'eleve';
  const autres = s.autresBoites?.length ? `<div class="secu-note">${icon('team', 14)}Reçu aussi par ${s.autresBoites.map(escapeHtml).join(', ')} : chacun est prévenu.</div>` : '';
  const surv = s.surveillanceJusqua ? `<div class="secu-note">${icon('eye', 14)}Expéditeur sous surveillance renforcée jusqu'au ${escapeHtml(dateCourte(s.surveillanceJusqua))}.</div>` : '';
  const protege = s.protege ? '<div class="secu-note">N\'ouvre aucun lien ni aucune pièce jointe. ATLAS ne fait rien d\'automatique sur ce mail (ni dépôt, ni brouillon, ni classement).</div>' : '';
  const quarantaine = s.quarantaine
    ? `<div class="secu-note">${icon('check', 14)}Mis en quarantaine le ${escapeHtml(dateCourte(s.quarantaine.le))}.</div>`
    : '';
  return `
    <div class="secu-bandeau ${eleve ? 'is-eleve' : 'is-verifier'}" role="alert">
      <div class="secu-titre">${icon('alert', 16)}${eleve ? 'Mail à risque élevé' : 'Mail à vérifier'}</div>
      <ul class="secu-raisons">${s.raisons.map(r => `<li>${escapeHtml(r)}</li>`).join('')}</ul>
      ${autres}${surv}${protege}${quarantaine}
      <div class="secu-btns">
        ${s.quarantaine ? '' : `<button type="button" class="btn btn-danger agent-btn" data-secu="quarantaine" ${quarantaineDisponible ? '' : 'disabled title="Seule la propriétaire de la boîte peut mettre ce mail en quarantaine"'}>${icon('lock', 14)}Mettre en quarantaine</button>`}
        <button type="button" class="btn btn-secondary agent-btn" data-secu="sur">${icon('check', 14)}Ce mail est sûr</button>
        <button type="button" class="btn btn-secondary agent-btn" data-secu="prevenir">${icon('user', 14)}Prévenir l'expéditeur</button>
      </div>
      <div class="secu-res" data-secu-res hidden></div>
      <div class="secu-prevenir" data-secu-prevenir hidden></div>
    </div>`;
}

/** Bandeau de sécurité du mail (masqué si le risque est faible). Ne lève jamais. */
export async function renderSecurite(host: HTMLElement, ctx: CtxSecurite): Promise<void> {
  host.hidden = true;
  host.innerHTML = '';
  let r;
  try { r = await fetchSecurite(ctx.messageId, ctx.mailbox); }
  catch { return; } // sécurité indisponible : pas de bandeau (le reste du panneau fonctionne)
  const s = r.securite;
  ctx.onProtege?.(!!r.protege);
  if (!s || s.niveau === 'faible') return;
  const mailbox = r.mailbox || ctx.mailbox;
  host.hidden = false;
  host.innerHTML = bandeauHtml(s, r.quarantaineDisponible);
  const res = host.querySelector<HTMLElement>('[data-secu-res]')!;
  const zone = host.querySelector<HTMLElement>('[data-secu-prevenir]')!;
  const bouton = (k: string) => host.querySelector<HTMLButtonElement>(`[data-secu="${k}"]`);

  bouton('quarantaine')?.addEventListener('click', async () => {
    const b = bouton('quarantaine')!;
    b.disabled = true;
    b.textContent = 'Quarantaine…';
    try {
      const q = await mettreEnQuarantaine(ctx.messageId, mailbox);
      res.hidden = false;
      res.innerHTML = `<div>${icon('check', 14)}Rangé dans « ${escapeHtml(q.dossier.chemin)} ».${q.surveillanceJusqua ? ` Expéditeur sous surveillance jusqu'au ${escapeHtml(dateCourte(q.surveillanceJusqua))}.` : ''}</div>
        <button type="button" class="agent-link" data-secu-annuler>Annuler</button>`;
      b.remove();
      ctx.onInfo?.('Mail mis en quarantaine', 'success');
      res.querySelector<HTMLButtonElement>('[data-secu-annuler]')?.addEventListener('click', async ev => {
        const a = ev.currentTarget as HTMLButtonElement;
        a.disabled = true;
        try { await annulerAction(q.actionId); a.replaceWith(Object.assign(document.createElement('span'), { className: 'agent-muted', textContent: 'Annulé : mail remis à sa place' })); ctx.onInfo?.('Quarantaine annulée', 'success'); }
        catch (e) { a.disabled = false; ctx.onInfo?.(`Annulation impossible : ${humanError(e)}`, 'error'); }
      });
    } catch (e) {
      b.disabled = false;
      b.innerHTML = `${icon('lock', 14)}Mettre en quarantaine`;
      ctx.onInfo?.(humanError(e, 'Quarantaine impossible pour le moment.'), 'error');
    }
  });

  bouton('sur')?.addEventListener('click', async () => {
    const b = bouton('sur')!;
    b.disabled = true;
    try {
      await marquerSur(ctx.messageId, mailbox);
      host.innerHTML = `<div class="secu-ok">${icon('check', 14)}Marqué sûr : plus d'alerte pour ces signaux de cet expéditeur.</div>`;
      ctx.onProtege?.(false);
      ctx.onInfo?.('Mail marqué sûr', 'success');
    } catch (e) {
      b.disabled = false;
      ctx.onInfo?.(humanError(e), 'error');
    }
  });

  bouton('prevenir')?.addEventListener('click', async () => {
    if (!zone.hidden) { zone.hidden = true; return; }
    zone.hidden = false;
    zone.innerHTML = '<div class="agent-loading"><div class="spinner"></div><span>Coordonnées dans ATLAS…</span></div>';
    try { renderPrevenir(zone, await fetchPrevenir(ctx.messageId, mailbox), ctx.onInfo); }
    catch (e) { zone.innerHTML = `<div class="agent-muted">${escapeHtml(humanError(e))}</div>`; }
  });
}

function renderPrevenir(zone: HTMLElement, p: Prevenir, onInfo?: InfoFn): void {
  const tels = p.telephones.map((t, i) => `
    <div class="secu-ligne"><span>${escapeHtml(t.libelle)} : <strong>${escapeHtml(t.numero)}</strong></span>
      <button type="button" class="agent-link" data-copier-tel="${i}">${icon('copy', 14)}Copier</button></div>`).join('');
  const coll = p.collegues.map((c, i) => `
    <div class="secu-ligne"><span>${escapeHtml(c.nom)} <span class="agent-muted">${escapeHtml(c.email)}</span></span>
      <button type="button" class="agent-link" data-mail-collegue="${i}">${icon('mail', 14)}Préparer le mail</button></div>`).join('');
  zone.innerHTML = `
    <div class="secu-sous-titre">Prévenir ${escapeHtml(p.contact.nom)}${p.contact.societe ? ` (${escapeHtml(p.contact.societe)})` : ''} sans passer par son adresse</div>
    <div class="agent-muted">Sa boîte est peut-être piratée : ne réponds pas à ce mail.</div>
    ${tels ? `<div class="secu-bloc"><div class="secu-label">Par téléphone</div>${tels}</div>` : ''}
    <div class="secu-bloc">
      <div class="secu-label">Texte à lire ou à envoyer par SMS depuis ton téléphone</div>
      <textarea class="agent-input agent-texte" rows="6" readonly data-texte-tel>${escapeHtml(p.texteTelephone)}</textarea>
      <button type="button" class="btn btn-secondary agent-btn" data-copier-texte>${icon('copy', 14)}Copier le texte</button>
    </div>
    ${coll ? `<div class="secu-bloc"><div class="secu-label">Par un collègue de la même société</div>${coll}</div>` : ''}
    ${!tels && !coll ? '<div class="agent-muted">Aucun téléphone ni collègue dans ATLAS : appelle un numéro que tu connais déjà, jamais celui donné dans le mail.</div>' : ''}`;
  zone.querySelectorAll<HTMLButtonElement>('[data-copier-tel]').forEach(b => b.addEventListener('click', async () => {
    const t = p.telephones[Number(b.dataset.copierTel)];
    const ok = t ? await copierTexte(t.numero) : false;
    onInfo?.(ok ? 'Numéro copié' : 'Copie impossible', ok ? 'success' : 'error');
  }));
  zone.querySelector<HTMLButtonElement>('[data-copier-texte]')?.addEventListener('click', async () => {
    const ok = await copierTexte(p.texteTelephone, zone.querySelector<HTMLTextAreaElement>('[data-texte-tel]'));
    onInfo?.(ok ? 'Texte copié' : 'Copie impossible : sélectionne le texte et copie-le', ok ? 'success' : 'error');
  });
  zone.querySelectorAll<HTMLButtonElement>('[data-mail-collegue]').forEach(b => b.addEventListener('click', async () => {
    const c = p.collegues[Number(b.dataset.mailCollegue)];
    if (!c) return;
    if (peutRediger()) {
      // Fenêtre de rédaction d'Outlook pré-remplie : rien n'est envoyé, la personne relit et envoie.
      try {
        (globalThis as any).Office.context.mailbox.displayNewMessageForm({ toRecipients: [c.email], subject: c.objet, htmlBody: c.html });
        onInfo?.('Mail préparé : relis-le puis envoie-le', 'success');
        return;
      } catch { /* repli ci-dessous */ }
    }
    const ok = await copierTexte(c.texte);
    openExternal(`mailto:${encodeURIComponent(c.email)}?subject=${encodeURIComponent(c.objet)}`);
    onInfo?.(ok ? 'Texte du mail copié : colle-le dans le nouveau message' : 'Ouvre un nouveau message et recopie le texte', ok ? 'success' : 'info');
  }));
}
