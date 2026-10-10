/**
 * reponse-prete.ts · « Réponse prête » dans le volet du mail (lot 5 du complément, 10/10/2026).
 *
 * L'agent a rédigé la réponse à l'arrivée du mail, dans ton style (profil de conversation, tu / vous
 * selon le profil ARGO, prénom seul en tutoiement, sans signature : Exclaimer l'ajoute à l'envoi).
 * Une seule relecture, ici :
 *   - le texte est modifiable ; les réponses courtes le remplacent d'un clic ;
 *   - « Envoyer » : le worker prépare la réponse dans ta boîte, tu confirmes « Envoyer à X ? », le mail
 *     part en quelques secondes et « Annuler » le retient tant qu'il n'est pas parti (même route que
 *     le tableau de bord) ;
 *   - « Ouvrir dans Outlook » : formulaire de réponse de l'application, prérempli (rien n'est envoyé).
 * Rien ne part sans ton clic.
 */
import { abandonnerEnvoi, confirmerEnvoi, preparerEnvoi } from '../api/tableau';
import { questionEnvoi } from '../tableau/envoi-direct';
import { humanError } from '../api/net';
import type { ReponsePrete } from '../api/vue-mail';
import { escapeHtml } from '../utils/html';
import { icon } from '../ui/icons';

type InfoFn = (msg: string, kind?: 'success' | 'error' | 'info') => void;

export interface CtxReponsePrete {
  messageId: string;
  mailbox: string;
  mobile: boolean;
  onInfo: InfoFn;
  /** Formulaire de réponse d'Outlook prérempli avec ce texte (bureau ; absent sur mobile). */
  ouvrirDansOutlook?: (texte: string) => void;
}

/** Réponse prête (et réponses courtes) ; masquée s'il n'y a ni l'une ni les autres. */
export function renderReponsePrete(host: HTMLElement, r: ReponsePrete | null, courtes: string[], ctx: CtxReponsePrete): boolean {
  if (!r && !courtes.length) { host.hidden = true; host.innerHTML = ''; return false; }
  host.hidden = false;
  const alertes = (r?.alertes || []).filter(a => a?.message);
  host.innerHTML = `
    <div class="agent-section-title">Réponse prête</div>
    ${r?.resumeIntention ? `<p class="agent-resume">« ${escapeHtml(r.resumeIntention)} »</p>` : ''}
    ${alertes.length ? `<ul class="agent-avertissements" aria-label="Contrôle avant envoi">${alertes.map(a => `<li class="${a.gravite === 'bloquant' ? 'is-bloquant' : ''}">${a.gravite === 'bloquant' ? '<strong>À corriger :</strong> ' : ''}${escapeHtml(a.message)}</li>`).join('')}</ul>` : ''}
    ${courtes.length ? `<div class="agent-liens" role="group" aria-label="Réponses courtes">${courtes.map((t, i) => `<button type="button" class="agent-link" data-courte="${i}">${escapeHtml(t.length > 60 ? `${t.slice(0, 57)}…` : t)}</button>`).join('')}</div>` : ''}
    <textarea class="agent-input" data-texte rows="8" aria-label="Texte de la réponse">${escapeHtml(r?.texte || '')}</textarea>
    <button type="button" class="btn btn-primary btn-block agent-btn" data-envoyer>${icon('reply', 14)} Envoyer</button>
    ${ctx.ouvrirDansOutlook ? '<button type="button" class="btn btn-secondary btn-block agent-btn" data-ouvrir>Ouvrir dans Outlook</button>' : ''}
    <div class="agent-muted">Sans signature : elle est ajoutée à l'envoi. Rien ne part sans ton clic.</div>
    <div data-confirmer></div>`;
  const zone = host.querySelector<HTMLTextAreaElement>('[data-texte]')!;
  const confirmer = host.querySelector<HTMLElement>('[data-confirmer]')!;
  const envoyer = host.querySelector<HTMLButtonElement>('[data-envoyer]')!;
  host.querySelectorAll<HTMLButtonElement>('[data-courte]').forEach(b => b.addEventListener('click', () => {
    zone.value = courtes[Number(b.dataset.courte)];
    zone.focus();
  }));
  host.querySelector('[data-ouvrir]')?.addEventListener('click', () => ctx.ouvrirDansOutlook?.(zone.value));
  envoyer.addEventListener('click', async () => {
    const texte = zone.value.trim();
    if (!texte) { ctx.onInfo('Écris d\'abord ta réponse', 'error'); return; }
    envoyer.disabled = true;
    try {
      const p = await preparerEnvoi(ctx.messageId, ctx.mailbox, texte);
      if (!p?.pret || !p.brouillonId) throw new Error('Envoi indisponible : réponds depuis Outlook');
      confirmer.innerHTML = `
        <p><strong>${escapeHtml(questionEnvoi(p))}</strong></p>
        <button type="button" class="btn btn-primary btn-block agent-btn" data-ok>Envoyer</button>
        <button type="button" class="btn btn-secondary btn-block agent-btn" data-non>Annuler</button>`;
      const ok = await new Promise<boolean>(res => {
        confirmer.querySelector('[data-ok]')?.addEventListener('click', () => res(true));
        confirmer.querySelector('[data-non]')?.addEventListener('click', () => res(false));
      });
      confirmer.innerHTML = '';
      if (!ok) { await abandonnerEnvoi(p.brouillonId).catch(() => undefined); envoyer.disabled = false; return; }
      const r2 = await confirmerEnvoi(p.brouillonId, ctx.messageId);
      if (!r2?.envoye) throw new Error('Envoi refusé par le worker');
      host.innerHTML = `
        <div class="agent-section-title">Réponse prête</div>
        <p class="status-linked">${icon('check-circle', 14)}Envoyé. Part dans quelques secondes.</p>
        <button type="button" class="agent-link" data-retenir>Annuler l'envoi</button>`;
      const retenir = host.querySelector<HTMLButtonElement>('[data-retenir]')!;
      const fin = window.setTimeout(() => retenir.remove(), Math.max(3000, (r2.annulableS || 10) * 1000 - 2000));
      retenir.addEventListener('click', async () => {
        window.clearTimeout(fin);
        retenir.disabled = true;
        try { await abandonnerEnvoi(p.brouillonId); ctx.onInfo('Envoi annulé : le mail n\'est pas parti', 'info'); renderReponsePrete(host, { ...(r || { resumeIntention: '', depose: false }), texte }, courtes, ctx); }
        catch (e) { ctx.onInfo(/déjà parti/i.test(humanError(e)) ? 'Trop tard : le mail est déjà parti' : humanError(e), 'error'); retenir.remove(); }
      });
      ctx.onInfo('Réponse envoyée', 'success');
    } catch (e) {
      confirmer.innerHTML = '';
      envoyer.disabled = false;
      ctx.onInfo(humanError(e), 'error');
    }
  });
  return true;
}
