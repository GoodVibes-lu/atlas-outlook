/**
 * agent-dossiers.ts · Parité Outlook avec l'Inbox ATLAS (demande de Charles du 07/10/2026) :
 *
 *  • « Dossier du projet » : dossier Outlook EXISTANT du projet, n'importe où dans l'arbre complet
 *    (n° de projet d'abord, puis nom), sinon « Créer le dossier et ranger » au chemin proposé, à la
 *    même place et avec le même nom que l'Inbox ATLAS (« Clients/<Client>/#755 Nom », modifiable) ;
 *  • « Classer dans Outlook » : recherche dans TOUTE l'arborescence de la boîte, rangement en un
 *    clic, « Nouveau dossier » sous le dossier choisi (ou à la racine), comme FileEmailModal ;
 *  • « Offre fournisseur reçue » : comme le bouton « Offre reçue » de l'Inbox ATLAS : projet
 *    pré-sélectionné (fil lié, « #847 » dans l'objet), demande de devis notée sur les mêmes
 *    signaux, pièce jointe choisie, dépôt par l'action « devis-fournisseur » (journal, Annuler),
 *    puis remerciement ARGO repris dans la réponse (jamais envoyé seul).
 *
 * Écritures : par le worker (double verrou de la boîte). Verrou fermé : si la boîte est celle de la
 * personne, repli avec SON jeton Outlook (`delegue`), exactement comme l'Inbox ATLAS ; sinon le
 * refus est dit. Erreurs toujours affichées, jamais avalées.
 */
import {
  fetchDossiers, fetchDossierProjet, rangerDansDossier, fetchOffreFournisseur, preparerRemerciementOffre, executerAction,
  annulerAction, estVerrouFerme, type DossierOutlook, type OffreFournisseurContexte,
} from '../api/agent';
import { getAllProjets } from '../api/airtable';
import { ensureFolderPath, moveMessageToFolder } from '../api/graph';
import { escapeHtml } from '../utils/html';
import { humanError } from '../api/net';
import { icon } from '../ui/icons';

export type InfoFn = (message: string, type?: 'success' | 'error' | 'info') => void;

export interface CtxDossiers {
  messageId: string;
  mailbox: string;
  onInfo?: InfoFn;
  /** Repli « jeton de la personne » (sa boîte seulement) quand l'agent ne peut pas écrire. */
  delegue?: () => Promise<{ token: string; restId: string } | null>;
}

const erreur = (quoi: string, e: unknown) => `<p class="agent-error" role="alert">${escapeHtml(quoi)} : ${escapeHtml(humanError(e))}</p>`;

/**
 * Range le mail (dossier existant ou chemin à créer) : worker d'abord ; verrou fermé → repli avec le
 * jeton de la personne si permis. Renvoie le chemin final et l'action à annuler (worker seulement).
 */
export async function rangerMail(ctx: CtxDossiers, cible: { dossier?: DossierOutlook; creer?: { chemin: string; projetId?: string; mandatId?: string }; projetId?: string }): Promise<{ chemin: string; cree: boolean; actionId?: string; parOutlook?: boolean }> {
  try {
    const r = await rangerDansDossier({
      messageId: ctx.messageId, ...(ctx.mailbox ? { mailbox: ctx.mailbox } : {}),
      ...(cible.dossier ? { dossierId: cible.dossier.id } : {}), ...(cible.creer ? { creer: cible.creer } : {}),
      ...(cible.projetId ? { projetId: cible.projetId } : {}),
    });
    return { chemin: r.dossier.chemin, cree: r.cree, ...(r.actionId ? { actionId: r.actionId } : {}) };
  } catch (e) {
    if (!estVerrouFerme(e) || !ctx.delegue) throw e;
    const d = await ctx.delegue();
    if (!d) throw e;
    const chemin = cible.creer?.chemin || cible.dossier?.chemin || '';
    const id = cible.creer ? await ensureFolderPath(d.token, cible.creer.chemin) : cible.dossier!.id;
    await moveMessageToFolder(d.token, d.restId, id);
    return { chemin, cree: !!cible.creer, parOutlook: true };
  }
}

function faitHtml(r: { chemin: string; cree: boolean; actionId?: string }): string {
  return `<div class="agent-fait" role="status"><span>${icon('check-circle', 14)} ${r.cree ? 'Dossier créé et mail rangé' : 'Mail rangé'} dans « ${escapeHtml(r.chemin)} »</span>
    ${r.actionId ? '<span class="agent-fait-btns"><button type="button" class="btn btn-secondary agent-btn" data-annuler-rangement>Annuler</button></span>' : ''}</div>`;
}

function brancherAnnuler(host: HTMLElement, actionId: string | undefined, onInfo?: InfoFn): void {
  const b = host.querySelector<HTMLButtonElement>('[data-annuler-rangement]');
  if (!b || !actionId) return;
  b.addEventListener('click', async () => {
    b.disabled = true;
    try { await annulerAction(actionId); b.replaceWith(Object.assign(document.createElement('span'), { className: 'agent-muted', textContent: 'Annulé : mail remis à sa place' })); onInfo?.('Rangement annulé', 'success'); }
    catch (e) { b.disabled = false; onInfo?.(`Annulation impossible : ${humanError(e)}`, 'error'); }
  });
}

// ── Dossier du projet ──────────────────────────────────────────────────────────

/**
 * Carte « Dossier du projet » : existant → « Ranger dans … » ; absent → chemin proposé (modifiable)
 * + « Créer le dossier et ranger ». `propose` : chemin déjà calculé (résultat d'action).
 */
export async function renderDossierProjet(host: HTMLElement, ctx: CtxDossiers & { projetId?: string; mandatId?: string; libelle?: string; propose?: string }): Promise<void> {
  host.hidden = false;
  host.innerHTML = '<div class="agent-loading"><div class="spinner"></div><span>Dossier Outlook du projet…</span></div>';
  let existant: DossierOutlook | null = null;
  let propose = ctx.propose || '';
  if (!propose && ctx.projetId) {
    try { const d = await fetchDossierProjet(ctx.projetId, ctx.mailbox || undefined); existant = d.existant; propose = d.propose || ''; }
    catch (e) { if (host.isConnected) host.innerHTML = erreur('Dossiers Outlook indisponibles', e); return; }
  }
  if (!host.isConnected) return;
  host.innerHTML = existant
    ? `<div class="agent-carte"><div class="agent-carte-titre">${icon('folder', 14)} Dossier du projet${ctx.libelle ? ` ${escapeHtml(ctx.libelle)}` : ''}</div>
        <div class="agent-carte-apercu">« ${escapeHtml(existant.chemin)} »</div>
        <button type="button" class="btn btn-primary btn-block agent-btn" data-ranger>Ranger ce mail dans ce dossier</button>
        <div data-res hidden></div></div>`
    : `<div class="agent-carte"><div class="agent-carte-titre">${icon('folder', 14)} Aucun dossier Outlook pour ce ${ctx.mandatId ? 'mandat' : 'projet'}</div>
        <label class="agent-muted" for="dp-chemin">Dossier à créer (même place et même nom que dans ATLAS)</label>
        <input type="text" class="agent-input" id="dp-chemin" value="${escapeHtml(propose)}" maxlength="600">
        <button type="button" class="btn btn-primary btn-block agent-btn" data-creer>Créer le dossier et ranger</button>
        <div data-res hidden></div></div>`;
  const res = host.querySelector<HTMLElement>('[data-res]')!;
  const go = async (btn: HTMLButtonElement, cible: Parameters<typeof rangerMail>[1]) => {
    btn.disabled = true;
    const lib = btn.textContent || '';
    btn.textContent = 'En cours…';
    try {
      const r = await rangerMail(ctx, cible);
      host.innerHTML = faitHtml(r);
      brancherAnnuler(host, r.actionId, ctx.onInfo);
      ctx.onInfo?.(`Rangé dans ${r.chemin}`, 'success');
    } catch (e) {
      btn.disabled = false; btn.textContent = lib;
      res.hidden = false;
      res.innerHTML = estVerrouFerme(e) ? '<p class="agent-error" role="alert">L\'agent n\'écrit pas encore dans cette boîte (double verrou fermé) : rien n\'a été créé ni déplacé.</p>' : erreur('Rangement impossible', e);
    }
  };
  host.querySelector<HTMLButtonElement>('[data-ranger]')?.addEventListener('click', e => void go(e.currentTarget as HTMLButtonElement, { dossier: existant!, ...(ctx.projetId ? { projetId: ctx.projetId } : {}) }));
  host.querySelector<HTMLButtonElement>('[data-creer]')?.addEventListener('click', e => {
    const chemin = (host.querySelector<HTMLInputElement>('#dp-chemin')?.value || '').trim();
    if (!chemin) { res.hidden = false; res.innerHTML = '<p class="agent-error" role="alert">Indique le dossier à créer.</p>'; return; }
    void go(e.currentTarget as HTMLButtonElement, { creer: { chemin, ...(ctx.projetId ? { projetId: ctx.projetId } : {}), ...(ctx.mandatId ? { mandatId: ctx.mandatId } : {}) } });
  });
}

// ── Classer dans Outlook (arbre complet + nouveau dossier) ─────────────────────

export function renderClasserOutlook(host: HTMLElement, ctx: CtxDossiers): void {
  host.hidden = false;
  host.innerHTML = `
    <div class="agent-section-title">Classer dans Outlook</div>
    <button type="button" class="btn btn-secondary btn-block agent-btn" data-ouvrir>${icon('folder', 14)}Choisir un dossier…</button>
    <div data-zone hidden>
      <input type="search" class="agent-input" data-recherche placeholder="Rechercher dans tous les dossiers (ex. 755 gala)" aria-label="Rechercher un dossier Outlook">
      <div data-liste></div>
      <div class="agent-muted">Dossier choisi : <span data-choisi>aucun (racine de la boîte)</span></div>
      <div class="btn-row">
        <button type="button" class="btn btn-primary agent-btn" data-ranger disabled>Ranger ici</button>
        <button type="button" class="btn btn-secondary agent-btn" data-nouveau>Nouveau dossier…</button>
      </div>
      <div data-nouveau-zone hidden>
        <input type="text" class="agent-input" data-nom placeholder="Nom du nouveau dossier" maxlength="120" aria-label="Nom du nouveau dossier">
        <button type="button" class="btn btn-primary btn-block agent-btn" data-creer>Créer et ranger</button>
      </div>
    </div>
    <div data-res hidden></div>`;
  const zone = host.querySelector<HTMLElement>('[data-zone]')!;
  const liste = host.querySelector<HTMLElement>('[data-liste]')!;
  const choisiEl = host.querySelector<HTMLElement>('[data-choisi]')!;
  const res = host.querySelector<HTMLElement>('[data-res]')!;
  const btnRanger = host.querySelector<HTMLButtonElement>('[data-ranger]')!;
  let choisi: DossierOutlook | null = null;
  let minuterie: ReturnType<typeof setTimeout> | undefined;
  let n = 0;
  const afficher = async (q: string) => {
    const k = ++n;
    liste.innerHTML = '<div class="agent-loading"><div class="spinner"></div><span>Dossiers Outlook…</span></div>';
    try {
      const ds = await fetchDossiers(ctx.mailbox || undefined, q || undefined);
      if (k !== n || !liste.isConnected) return;
      liste.innerHTML = ds.length
        ? `<ul class="agent-faites">${ds.map((d, i) => `<li><button type="button" class="agent-link" data-k="${i}">${escapeHtml(d.chemin)}</button></li>`).join('')}</ul>`
        : '<div class="agent-muted">Aucun dossier ne correspond.</div>';
      liste.querySelectorAll<HTMLButtonElement>('[data-k]').forEach(b => b.addEventListener('click', () => {
        choisi = ds[Number(b.dataset.k)] || null;
        choisiEl.textContent = choisi?.chemin || 'aucun (racine de la boîte)';
        btnRanger.disabled = !choisi;
      }));
    } catch (e) { if (k === n) liste.innerHTML = erreur('Dossiers Outlook indisponibles', e); }
  };
  host.querySelector<HTMLButtonElement>('[data-ouvrir]')!.addEventListener('click', () => {
    zone.hidden = !zone.hidden;
    if (!zone.hidden) { void afficher(''); host.querySelector<HTMLInputElement>('[data-recherche]')?.focus(); }
  });
  host.querySelector<HTMLInputElement>('[data-recherche]')!.addEventListener('input', ev => {
    if (minuterie) clearTimeout(minuterie);
    const v = (ev.target as HTMLInputElement).value.trim();
    minuterie = setTimeout(() => void afficher(v), 250);
  });
  host.querySelector<HTMLButtonElement>('[data-nouveau]')!.addEventListener('click', () => {
    const z = host.querySelector<HTMLElement>('[data-nouveau-zone]')!;
    z.hidden = !z.hidden;
    if (!z.hidden) host.querySelector<HTMLInputElement>('[data-nom]')?.focus();
  });
  const go = async (btn: HTMLButtonElement, cible: Parameters<typeof rangerMail>[1]) => {
    btn.disabled = true;
    res.hidden = true;
    try {
      const r = await rangerMail(ctx, cible);
      zone.hidden = true;
      res.hidden = false;
      res.innerHTML = faitHtml(r);
      brancherAnnuler(res, r.actionId, ctx.onInfo);
      ctx.onInfo?.(`Rangé dans ${r.chemin}`, 'success');
    } catch (e) {
      btn.disabled = false;
      res.hidden = false;
      res.innerHTML = estVerrouFerme(e) ? '<p class="agent-error" role="alert">L\'agent n\'écrit pas encore dans cette boîte (double verrou fermé) : rien n\'a été créé ni déplacé.</p>' : erreur('Classement impossible', e);
    }
  };
  btnRanger.addEventListener('click', () => { if (choisi) void go(btnRanger, { dossier: choisi }); });
  host.querySelector<HTMLButtonElement>('[data-creer]')!.addEventListener('click', ev => {
    const nom = (host.querySelector<HTMLInputElement>('[data-nom]')?.value || '').replace(/[\\/]/g, '-').trim();
    if (!nom) { res.hidden = false; res.innerHTML = '<p class="agent-error" role="alert">Indique le nom du dossier.</p>'; return; }
    void go(ev.currentTarget as HTMLButtonElement, { creer: { chemin: choisi ? `${choisi.chemin}/${nom}` : nom } });
  });
}

// ── Offre fournisseur reçue ────────────────────────────────────────────────────

export function renderOffreRecue(host: HTMLElement, ctx: CtxDossiers & { repondre?: (html: string) => void }): void {
  host.hidden = false;
  host.innerHTML = `
    <div class="agent-section-title">Offre fournisseur reçue</div>
    <button type="button" class="btn btn-secondary btn-block agent-btn" data-ouvrir>${icon('paperclip', 14)}Déposer l'offre sur une demande de devis…</button>
    <div data-zone hidden></div>`;
  const zone = host.querySelector<HTMLElement>('[data-zone]')!;
  let c: OffreFournisseurContexte | null = null;

  const dessiner = () => {
    if (!c) return;
    const opt = (v: string, l: string, sel: boolean) => `<option value="${escapeHtml(v)}" ${sel ? 'selected' : ''}>${escapeHtml(l)}</option>`;
    zone.innerHTML = `
      <div class="agent-carte">
        <label class="agent-muted" for="of-projet">Projet</label>
        <select class="agent-select" id="of-projet">
          <option value="">Choisir un projet…</option>
          ${c.projets.map(p => opt(p.id, p.libelle, p.id === c!.projetId)).join('')}
          <option value="__tous">Autre projet (liste complète)…</option>
        </select>
        <label class="agent-muted" for="of-devis">Demande de devis</label>
        <select class="agent-select" id="of-devis" ${c.demandes.length ? '' : 'disabled'}>
          ${c.demandes.length ? `<option value="">Choisir la demande…</option>${c.demandes.map(d => opt(d.id, `${d.libelle}${d.statut ? ` (${d.statut})` : ''}`, d.id === c!.devisId)).join('')}` : `<option value="">${c.projetId ? 'Aucune demande de devis sur ce projet' : 'Choisir d\'abord le projet'}</option>`}
        </select>
        <label class="agent-muted" for="of-pj">Pièce jointe (offre)</label>
        <select class="agent-select" id="of-pj">
          <option value="">Aucune (offre dans le texte du mail)</option>
          ${c.pieces.map(p => opt(p.id, `${p.nom}${p.devis ? ' (devis)' : ''}`, p.id === c!.pieceJointeId)).join('')}
        </select>
        <button type="button" class="btn btn-primary btn-block agent-btn" data-deposer>Déposer l'offre</button>
        <div data-res hidden></div>
      </div>`;
    const projetSel = zone.querySelector<HTMLSelectElement>('#of-projet')!;
    projetSel.addEventListener('change', async () => {
      if (projetSel.value === '__tous') { await choisirDansTous(); return; }
      await charger(projetSel.value || undefined);
    });
    zone.querySelector<HTMLButtonElement>('[data-deposer]')!.addEventListener('click', ev => void deposer(ev.currentTarget as HTMLButtonElement));
  };

  const charger = async (projetId?: string) => {
    zone.hidden = false;
    zone.innerHTML = '<div class="agent-loading"><div class="spinner"></div><span>Demandes de devis…</span></div>';
    try {
      const avant = c?.projets || [];
      c = await fetchOffreFournisseur(ctx.messageId, ctx.mailbox || undefined, projetId);
      for (const p of avant) if (!c.projets.some(x => x.id === p.id)) c.projets.push(p);
      dessiner();
    } catch (e) { zone.innerHTML = erreur('Offre fournisseur indisponible', e); }
  };

  const choisirDansTous = async () => {
    try {
      const tous = (await getAllProjets()).filter(p => !/clôtur|clotur|annul/i.test(p.statut || ''));
      const sel = zone.querySelector<HTMLSelectElement>('#of-projet')!;
      sel.innerHTML = `<option value="">Choisir un projet…</option>${tous.map(p => `<option value="${escapeHtml(p.id)}">${escapeHtml(`#${p.noProjet || p.refProjet} ${p.denomination}${p.client ? ` · ${p.client}` : ''}`)}</option>`).join('')}`;
      if (c) c.projets = tous.map(p => ({ id: p.id, libelle: `#${p.noProjet || p.refProjet} ${p.denomination}` }));
      sel.focus();
    } catch (e) { ctx.onInfo?.(`Projets indisponibles : ${humanError(e)}`, 'error'); }
  };

  const deposer = async (btn: HTMLButtonElement) => {
    const res = zone.querySelector<HTMLElement>('[data-res]')!;
    const devisId = zone.querySelector<HTMLSelectElement>('#of-devis')?.value || '';
    const pieceJointeId = zone.querySelector<HTMLSelectElement>('#of-pj')?.value || '';
    if (!devisId) { res.hidden = false; res.innerHTML = '<p class="agent-error" role="alert">Choisis d\'abord la demande de devis.</p>'; return; }
    btn.disabled = true;
    btn.textContent = 'Dépôt…';
    try {
      const r = await executerAction({ messageId: ctx.messageId, mailbox: ctx.mailbox || undefined, type: 'devis-fournisseur', choix: { devisId, ...(pieceJointeId ? { pieceJointeId } : {}) } });
      if (!r.ok) throw new Error(r.resume || 'Dépôt impossible');
      zone.innerHTML = `<div class="agent-carte is-fait"><div class="agent-fait" role="status"><span>${r.simule ? 'Simulé (agent en observation)' : 'Fait'} : ${escapeHtml(r.resume)}</span></div>
        ${r.simule ? '' : `<button type="button" class="btn btn-secondary btn-block agent-btn" data-merci>${icon('bolt', 14)}Préparer le remerciement (ARGO)</button>`}
        <div data-res hidden></div></div>`;
      ctx.onInfo?.('Offre déposée', 'success');
      zone.querySelector<HTMLButtonElement>('[data-merci]')?.addEventListener('click', async ev => {
        const b = ev.currentTarget as HTMLButtonElement;
        const out = zone.querySelector<HTMLElement>('[data-res]')!;
        b.disabled = true; b.textContent = 'ARGO rédige…';
        try {
          const m = await preparerRemerciementOffre(ctx.messageId, ctx.mailbox || undefined, devisId);
          if (ctx.repondre && m.html) { ctx.repondre(m.html); b.textContent = 'Réponse ouverte (relis puis envoie)'; }
          else { out.hidden = false; out.innerHTML = `<textarea class="agent-input" rows="8" readonly>${escapeHtml(m.texte)}</textarea>`; b.textContent = 'Texte prêt (à copier)'; }
        } catch (e) { b.disabled = false; b.textContent = 'Préparer le remerciement (ARGO)'; out.hidden = false; out.innerHTML = erreur('Remerciement indisponible', e); }
      });
    } catch (e) {
      btn.disabled = false;
      btn.textContent = 'Déposer l\'offre';
      res.hidden = false;
      res.innerHTML = erreur('Dépôt impossible', e);
    }
  };

  host.querySelector<HTMLButtonElement>('[data-ouvrir]')!.addEventListener('click', () => {
    if (!zone.hidden) { zone.hidden = true; return; }
    void charger();
  });
}
