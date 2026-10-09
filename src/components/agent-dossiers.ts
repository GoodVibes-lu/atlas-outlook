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
 * Écritures : par le worker, sur le clic de la personne dans SA boîte (graph.ts › rangementClicAutorise) ;
 * ailleurs, le refus du worker est dit tel quel. Erreurs toujours affichées, jamais avalées.
 */
import { changerProjet, delierProjet } from '../api/projet-mail';
import { monterChoixProjet } from './carte-pret';
import {
  fetchDossiers, fetchDossierProjet, rangerDansDossier, fetchOffreFournisseur, preparerRemerciementOffre, executerAction,
  annulerAction, estVerrouFerme, type DossierOutlook, type OffreFournisseurContexte,
} from '../api/agent';
import { getAllProjets } from '../api/airtable';
import { renderImportDossier, renderContactDevis } from './agent-parite';
import { fetchProjetDuMail } from '../api/parite';
import { annulerTraite, phraseSuggestion, rangerTraite, sourceSuggestion, type TraiteARanger } from '../api/tableau';
import { escapeHtml } from '../utils/html';
import { humanError } from '../api/net';
import { icon } from '../ui/icons';

export type InfoFn = (message: string, type?: 'success' | 'error' | 'info') => void;

export interface CtxDossiers {
  messageId: string;
  mailbox: string;
  onInfo?: InfoFn;
  /** Ancien repli « jeton de la personne » : plus utilisé pour ranger (Microsoft ne donne plus ce jeton). */
  delegue?: () => Promise<{ token: string; restId: string } | null>;
}

const erreur = (quoi: string, e: unknown) => `<p class="agent-error" role="alert">${escapeHtml(quoi)} : ${escapeHtml(humanError(e))}</p>`;

/**
 * Range le mail (dossier existant ou chemin à créer) par le worker, sur le clic de la personne (sa
 * boîte : graph.ts › rangementClicAutorise). Plus de repli par le jeton Outlook du complément :
 * Microsoft ne le donne plus (il échouait avec un faux « rouvre ta session »). Renvoie le chemin
 * final et l'action à annuler.
 */
export async function rangerMail(ctx: CtxDossiers, cible: { dossier?: DossierOutlook; creer?: { chemin: string; projetId?: string; mandatId?: string }; projetId?: string }): Promise<{ chemin: string; cree: boolean; actionId?: string; parOutlook?: boolean; dossierId?: string }> {
  const r = await rangerDansDossier({
    messageId: ctx.messageId, ...(ctx.mailbox ? { mailbox: ctx.mailbox } : {}),
    ...(cible.dossier ? { dossierId: cible.dossier.id } : {}), ...(cible.creer ? { creer: cible.creer } : {}),
    ...(cible.projetId ? { projetId: cible.projetId } : {}),
  });
  return { chemin: r.dossier.chemin, cree: r.cree, ...(r.dossier.id ? { dossierId: r.dossier.id } : {}), ...(r.actionId ? { actionId: r.actionId } : {}) };
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
        <div data-import hidden></div>
        <div data-res hidden></div></div>`
    : `<div class="agent-carte"><div class="agent-carte-titre">${icon('folder', 14)} Aucun dossier Outlook pour ce ${ctx.mandatId ? 'mandat' : 'projet'}</div>
        <label class="agent-muted" for="dp-chemin">Dossier à créer (même place et même nom que dans ATLAS)</label>
        <input type="text" class="agent-input" id="dp-chemin" value="${escapeHtml(propose)}" maxlength="600">
        <button type="button" class="btn btn-primary btn-block agent-btn" data-creer>Créer le dossier et ranger</button>
        <div data-res hidden></div></div>`;
  const res = host.querySelector<HTMLElement>('[data-res]')!;
  // 07/10/2026 (parité FileEmailModal › handleBulkImport) : mails déjà dans le dossier du projet → importés dans le projet.
  const zoneImport = host.querySelector<HTMLElement>('[data-import]');
  if (existant && zoneImport && ctx.projetId) renderImportDossier(zoneImport, { ...ctx, projetId: ctx.projetId, dossier: existant });
  const go = async (btn: HTMLButtonElement, cible: Parameters<typeof rangerMail>[1]) => {
    btn.disabled = true;
    const lib = btn.textContent || '';
    btn.textContent = 'En cours…';
    try {
      const r = await rangerMail(ctx, cible);
      host.innerHTML = `${faitHtml(r)}<div data-import hidden></div>`;
      brancherAnnuler(host, r.actionId, ctx.onInfo);
      ctx.onInfo?.(`Rangé dans ${r.chemin}`, 'success');
      const zi = host.querySelector<HTMLElement>('[data-import]');
      if (zi && ctx.projetId && r.dossierId) renderImportDossier(zi, { ...ctx, projetId: ctx.projetId, dossier: { id: r.dossierId, chemin: r.chemin } });
    } catch (e) {
      btn.disabled = false; btn.textContent = lib;
      res.hidden = false;
      res.innerHTML = `<p class="agent-error" role="alert">Rangement impossible : ${escapeHtml(texteErreur(e))}</p>`;
    }
  };
  host.querySelector<HTMLButtonElement>('[data-ranger]')?.addEventListener('click', e => void go(e.currentTarget as HTMLButtonElement, { dossier: existant!, ...(ctx.projetId ? { projetId: ctx.projetId } : {}) }));
  host.querySelector<HTMLButtonElement>('[data-creer]')?.addEventListener('click', e => {
    const chemin = (host.querySelector<HTMLInputElement>('#dp-chemin')?.value || '').trim();
    if (!chemin) { res.hidden = false; res.innerHTML = '<p class="agent-error" role="alert">Indique le dossier à créer.</p>'; return; }
    void go(e.currentTarget as HTMLButtonElement, { creer: { chemin, ...(ctx.projetId ? { projetId: ctx.projetId } : {}), ...(ctx.mandatId ? { mandatId: ctx.mandatId } : {}) } });
  });
}

// ── Classer ce mail (une seule carte, 07/10/2026) ──────────────────────────────
//
// Retour de Charles : l'ancien bloc (liste de 80 dossiers, « Dossier choisi : aucun (racine) »,
// « Ranger ici » grisé, « Nouveau dossier… » qui créait à la racine) n'était pas utilisable. Ici :
//   1. le BON dossier d'abord : celui du projet du mail (appris dans ATLAS, sinon trouvé dans tout
//      l'arbre), en un clic ; pas de dossier ? le chemin d'ATLAS « Clients/<Client>/#871 Nom »,
//      « Créer et classer » en un clic (modifiable) ;
//   2. un autre dossier : une seule zone, on tape, les dossiers correspondants s'affichent, un clic
//      classe ; ce qu'on tape peut aussi devenir un nouveau dossier (« A/B/C » = niveaux) ;
//   3. après : « Classé dans … » avec « Annuler ». Le dossier du projet est retenu pour ATLAS (même
//      table que le widget Inbox) : rien à ressaisir dans ATLAS.

export interface CtxClasser extends CtxDossiers {
  /** Projet du mail s'il est déjà connu (sinon lu : fil lié dans ATLAS, ou appris par l'agent). */
  projetId?: string;
  projetLibelle?: string;
  /**
   * Traités à ranger (08/10/2026) : la personne a DÉJÀ répondu à ce mail ; la carte passe en tête avec
   * la suggestion du worker mise en avant (« Répondu : le ranger dans … ? »), un clic range.
   */
  repondu?: TraiteARanger;
  /** Fil du mail (Office.js) : le worker trouve le projet sans relire le mail. */
  conversationId?: string;
  /** Projet reconnu par l'agent (pas encore lié dans ATLAS) : présenté comme « probable ». */
  projetPropose?: boolean;
  /** La vue du mail a déjà répondu « aucun projet » : rien à redemander. */
  sansProjet?: boolean;
  /** Dossier du projet préparé par l'agent à l'arrivée (lot 2, 10/10/2026) : plus d'appel séparé. */
  dossier?: { existant: { id: string; chemin: string } | null; propose: string | null } | null;
}

/** Message d'un refus du worker (texte du serveur s'il en donne un), sinon message lisible. */
function texteErreur(e: unknown): string {
  const d = (e as any)?.data;
  return estVerrouFerme(e) && typeof d?.error === 'string' ? d.error : humanError(e);
}

export function renderClasserOutlook(host: HTMLElement, ctx: CtxClasser): void {
  host.hidden = false;
  const rep = ctx.repondu;
  host.innerHTML = `
    <div class="agent-section-title">${rep ? 'Répondu : ranger ce mail' : 'Classer ce mail'}</div>
    <div class="agent-carte${rep ? ' is-repondu' : ''}">
      ${rep ? '<div data-repondu></div>' : ''}
      <div data-suggestion${rep && rep.destination.type === 'projet' ? ' hidden' : ''}><div class="agent-loading"><div class="spinner"></div><span>Dossier du projet…</span></div></div>
      <label class="agent-muted" for="cl-q-${ctx.messageId.length}">Autre dossier</label>
      <input type="search" class="agent-input" id="cl-q-${ctx.messageId.length}" data-recherche autocomplete="off"
        placeholder="Tape un nom ou un n° de projet (ex. 871, DealsUp)" aria-label="Chercher un dossier Outlook">
      <div data-liste></div>
      <div data-res hidden></div>
    </div>`;
  const sugg = host.querySelector<HTMLElement>('[data-suggestion]')!;
  const liste = host.querySelector<HTMLElement>('[data-liste]')!;
  const res = host.querySelector<HTMLElement>('[data-res]')!;
  const recherche = host.querySelector<HTMLInputElement>('[data-recherche]')!;
  let projet: { id: string; libelle: string } | null = ctx.projetId ? { id: ctx.projetId, libelle: ctx.projetLibelle || '' } : null;
  let occupe = false;

  const classer = async (cible: Parameters<typeof rangerMail>[1], btn?: HTMLButtonElement) => {
    if (occupe) return;
    occupe = true;
    const lib = btn?.textContent || '';
    if (btn) { btn.disabled = true; btn.textContent = 'Classement…'; }
    res.hidden = true;
    try {
      const r = await rangerMail(ctx, cible);
      host.querySelector('.agent-carte')!.innerHTML = `${faitHtml(r)}<div data-import hidden></div>`;
      brancherAnnuler(host, r.actionId, ctx.onInfo);
      ctx.onInfo?.(`Classé dans ${r.chemin}`, 'success');
      // Mails déjà présents dans le dossier du projet : proposés à l'import dans le projet (comme ATLAS).
      const zi = host.querySelector<HTMLElement>('[data-import]');
      if (zi && cible.projetId && r.dossierId) renderImportDossier(zi, { ...ctx, projetId: cible.projetId, dossier: { id: r.dossierId, chemin: r.chemin } });
    } catch (e) {
      if (btn) { btn.disabled = false; btn.textContent = lib; }
      res.hidden = false;
      res.innerHTML = `<p class="agent-error" role="alert">Classement impossible : ${escapeHtml(texteErreur(e))}</p>`;
    } finally { occupe = false; }
  };

  // 0. Déjà répondu (traités à ranger) : la suggestion du worker d'abord, un clic range (statut tenu par le worker, annulable).
  if (rep) {
    const zone = host.querySelector<HTMLElement>('[data-repondu]')!;
    const d = rep.destination;
    const src = sourceSuggestion(d);
    zone.innerHTML = `<div class="agent-carte-titre">${icon('check-circle', 14)} ${escapeHtml(phraseSuggestion(d))}</div>
      ${src ? `<div class="agent-muted">${escapeHtml(src)}</div>` : '<div class="agent-muted">Aucun dossier proposé : choisis-en un ci-dessous.</div>'}
      ${d.type !== 'aucun' && d.chemin ? '<button type="button" class="btn btn-primary btn-block agent-btn" data-ranger-repondu>Ranger</button>' : ''}`;
    zone.querySelector<HTMLButtonElement>('[data-ranger-repondu]')?.addEventListener('click', async ev => {
      if (occupe) return;
      occupe = true;
      const btn = ev.currentTarget as HTMLButtonElement;
      btn.disabled = true; btn.textContent = 'Rangement…';
      res.hidden = true;
      try {
        const r = await rangerTraite({ messageId: rep.messageId });
        host.querySelector('.agent-carte')!.innerHTML = faitHtml({ chemin: r.dossier.chemin, cree: r.cree, actionId: r.actionId });
        const b = host.querySelector<HTMLButtonElement>('[data-annuler-rangement]');
        b?.addEventListener('click', async () => {
          b.disabled = true;
          try { await annulerTraite(rep.messageId); b.replaceWith(Object.assign(document.createElement('span'), { className: 'agent-muted', textContent: 'Annulé : mail remis à sa place' })); ctx.onInfo?.('Rangement annulé', 'success'); }
          catch (e) { b.disabled = false; ctx.onInfo?.(`Annulation impossible : ${humanError(e)}`, 'error'); }
        });
        ctx.onInfo?.(`Rangé dans ${r.dossier.chemin}`, 'success');
      } catch (e) {
        btn.disabled = false; btn.textContent = 'Ranger';
        res.hidden = false;
        res.innerHTML = `<p class="agent-error" role="alert">Rangement impossible : ${escapeHtml(texteErreur(e))}</p>`;
      } finally { occupe = false; }
    });
  }

  // Changer de projet / Délier (correction mémorisée par le worker : elle prime ensuite).
  const brancherProjet = () => {
    const ancien = projet;
    sugg.querySelector('[data-changer-projet]')?.addEventListener('click', () => {
      const zone = sugg.querySelector<HTMLElement>('[data-choix-projet]');
      if (!zone) return;
      monterChoixProjet(zone, async p => {
        try {
          const r = await changerProjet({ messageId: ctx.messageId, projetId: p.id, ...(ancien && !ctx.projetPropose ? { ancienProjetId: ancien.id } : {}), ...(ctx.mailbox ? { mailbox: ctx.mailbox } : {}) });
          ctx.onInfo?.(`Lié à ${r.projet.libelle}${r.rangement?.dossier ? ` et classé dans ${r.rangement.dossier.chemin}` : ''}`, 'success');
          renderClasserOutlook(host, { ...ctx, repondu: undefined, projetId: r.projet.id, projetLibelle: r.projet.libelle, projetPropose: false, sansProjet: false, dossier: undefined });
        } catch (e) { res.hidden = false; res.innerHTML = `<p class="agent-error" role="alert">Changement impossible : ${escapeHtml(texteErreur(e))}</p>`; }
      }, () => undefined);
    });
    sugg.querySelector<HTMLButtonElement>('[data-delier-projet]')?.addEventListener('click', async ev => {
      const b = ev.currentTarget as HTMLButtonElement;
      b.disabled = true;
      try {
        await delierProjet({ messageId: ctx.messageId, ...(ancien ? { projetId: ancien.id } : {}), ...(ctx.conversationId ? { conversationId: ctx.conversationId } : {}), ...(ctx.mailbox ? { mailbox: ctx.mailbox } : {}) });
        ctx.onInfo?.('Délié du projet : ce choix est retenu pour ce fil', 'success');
        renderClasserOutlook(host, { ...ctx, repondu: undefined, projetId: undefined, projetLibelle: undefined, projetPropose: false, sansProjet: true, dossier: undefined });
      } catch (e) { b.disabled = false; res.hidden = false; res.innerHTML = `<p class="agent-error" role="alert">Impossible de délier : ${escapeHtml(texteErreur(e))}</p>`; }
    });
  };

  // 1. Dossier du projet
  const suggestion = async () => {
    try {
      if (!projet && !ctx.sansProjet) projet = await fetchProjetDuMail(ctx.messageId, ctx.mailbox || undefined, ctx.conversationId).catch(() => null);
      if (!host.isConnected) return;
      if (!projet) { sugg.innerHTML = '<p class="agent-muted">Mail lié à aucun projet : choisis le dossier ci-dessous.</p>'; return; }
      // Dossier préparé à l'arrivée (existant) : affiché tout de suite ; « déjà classé » se voit après le clic.
      const d = ctx.dossier && ctx.projetId === projet.id
        ? { existant: ctx.dossier.existant, propose: ctx.dossier.propose, verrou: true, dejaRange: false }
        : await fetchDossierProjet(projet.id, ctx.mailbox || undefined, ctx.messageId);
      if (!host.isConnected) return;
      // Lot 3 (10/10/2026) : « Changer de projet » et « Délier » partout où le projet du mail est affiché.
      const titre = `<div class="agent-carte-titre">${icon('folder', 14)} ${ctx.projetPropose ? 'Projet probable' : 'Projet'} ${escapeHtml(projet.libelle)}</div>
        <div class="agent-liens"><button type="button" class="agent-link" data-changer-projet>Changer de projet</button>${ctx.projetPropose ? '' : '<button type="button" class="agent-link" data-delier-projet>Délier</button>'}</div>
        <div data-choix-projet hidden></div>`;
      queueMicrotask(() => brancherProjet());
      if (d.existant && d.dejaRange) {
        sugg.innerHTML = `${titre}<p class="status-linked">${icon('check-circle', 14)}Déjà classé dans « ${escapeHtml(d.existant.chemin)} »</p>`;
        return;
      }
      if (d.existant) {
        sugg.innerHTML = `${titre}<div class="agent-carte-apercu">« ${escapeHtml(d.existant.chemin)} »</div>
          <button type="button" class="btn btn-primary btn-block agent-btn" data-go>Classer dans ce dossier</button>`;
        sugg.querySelector<HTMLButtonElement>('[data-go]')!.addEventListener('click', ev =>
          void classer({ dossier: d.existant!, projetId: projet!.id }, ev.currentTarget as HTMLButtonElement));
        return;
      }
      const chemin = d.propose || '';
      sugg.innerHTML = `${titre}<div class="agent-muted">Pas encore de dossier Outlook pour ce projet. Il sera créé ici :</div>
        <input type="text" class="agent-input" data-chemin value="${escapeHtml(chemin)}" maxlength="600" aria-label="Dossier à créer">
        <button type="button" class="btn btn-primary btn-block agent-btn" data-go>Créer le dossier et classer</button>`;
      sugg.querySelector<HTMLButtonElement>('[data-go]')!.addEventListener('click', ev => {
        const c = (sugg.querySelector<HTMLInputElement>('[data-chemin]')?.value || '').trim();
        if (!c) { res.hidden = false; res.innerHTML = '<p class="agent-error" role="alert">Indique le dossier à créer.</p>'; return; }
        void classer({ creer: { chemin: c, projetId: projet!.id }, projetId: projet!.id }, ev.currentTarget as HTMLButtonElement);
      });
    } catch (e) {
      if (host.isConnected) sugg.innerHTML = `<p class="agent-error" role="alert">Dossier du projet indisponible : ${escapeHtml(texteErreur(e))}</p>`;
    }
  };
  void suggestion();

  // 2. Autre dossier : recherche dans tout l'arbre, un clic = classé ; ce qui est tapé peut devenir un dossier.
  let minuterie: ReturnType<typeof setTimeout> | undefined;
  let n = 0;
  const afficher = async (q: string) => {
    const k = ++n;
    if (!q) { liste.innerHTML = ''; return; }
    liste.innerHTML = '<div class="agent-loading"><div class="spinner"></div><span>Recherche…</span></div>';
    try {
      const ds = (await fetchDossiers(ctx.mailbox || undefined, q)).slice(0, 8);
      if (k !== n || !liste.isConnected) return;
      const cheminTape = q.split('/').map(x => x.trim()).filter(Boolean).join('/');
      const exact = ds.some(d => d.chemin.toLowerCase() === cheminTape.toLowerCase() || d.chemin.toLowerCase().endsWith(`/${cheminTape.toLowerCase()}`));
      liste.innerHTML = `<ul class="agent-faites">
        ${ds.map((d, i) => `<li><button type="button" class="agent-link" data-k="${i}" title="Classer dans ce dossier">${icon('folder', 12)} ${escapeHtml(d.chemin)}</button></li>`).join('')}
        ${cheminTape && !exact ? `<li><button type="button" class="agent-link" data-nouveau>${icon('plus', 12)} Créer « ${escapeHtml(cheminTape)} » et classer</button></li>` : ''}
      </ul>${ds.length ? '' : '<div class="agent-muted">Aucun dossier ne correspond.</div>'}`;
      liste.querySelectorAll<HTMLButtonElement>('[data-k]').forEach(b => b.addEventListener('click', () => {
        const d = ds[Number(b.dataset.k)];
        if (d) void classer({ dossier: d });
      }));
      liste.querySelector<HTMLButtonElement>('[data-nouveau]')?.addEventListener('click', () => void classer({ creer: { chemin: cheminTape } }));
    } catch (e) { if (k === n) liste.innerHTML = `<p class="agent-error" role="alert">Dossiers Outlook indisponibles : ${escapeHtml(texteErreur(e))}</p>`; }
  };
  recherche.addEventListener('input', () => {
    if (minuterie) clearTimeout(minuterie);
    const v = recherche.value.trim();
    minuterie = setTimeout(() => void afficher(v), 250);
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
        <div class="agent-muted">Pièces jointes de l'offre (aucune cochée : offre dans le texte du mail)</div>
        ${c.pieces.length ? c.pieces.map((p, i) => `<label class="agent-check"><input type="checkbox" data-of-pj="${i}" ${p.id === c!.pieceJointeId ? 'checked' : ''}/> ${escapeHtml(p.nom)}${p.devis ? ' <span class="agent-muted">(devis)</span>' : ''}</label>`).join('') : '<div class="agent-muted">Aucune pièce jointe.</div>'}
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
    // 07/10/2026 : plusieurs pièces jointes, chacune ajoutée à l'offre (comme l'Inbox ATLAS).
    const ids = [...zone.querySelectorAll<HTMLInputElement>('[data-of-pj]')].filter(x => x.checked).map(x => c?.pieces[Number(x.dataset.ofPj)]?.id).filter(Boolean) as string[];
    if (!devisId) { res.hidden = false; res.innerHTML = '<p class="agent-error" role="alert">Choisis d\'abord la demande de devis.</p>'; return; }
    btn.disabled = true;
    btn.textContent = 'Dépôt…';
    try {
      const choixPj = ids.length > 1 ? { pieceJointeIds: ids.join(',') } : ids.length ? { pieceJointeId: ids[0] } : {};
      const r = await executerAction({ messageId: ctx.messageId, mailbox: ctx.mailbox || undefined, type: 'devis-fournisseur', choix: { devisId, ...choixPj } });
      if (!r.ok) throw new Error(r.resume || 'Dépôt impossible');
      zone.innerHTML = `<div class="agent-carte is-fait"><div class="agent-fait" role="status"><span>${r.simule ? 'Simulé (agent en observation)' : 'Fait'} : ${escapeHtml(r.resume)}</span></div>
        <div data-contact hidden></div>
        ${r.simule ? '' : `<button type="button" class="btn btn-secondary btn-block agent-btn" data-merci>${icon('bolt', 14)}Préparer le remerciement (ARGO)</button>`}
        <div data-res hidden></div></div>`;
      ctx.onInfo?.('Offre déposée', 'success');
      // 07/10/2026 (parité handleRelinkContact) : le répondant n'est pas le contact de la demande ? à trancher avant le remerciement.
      const zc = zone.querySelector<HTMLElement>('[data-contact]');
      if (zc && !r.simule) void renderContactDevis(zc, { ...ctx, devisId });
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
