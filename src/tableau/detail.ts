/**
 * detail.ts · colonne de droite quand un mail est choisi : ce qu'ARGO en dit (résumé, pourquoi en
 * tête), ce que l'agent propose (actions, réutilise /agent/actions), la réponse (ARGO dans le style
 * de la personne ou modèle Communications) déposée dans le brouillon Outlook, l'envoi programmé,
 * la relance si pas de réponse, et le fil d'équipe des boîtes partagées (« Je prends »,
 * « Attribuer à… », commentaires internes, jamais dans le mail).
 */
import type { Ctx } from './app';
import type { DigestNewsletter, ReponseRedigee, TableauMail } from '../api/tableau';
import { deposerBrouillon, programmerEnvoi, redigerArgo, remplirModeleTableau } from '../api/tableau';
import {
  fetchEquipe, relacherMail, attribuerMail, ajouterCommentaire, fetchModeles, demanderRelance,
  type ModeleReponse,
} from '../api/agent';
import { copierTexte, renderActionsMetier } from '../components/agent-outils';
import { renderDossierProjet, renderOffreRecue, renderClasserOutlook } from '../components/agent-dossiers';
import { renderPlusActions } from '../components/agent-parite';
import { renderSecurite } from '../components/agent-securite';
import { humanError } from '../api/net';
import { icon } from '../ui/icons';
import { filtrerCommandes, ilYA, prenomDe, quandLisible } from './logique';
import { h, toast, choisirMoment, ouvrirCouche } from './ui';

const kbd = (k: string) => `<span class="tb-kbd">${h(k)}</span>`;

export function renderDetail(host: HTMLElement, m: TableauMail, ctx: Ctx): void {
  const t = ctx.etat.t;
  const moi = t?.moi || '';
  const perso = (m.mailbox || '').toLowerCase() === moi;
  // « En attente » : un ENVOI sans réponse (relance) ; un mail reçu garde les actions de réponse.
  const envoye = ctx.etat.section === 'enAttente' && (m.envoye === true || m.categorie === 'envoi');
  const deCote = !!m.plusTardJusqua;
  host.innerHTML = `
    <div class="tb-panel tb-securite" id="tb-securite" hidden></div>
    <div class="tb-panel is-raised">
      <div class="tb-detail-h">
        <h3>${h(m.subject || '(sans objet)')}</h3>
        <div class="tb-detail-meta"><span>${h(m.from?.name || '')} &lt;${h(m.from?.email || '')}&gt;</span><span>${h(ilYA(m.receivedAt, Date.now()))}</span>${perso ? '' : `<span>${h(m.mailbox)}</span>`}</div>
        ${m.raisons?.length ? `<div class="tb-chips" style="justify-content:flex-start">${m.raisons.map(r => `<span class="tb-chip">${h(r)}</span>`).join('')}</div>` : ''}
      </div>
      ${m.resume ? `<div class="tb-resume">${h(m.resume)}</div>` : ''}
      <div class="tb-actions">
        <button type="button" class="tb-btn" data-a="ouvrir">${icon('external', 14)}Ouvrir dans Outlook ${kbd('⏎')}</button>
        ${envoye ? '' : `<button type="button" class="tb-btn is-argo" data-raccourci="r">${icon('bolt', 14)}Répondre avec ARGO ${kbd('r')}</button>
        <button type="button" class="tb-btn" data-raccourci="t">${icon('template', 14)}Modèle ${kbd('t')}</button>`}
        ${perso && !envoye ? (deCote ? '<button type="button" class="tb-btn" data-a="remettre">Remettre maintenant</button>' : `<button type="button" class="tb-btn" data-a="cote">${icon('clock', 14)}De côté ${kbd('s')}</button>`) : ''}
        ${envoye ? `<button type="button" class="tb-btn" data-a="relance">${icon('refresh', 14)}Relancer si pas de réponse…</button>` : ''}
        ${m.famille === 'mandats' ? '' : `<button type="button" class="tb-btn is-ghost" data-a="projet">${icon('folder', 14)}Projet ${kbd('l')}</button>`}
      </div>
    </div>
    <div class="tb-panel" id="tb-propositions"><div class="tb-h">L'agent propose</div><div class="tb-note">Chargement…</div></div>
    <div class="tb-panel" id="tb-dossier-projet" hidden></div>
    ${envoye ? '' : `<div class="tb-panel" id="tb-offre"></div>`}
    <div class="tb-panel" id="tb-classer"></div>
    <div class="tb-panel" id="tb-parite"></div>
    <div class="tb-panel" id="tb-composer" hidden></div>
    <div class="tb-panel" id="tb-equipe" hidden></div>`;

  // Sécurité (backlog reczJ0zLhPXqkWF9S) : bandeau rouge en tête quand le mail est à vérifier ou à risque élevé.
  const secu = host.querySelector<HTMLElement>('#tb-securite');
  if (secu && !envoye) void renderSecurite(secu, { messageId: m.messageId, mailbox: m.mailbox, onInfo: (t, k) => toast(t, k || 'info') });
  const on = (sel: string, f: () => void) => host.querySelector<HTMLElement>(sel)?.addEventListener('click', f);
  on('[data-a="ouvrir"]', () => ctx.actions.ouvrir(m));
  on('[data-a="cote"]', () => void ctx.actions.deCote(m));
  on('[data-a="remettre"]', () => void ctx.actions.remettre(m));
  // 07/10/2026 (parité Inbox ATLAS) : après le rattachement, dossier Outlook du projet (existant n'importe
  // où dans l'arbre, sinon « Créer le dossier et ranger » au même endroit et avec le même nom qu'ATLAS).
  on('[data-a="projet"]', async () => {
    const projetId = await ctx.actions.lierProjet(m);
    const zone = host.querySelector<HTMLElement>('#tb-dossier-projet');
    if (projetId && zone) void renderDossierProjet(zone, { messageId: m.messageId, mailbox: m.mailbox, onInfo: toast, projetId });
  });
  on('[data-a="relance"]', async () => {
    const jours = await choisirJours();
    if (!jours) return;
    try { const r = await demanderRelance(m.messageId, jours); toast(r.relanceLe ? `Relance prévue le ${r.relanceLe.slice(8, 10)}/${r.relanceLe.slice(5, 7)} sans réponse` : 'Relance prévue', 'success'); }
    catch (e) { toast(humanError(e), 'error'); }
  });
  on('[data-raccourci="r"]', () => void composer(host, m, ctx, 'argo'));
  on('[data-raccourci="t"]', () => void composer(host, m, ctx, 'modele'));

  // Mandat / association : aucune action commerciale proposée.
  if (m.famille === 'mandats') host.querySelector('#tb-propositions')!.innerHTML = '<div class="tb-h">Mandat / association</div><p class="tb-note">Ni projet, ni client, ni prospect : aucune action commerciale proposée.</p>';
  else void propositions(host.querySelector('#tb-propositions')!, m);
  const offre = host.querySelector<HTMLElement>('#tb-offre');
  if (offre && m.famille !== 'mandats') renderOffreRecue(offre, { messageId: m.messageId, mailbox: m.mailbox, onInfo: toast });
  else if (offre) offre.hidden = true;
  const classer = host.querySelector<HTMLElement>('#tb-classer');
  if (classer) renderClasserOutlook(classer, { messageId: m.messageId, mailbox: m.mailbox, onInfo: toast });
  // 07/10/2026 (fin de la parité Inbox ATLAS) : reclasser, pièces → projet, RDV (déposé en brouillon), tiers, prospection.
  const parite = host.querySelector<HTMLElement>('#tb-parite');
  if (parite) {
    const [prenomExp, ...nomExp] = String(m.from?.name || '').trim().split(/\s+/);
    renderPlusActions(parite, {
      messageId: m.messageId, mailbox: m.mailbox, onInfo: toast,
      deposer: async (texte: string) => {
        const d = await deposerBrouillon(m.messageId, m.mailbox, texte);
        if (!d.depose) { const ok = await copierTexte(texte); toast(ok ? 'Dépôt indisponible (double verrou) : texte copié' : 'Dépôt indisponible (double verrou)', 'info'); return; }
        toast('Proposition de RDV déposée dans le brouillon de réponse', 'success');
        if (d.webLink) ctx.openLink(d.webLink);
      },
      prefillTiers: { contactPrenom: prenomExp || '', contactNom: nomExp.join(' '), contactEmail: m.from?.email || '' },
    });
  }
  if (!perso) void equipe(host.querySelector('#tb-equipe')!, m, ctx);
}

async function choisirJours(): Promise<number | null> {
  return new Promise(resolve => {
    const el = document.createElement('div');
    el.className = 'tb-modal';
    el.innerHTML = `<h3>Relancer si pas de réponse sous…</h3><div class="tb-presets">${[2, 3, 5, 10].map(n => `<button type="button" class="tb-btn" data-j="${n}">${n} jours ouvrés</button>`).join('')}</div><p class="tb-note">ARGO prépare la relance dans ton style ; rien n'est envoyé sans toi.</p>`;
    let v: number | null = null;
    const fermer = ouvrirCouche(el, () => resolve(v));
    el.querySelectorAll<HTMLButtonElement>('[data-j]').forEach(b => b.addEventListener('click', () => { v = Number(b.dataset.j); fermer(); }));
    el.querySelector<HTMLButtonElement>('[data-j]')?.focus();
  });
}

/**
 * « L'agent propose » : mêmes cartes que le panneau du mail (agent-outils › renderActionsMetier),
 * choix compris (projet, demande de devis, pièce jointe, dossier Outlook) et « Créer le dossier »
 * (07/10/2026 : plus de renvoi vers le panneau du mail pour un choix).
 */
async function propositions(host: HTMLElement, m: TableauMail): Promise<void> {
  host.innerHTML = '<div data-cartes></div>';
  const cartes = host.querySelector<HTMLElement>('[data-cartes]')!;
  await renderActionsMetier(cartes, { messageId: m.messageId, mailbox: m.mailbox, onInfo: toast });
  if (cartes.hidden) host.innerHTML = '<div class="tb-h">L\'agent propose</div><p class="tb-note">Rien de plus à faire dans ATLAS pour ce mail.</p>';
}

// ── Réponse : ARGO ou modèle, déposée dans le brouillon Outlook ──

async function choisirModele(): Promise<ModeleReponse | null> {
  let modeles: ModeleReponse[] = [];
  try { modeles = await fetchModeles(); } catch (e) { toast(humanError(e), 'error'); return null; }
  if (!modeles.length) { toast('Aucun modèle « Réponse type » dans Communications', 'info'); return null; }
  return new Promise(resolve => {
    let choisi: ModeleReponse | null = null;
    const el = document.createElement('div');
    el.className = 'tb-palette';
    el.innerHTML = '<input type="text" placeholder="Modèle de réponse (Communications)…" aria-label="Modèle"><ul role="listbox"></ul>';
    const input = el.querySelector('input')!, ul = el.querySelector('ul')!;
    const cmds = modeles.map(x => ({ id: x.id, libelle: x.nom, groupe: x.categorie || 'Modèles', motsCles: `${x.objet} ${x.apercu} ${x.destinataires.join(' ')}`, m: x }));
    let vis = cmds, idx = 0;
    const dessiner = () => {
      vis = filtrerCommandes(cmds, input.value, 14);
      idx = Math.min(idx, Math.max(0, vis.length - 1));
      ul.innerHTML = vis.map((c, i) => `<li role="option" data-i="${i}" aria-selected="${i === idx}"><span>${h(c.libelle)}</span><small>${h(c.m.categorie || '')}</small></li>`).join('') || '<li class="tb-grp">Aucun modèle</li>';
    };
    const fermer = ouvrirCouche(el, () => resolve(choisi));
    const valider = (i: number) => { if (vis[i]) { choisi = vis[i].m; fermer(); } };
    input.addEventListener('input', () => { idx = 0; dessiner(); });
    input.addEventListener('keydown', e => {
      if (e.key === 'ArrowDown') { e.preventDefault(); idx = Math.min(vis.length - 1, idx + 1); dessiner(); }
      else if (e.key === 'ArrowUp') { e.preventDefault(); idx = Math.max(0, idx - 1); dessiner(); }
      else if (e.key === 'Enter') { e.preventDefault(); valider(idx); }
    });
    ul.addEventListener('click', e => { const li = (e.target as HTMLElement).closest<HTMLElement>('[data-i]'); if (li) valider(Number(li.dataset.i)); });
    dessiner();
    input.focus();
  });
}

const SOURCES: Record<ReponseRedigee['source'], string> = {
  agent: 'Préparé par l\'agent dans ton style',
  argo: 'Rédigé par ARGO dans ton style (profil de conversation avec ce contact)',
  modele: 'Modèle Communications rempli pour ce contact',
};

/** Ouvre la zone de réponse : ARGO ou modèle ; aussi appelée depuis « Engagements » avec un texte prêt. */
export async function composer(host: HTMLElement, m: TableauMail, ctx: Ctx, source: 'argo' | 'modele', pret?: ReponseRedigee): Promise<void> {
  const zone = host.querySelector<HTMLElement>('#tb-composer');
  if (!zone) return;
  zone.hidden = false;
  zone.innerHTML = `<div class="tb-h">${source === 'argo' ? 'Réponse ARGO' : 'Réponse par modèle'}</div><div class="tb-skel"></div>`;
  zone.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
  let r: ReponseRedigee | null = pret || null;
  try {
    if (!r) {
      if (source === 'argo') r = await redigerArgo(m.messageId, m.mailbox);
      else {
        const modele = await choisirModele();
        if (!modele) { zone.hidden = true; return; }
        r = await remplirModeleTableau(m.messageId, m.mailbox, modele.id);
      }
    }
  } catch (e) {
    zone.innerHTML = `<div class="tb-h">Réponse</div><p class="tb-warn">${h(humanError(e))}</p>`;
    return;
  }
  const actives = !!ctx.etat.t?.ecrituresActives;
  const manque = [...(r.aCompleter || []), ...(r.manquantes || []).map(x => `{{${x}}}`)];
  zone.innerHTML = `
    <div class="tb-h">Réponse<span class="tb-h-n">${h(SOURCES[r.source])}</span></div>
    ${r.resumeIntention ? `<p class="tb-note">${h(r.resumeIntention)}</p>` : ''}
    ${manque.length ? `<p class="tb-warn">À compléter : ${manque.map(h).join(' · ')}</p>` : ''}
    <textarea class="tb-textarea" id="tb-texte" aria-label="Texte de la réponse">${h(r.texte)}</textarea>
    <label class="tb-note"><input type="checkbox" id="tb-tous"> Répondre à tous</label>
    <div class="tb-actions">
      ${actives ? `<button type="button" class="tb-btn is-primary" id="tb-deposer">${icon('reply', 14)}Déposer dans Outlook</button>` : ''}
      <button type="button" class="tb-btn" id="tb-copier">${icon('copy', 14)}Copier</button>
      <button type="button" class="tb-btn is-ghost" id="tb-fermer-rep">Fermer</button>
    </div>
    ${actives ? '' : '<p class="tb-note">Le dépôt direct dans Outlook n\'est pas encore activé : copie le texte, ouvre le mail (Entrée) et colle-le dans ta réponse.</p>'}
    <div id="tb-apres"></div>`;
  const texte = () => (zone.querySelector('#tb-texte') as HTMLTextAreaElement).value;
  zone.querySelector('#tb-copier')?.addEventListener('click', async () => {
    const ok = await copierTexte(texte(), zone.querySelector<HTMLTextAreaElement>('#tb-texte'));
    toast(ok ? 'Texte copié : colle-le dans ta réponse Outlook' : 'Copie impossible : sélectionne le texte', ok ? 'success' : 'error');
  });
  zone.querySelector('#tb-fermer-rep')?.addEventListener('click', () => { zone.hidden = true; });
  zone.querySelector('#tb-deposer')?.addEventListener('click', async ev => {
    const b = ev.currentTarget as HTMLButtonElement;
    b.disabled = true;
    try {
      const d = await deposerBrouillon(m.messageId, m.mailbox, texte(), (zone.querySelector('#tb-tous') as HTMLInputElement).checked);
      if (!d.depose) { toast('Dépôt indisponible : texte à copier', 'info'); b.disabled = false; return; }
      toast('Brouillon prêt dans Outlook', 'success');
      apresDepot(zone.querySelector('#tb-apres')!, m, ctx, d.brouillonId || '', d.webLink || '');
    } catch (e) { toast(humanError(e), 'error'); b.disabled = false; }
  });
  (zone.querySelector('#tb-texte') as HTMLTextAreaElement).focus();
}

/** Après le dépôt : ouvrir le brouillon, ou l'envoyer plus tard (remise différée Exchange) avec relance. */
function apresDepot(host: HTMLElement, m: TableauMail, ctx: Ctx, brouillonId: string, webLink: string): void {
  const perso = (m.mailbox || '').toLowerCase() === ctx.etat.t?.moi;
  host.innerHTML = `<div class="tb-sep"></div><div class="tb-actions">
    ${webLink ? `<button type="button" class="tb-btn" id="tb-ouvrir-br">${icon('external', 14)}Ouvrir le brouillon</button>` : ''}
    ${perso && brouillonId ? `<button type="button" class="tb-btn is-primary" id="tb-plus-tard">${icon('clock', 14)}Envoyer plus tard…</button>` : ''}
  </div>${perso ? '' : '<p class="tb-note">Boîte partagée : relis et envoie depuis Outlook.</p>'}`;
  host.querySelector('#tb-ouvrir-br')?.addEventListener('click', () => ctx.openLink(webLink));
  host.querySelector('#tb-plus-tard')?.addEventListener('click', async () => {
    const choix = await choisirMoment<number | undefined>('Envoyer plus tard', {
      envoi: true,
      extra: `<label class="tb-note"><input type="checkbox" id="tb-rel"> Relancer si pas de réponse sous <select id="tb-rel-j" class="tb-input">${[2, 3, 5, 10].map(n => `<option value="${n}"${n === 3 ? ' selected' : ''}>${n} jours ouvrés</option>`).join('')}</select></label>`,
      lireExtra: el => ((el.querySelector('#tb-rel') as HTMLInputElement)?.checked ? Number((el.querySelector('#tb-rel-j') as HTMLSelectElement).value) : undefined),
    });
    if (!choix) return;
    try {
      const r = await programmerEnvoi({ brouillonId, quand: choix.iso, relanceJours: choix.extra, enReponseA: m.messageId });
      toast(`Envoi programmé ${quandLisible(r.envoi.quand, Date.now())} : Exchange l'enverra même si ATLAS est fermé`, 'success');
      host.innerHTML = `<p class="tb-note">${icon('clock', 13)} Part ${h(quandLisible(r.envoi.quand, Date.now()))}${choix.extra ? `, relance sous ${choix.extra} jours ouvrés sans réponse` : ''}.</p>`;
      void ctx.rafraichir(true);
    } catch (e) { toast(humanError(e), 'error'); }
  });
}

// ── Fil d'équipe (boîtes partagées) ──

async function equipe(host: HTMLElement, m: TableauMail, ctx: Ctx): Promise<void> {
  let v: Awaited<ReturnType<typeof fetchEquipe>>;
  try { v = await fetchEquipe(m.messageId, m.mailbox); } catch { return; }
  if (!v) return;
  const e = v.equipe;
  const moi = v.moi;
  const prisParMoi = e.prisPar === moi;
  host.hidden = false;
  host.innerHTML = `
    <div class="tb-h">Équipe · ${h(v.mailbox.split('@')[0])}@</div>
    <p>${e.prisPar ? `${icon('user', 14)} Pris par <b>${h(e.prisParNom || prenomDe(e.prisPar))}</b>` : e.assigneA ? `Pour <b>${h(e.assigneNom || prenomDe(e.assigneA))}</b>` : 'À attribuer'}${e.motif ? ` <span class="tb-note">· ${h(e.motif)}</span>` : ''}</p>
    ${v.collision ? `<p class="tb-warn">${h(v.collision.message)}</p>` : ''}
    <div class="tb-actions">
      ${prisParMoi ? '<button type="button" class="tb-btn" data-e="relacher">Relâcher</button>' : e.prisPar ? '' : `<button type="button" class="tb-btn is-primary" data-e="prendre">Je prends ${kbd('p')}</button>`}
      <select class="tb-input" data-e="attribuer" aria-label="Attribuer à"><option value="">Attribuer à…</option>${v.membres.filter(x => x.email !== moi).map(x => `<option value="${h(x.email)}">${h(x.nom)}</option>`).join('')}</select>
    </div>
    <div class="tb-sep"></div>
    <div class="tb-thread" id="tb-thread">${(e.commentaires || []).map(c => `<div class="tb-com"><b>${h(c.auteurNom || prenomDe(c.auteur))}</b><time>${h(ilYA(c.at, Date.now()))}</time><div>${h(c.texte)}</div></div>`).join('') || '<p class="tb-note">Aucun commentaire. Visible des membres de la boîte seulement, jamais dans le mail.</p>'}</div>
    <div class="tb-actions"><input class="tb-input" style="flex:1" id="tb-com" maxlength="1000" placeholder="Commentaire interne…" aria-label="Commentaire interne"><button type="button" class="tb-btn" data-raccourci="c">Envoyer ${kbd('c')}</button></div>`;
  host.querySelector('[data-e="prendre"]')?.addEventListener('click', () => void ctx.actions.jePrends(m).then(() => equipe(host, m, ctx)));
  host.querySelector('[data-e="relacher"]')?.addEventListener('click', async () => {
    try { await relacherMail(m.messageId, m.mailbox); toast('Relâché', 'success'); void ctx.rafraichir(true); void equipe(host, m, ctx); } catch (err) { toast(humanError(err), 'error'); }
  });
  host.querySelector<HTMLSelectElement>('[data-e="attribuer"]')?.addEventListener('change', async ev => {
    const a = (ev.target as HTMLSelectElement).value;
    if (!a) return;
    try {
      const r = await attribuerMail(m.messageId, m.mailbox, a);
      toast(r.ok ? `Attribué à ${prenomDe(a)} : prévenu par une cloche` : (r.error || 'Attribution impossible'), r.ok ? 'success' : 'error');
      void ctx.rafraichir(true); void equipe(host, m, ctx);
    } catch (err) { toast(humanError(err), 'error'); }
  });
  const input = host.querySelector<HTMLInputElement>('#tb-com')!;
  const envoyer = async () => {
    const texte = input.value.trim();
    if (!texte) { input.focus(); return; }
    try { await ajouterCommentaire(m.messageId, m.mailbox, texte); input.value = ''; void equipe(host, m, ctx); } catch (err) { toast(humanError(err), 'error'); }
  };
  host.querySelector('[data-raccourci="c"]')?.addEventListener('click', () => (input.value.trim() ? void envoyer() : input.focus()));
  input.addEventListener('keydown', ev => { if (ev.key === 'Enter') { ev.preventDefault(); void envoyer(); } });
}

// ── Digest d'une newsletter ──

export function renderDigest(host: HTMLElement, d: DigestNewsletter, ctx: Ctx): void {
  host.innerHTML = `<div class="tb-panel is-raised">
    <div class="tb-detail-h"><h3>${h(d.nom)}</h3><div class="tb-detail-meta"><span>${h(d.expediteur)}</span><span>${d.nb} mail${d.nb > 1 ? 's' : ''} en 14 jours</span></div></div>
    <div class="tb-actions">${d.desinscription ? `<button type="button" class="tb-btn" id="tb-desinscr">${icon('x', 14)}Se désinscrire</button>` : ''}</div>
    <div>${d.mails.map((m, i) => `<div class="tb-eng"><p>${h(m.subject || '(sans objet)')}<br><small>${h(m.resume || '')}</small></p><button type="button" class="tb-btn is-ghost" data-i="${i}">${icon('external', 14)}</button></div>`).join('')}</div>
  </div>`;
  host.querySelector('#tb-desinscr')?.addEventListener('click', () => d.desinscription && ctx.openLink(d.desinscription));
  host.querySelectorAll<HTMLButtonElement>('[data-i]').forEach(b => b.addEventListener('click', () => ctx.actions.ouvrir(d.mails[Number(b.dataset.i)])));
}
