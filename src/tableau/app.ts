/**
 * app.ts · tableau de bord pleine largeur de l'application ATLAS (barre de gauche d'Outlook),
 * section B du cadrage du 06/10/2026 (`.claude/CADRAGE-OUTLOOK-DASHBOARD.md`, B, C et D).
 *
 * Sans framework (TypeScript + DOM, bundle Vite du complément, aucun CDN en plus de TeamsJS) :
 *  - boîte triée, toutes les boîtes permises (personnes et clients d'abord, notifications,
 *    newsletters en digest, factures), priorités ARGO expliquées, en attente, mis de côté ;
 *  - clavier d'abord : j/k, Entrée (ouvrir dans Outlook), r (réponse ARGO), t (modèle), e / s
 *    (07/10/2026 : x coche le mail, Maj-clic une plage, cases à cocher : rattacher le LOT à un projet
 *    ou à un contact, comme la liaison en lot de l'Inbox ATLAS),
 *    (mettre de côté), p (je prends), l (rattacher à un projet), c (commentaire), 1 à 8 (sections),
 *    ⌘K / Ctrl+K ou / (palette), ? (aide), Échap ;
 *  - glisser-déposer d'un mail vers « Je prends », une date ou un projet ; aperçu au survol ;
 *  - mise à jour en temps réel : empreinte relue toutes les 20 s (onglet visible), tableau
 *    rechargé seulement si elle change, nouveautés signalées, sélection conservée ;
 *  - colonne de droite : le mail choisi (détail, propositions de l'agent, réponse ARGO / modèle
 *    déposés dans le brouillon Outlook, envoi programmé, fil d'équipe), sinon « Aujourd'hui »
 *    (prochains RDV préparés, engagements, réactivité aux clients).
 */
import { fetchTableau, fetchVersion, type Tableau, type TableauMail, type DigestNewsletter } from '../api/tableau';
import { mettrePlusTard, retirerPlusTard, prendreMail, relacherMail } from '../api/agent';
import { autoriserConnexionInteractive, callAtlasWorker, changerDeCompte, compteConnecte } from '../api/worker';
import { getAllProjets, getAllContacts } from '../api/airtable';
import { openAgentMail } from '../components/agent-lists';
import { AtlasError, humanError } from '../api/net';
import { icon } from '../ui/icons';
import { deplacer, ilYA, initiales, nouveautes, quandLisible, presetsMoments, prenomDe, basculerSelection, selectionPlage, selectionValide, resumeLot, sousTitreEnAttente } from './logique';
import { h, toast, animerChiffre, choisirMoment, ouvrirCouche, mouvementReduit } from './ui';
import { renderDetail, renderDigest } from './detail';
import { renderAujourdhui } from './aujourdhui';
import { ouvrirPalette, type ActionPalette } from './palette';
import { ouvrirSeance } from './seance';
import { effacerReprise, lireReprise, messageIdDeCle, sauverReprise, selectionReprise, suivantsDe, type Reprise, type RepriseSeance } from './reprise';
import { fetchStatsSeance, fetchStatistiques, type SeanceSemaine, type StatistiquesBoite } from '../api/seance';

export type Section = 'priorites' | 'personnes' | 'clients' | 'mandats' | 'enAttente' | 'deCote' | 'factures' | 'newsletters' | 'notifications';

export const SECTIONS: Array<{ id: Section; libelle: string; titre: string; vide: string }> = [
  { id: 'priorites', libelle: 'Priorités ARGO', titre: 'Priorités ARGO', vide: 'Rien d\'urgent : ARGO ne voit aucun mail qui doive passer avant le reste.' },
  { id: 'personnes', libelle: 'Personnes', titre: 'Personnes et clients', vide: 'Tous les mails de personnes ont une réponse. Rien n\'attend de toi.' },
  { id: 'clients', libelle: 'Clients', titre: 'Clients actifs', vide: 'Aucun client n\'attend de réponse.' },
  { id: 'mandats', libelle: 'Mandats', titre: 'Mandats et associations', vide: 'Aucun mail de mandat ou d\'association.' },
  { id: 'enAttente', libelle: 'En attente', titre: 'En attente d\'une réponse', vide: 'Aucun fil n\'attend de réponse de l\'autre côté.' },
  { id: 'deCote', libelle: 'De côté', titre: 'Mis de côté', vide: 'Rien de mis de côté. Glisse un mail vers une date pour le faire revenir plus tard.' },
  { id: 'factures', libelle: 'Factures', titre: 'Factures', vide: 'Aucune facture dans la boîte de réception (14 derniers jours).' },
  { id: 'newsletters', libelle: 'Newsletters', titre: 'Newsletters (digest)', vide: 'Aucune newsletter : la boîte est calme.' },
  { id: 'notifications', libelle: 'Notifications', titre: 'Notifications', vide: 'Aucune notification de service.' },
];

export interface Etat {
  boite: string;
  section: Section;
  t: Tableau | null;
  /** Clé du mail choisi : `${mailbox}|${messageId}` ; digest : `nl|<expéditeur>`. */
  sel: string | null;
  vus: string[];
  derniereSync: number;
  enLigne: boolean;
  /** Mails cochés pour une action en lot (clés `${mailbox}|${messageId}`) et ancre du Maj-clic. */
  lot: string[];
  ancre: string | null;
}

export interface Ctx {
  etat: Etat;
  openLink: (url: string) => void;
  rafraichir: (force?: boolean) => Promise<void>;
  choisir: (cle: string | null) => void;
  /** Actions communes (clavier, glisser-déposer, détail, palette). */
  actions: ReturnType<typeof creerActions>;
}

const POLL_MS = 20_000;
const cleDe = (m: { mailbox?: string; messageId: string }) => `${(m.mailbox || '').toLowerCase()}|${m.messageId}`;

export function lignesDe(t: Tableau | null, s: Section): TableauMail[] {
  if (!t) return [];
  switch (s) {
    case 'priorites': return t.priorites;
    case 'personnes': return t.personnes;
    case 'clients': return t.personnes.filter(m => m.correspondant === 'client');
    case 'mandats': return t.mandats || [];
    case 'enAttente': return t.enAttente.map(e => ({ ...e, mailbox: e.mailbox || t.moi, conversationId: e.conversationId || '', famille: 'personnes', priorite: 0, raisons: e.envoye ? ['Envoyé, sans réponse'] : [] }) as TableauMail);
    case 'deCote': return t.deCote;
    case 'factures': return t.factures;
    case 'notifications': return t.notifications;
    case 'newsletters': return [];
  }
}

const compteDe = (t: Tableau, s: Section): number => (s === 'newsletters' ? t.newsletters.length : s === 'mandats' ? (t.mandats || []).length : t.compteurs[s as keyof Tableau['compteurs']] ?? 0);

/** Sections affichées : « Mandats » seulement s'il y en a (classement de l'agent, repli propre). */
export function sectionsVisibles(t: Tableau | null): typeof SECTIONS {
  return SECTIONS.filter(s => s.id !== 'mandats' || (t ? compteDe(t, 'mandats') > 0 : false));
}

// ── Actions communes ───────────────────────────────────────────────────────────────

function creerActions(ctx: () => Ctx) {
  const c = () => ctx();
  const personnelle = (m: TableauMail) => (m.mailbox || '').toLowerCase() === (c().etat.t?.moi || '');
  return {
    ouvrir(m: TableauMail) {
      if (!openAgentMail(m, c().openLink)) toast('Lien du mail indisponible : ouvre-le depuis Outlook', 'error');
    },
    async deCote(m: TableauMail, iso?: string) {
      if (!personnelle(m)) { toast('« Mettre de côté » : mails de ta boîte personnelle seulement', 'error'); return; }
      const quand = iso || (await choisirMoment('Mettre de côté jusqu\'à…'))?.iso;
      if (!quand) return;
      try {
        await mettrePlusTard(m.messageId, quand);
        toast(`De côté jusqu'à ${quandLisible(quand, Date.now())}`, 'success', async () => {
          try { await retirerPlusTard(m.messageId); await c().rafraichir(true); } catch (e) { toast(humanError(e), 'error'); }
        });
        await c().rafraichir(true);
      } catch (e) { toast(humanError(e), 'error'); }
    },
    async remettre(m: TableauMail) {
      try { await retirerPlusTard(m.messageId); toast('Revenu dans la boîte', 'success'); await c().rafraichir(true); } catch (e) { toast(humanError(e), 'error'); }
    },
    async jePrends(m: TableauMail) {
      try {
        const r = await prendreMail(m.messageId, m.mailbox);
        if (!r.ok) { toast(r.equipe?.prisParNom ? `Déjà pris par ${r.equipe.prisParNom}` : (r.error || 'Impossible de prendre ce mail'), 'error'); return; }
        toast('Tu prends ce mail : l\'équipe est prévenue', 'success', async () => { try { await relacherMail(m.messageId, m.mailbox); await c().rafraichir(true); } catch { /* rien */ } });
        await c().rafraichir(true);
      } catch (e) { toast(humanError(e), 'error'); }
    },
    /**
     * Liaison EN LOT (parité handleBulkLinkToProject / handleBulkLinkToContact de l'Inbox ATLAS) : même
     * route que le rattachement d'un mail (relecture Graph, sans doublon), un mail après l'autre avec la
     * progression ; un échec n'arrête pas le lot. Renvoie true si au moins un mail a été rattaché.
     */
    async lierLot(mails: TableauMail[], quoi: 'projet' | 'contact', progres: (texte: string) => void): Promise<boolean> {
      const cible = quoi === 'projet' ? await choisirProjetLibelle() : await choisirContact();
      if (!cible) return false;
      let ok = 0, ko = 0;
      for (let i = 0; i < mails.length; i++) {
        const m = mails[i];
        progres(`${i + 1}/${mails.length}…`);
        try {
          const email = { id: m.graphId, internetMessageId: m.messageId, subject: m.subject, from: m.from, receivedAt: m.receivedAt };
          if (quoi === 'projet') await callAtlasWorker('emails/link-projet', { email, projetId: cible.id, linkedByName: prenomDe(c().etat.t?.moi || ''), direction: 'reçu' });
          else await callAtlasWorker('emails/link-contact', { email, contactName: cible.id, tiersName: cible.tiers || undefined, linkedByName: prenomDe(c().etat.t?.moi || ''), direction: 'reçu' });
          ok++;
        } catch { ko++; }
      }
      toast(resumeLot(ok, ko, cible.libelle), ko ? (ok ? 'info' : 'error') : 'success');
      return ok > 0;
    },
    /** Rattache le mail à un projet ; renvoie le projet (null si rien n'est fait) pour proposer son dossier Outlook. */
    async lierProjet(m: TableauMail): Promise<string | null> {
      if (m.famille === 'mandats') { toast('Mandat / association : ni projet ni client', 'info'); return null; }
      const projetId = await choisirProjet();
      if (!projetId) return null;
      try {
        await callAtlasWorker('emails/link-projet', {
          email: { id: m.graphId, internetMessageId: m.messageId, subject: m.subject, from: m.from, receivedAt: m.receivedAt },
          projetId, linkedByName: prenomDe(c().etat.t?.moi || ''), direction: 'reçu',
        });
        toast('Mail rattaché au projet', 'success');
        return projetId;
      } catch (e) { toast(humanError(e), 'error'); return null; }
    },
  };
}

/** Choix d'un projet (liste ATLAS, recherche) : recId ou null. */
async function choisirProjet(): Promise<string | null> {
  let projets: Awaited<ReturnType<typeof getAllProjets>> = [];
  try { projets = await getAllProjets(); } catch (e) { toast(humanError(e), 'error'); return null; }
  const actifs = projets.filter(p => !/clôtur|clotur|annul|refus|factur/i.test(p.statut || ''));
  return new Promise(resolve => {
    let choisi: string | null = null;
    const el = document.createElement('div');
    el.className = 'tb-palette';
    el.innerHTML = `<input type="text" placeholder="Rattacher à un projet : numéro, nom, client…" aria-label="Projet"><ul role="listbox"></ul>`;
    const input = el.querySelector('input')!, ul = el.querySelector('ul')!;
    let idx = 0, vis = actifs.slice(0, 12);
    const dessiner = () => {
      const q = input.value.trim().toLowerCase();
      vis = (q ? actifs.filter(p => `${p.noProjet} ${p.refProjet} ${p.denomination} ${p.client}`.toLowerCase().includes(q)) : actifs).slice(0, 12);
      idx = Math.min(idx, Math.max(0, vis.length - 1));
      ul.innerHTML = vis.length
        ? vis.map((p, i) => `<li role="option" data-i="${i}" aria-selected="${i === idx}">${icon('folder', 14)}<span>#${h(p.noProjet || p.refProjet)} ${h(p.denomination)}</span><small>${h(p.client)}</small></li>`).join('')
        : '<li class="tb-grp">Aucun projet actif ne correspond</li>';
    };
    const fermer = ouvrirCouche(el, () => resolve(choisi));
    const valider = (i: number) => { const p = vis[i]; if (!p) return; choisi = p.id; fermer(); };
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

/** Projet choisi avec son libellé (liaison en lot). */
async function choisirProjetLibelle(): Promise<{ id: string; libelle: string; tiers?: string } | null> {
  const id = await choisirProjet();
  if (!id) return null;
  const p = (await getAllProjets().catch(() => [])).find(x => x.id === id);
  return { id, libelle: p ? `#${p.noProjet || p.refProjet} ${p.denomination}` : 'ce projet' };
}

/** Contact choisi (liste ATLAS, recherche) : `id` = nom de la fiche (« Personne de contact »), comme la liaison de l'Inbox. */
async function choisirContact(): Promise<{ id: string; libelle: string; tiers?: string } | null> {
  let contacts: Awaited<ReturnType<typeof getAllContacts>> = [];
  try { contacts = (await getAllContacts()).filter(x => x.personneDeContact); } catch (e) { toast(humanError(e), 'error'); return null; }
  return new Promise(resolve => {
    let choisi: { id: string; libelle: string; tiers?: string } | null = null;
    const el = document.createElement('div');
    el.className = 'tb-palette';
    el.innerHTML = `<input type="text" placeholder="Rattacher à un contact : nom, société, e-mail…" aria-label="Contact"><ul role="listbox"></ul>`;
    const input = el.querySelector('input')!, ul = el.querySelector('ul')!;
    let idx = 0, vis = contacts.slice(0, 12);
    const tiersDe = (x: typeof contacts[number]) => Array.isArray(x.relationSociete) ? String(x.relationSociete[0] || '') : String(x.relationSociete || '');
    const dessiner = () => {
      const q = input.value.trim().toLowerCase();
      vis = (q ? contacts.filter(x => `${x.personneDeContact} ${tiersDe(x)} ${x.email}`.toLowerCase().includes(q)) : contacts).slice(0, 12);
      idx = Math.min(idx, Math.max(0, vis.length - 1));
      ul.innerHTML = vis.length
        ? vis.map((x, i) => `<li role="option" data-i="${i}" aria-selected="${i === idx}">${icon('user', 14)}<span>${h(x.personneDeContact)}</span><small>${h(tiersDe(x) || x.email || '')}</small></li>`).join('')
        : '<li class="tb-grp">Aucun contact ne correspond</li>';
    };
    const fermer = ouvrirCouche(el, () => resolve(choisi));
    const valider = (i: number) => { const x = vis[i]; if (!x) return; choisi = { id: x.personneDeContact, libelle: x.personneDeContact, tiers: tiersDe(x) || undefined }; fermer(); };
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

// ── Application ────────────────────────────────────────────────────────────────────

/**
 * Écran tactile sans souris (Outlook iPhone / Android, tablette) : pas de glisser-déposer (le dock
 * n'a pas de sens au doigt), pas d'aperçu au survol, pas d'aides clavier (07/10/2026).
 */
export function ecranTactile(): boolean {
  try { return window.matchMedia('(hover: none) and (pointer: coarse)').matches; } catch { return false; }
}

export function demarrerTableau(root: HTMLElement, opts: { openLink: (url: string) => void; ouvrirDansNavigateur?: () => void; repondreDansOutlook?: (messageId: string, html: string) => Promise<boolean> }): void {
  const tactile = ecranTactile();
  let boiteMemo = 'toutes';
  try { boiteMemo = localStorage.getItem('atlas.tdb.boite') || 'toutes'; } catch { /* stockage indisponible */ }
  // Reprise (08/10/2026) : fenêtre rouverte dans les 30 minutes → même boîte, même section, même mail ou séance.
  let repriseAttendue: Reprise | null = lireReprise();
  const sectionReprise = repriseAttendue && SECTIONS.some(x => x.id === repriseAttendue!.section) ? repriseAttendue.section as Section : 'priorites';
  const etat: Etat = { boite: repriseAttendue?.boite || boiteMemo, section: sectionReprise, t: null, sel: null, vus: [], derniereSync: 0, enLigne: true, lot: [], ancre: null };
  let seanceCourante: RepriseSeance | null = null;
  /** Note où la personne en est (section, mail choisi et ses suivants, séance), pour la reprise. */
  function sauver(): void {
    sauverReprise({ boite: etat.boite, section: etat.section, sel: etat.sel, suivants: suivantsDe(etat.sel, lignesDe(etat.t, etat.section).map(cleDe)), seance: seanceCourante });
  }
  let ctx: Ctx;
  const actions = creerActions(() => ctx);
  ctx = { etat, openLink: opts.openLink, rafraichir, choisir, actions };

  document.body.classList.add('tb');
  document.body.classList.toggle('tb-tactile', tactile);
  root.innerHTML = `
    <div class="tb-app" id="tb-app">
      <header class="tb-top">
        <span class="tb-mark">ATLAS</span>
        <div class="tb-boites" id="tb-boites" role="group" aria-label="Boîtes"></div>
        <span class="tb-spacer"></span>
        <button type="button" class="tb-btn is-primary tb-seance-btn" id="tb-seance-btn" title="Séance de tri : une carte par mail, l'action est déjà préparée">${icon('inbox', 14)}<span>Séance de tri</span></button>
        <button type="button" class="tb-cmdk" id="tb-cmdk" aria-label="Rechercher ou agir">${icon('search', 14)}<span>Rechercher, agir…</span><span class="tb-spacer"></span><span class="tb-kbd">⌘K</span></button>
        <span class="tb-live" id="tb-live" aria-live="polite"><i></i><span>connexion…</span></span>
        <span class="tb-compte" id="tb-compte" hidden><span class="tb-compte-nom" id="tb-compte-nom"></span><button type="button" class="tb-lien" id="tb-compte-changer">Changer de compte</button></span>
        <button type="button" class="tb-iconbtn tb-seul-etroit" id="tb-jour" title="Aujourd'hui" aria-label="Aujourd'hui : rendez-vous, engagements, réactivité">${icon('clock', 16)}</button>
        <button type="button" class="tb-iconbtn tb-clavier" id="tb-aide" title="Raccourcis (?)" aria-label="Raccourcis clavier">${icon('question', 16)}</button>
        <button type="button" class="tb-iconbtn" id="tb-refresh" title="Actualiser" aria-label="Actualiser">${icon('refresh', 16)}</button>
      </header>
      <nav class="tb-stats" id="tb-stats" role="tablist" aria-label="Sections"></nav>
      <main class="tb-main">
        <aside class="tb-col is-left tb-scroll" id="tb-left"></aside>
        <section class="tb-list-wrap tb-scroll" id="tb-center" aria-label="Mails">
          <div class="tb-list-head"><h2 id="tb-titre"></h2><p id="tb-sous"></p></div>
          <div class="tb-lot" id="tb-lot" role="toolbar" aria-label="Mails cochés" hidden></div>
          <div class="tb-list" id="tb-list" role="listbox" aria-label="Mails"></div>
        </section>
        <aside class="tb-col is-right tb-scroll" id="tb-right"></aside>
      </main>
      <button type="button" class="tb-btn tb-retour" id="tb-retour" aria-label="Retour à la liste">${icon('chevron-right', 14)}<span>Retour</span></button>
    </div>
    <div class="tb-dock" id="tb-dock" aria-hidden="true"></div>
    <div class="tb-peek" id="tb-peek" aria-hidden="true"></div>`;

  const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

  // ── Rendu ──
  function renderBoites(): void {
    const t = etat.t;
    const boites = t?.boites || [];
    const host = $('tb-boites');
    if (boites.length < 2) { host.hidden = true; return; }
    host.hidden = false;
    const opts2 = ['toutes', ...boites];
    host.innerHTML = opts2.map(b => `<button type="button" data-b="${h(b)}" aria-pressed="${etat.boite === b}">${b === 'toutes' ? 'Toutes' : `${h(b.split('@')[0])}@`}</button>`).join('');
    host.querySelectorAll<HTMLButtonElement>('[data-b]').forEach(btn => btn.addEventListener('click', () => {
      etat.boite = btn.dataset.b!;
      try { localStorage.setItem('atlas.tdb.boite', etat.boite); } catch { /* rien */ }
      etat.sel = null;
      void rafraichir(true);
    }));
  }

  function renderStats(): void {
    const host = $('tb-stats');
    const t = etat.t;
    const vis = sectionsVisibles(t);
    if (host.dataset.ids !== vis.map(s => s.id).join(',')) {
      host.dataset.ids = vis.map(s => s.id).join(',');
      host.innerHTML = vis.map((s, i) => `<button type="button" role="tab" class="tb-stat${s.id === 'priorites' ? ' is-argo' : ''}" data-s="${s.id}" aria-selected="false"><span class="tb-stat-k">${i + 1}</span><span class="tb-stat-v" data-v="0">0</span><span class="tb-stat-l">${h(s.libelle)}</span></button>`).join('');
      host.querySelectorAll<HTMLButtonElement>('[data-s]').forEach(b => b.addEventListener('click', () => allerSection(b.dataset.s as Section)));
      // Glisser un mail sur « De côté » : choix de la date.
      const dc = host.querySelector<HTMLElement>('[data-s="deCote"]');
      if (dc) brancherDepot(dc, m => actions.deCote(m));
    }
    host.querySelectorAll<HTMLButtonElement>('[data-s]').forEach(b => {
      const s = b.dataset.s as Section;
      b.setAttribute('aria-selected', String(s === etat.section));
      const v = t ? compteDe(t, s) : 0;
      animerChiffre(b.querySelector('.tb-stat-v')!, v);
      b.classList.toggle('is-hot', (s === 'clients' && v > 0) || (s === 'personnes' && (t?.personnes.some(m => m.urgence >= 3) ?? false)));
    });
  }

  function renderLeft(): void {
    const t = etat.t;
    const host = $('tb-left');
    const qui = t?.quiTraite || [];
    const max = Math.max(1, ...qui.map(q => q.mails));
    host.innerHTML = `
      <div class="tb-panel"><div class="tb-h">Sections</div><nav class="tb-nav">${sectionsVisibles(t).map(s => `<button type="button" data-s="${s.id}" aria-current="${s.id === etat.section}">${h(s.libelle)}<span class="n">${t ? compteDe(t, s.id) : ''}</span></button>`).join('')}</nav></div>
      ${qui.length ? `<div class="tb-panel"><div class="tb-h">Qui traite quoi</div><div class="tb-mini">${qui.map(q => `<div class="tb-mini-row"><span>${h(q.nom)}</span><span class="tb-bar"><i style="width:${Math.round((q.mails / max) * 100)}%"></i></span><span class="mono">${q.mails}</span></div>`).join('')}</div></div>` : ''}
      ${t?.programmes.length ? `<div class="tb-panel"><div class="tb-h">Envois programmés<span class="tb-h-n">${t.programmes.length}</span></div><div class="tb-mini">${t.programmes.slice(0, 6).map(p => `<div class="tb-mini-row" title="${h(p.a.join(', '))}">${icon('clock', 13)}<span>${h(p.sujet || '(sans objet)')}</span><span class="tb-when">${h(quandLisible(p.quand, Date.now()))}</span></div>`).join('')}</div></div>` : ''}
      ${statsSeanceHtml()}
      ${statistiquesHtml()}
      <div class="tb-panel tb-clavier"><div class="tb-h">Glisser un mail vers</div><p class="tb-note">« Je prends », une date ou un projet : le dock apparaît en bas pendant le glisser.</p></div>`;
    host.querySelectorAll<HTMLButtonElement>('[data-s]').forEach(b => b.addEventListener('click', () => allerSection(b.dataset.s as Section)));
    host.querySelector<HTMLButtonElement>('[data-seance]')?.addEventListener('click', () => seance());
  }

  // Séance de tri (07/10/2026) : statistiques de la semaine (mails restants à la clôture, temps passé).
  let statsSeance: { stats: SeanceSemaine[]; dansLaBoite: number | null; objectif: number } | null = null;
  function statsSeanceHtml(): string {
    const sem = statsSeance?.stats?.[0];
    const n = statsSeance?.dansLaBoite;
    if (!statsSeance) return '';
    return `<div class="tb-panel"><div class="tb-h">Séance de tri</div><div class="tb-mini">`
      + `${n != null ? `<div class="tb-mini-row"><span>Dans la boîte</span><span class="mono">${n} / objectif ${statsSeance.objectif}</span></div>` : ''}`
      + `${sem ? `<div class="tb-mini-row"><span>Cette semaine</span><span class="mono">${sem.traites} triés · ${sem.minutes} min</span></div>`
        + `<div class="tb-mini-row"><span>Clôture à 5 ou moins</span><span class="mono">${sem.objectifAtteint}/${sem.joursReleves}</span></div>` : '<p class="tb-note">Aucune séance cette semaine.</p>'}`
      + `</div><button type="button" class="tb-btn" data-seance>${icon('inbox', 14)}Lancer la séance</button></div>`;
  }
  // Assistant inbox (07/10/2026) : expéditeurs, mails hors horaires, temps par client, courbe vers 5 mails.
  let statistiques: StatistiquesBoite | null = null;
  function courbeSvg(st: StatistiquesBoite): string {
    const pts = st.courbe;
    if (pts.length < 2) return '';
    const max = Math.max(st.objectif + 1, ...pts.map(p => p.restants));
    const x = (k: number) => Math.round((k / (pts.length - 1)) * 200);
    const y = (v: number) => Math.round(46 - (v / max) * 42);
    const d = pts.map((p, k) => `${k ? 'L' : 'M'}${x(k)} ${y(p.restants)}`).join(' ');
    const dernier = pts[pts.length - 1];
    return `<svg class="tb-courbe" viewBox="0 0 200 48" preserveAspectRatio="none" role="img" aria-label="Mails restants à la clôture, ${pts.length} derniers relevés, dernier ${dernier.restants}, objectif ${st.objectif}"><line class="obj" x1="0" x2="200" y1="${y(st.objectif)}" y2="${y(st.objectif)}"/><path class="ligne" d="${d}"/></svg>`;
  }
  function statistiquesHtml(): string {
    const st = statistiques;
    if (!st || (!st.total && !st.courbe.length)) return '';
    const top = st.topExpediteurs.slice(0, 5);
    const maxTop = Math.max(1, ...top.map(t => t.n));
    return `<div class="tb-panel"><div class="tb-h">Ta boîte, 30 jours</div><div class="tb-mini">`
      + `${courbeSvg(st)}${st.courbe.length >= 2 ? `<p class="tb-note">Mails restants à la clôture (pointillé : objectif ${st.objectif}).</p>` : ''}`
      + `<div class="tb-mini-row"><span>Reçus hors horaires</span><span class="mono">${st.horsHoraires.n} · ${st.horsHoraires.part} %</span></div>`
      + top.map(t => `<div class="tb-mini-row" title="${h(t.email)}"><span>${h(t.nom || t.email)}</span><span class="tb-bar"><i style="width:${Math.round((t.n / maxTop) * 100)}%"></i></span><span class="mono">${t.n}</span></div>`).join('')
      + (st.tempsParClient.length ? `<div class="tb-h" style="margin-top:6px">Temps de tri par client</div>${st.tempsParClient.slice(0, 5).map(c => `<div class="tb-mini-row"><span>${h(c.client)}</span><span class="mono">${c.minutes} min</span></div>`).join('')}` : '')
      + '</div></div>';
  }
  async function chargerStatsSeance(): Promise<void> {
    try { statsSeance = await fetchStatsSeance(); renderLeft(); } catch { /* facultatif (boîte non suivie) */ }
    try { statistiques = await fetchStatistiques(); renderLeft(); } catch { /* facultatif */ }
  }
  function seance(reprise: RepriseSeance | null = null): void {
    seanceCourante = reprise || { messageId: null, suivants: [] };
    sauver();
    ouvrirSeance({
      openLink: opts.openLink, repondreDansOutlook: opts.repondreDansOutlook, reprise,
      onPosition: p => { seanceCourante = p; sauver(); },
      onFerme: () => { seanceCourante = null; sauver(); void rafraichir(true); void chargerStatsSeance(); },
    });
  }

  function renderListe(nouv: string[] = []): void {
    const sec = SECTIONS.find(s => s.id === etat.section)!;
    $('tb-titre').textContent = sec.titre;
    const t = etat.t;
    const list = $('tb-list');
    if (!t) { list.innerHTML = '<div class="tb-skel"></div><div class="tb-skel"></div><div class="tb-skel"></div><div class="tb-skel"></div>'; return; }
    if (etat.section === 'newsletters') {
      $('tb-sous').textContent = `${t.compteurs.newsletters} mails de ${t.newsletters.length} expéditeurs`;
      list.innerHTML = t.newsletters.length ? t.newsletters.map(d => digestHtml(d)).join('') : videHtml(sec.vide);
      list.querySelectorAll<HTMLElement>('[data-cle]').forEach(el => el.addEventListener('click', () => choisir(el.dataset.cle!)));
      return;
    }
    const lignes = lignesDe(t, etat.section);
    const sous = etat.section === 'priorites' ? 'Classés par ARGO : urgence, client, ton, attente'
      : etat.section === 'deCote' ? 'Hors de la boîte de réception : reviennent d\'eux-mêmes à la date choisie'
      : etat.section === 'enAttente' ? sousTitreEnAttente(t.enAttente)
      : `${lignes.length} mail${lignes.length > 1 ? 's' : ''} dans la boîte de réception`;
    $('tb-sous').textContent = sous + (t.inboxNonVerifiee?.length ? ' · boîte de réception non relue, rien n\'est masqué' : '');
    list.innerHTML = lignes.length ? lignes.map(m => ligneHtml(m, nouv.includes(cleDe(m)))).join('') : videHtml(sec.vide);
    list.querySelectorAll<HTMLElement>('.tb-row').forEach(el => brancherLigne(el));
    marquerSelection();
    renderLot();
  }

  // ── Lot (mails cochés) ──
  function mailsDuLot(): TableauMail[] { return etat.lot.map(k => trouver(k)).filter((m): m is TableauMail => !!m); }
  function cocher(cle: string, plage = false): void {
    const lignes = lignesDe(etat.t, etat.section).map(cleDe);
    etat.lot = plage ? selectionPlage(lignes, etat.ancre, cle, etat.lot) : basculerSelection(etat.lot, cle);
    etat.ancre = cle;
    $('tb-list').querySelectorAll<HTMLElement>('.tb-row').forEach(el => {
      const on = etat.lot.includes(el.dataset.cle || '');
      el.classList.toggle('is-coche', on);
      const cb = el.querySelector<HTMLInputElement>('[data-cocher]');
      if (cb) cb.checked = on;
    });
    renderLot();
  }
  function viderLot(): void { etat.lot = []; etat.ancre = null; renderListe(); }
  function renderLot(): void {
    const bar = $('tb-lot');
    const n = etat.lot.length;
    bar.hidden = !n;
    if (!n) { bar.innerHTML = ''; return; }
    bar.innerHTML = `<strong>${n} mail${n > 1 ? 's' : ''} coché${n > 1 ? 's' : ''}</strong>
      <button type="button" class="tb-btn is-primary" data-lot="projet">${icon('folder', 14)}Rattacher à un projet</button>
      <button type="button" class="tb-btn" data-lot="contact">${icon('user', 14)}Rattacher à un contact</button>
      <button type="button" class="tb-btn is-ghost" data-lot="vider">Tout décocher</button>
      <span class="tb-note" data-lot-progres aria-live="polite"></span>`;
    bar.querySelectorAll<HTMLButtonElement>('[data-lot]').forEach(b => b.addEventListener('click', () => void agirLot(b.dataset.lot as 'projet' | 'contact' | 'vider')));
  }
  async function agirLot(quoi: 'projet' | 'contact' | 'vider'): Promise<void> {
    if (quoi === 'vider') { viderLot(); return; }
    const bar = $('tb-lot');
    const progres = bar.querySelector<HTMLElement>('[data-lot-progres]');
    const mails = mailsDuLot().filter(m => m.famille !== 'mandats');
    if (!mails.length) { toast('Mandats et associations : ni projet ni contact', 'info'); return; }
    bar.querySelectorAll<HTMLButtonElement>('button').forEach(b => { b.disabled = true; });
    try {
      const fait = await actions.lierLot(mails, quoi, t => { if (progres) progres.textContent = t; });
      if (fait) { etat.lot = []; etat.ancre = null; await rafraichir(true); }
    } finally { renderLot(); }
  }

  const videHtml = (texte: string) => `<div class="tb-empty"><span class="tb-empty-mark">${icon('check', 20)}</span><strong>C'est vide, et c'est bien</strong><span>${h(texte)}</span><span class="tb-note tb-clavier">Appuie sur <span class="tb-kbd">⌘K</span> pour poser une question à ta boîte ou rattraper ce que tu as manqué.</span></div>`;

  function ligneHtml(m: TableauMail, nouveau: boolean): string {
    const cle = cleDe(m);
    const chips: string[] = [];
    if (m.correspondant === 'client') chips.push('<span class="tb-chip">Client</span>');
    if (m.brouillonPret) chips.push('<span class="tb-chip is-argo">Brouillon prêt</span>');
    if (m.equipe?.prisPar) chips.push(`<span class="tb-chip is-ok">${h(m.equipe.prisParNom || prenomDe(m.equipe.prisPar))}</span>`);
    else if (m.equipe?.assigneA) chips.push(`<span class="tb-chip">Pour ${h(m.equipe.assigneNom || prenomDe(m.equipe.assigneA))}</span>`);
    if (m.plusTardJusqua) chips.push(`<span class="tb-chip">${h(quandLisible(m.plusTardJusqua, Date.now()))}</span>`);
    if (m.relanceLe) chips.push(`<span class="tb-chip">Relance ${h(m.relanceLe.slice(8, 10))}/${h(m.relanceLe.slice(5, 7))}</span>`);
    else if (m.relanceGeree) chips.push(`<span class="tb-chip" title="${h(m.relanceGeree)}">${h(m.relanceGeree)}</span>`);
    if (m.urgence >= 3) chips.push('<span class="tb-chip is-hot">Urgent</span>');
    const boite = etat.boite === 'toutes' && (etat.t?.boites.length || 0) > 1 && m.mailbox && m.mailbox !== etat.t?.moi ? ` · ${h(m.mailbox.split('@')[0])}@` : '';
    const coche = etat.lot.includes(cle);
    return `<div class="tb-row${nouveau ? ' is-new' : ''}${coche ? ' is-coche' : ''}" role="option" tabindex="-1" draggable="${tactile ? 'false' : 'true'}" data-cle="${h(cle)}" aria-selected="false">
      <input type="checkbox" class="tb-sel" data-cocher ${coche ? 'checked' : ''} aria-label="Cocher ce mail (x)" title="Cocher (x, Maj-clic pour une plage)">
      ${m.urgence >= 2 ? `<span class="tb-prio p${Math.min(3, m.urgence)}"></span>` : ''}
      <span class="tb-av${m.correspondant === 'client' ? ' is-client' : ''}" aria-hidden="true">${h(initiales(m.from?.name || '', m.from?.email || ''))}</span>
      <span class="tb-row-main">
        <span class="tb-row-l1"><span class="tb-from">${h(m.from?.name || m.from?.email || '')}</span><span class="tb-subj">${h(m.subject || '(sans objet)')}</span></span>
        <span class="tb-row-l2">${h(m.resume || '')}${boite}</span>
      </span>
      <span class="tb-row-side"><span class="tb-when">${h(ilYA(m.receivedAt, Date.now()))}</span><span class="tb-chips">${chips.slice(0, 3).join('')}</span></span>
    </div>`;
  }

  function digestHtml(d: DigestNewsletter): string {
    return `<div class="tb-digest tb-row" data-cle="nl|${h(d.expediteur)}" role="option" tabindex="-1" aria-selected="false">
      <span class="tb-av" aria-hidden="true">${h(initiales(d.nom, d.expediteur))}</span>
      <span class="tb-row-main"><span class="tb-row-l1"><span class="tb-from">${h(d.nom)}</span><span class="tb-chip">${d.nb}</span></span><ul>${d.sujets.map(s => `<li>${h(s)}</li>`).join('')}</ul></span>
      <span class="tb-row-side"><span class="tb-when">${h(ilYA(d.dernier, Date.now()))}</span></span>
    </div>`;
  }

  function trouver(cle: string | null): TableauMail | null {
    if (!cle || !etat.t) return null;
    const t = etat.t;
    for (const m of [...lignesDe(t, etat.section), ...t.personnes, ...(t.mandats || []), ...t.notifications, ...t.factures, ...t.deCote, ...lignesDe(t, 'enAttente')]) if (cleDe(m) === cle) return m;
    return null;
  }

  function marquerSelection(): void {
    $('tb-list').querySelectorAll<HTMLElement>('.tb-row').forEach(el => el.setAttribute('aria-selected', String(el.dataset.cle === etat.sel)));
    $('tb-app').classList.toggle('has-detail', !!etat.sel);
  }

  function renderDroite(): void {
    const host = $('tb-right');
    if (etat.sel?.startsWith('nl|') && etat.t) {
      const d = etat.t.newsletters.find(x => `nl|${x.expediteur}` === etat.sel);
      if (d) { renderDigest(host, d, ctx); return; }
    }
    const m = trouver(etat.sel);
    if (m) renderDetail(host, m, ctx);
    else renderAujourdhui(host, ctx);
  }

  function choisir(cle: string | null): void {
    if (cle) $('tb-app').classList.remove('has-jour');
    if (cle === etat.sel) return;
    etat.sel = cle;
    marquerSelection();
    renderDroite();
    sauver();
    if (cle) $('tb-list').querySelector<HTMLElement>(`[data-cle="${CSS.escape(cle)}"]`)?.scrollIntoView({ block: 'nearest', behavior: mouvementReduit() ? 'auto' : 'smooth' });
  }

  function allerSection(s: Section): void {
    etat.section = s;
    etat.sel = null;
    renderStats(); renderLeft(); renderListe(); renderDroite();
    sauver();
  }

  /**
   * Reprise (08/10/2026) : après le premier tableau reçu, revenir là où la personne était : séance à la
   * même carte ; sinon le même mail, ou le suivant s'il a été traité ; le mail auquel elle vient de
   * répondre est montré dans « Traités à ranger » (« Aujourd'hui »).
   */
  function appliquerReprise(t: Tableau): void {
    const r = repriseAttendue;
    if (!r) return;
    repriseAttendue = null;
    if (!r.seance && !r.sel && r.section === 'priorites') return; // rien à reprendre : pas de toast
    if (r.seance) { seance(r.seance); toast('Reprise là où tu étais', 'info'); return; }
    const mid = messageIdDeCle(r.sel);
    const traite = !!mid && (t.traites || []).some(x => x.messageId === mid);
    if (traite) {
      choisir(null);
      const carte = document.querySelector<HTMLElement>(`#tb-traites [data-id="${CSS.escape(mid)}"]`);
      if (carte) { carte.classList.add('is-nouveau'); carte.scrollIntoView({ block: 'center', behavior: mouvementReduit() ? 'auto' : 'smooth' }); }
    } else {
      choisir(selectionReprise(r.sel, r.suivants, lignesDe(t, etat.section).map(cleDe)));
    }
    toast('Reprise là où tu étais', 'info');
  }

  // ── Glisser-déposer, aperçu au survol ──
  let glisse: TableauMail | null = null;
  function brancherLigne(el: HTMLElement): void {
    el.querySelector<HTMLInputElement>('[data-cocher]')?.addEventListener('click', e => {
      e.stopPropagation();
      cocher(el.dataset.cle!, (e as MouseEvent).shiftKey);
    });
    el.addEventListener('click', e => {
      // Maj-clic : plage ; ⌘ / Ctrl-clic : coche ce mail (sans changer le mail affiché).
      if (e.shiftKey) { e.preventDefault(); cocher(el.dataset.cle!, true); return; }
      if (e.metaKey || e.ctrlKey) { e.preventDefault(); cocher(el.dataset.cle!); return; }
      choisir(el.dataset.cle!);
    });
    el.addEventListener('dblclick', () => { const m = trouver(el.dataset.cle!); if (m) actions.ouvrir(m); });
    el.addEventListener('dragstart', e => {
      glisse = trouver(el.dataset.cle!);
      if (!glisse) return;
      el.classList.add('is-dragging');
      e.dataTransfer?.setData('text/plain', glisse.subject || '');
      if (e.dataTransfer) e.dataTransfer.effectAllowed = 'move';
      cacherApercu();
      ouvrirDock(glisse);
    });
    el.addEventListener('dragend', () => { el.classList.remove('is-dragging'); fermerDock(); glisse = null; });
    if (!tactile) {
      el.addEventListener('mouseenter', () => planifierApercu(el));
      el.addEventListener('mouseleave', cacherApercu);
    }
  }

  function brancherDepot(zone: HTMLElement, faire: (m: TableauMail) => void): void {
    zone.addEventListener('dragover', e => { if (glisse) { e.preventDefault(); zone.classList.add('is-over'); } });
    zone.addEventListener('dragleave', () => zone.classList.remove('is-over'));
    zone.addEventListener('drop', e => {
      e.preventDefault();
      zone.classList.remove('is-over');
      const m = glisse;
      fermerDock();
      if (m) faire(m);
    });
  }

  function ouvrirDock(m: TableauMail): void {
    const dock = $('tb-dock');
    const perso = (m.mailbox || '').toLowerCase() === etat.t?.moi;
    const partage = !perso;
    const presets = presetsMoments(Date.now()).slice(0, 3);
    dock.innerHTML = [
      partage ? `<div class="tb-drop" data-d="prendre">${icon('user', 18)}Je prends</div>` : '',
      ...(perso ? presets.map(p => `<div class="tb-drop" data-d="date" data-iso="${p.iso}">${icon('clock', 18)}${h(p.libelle)}</div>`) : []),
      perso ? `<div class="tb-drop" data-d="date-libre">${icon('calendar-plus', 18)}Choisir une date</div>` : '',
      `<div class="tb-drop" data-d="projet">${icon('folder', 18)}Rattacher à un projet</div>`,
    ].join('');
    dock.querySelectorAll<HTMLElement>('.tb-drop').forEach(z => brancherDepot(z, mm => {
      const d = z.dataset.d;
      if (d === 'prendre') void actions.jePrends(mm);
      else if (d === 'date') void actions.deCote(mm, z.dataset.iso);
      else if (d === 'date-libre') void actions.deCote(mm);
      else if (d === 'projet') void actions.lierProjet(mm);
    }));
    dock.classList.add('is-on');
    dock.setAttribute('aria-hidden', 'false');
  }
  function fermerDock(): void { const d = $('tb-dock'); d.classList.remove('is-on'); d.setAttribute('aria-hidden', 'true'); }

  let apercuTimer: number | null = null;
  function planifierApercu(el: HTMLElement): void {
    if (apercuTimer) window.clearTimeout(apercuTimer);
    apercuTimer = window.setTimeout(() => {
      const m = trouver(el.dataset.cle!);
      if (!m || el.dataset.cle === etat.sel || glisse) return;
      const peek = $('tb-peek');
      const r = el.getBoundingClientRect();
      peek.innerHTML = `<h4>${h(m.subject || '(sans objet)')}</h4><p>${h(m.from?.name || '')} &lt;${h(m.from?.email || '')}&gt;</p>${m.resume ? `<p>${h(m.resume)}</p>` : ''}${m.raisons?.length ? `<div class="tb-chips" style="justify-content:flex-start">${m.raisons.map(x => `<span class="tb-chip">${h(x)}</span>`).join('')}</div>` : ''}${m.actions?.length ? `<p>ARGO propose : ${m.actions.map(a => h(a.libelle)).join(' · ')}</p>` : ''}<p class="tb-note">Entrée : ouvrir · r : répondre · e : de côté</p>`;
      const top = Math.min(window.innerHeight - 220, Math.max(12, r.top));
      const left = r.right + 12 + 360 < window.innerWidth ? r.right + 12 : Math.max(12, r.left - 372);
      peek.style.top = `${top}px`;
      peek.style.left = `${left}px`;
      peek.classList.add('is-on');
    }, 420);
  }
  function cacherApercu(): void { if (apercuTimer) window.clearTimeout(apercuTimer); $('tb-peek').classList.remove('is-on'); }

  // ── Clavier ──
  document.addEventListener('keydown', e => {
    if (root.closest('[hidden]')) return; // volet Outlook revenu sur « Ce mail » : le tableau dort
    const cible = e.target as HTMLElement;
    const saisie = cible instanceof HTMLInputElement || cible instanceof HTMLTextAreaElement || cible instanceof HTMLSelectElement || cible?.isContentEditable;
    if (document.getElementById('tb-seance')) return; // séance de tri ouverte : elle a son propre clavier
    if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'k') { e.preventDefault(); palette(); return; }
    if (saisie || e.metaKey || e.ctrlKey || e.altKey || document.querySelector('.tb-overlay')) return;
    const lignes = etat.section === 'newsletters' ? (etat.t?.newsletters.map(d => `nl|${d.expediteur}`) || []) : lignesDe(etat.t, etat.section).map(cleDe);
    const i = etat.sel ? lignes.indexOf(etat.sel) : -1;
    const m = trouver(etat.sel);
    const k = e.key;
    if (k === 'j' || k === 'ArrowDown') { e.preventDefault(); const n = deplacer(i, 1, lignes.length); if (n >= 0) choisir(lignes[n]); }
    else if (k === 'k' || k === 'ArrowUp') { e.preventDefault(); const n = deplacer(i, -1, lignes.length); if (n >= 0) choisir(lignes[n]); }
    else if (k === 'Enter' && m) { e.preventDefault(); actions.ouvrir(m); }
    else if (k === 'Escape') { if (etat.lot.length) { e.preventDefault(); viderLot(); } else if (etat.sel) { e.preventDefault(); choisir(null); } }
    else if (k === 'x' && etat.sel && !etat.sel.startsWith('nl|')) { e.preventDefault(); cocher(etat.sel); }
    else if (k === 'X' && etat.sel && !etat.sel.startsWith('nl|')) { e.preventDefault(); cocher(etat.sel, true); }
    else if (k === 'l' && etat.lot.length) { e.preventDefault(); void agirLot('projet'); }
    else if (k === '/' ) { e.preventDefault(); palette(); }
    else if (k === '?') { e.preventDefault(); aide(); }
    else if (/^[1-9]$/.test(k) && sectionsVisibles(etat.t)[Number(k) - 1]) { e.preventDefault(); allerSection(sectionsVisibles(etat.t)[Number(k) - 1].id); }
    else if (m && k === 'e') { e.preventDefault(); void actions.deCote(m, presetsMoments(Date.now()).find(p => p.cle === 'demain')?.iso); }
    else if (m && k === 's') { e.preventDefault(); void actions.deCote(m); }
    else if (m && k === 'p') { e.preventDefault(); void actions.jePrends(m); }
    else if (m && k === 'l') { e.preventDefault(); void actions.lierProjet(m); }
    else if (m && ['r', 't', 'c'].includes(k)) { e.preventDefault(); document.querySelector<HTMLElement>(`[data-raccourci="${k}"]`)?.click(); }
    else if (k === 'g') { e.preventDefault(); void rafraichir(true); }
    else if (k === 'b') { e.preventDefault(); seance(); }
  });

  function commandesPalette(): ActionPalette[] {
    const m = trouver(etat.sel);
    const cmds: ActionPalette[] = [];
    if (m) {
      cmds.push(
        { id: 'ouvrir', groupe: 'Ce mail', libelle: 'Ouvrir dans Outlook', raccourci: 'Entrée', faire: () => actions.ouvrir(m) },
        { id: 'argo', groupe: 'Ce mail', libelle: 'Répondre avec ARGO', motsCles: 'réponse brouillon', raccourci: 'r', faire: () => document.querySelector<HTMLElement>('[data-raccourci="r"]')?.click() },
        { id: 'modele', groupe: 'Ce mail', libelle: 'Répondre avec un modèle', motsCles: 'template communications', raccourci: 't', faire: () => document.querySelector<HTMLElement>('[data-raccourci="t"]')?.click() },
        { id: 'decote', groupe: 'Ce mail', libelle: 'Mettre de côté jusqu\'à…', motsCles: 'plus tard rappel revenir snooze', raccourci: 's', faire: () => actions.deCote(m) },
        { id: 'projet', groupe: 'Ce mail', libelle: 'Rattacher à un projet', motsCles: 'lier classer', raccourci: 'l', faire: () => actions.lierProjet(m) },
      );
      if ((m.mailbox || '').toLowerCase() !== etat.t?.moi) cmds.push({ id: 'prendre', groupe: 'Ce mail', libelle: 'Je prends', motsCles: 'attribuer équipe', raccourci: 'p', faire: () => actions.jePrends(m) });
    }
    sectionsVisibles(etat.t).forEach((s, i) => cmds.push({ id: `s-${s.id}`, groupe: 'Aller à', libelle: s.titre, raccourci: String(i + 1), faire: () => allerSection(s.id) }));
    cmds.push(
      { id: 'seance', groupe: 'Général', libelle: 'Séance de tri : vider la boîte, une carte par mail', motsCles: 'trier inbox zéro objectif 5 fermer', raccourci: 'b', faire: () => seance() },
      { id: 'aujourdhui', groupe: 'Aller à', libelle: 'Aujourd\'hui : rendez-vous, engagements, réactivité', motsCles: 'agenda rdv promesses délai', faire: () => choisir(null) },
      { id: 'question', groupe: 'ARGO', libelle: 'Poser une question à ma boîte', motsCles: 'chercher qui quand', faire: () => outil('question') },
      { id: 'rattrapage', groupe: 'ARGO', libelle: 'Rattraper ce que j\'ai manqué', motsCles: 'absence retour résumé', faire: () => outil('rattrapage') },
      { id: 'refresh', groupe: 'Général', libelle: 'Actualiser', raccourci: 'g', faire: () => rafraichir(true) },
      { id: 'aide', groupe: 'Général', libelle: 'Raccourcis clavier', raccourci: '?', faire: () => aide() },
    );
    // Mails chargés : recherche par expéditeur et objet.
    const t = etat.t;
    if (t) {
      const vus = new Set<string>();
      for (const x of [...t.personnes, ...(t.mandats || []), ...t.factures, ...t.deCote, ...t.notifications]) {
        const cle = cleDe(x);
        if (vus.has(cle)) continue;
        vus.add(cle);
        cmds.push({ id: `m-${cle}`, groupe: 'Mails', libelle: `${x.from?.name || x.from?.email} : ${x.subject || '(sans objet)'}`, motsCles: `${x.from?.email} ${x.resume || ''}`, faire: () => { const s = SECTIONS.find(sx => lignesDe(t, sx.id).some(y => cleDe(y) === cle))?.id; if (s && s !== etat.section) allerSection(s); choisir(cle); } });
      }
    }
    return cmds;
  }
  const palette = () => ouvrirPalette(commandesPalette());
  /** Outils ARGO de la colonne « Aujourd'hui » (question à sa boîte, rattrapage). */
  function outil(quoi: 'question' | 'rattrapage'): void {
    if (etat.sel) choisir(null);
    window.setTimeout(() => document.querySelector<HTMLElement>(`[data-outil="${quoi}"]`)?.click(), 0);
  }

  function aide(): void {
    const el = document.createElement('div');
    el.className = 'tb-modal';
    el.setAttribute('role', 'dialog');
    const lignes: Array<[string, string]> = [
      ['j / k', 'Mail suivant / précédent'], ['Entrée', 'Ouvrir dans Outlook'], ['r', 'Répondre avec ARGO'], ['t', 'Répondre avec un modèle'],
      ['e', 'De côté jusqu\'à demain 8 h'], ['s', 'De côté jusqu\'à…'], ['p', 'Je prends (boîte partagée)'], ['l', 'Rattacher à un projet'],
      ['b', 'Séance de tri (une carte par mail)'], ['c', 'Commentaire interne'], ['x', 'Cocher le mail (Maj-x / Maj-clic : plage)'], ['l (mails cochés)', 'Rattacher le lot à un projet'], ['1 à 9', 'Sections'], ['⌘K ou /', 'Palette : chercher, agir'], ['g', 'Actualiser'], ['Échap', 'Fermer, revenir à Aujourd\'hui'],
    ];
    el.innerHTML = `<h3>Raccourcis clavier</h3><div class="tb-help">${lignes.map(([k, v]) => `<span class="tb-kbd">${h(k)}</span><span>${h(v)}</span>`).join('')}</div><p class="tb-note">Glisser un mail : « Je prends », une date ou un projet.</p>`;
    ouvrirCouche(el);
  }

  $('tb-cmdk').addEventListener('click', palette);
  $('tb-seance-btn').addEventListener('click', () => seance());
  // Écran étroit (téléphone) : la colonne de droite devient un volet ; « Aujourd'hui » l'ouvre, « Retour » le ferme.
  $('tb-jour').addEventListener('click', () => {
    if (etat.sel) choisir(null);
    $('tb-app').classList.toggle('has-jour');
  });
  $('tb-retour').addEventListener('click', () => {
    $('tb-app').classList.remove('has-jour');
    if (etat.sel) choisir(null);
  });
  $('tb-aide').addEventListener('click', aide);
  $('tb-refresh').addEventListener('click', () => rafraichir(true));
  // Compte connecté (07/10/2026 : sur Outlook iPhone, le jeton venait d'une boîte d'équipe ajoutée
  // dans Outlook) : affiché en permanence, « Changer de compte » propose le choix du compte.
  function majCompte(): void {
    const c = compteConnecte();
    $('tb-compte').hidden = !c;
    $('tb-compte-nom').textContent = c ? `Connecté en tant que ${c}` : '';
    $('tb-compte-nom').title = c;
  }
  $('tb-compte-changer').addEventListener('click', async () => {
    changerDeCompte();
    await rafraichir(true);
    majCompte();
  });

  // ── Données et temps réel ──
  function majLive(statut: 'sync' | 'ok' | 'off', texte?: string): void {
    const el = $('tb-live');
    el.classList.toggle('is-sync', statut === 'sync');
    el.classList.toggle('is-off', statut === 'off');
    el.querySelector('span')!.textContent = texte || (statut === 'sync' ? 'mise à jour…' : statut === 'off' ? 'hors ligne' : `à jour ${ilYA(new Date(etat.derniereSync).toISOString(), Date.now())}`);
  }

  async function rafraichir(force = false): Promise<void> {
    majLive('sync');
    try {
      if (!force && etat.t) {
        const v = await fetchVersion(etat.boite);
        if (v.version === etat.t.version) { etat.derniereSync = Date.now(); etat.enLigne = true; majLive('ok'); return; }
      }
      const t = await fetchTableau(etat.boite, force);
      const avant = etat.vus;
      const tous = [...t.personnes, ...(t.mandats || []), ...t.factures, ...t.notifications].map(cleDe);
      if (etat.section === 'mandats' && !(t.mandats || []).length) etat.section = 'priorites';
      const nouv = avant.length ? nouveautes(avant, tous) : [];
      etat.t = t;
      etat.vus = tous;
      etat.lot = selectionValide(etat.lot, tous);
      if (etat.boite !== 'toutes' && !t.boites.includes(etat.boite)) etat.boite = 'toutes';
      etat.derniereSync = Date.now();
      etat.enLigne = true;
      renderBoites(); renderStats(); renderLeft(); renderListe(nouv); majCompte();
      // La colonne de droite n'est redessinée que si le mail choisi a disparu (saisie en cours protégée).
      if (etat.sel && !trouver(etat.sel) && !etat.sel.startsWith('nl|')) { etat.sel = null; renderDroite(); }
      // « Aujourd'hui » relu seulement si la personne n'y écrit pas (question à sa boîte, etc.).
      else if (!etat.sel && !$('tb-right').contains(document.activeElement)) renderDroite();
      if (nouv.length && avant.length) toast(`${nouv.length} nouveau${nouv.length > 1 ? 'x' : ''} mail${nouv.length > 1 ? 's' : ''}`, 'info');
      if (repriseAttendue) appliquerReprise(t);
      majLive('ok');
    } catch (e) {
      etat.enLigne = false;
      majLive('off', 'hors ligne');
      const session = e instanceof AtlasError && e.kind === 'session';
      const choix = session && !!(e as AtlasError).data?.choixCompte;
      if (!etat.t) {
        $('tb-list').innerHTML = `<div class="tb-empty"><strong>Tableau indisponible</strong><span>${h(humanError(e))}</span>`
          + `<span class="tb-actions"><button type="button" class="tb-btn is-primary" id="tb-retry">${choix ? 'Choisir mon compte' : session ? 'Se connecter' : 'Réessayer'}</button>`
          + `${session && !choix ? '<button type="button" class="tb-btn" id="tb-choix">Changer de compte</button>' : ''}`
          + `${session && opts.ouvrirDansNavigateur ? '<button type="button" class="tb-btn" id="tb-navigateur">Ouvrir dans le navigateur</button>' : ''}</span></div>`;
        document.getElementById('tb-retry')?.addEventListener('click', () => {
          // Session refusée : le clic autorise la fenêtre de connexion de l'hôte (jamais ouverte toute seule),
          // avec le choix du compte si le compte proposé n'est pas un compte employé.
          if (choix) changerDeCompte(); else if (session) autoriserConnexionInteractive();
          void rafraichir(true);
        });
        document.getElementById('tb-choix')?.addEventListener('click', () => { changerDeCompte(); void rafraichir(true); });
        document.getElementById('tb-navigateur')?.addEventListener('click', () => opts.ouvrirDansNavigateur?.());
      } else if (session && force) {
        toast(humanError(e), 'error');
      }
      majCompte();
    }
  }

  renderStats();
  renderListe();
  renderDroite();
  void rafraichir(true);
  void chargerStatsSeance();
  // Lien de la cloche de fin de journée (`?seance=1`, gardé pendant la connexion Microsoft) : séance ouverte d'emblée.
  let lienSeance = new URLSearchParams(window.location.search).get('seance') === '1';
  try { lienSeance = lienSeance || sessionStorage.getItem('atlas_tdb_seance') === '1'; sessionStorage.removeItem('atlas_tdb_seance'); } catch { /* stockage indisponible */ }
  if (lienSeance) { repriseAttendue = null; effacerReprise(); seance(); }
  window.setInterval(() => { if (document.visibilityState === 'visible' && !root.closest('[hidden]')) void rafraichir(false); }, POLL_MS);
  window.setInterval(() => { if (etat.enLigne && etat.derniereSync) majLive('ok'); }, 15_000);
  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible' && Date.now() - etat.derniereSync > POLL_MS) void rafraichir(false); });
  window.addEventListener('focus', () => { if (Date.now() - etat.derniereSync > 5_000) void rafraichir(false); });
}
