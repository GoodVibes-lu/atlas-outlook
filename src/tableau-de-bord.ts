/**
 * tableau-de-bord.ts — Tableau de bord de l'agent d'inbox dans la BARRE DE GAUCHE d'Outlook
 * (application Microsoft 365 « ATLAS », onglet personnel : teams-app/manifest.json). Phase 2,
 * décision du 03/10/2026 (`.claude/BACKLOG-AGENT-INBOX.md` §7). Bureau (nouvel Outlook Windows,
 * Mac à confirmer) et web ; Outlook classique et mobile gardent le bandeau « Ma journée » du panneau.
 *
 * Vue d'ensemble « Ma journée » : piles (à traiter, en attente, à filtrer, pour info, bruit),
 * clients qui attendent depuis plus de 48 h, relances dues ; listes cliquables (GET
 * /api/plugin/agent/liste, clic = ouvre le mail dans Outlook sur le web via `webLink`) ; filtre
 * des nouveaux expéditeurs (Accepter / Refuser) ; journal récent avec Annuler. Aucun appel IA.
 *
 * Phase 3 (04/10/2026), à la demande seulement (droit à la déconnexion, aucune notification) :
 * pile « Plus tard » (tuile + « Remettre maintenant »), « Poser une question à ma boîte »
 * (POST /api/plugin/agent/question, réponse + mails cités) et « Rattraper depuis… »
 * (GET /api/plugin/agent/rattrapage, sections de mails cliquables) ; ces deux appels utilisent l'IA
 * côté worker, uniquement sur clic.
 *
 * Phase 5 (membres de good@ seulement) : tuiles « good@ à attribuer » et « good@ pour toi »
 * (listes `equipe_a_attribuer` / `equipe_pour_moi`, avec « Pour X » / « Pris par X »).
 *
 * Connexion : même `getWorkerToken` que le complément (worker.ts), voie « nested app
 * authentication » activée hors Office.js par `enableHostedNaa()` après l'initialisation de
 * TeamsJS. Limite : si l'hôte ne fournit pas la connexion automatique (Outlook ou navigateur
 * trop ancien, page ouverte hors d'Outlook / Teams), MSAL tente une fenêtre de connexion, sinon
 * le tableau de bord affiche l'erreur et renvoie vers ATLAS (docs/agent-inbox-essai-complement.md §5).
 */

import { enableHostedNaa, getWorkerToken } from './api/worker';
import { fetchJournee, invalidateJournee, type AgentJournee, type AgentListePile } from './api/agent';
import { renderMailList, renderJournal, escapeHtml, PILE_LIBELLES } from './components/agent-lists';
import { ATLAS_BASE } from './api/platform';
import { renderQuestion, renderRattrapage } from './components/agent-outils';
import { humanError } from './api/net';

/** TeamsJS (chargé par tableau-de-bord.html depuis le CDN Microsoft), facultatif. */
const teams: any = (window as any).microsoftTeams;

const REFRESH_MS = 5 * 60 * 1000;

let pileActive: AgentListePile = 'a_traiter';
let refreshTimer: number | null = null;

// ── Hôte (TeamsJS) ──

async function initHost(): Promise<void> {
  if (!teams?.app?.initialize) return;
  try {
    await Promise.race([
      teams.app.initialize(),
      new Promise((_, reject) => setTimeout(() => reject(new Error('délai dépassé')), 5000)),
    ]);
    try {
      const ctx = await teams.app.getContext();
      applyTheme(String(ctx?.app?.theme || 'default'));
      teams.app.registerOnThemeChangeHandler?.((t: string) => applyTheme(t));
    } catch { /* thème par défaut */ }
    teams.app.notifySuccess?.();
  } catch (e) {
    // Page ouverte hors d'un hôte Microsoft 365 (navigateur seul) : on continue sans TeamsJS.
    console.warn('[tableau-de-bord] TeamsJS indisponible :', (e as Error)?.message || e);
  }
}

function applyTheme(theme: string): void {
  const t = theme === 'dark' || theme === 'contrast' ? 'dark' : 'light';
  document.body.dataset.theme = t;
  document.documentElement.dataset.theme = t; // jetons de couleur (styles.css, :root[data-theme])
}

/** Ouvre une adresse depuis l'onglet (Outlook sur le web pour un mail, ATLAS pour une fiche). */
function openLink(url: string): void {
  try {
    if (teams?.app?.openLink) { teams.app.openLink(url); return; }
  } catch { /* repli ci-dessous */ }
  window.open(url, '_blank', 'noopener');
}

function toast(message: string, type: 'success' | 'error' | 'info' = 'info'): void {
  document.querySelectorAll('.toast').forEach(el => el.remove());
  const el = document.createElement('div');
  el.className = `toast toast-${type}`;
  el.textContent = message;
  document.body.appendChild(el);
  setTimeout(() => el.remove(), 3500);
}

// ── Rendu ──

interface Tuile { cle: string; valeur: number; libelle: string; pile?: AgentListePile; alerte?: boolean }

function tuiles(j: AgentJournee): Tuile[] {
  return [
    { cle: 'a_traiter', valeur: j.aTraiter, libelle: 'à traiter', pile: 'a_traiter', alerte: j.aTraiter > 0 },
    { cle: 'en_attente', valeur: j.enAttente, libelle: 'en attente', pile: 'en_attente' },
    { cle: 'a_filtrer', valeur: j.aFiltrer, libelle: 'à filtrer', pile: 'a_filtrer', alerte: j.aFiltrer > 0 },
    { cle: 'clients', valeur: j.clientsPlus48h, libelle: j.clientsPlus48h > 1 ? 'clients > 48 h' : 'client > 48 h', alerte: j.clientsPlus48h > 0 },
    { cle: 'relances', valeur: j.relancesDues, libelle: j.relancesDues > 1 ? 'relances dues' : 'relance due' },
    { cle: 'pour_info', valeur: j.pourInfo, libelle: 'pour info', pile: 'pour_info' },
    { cle: 'bruit', valeur: j.bruit, libelle: 'bruit', pile: 'bruit' },
    { cle: 'plus_tard', valeur: j.plusTard, libelle: 'plus tard', pile: 'plus_tard' },
    // Phase 5 : good@ (membres autorisés seulement ; absent sinon).
    ...(j.equipe ? [
      { cle: 'equipe_a_attribuer', valeur: j.equipe.aAttribuer, libelle: 'good@ à attribuer', pile: 'equipe_a_attribuer' as const, alerte: j.equipe.aAttribuer > 0 },
      { cle: 'equipe_pour_moi', valeur: j.equipe.pourMoi, libelle: 'good@ pour toi', pile: 'equipe_pour_moi' as const, alerte: j.equipe.pourMoi > 0 },
    ] : []),
  ];
}

function renderShell(): void {
  const app = document.getElementById('app')!;
  app.innerHTML = `
    <div class="header tdb-header">
      <span class="header-logo">ATLAS</span>
      <span class="header-subtitle">Ma journée</span>
      <span class="tdb-spacer"></span>
      <button type="button" class="btn btn-secondary tdb-refresh" id="tdb-refresh" title="Actualiser">Actualiser</button>
    </div>
    <div class="tdb-body">
      <div id="tdb-tiles" class="tdb-tiles" aria-label="Compteurs"></div>
      <div class="tdb-columns tdb-outils">
        <section class="tdb-card" id="tdb-question"></section>
        <section class="tdb-card">
          <div class="tdb-card-title">Rattrapage</div>
          <div id="tdb-rattrapage"></div>
        </section>
      </div>
      <div class="tdb-columns">
        <section class="tdb-card tdb-main">
          <div class="tdb-card-title" id="tdb-list-title"></div>
          <div id="tdb-list"></div>
        </section>
        <aside class="tdb-card tdb-side">
          <div class="tdb-card-title">Ce que l'agent a fait</div>
          <div id="tdb-journal"></div>
          <p class="agent-muted tdb-note">Clients qui attendent et relances dues : détail dans le brief du matin et dans <a href="${ATLAS_BASE}" id="tdb-atlas">ATLAS</a>.</p>
        </aside>
      </div>
    </div>
  `;
  document.getElementById('tdb-refresh')?.addEventListener('click', () => refreshAll(true));
  document.getElementById('tdb-atlas')?.addEventListener('click', (e) => { e.preventDefault(); openLink(ATLAS_BASE); });
  const question = document.getElementById('tdb-question');
  if (question) renderQuestion(question, { openLink, onInfo: toast });
  const rattrapage = document.getElementById('tdb-rattrapage');
  if (rattrapage) renderRattrapage(rattrapage, { openLink, onInfo: toast });
}

async function renderTiles(force = false): Promise<boolean> {
  const host = document.getElementById('tdb-tiles');
  if (!host) return false;
  if (force) invalidateJournee();
  const j = await fetchJournee();
  if (!j) return false;
  host.innerHTML = tuiles(j).map(t => {
    const cls = `tdb-tile${t.alerte ? ' is-alert' : ''}${t.pile ? ' is-clickable' : ''}${t.pile === pileActive ? ' active' : ''}`;
    const inner = `<span class="tdb-tile-value">${t.valeur}</span><span class="tdb-tile-label">${escapeHtml(t.libelle)}</span>`;
    return t.pile
      ? `<button type="button" class="${cls}" data-pile="${t.pile}" aria-pressed="${t.pile === pileActive}">${inner}</button>`
      : `<div class="${cls}">${inner}</div>`;
  }).join('');
  host.querySelectorAll<HTMLButtonElement>('button[data-pile]').forEach(btn => {
    btn.addEventListener('click', () => {
      pileActive = btn.dataset.pile as AgentListePile;
      host.querySelectorAll<HTMLButtonElement>('button[data-pile]').forEach(b => {
        const on = b.dataset.pile === pileActive;
        b.classList.toggle('active', on);
        b.setAttribute('aria-pressed', String(on));
      });
      renderList();
    });
  });
  return true;
}

function renderList(): void {
  const title = document.getElementById('tdb-list-title');
  const host = document.getElementById('tdb-list');
  if (!title || !host) return;
  title.textContent = PILE_LIBELLES[pileActive] || pileActive;
  renderMailList(host, pileActive, {
    limite: 50,
    openLink,
    onInfo: toast,
    onChange: () => { renderTiles(true).catch(() => { /* compteurs facultatifs */ }); },
  });
}

function renderJournalSide(): void {
  const host = document.getElementById('tdb-journal');
  if (!host) return;
  renderJournal(host, {
    limite: 15,
    onInfo: toast,
    onChange: () => { renderTiles(true).catch(() => { /* compteurs facultatifs */ }); renderList(); },
  });
}

function renderError(message: string): void {
  const app = document.getElementById('app')!;
  app.innerHTML = `
    <div class="header tdb-header"><span class="header-logo">ATLAS</span><span class="header-subtitle">Ma journée</span></div>
    <div class="tdb-body">
      <div class="tdb-card">
        <div class="tdb-card-title">Connexion impossible</div>
        <p>${escapeHtml(message)}</p>
        <p class="agent-muted">La connexion automatique (nested app authentication) n'est peut-être pas disponible dans cette version d'Outlook. Le panneau ATLAS d'un mail (bandeau « Ma journée ») reste utilisable.</p>
        <div class="tdb-actions">
          <button type="button" class="btn btn-primary" id="tdb-retry">Réessayer</button>
          <button type="button" class="btn btn-secondary" id="tdb-open-atlas">Ouvrir ATLAS</button>
        </div>
      </div>
    </div>
  `;
  document.getElementById('tdb-retry')?.addEventListener('click', () => start());
  document.getElementById('tdb-open-atlas')?.addEventListener('click', () => openLink(ATLAS_BASE));
}

async function refreshAll(force = false): Promise<void> {
  const ok = await renderTiles(force);
  if (!ok) throw new Error('Compteurs de l\'agent indisponibles (connexion ou route /api/plugin/agent/journee).');
  renderList();
  renderJournalSide();
}

async function start(): Promise<void> {
  renderShell();
  try {
    // Connexion d'abord : son erreur exacte est plus parlante que « compteurs indisponibles ».
    await getWorkerToken();
    await refreshAll(true);
  } catch (e) {
    renderError(humanError(e));
    return;
  }
  if (refreshTimer === null) {
    // Rafraîchissement discret quand l'onglet est visible (aucune notification).
    refreshTimer = window.setInterval(() => {
      if (document.visibilityState === 'visible') renderTiles(true).catch(() => { /* facultatif */ });
    }, REFRESH_MS);
  }
}

(async () => {
  applyTheme('default');
  await initHost();
  enableHostedNaa();
  await start();
})();
