/**
 * ATLAS Outlook Add-in — Main Taskpane Entry Point
 *
 * Detects context (read/compose) and renders the appropriate panel.
 * Manages navigation between tabs and handles Office.js initialization.
 * Supports ?mode=read|compose, ?tab=agent|ia|link|info|create|compose|quick|check|reply|settings and
 * ?surface=mobile (manifeste mobile) URL params.
 *
 * Phase 1.5 (complément unifié, 03/10/2026) :
 *   • lecture : onglet « Agent » par défaut (état calculé par le serveur, AUCUN appel IA à
 *     l'ouverture) + bandeau « Ma journée » en tête du panneau ;
 *   • rédaction (bureau / web) : onglet « ⚡ Rapide » repris d'ATLAS Assistant (réponse rapide,
 *     brouillon automatique) ;
 *   • mobile : une colonne, sans rédaction ni balayage automatique (fait côté serveur).
 *
 * Phase 2 : onglet « 🛡️ Vérifier » en rédaction (contrôles avant l'envoi, bureau / web) ;
 * compteurs du bandeau cliquables (liste par pile, filtre des nouveaux expéditeurs).
 */

// Lot 1 « vitesse » (09/10/2026) : seul l'onglet « Ce mail » (AgentPanel) est chargé au démarrage ;
// les autres onglets, les réglages et la fenêtre du tableau de bord sont chargés à la demande.
// L'ancien balayage du volet (auto-sweep, accès direct à la boîte coupé par Microsoft) est retiré :
// le rangement automatique est fait par l'agent du worker.
import { AgentPanel } from './components/agent-panel';
import { lireElementReunion } from './api/reunion';
import { refreshJourneeBanner } from './components/journee-banner';
import { isMobile } from './api/platform';
import { initRoamingStorage } from './api/roaming-storage';
import { purgeLegacySecrets, prechaufferConnexion } from './api/worker';
import { icon } from './ui/icons';
import { humanError } from './api/net';
import type { AddinMode } from './types';

// ── State ──

let currentPanel: { destroy: () => void } | null = null;
let currentTab = '';
let previousTab = '';

// ── Toast ──

export function showToast(message: string, type: 'success' | 'error' | 'info' = 'info'): void {
  // Remove existing toast
  document.querySelectorAll('.toast').forEach(el => el.remove());

  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  toast.setAttribute('role', type === 'error' ? 'alert' : 'status');
  // Jamais de texte technique brut, et plus d'emoji décoratifs en tête de message.
  const text = String(message || '')
    .replace(/[☀-➿\u{1F300}-\u{1FAFF}️‍✓✔✗⏰⏳]/gu, '')
    .replace(/\s{2,}/g, ' ').replace(/^[\s:·-]+|[\s:·-]+$/g, '').trim();
  toast.innerHTML = `${icon(type === 'success' ? 'check-circle' : type === 'error' ? 'alert' : 'info', 16)}<span></span>`;
  toast.querySelector('span')!.textContent = /load failed|failed to fetch|typeerror|networkerror/i.test(text) ? humanError(new Error(text)) : text;
  document.body.appendChild(toast);

  setTimeout(() => toast.remove(), type === 'error' ? 6000 : 3500);
}

// ── Helpers ──

// Plus de configuration initiale (03/10/2026) : aucun secret à saisir, les données ATLAS et l'IA
// passent par le worker avec le compte Microsoft de l'utilisateur.
function getUserName(): string {
  let profileName = '';
  try { profileName = Office.context?.mailbox?.userProfile?.displayName || ''; } catch { /* hors Outlook */ }
  return localStorage.getItem('atlas_addin_user_name') || profileName || 'Utilisateur';
}

function getUrlParams(): { mode: AddinMode; tab: string | null } {
  const params = new URLSearchParams(window.location.search);
  const mode = (params.get('mode') as AddinMode) || 'read';
  const tab = params.get('tab');
  return { mode, tab };
}

/**
 * Detect if this email references an existing project (#NNN in subject).
 * "Créer" tab is hidden ONLY when a real project is detected, not for Re:/Fwd:.
 */
function detectProjectInSubject(): boolean {
  try {
    const item = Office.context.mailbox?.item;
    if (!item) return false;
    const subject: string | undefined = (item as any).subject;
    if (typeof subject === 'string' && /#\s*\d{2,4}/.test(subject)) return true;
  } catch { /* ignore */ }
  return false;
}

// ── Rendering ──

function renderApp(): void {
  const app = document.getElementById('app')!;
  const { mode, tab } = getUrlParams();

  // Detect project in subject before building tabs
  const mobile = isMobile();
  // Pas de surface de rédaction sur mobile (manifeste) : on reste en lecture quoi qu'il arrive.
  const isCompose = mode === 'compose' && !mobile;
  const hasDetectedProject = !isCompose ? detectProjectInSubject() : false;
  document.body.classList.toggle('is-mobile', mobile);
  applyOfficeTheme();

  // Onglets : libellé court lisible + icône SVG + description (infobulle, lecteur d'écran).
  // Réglages : bouton dédié en en-tête (plus d'onglet « ⚙ »). Refonte 05/10/2026.
  type Tab = { id: string; label: string; icon: string; title: string };
  const T: Record<string, Tab> = {
    agent: { id: 'agent', label: 'Ce mail', icon: 'inbox', title: 'Ce mail : résumé, quoi faire, correspondant' },
    ia: { id: 'ia', label: 'Classer', icon: 'tag', title: 'Classer avec ARGO : thème, urgence, dossier' },
    link: { id: 'link', label: 'Lier', icon: 'link', title: 'Lier ce mail à un projet ou un contact ATLAS' },
    info: { id: 'info', label: 'Projet', icon: 'folder', title: 'Le projet ATLAS de ce mail' },
    create: { id: 'create', label: 'Créer', icon: 'plus', title: 'Créer un projet à partir de ce mail' },
    reply: { id: 'reply', label: 'Répondre', icon: 'reply', title: 'Répondre avec un modèle' },
    compose: { id: 'compose', label: 'Modèles', icon: 'template', title: 'Insérer un modèle de mail' },
    quick: { id: 'quick', label: 'Rédiger', icon: 'bolt', title: 'Réponse rapide et brouillon ARGO' },
    check: { id: 'check', label: 'Vérifier', icon: 'shield-check', title: 'Vérifier avant l\'envoi' },
    reunion: { id: 'reunion', label: 'Point', icon: 'calendar-plus', title: 'Proposer un point à l\'ordre du jour' },
  };
  let tabs: Tab[];
  if (isCompose) tabs = [T.compose, T.quick, T.check];
  // Mobile : lecture seule, pas de rédaction (ni « Répondre » ni « Créer »).
  else if (mobile) tabs = [T.agent, T.ia, T.link, T.info];
  // Projet détecté dans le sujet : pas de « Créer ».
  else if (hasDetectedProject) tabs = [T.agent, T.ia, T.link, T.info, T.reply];
  else tabs = [T.agent, T.ia, T.link, T.info, T.create, T.reply];

  // Invitation / rendez-vous de réunion (04/10/2026) : onglet « Proposer un point » (ordre du jour
  // collaboratif), par défaut ; le reste du panneau reste disponible.
  const reunionItem = !isCompose && !!lireElementReunion();
  if (reunionItem) tabs = [T.reunion, ...tabs];

  // Onglet actif : paramètre d'URL > Agent (lecture, état serveur sans IA) > Modèles (rédaction).
  let defaultTab: string;
  if (tab && (tab === 'settings' || tabs.some(t => t.id === tab))) {
    defaultTab = tab;
  } else {
    defaultTab = isCompose ? 'compose' : reunionItem ? 'reunion' : 'agent';
  }

  app.innerHTML = `
    <header class="header">
      <span class="brand" aria-label="ATLAS, GOOD VIBES"><span class="brand-mark" aria-hidden="true"></span><span class="brand-name">ATLAS</span></span>
      <span class="header-context">${isCompose ? 'Rédaction' : 'Lecture'}</span>
      ${mobile ? '' : `<button type="button" class="header-tdb" id="btn-tableau" title="Afficher le tableau de bord ATLAS dans ce volet">${icon('inbox', 16)}<span>Tableau de bord</span></button>`}
      <button type="button" class="icon-btn" id="btn-settings" data-tab="settings" aria-label="Réglages" title="Réglages">${icon('settings', 18)}</button>
    </header>
    ${isCompose ? '' : '<div id="journee-host" hidden></div>'}
    <nav class="nav-tabs" role="tablist" aria-label="Sections du panneau" style="--tab-count:${tabs.length}">
      ${tabs.map(t => `
        <button type="button" class="nav-tab" role="tab" id="tab-${t.id}" data-tab="${t.id}" title="${t.title}"
          aria-controls="panel-content" aria-selected="false" tabindex="-1">
          ${icon(t.icon, 18)}<span class="nav-tab-label">${t.label}</span>
        </button>
      `).join('')}
    </nav>
    <main id="panel-content" class="content" role="tabpanel" tabindex="-1"></main>
  `;

  // Tableau de bord DANS ce volet (10/10/2026) : la fenêtre de dialogue restait au-dessus de toutes
  // les applications sur Outlook Mac. Le volet bascule en entier, « Ce mail » le ramène.
  app.querySelector('#btn-tableau')?.addEventListener('click', () => ouvrirTableau());

  app.querySelector('#btn-settings')?.addEventListener('click', () => {
    switchTab(currentTab === 'settings' ? (previousTab || defaultTab) : 'settings');
  });

  // Navigation clavier entre onglets (flèches, Début, Fin).
  const tabEls = Array.from(app.querySelectorAll<HTMLButtonElement>('.nav-tab'));
  tabEls.forEach((el, i) => {
    el.addEventListener('keydown', (ev) => {
      let j = -1;
      if (ev.key === 'ArrowRight') j = (i + 1) % tabEls.length;
      else if (ev.key === 'ArrowLeft') j = (i - 1 + tabEls.length) % tabEls.length;
      else if (ev.key === 'Home') j = 0;
      else if (ev.key === 'End') j = tabEls.length - 1;
      if (j < 0) return;
      ev.preventDefault();
      tabEls[j].focus();
      tabEls[j].click();
    });
  });

  // Tab click handlers
  tabEls.forEach(tabEl => {
    tabEl.addEventListener('click', () => {
      const tabId = tabEl.getAttribute('data-tab')!;
      switchTab(tabId);
    });
  });

  switchTab(defaultTab);

  // Bandeau « Ma journée » (lecture) : à l'ouverture et à chaque changement de mail, au plus une
  // fois par minute (throttle dans fetchJournee). Aucun appel IA.
  if (!isCompose) {
    refreshJourneeBanner(document.getElementById('journee-host'), { onInfo: showToast }).catch(() => { /* bandeau facultatif */ });
  }
}

/**
 * Thème : suit le thème d'Outlook quand Office.js le donne (fond du corps sombre ⇒ sombre), sinon
 * la préférence du système (CSS prefers-color-scheme).
 */
function applyOfficeTheme(): void {
  try {
    const bg = (Office.context as any)?.officeTheme?.bodyBackgroundColor as string | undefined;
    const m = bg && /^#?([0-9a-f]{6})$/i.exec(bg.trim());
    if (!m) return;
    const n = parseInt(m[1], 16);
    const lum = (0.2126 * ((n >> 16) & 255) + 0.7152 * ((n >> 8) & 255) + 0.0722 * (n & 255)) / 255;
    document.documentElement.setAttribute('data-theme', lum < 0.5 ? 'dark' : 'light');
  } catch { /* hors Outlook : préférence du système */ }
}

function switchTab(tabId: string): void {
  if (tabId === currentTab) return;
  if (currentTab && currentTab !== 'settings') previousTab = currentTab;
  currentTab = tabId;

  // Onglet actif (classe + ARIA) ; Réglages = bouton d'en-tête enfoncé.
  document.querySelectorAll<HTMLElement>('.nav-tab').forEach(tabEl => {
    const on = tabEl.getAttribute('data-tab') === tabId;
    tabEl.classList.toggle('active', on);
    tabEl.setAttribute('aria-selected', on ? 'true' : 'false');
    tabEl.tabIndex = on ? 0 : -1;
  });
  document.getElementById('btn-settings')?.setAttribute('aria-pressed', tabId === 'settings' ? 'true' : 'false');
  const panel = document.getElementById('panel-content');
  if (panel) {
    if (tabId === 'settings') { panel.setAttribute('aria-label', 'Réglages'); panel.removeAttribute('aria-labelledby'); }
    else { panel.setAttribute('aria-labelledby', `tab-${tabId}`); panel.removeAttribute('aria-label'); }
  }

  // Destroy current panel
  currentPanel?.destroy();
  currentPanel = null;

  const content = document.getElementById('panel-content')!;
  const userName = getUserName();

  if (tabId === 'agent') {
    currentPanel = new AgentPanel(content, (target) => switchTab(target));
    return;
  }
  // Onglets chargés à la demande : squelette tout de suite, panneau dès que son code est là.
  content.innerHTML = '<div class="panel-scroll"><div class="state-loading" aria-busy="true" aria-label="Chargement"><div class="skeleton-line" style="width:92%"></div><div class="skeleton-line" style="width:76%"></div><div class="skeleton-line" style="width:58%"></div></div></div>';
  const demande = tabId;
  void chargerPanneau(tabId, content, userName).then(p => {
    if (currentTab !== demande) { p?.destroy(); return; }
    currentPanel = p;
  }).catch(e => {
    if (currentTab === demande) content.innerHTML = `<div class="panel-scroll"><p class="agent-error" role="alert">${humanError(e)}</p></div>`;
  });
}

async function chargerPanneau(tabId: string, content: HTMLElement, userName: string): Promise<{ destroy: () => void } | null> {
  switch (tabId) {
    case 'reunion': return new (await import('./components/reunion-point')).ReunionPointPanel(content);
    case 'quick': return new (await import('./components/quick-draft')).QuickDraftPanel(content);
    case 'check': return new (await import('./components/send-check-panel')).SendCheckPanel(content);
    case 'ia': return new (await import('./components/ia-panel')).IAPanel(content);
    case 'link': return new (await import('./components/link-panel')).LinkPanel(content, userName);
    case 'info': return new (await import('./components/project-info')).ProjectInfoPanel(content);
    case 'create': return new (await import('./components/create-project')).CreateProjectPanel(content, userName);
    case 'compose': return new (await import('./components/compose-panel')).ComposePanel(content, userName);
    // Reply tab reuses ComposePanel in reply mode
    case 'reply': return new (await import('./components/compose-panel')).ComposePanel(content, userName, { isReply: true });
    case 'settings': return new (await import('./components/settings')).SettingsPanel(content, () => {
      // After saving settings, re-render to show main panels
      renderApp();
    });
  }
  return null;
}

// ── Mode « Tableau de bord » du volet ──

/** Vrai tant que le volet montre le tableau de bord (le mode « Ce mail » est caché). */
let enTableau = false;
/** Le mail a changé pendant que le tableau était affiché : « Ce mail » se redessine au retour. */
let mailChange = false;

function ouvrirTableau(seance = false): void {
  enTableau = true;
  void import('./tableau-volet')
    .then(m => m.afficherTableauVolet({ onRetour: revenirAuMail, onInfo: showToast, seance }))
    .catch(e => { enTableau = false; showToast(humanError(e), 'error'); });
}

function revenirAuMail(): void {
  enTableau = false;
  void import('./tableau-volet').then(m => m.masquerTableauVolet());
  if (mailChange) { mailChange = false; currentTab = ''; renderApp(); }
}

// Bouton « Séance de tri » du bandeau « Ma journée » (components/journee-banner.ts).
window.addEventListener('atlas:tableau', (e) => ouvrirTableau(!!(e as CustomEvent).detail?.seance));

// ── Office.js Initialization ──

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Outlook) {
    console.log('[ATLAS] Outlook Add-in loaded');
    // Hydrate localStorage depuis roamingSettings AVANT le 1er renderApp, et efface les
    // anciens secrets (clé Anthropic, jeton Airtable) du poste et de la boîte.
    // Connexion lancée tout de suite (jeton de la session si encore valable, sinon connexion
    // silencieuse), en parallèle du premier rendu : la vue du mail l'attend déjà en cours.
    prechaufferConnexion();
    await initRoamingStorage();
    renderApp();
    // Commande « Tableau de bord » du ruban : volet ouvert directement sur le tableau.
    if (getUrlParams().tab === 'tableau' && !isMobile()) ouvrirTableau();

    // Quand le task-pane est ÉPINGLÉ par l'utilisateur (icône 📌 en haut),
    // il reste ouvert pendant qu'il change de mail. Sans handler, le contenu
    // reste figé sur l'ancien mail. On écoute ItemChanged et on re-render
    // pour rafraîchir tous les panels (tags IA, projet lié, etc.).
    try {
      Office.context.mailbox.addHandlerAsync(
        Office.EventType.ItemChanged,
        () => {
          // Tableau affiché : il reste tel quel (pas de rechargement), « Ce mail » attend le retour.
          if (enTableau) { mailChange = true; return; }
          currentTab = '';
          renderApp();
        },
      );
    } catch (e) {
      console.warn('[ATLAS] ItemChanged handler registration failed:', e);
    }
  } else {
    // Running outside Office (dev mode)
    console.log('[ATLAS] Running outside Office.js — dev mode');
    await purgeLegacySecrets();
    renderApp();
  }
});
