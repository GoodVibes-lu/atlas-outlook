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

import { LinkPanel } from './components/link-panel';
import { ComposePanel } from './components/compose-panel';
import { ProjectInfoPanel } from './components/project-info';
import { CreateProjectPanel } from './components/create-project';
import { SettingsPanel } from './components/settings';
import { IAPanel } from './components/ia-panel';
import { AgentPanel } from './components/agent-panel';
import { QuickDraftPanel } from './components/quick-draft';
import { SendCheckPanel } from './components/send-check-panel';
import { ReunionPointPanel } from './components/reunion-point';
import { lireElementReunion } from './api/reunion';
import { refreshJourneeBanner } from './components/journee-banner';
import { isMobile } from './api/platform';
import { initRoamingStorage } from './api/roaming-storage';
import { purgeLegacySecrets } from './api/worker';
import { maybeAutoSweep } from './api/auto-sweep';
import type { AddinMode } from './types';

// ── State ──

let currentPanel: { destroy: () => void } | null = null;
let currentTab = '';

// ── Toast ──

export function showToast(message: string, type: 'success' | 'error' | 'info' = 'info'): void {
  // Remove existing toast
  document.querySelectorAll('.toast').forEach(el => el.remove());

  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  toast.textContent = message;
  document.body.appendChild(toast);

  setTimeout(() => toast.remove(), 3500);
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

  // Determine available tabs based on mode
  let tabs: Array<{ id: string; label: string; icon: string }>;

  if (isCompose) {
    tabs = [
      { id: 'compose', label: '\uD83D\uDCDD Templates', icon: '' },
      { id: 'quick', label: '\u26A1 Rapide', icon: '' },
      { id: 'check', label: '\uD83D\uDEE1\uFE0F V\u00e9rifier', icon: '' },
      { id: 'settings', label: '\u2699\uFE0F', icon: '' },
    ];
  } else if (mobile) {
    // Mobile : lecture seule, une colonne, pas de rédaction (ni « Répondre » ni « Créer »).
    tabs = [
      { id: 'agent', label: '\uD83E\uDDED Agent', icon: '' },
      { id: 'ia', label: '\u2728 IA', icon: '' },
      { id: 'link', label: '\uD83D\uDD17 Lier', icon: '' },
      { id: 'info', label: '\uD83D\uDCC1 Projet', icon: '' },
      { id: 'settings', label: '\u2699\uFE0F', icon: '' },
    ];
  } else if (hasDetectedProject) {
    // Project detected in subject: remove "Créer" tab, default to link
    tabs = [
      { id: 'agent', label: '\uD83E\uDDED Agent', icon: '' },
      { id: 'ia', label: '\u2728 IA', icon: '' },
      { id: 'link', label: '\uD83D\uDD17 Lier', icon: '' },
      { id: 'info', label: '\uD83D\uDCC1 Projet', icon: '' },
      { id: 'reply', label: '\uD83D\uDCAC R\u00e9pondre', icon: '' },
      { id: 'settings', label: '\u2699\uFE0F', icon: '' },
    ];
  } else {
    tabs = [
      { id: 'agent', label: '\uD83E\uDDED Agent', icon: '' },
      { id: 'ia', label: '\u2728 IA', icon: '' },
      { id: 'link', label: '\uD83D\uDD17 Lier', icon: '' },
      { id: 'info', label: '\uD83D\uDCC1 Projet', icon: '' },
      { id: 'create', label: '\u2795 Cr\u00e9er', icon: '' },
      { id: 'reply', label: '\uD83D\uDCAC R\u00e9pondre', icon: '' },
      { id: 'settings', label: '\u2699\uFE0F', icon: '' },
    ];
  }

  // Invitation / rendez-vous de réunion (04/10/2026) : onglet « Proposer un point » (ordre du jour
  // collaboratif), par défaut ; le reste du panneau reste disponible.
  const reunionItem = !isCompose && !!lireElementReunion();
  if (reunionItem) tabs = [{ id: 'reunion', label: '\uD83D\uDCCB Proposer un point', icon: '' }, ...tabs];

  // Onglet actif : paramètre d'URL > Agent (lecture, état serveur sans IA) > Templates (rédaction).
  let defaultTab: string;
  if (tab && tabs.some(t => t.id === tab)) {
    defaultTab = tab;
  } else {
    defaultTab = isCompose ? 'compose' : reunionItem ? 'reunion' : 'agent';
  }

  app.innerHTML = `
    <div class="header">
      <span class="header-logo">ATLAS</span>
      <span class="header-subtitle">GOOD VIBES</span>
    </div>
    ${isCompose ? '' : '<div id="journee-host" style="display:none;"></div>'}
    <div class="nav-tabs">
      ${tabs.map(t => `
        <button class="nav-tab ${t.id === defaultTab ? 'active' : ''}" data-tab="${t.id}">
          ${t.label}
        </button>
      `).join('')}
    </div>
    <div id="panel-content" class="content"></div>
  `;

  // Tab click handlers
  app.querySelectorAll('.nav-tab').forEach(tabEl => {
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

function switchTab(tabId: string): void {
  if (tabId === currentTab) return;
  currentTab = tabId;

  // Update active tab styling
  document.querySelectorAll('.nav-tab').forEach(tabEl => {
    tabEl.classList.toggle('active', tabEl.getAttribute('data-tab') === tabId);
  });

  // Destroy current panel
  currentPanel?.destroy();
  currentPanel = null;

  const content = document.getElementById('panel-content')!;
  const userName = getUserName();

  switch (tabId) {
    case 'agent':
      currentPanel = new AgentPanel(content, (target) => switchTab(target));
      break;
    case 'reunion':
      currentPanel = new ReunionPointPanel(content);
      break;
    case 'quick':
      currentPanel = new QuickDraftPanel(content);
      break;
    case 'check':
      currentPanel = new SendCheckPanel(content);
      break;
    case 'ia':
      currentPanel = new IAPanel(content);
      break;
    case 'link':
      currentPanel = new LinkPanel(content, userName);
      break;
    case 'info':
      currentPanel = new ProjectInfoPanel(content);
      break;
    case 'create':
      currentPanel = new CreateProjectPanel(content, userName);
      break;
    case 'compose':
      currentPanel = new ComposePanel(content, userName);
      break;
    case 'reply':
      // Reply tab reuses ComposePanel in reply mode
      currentPanel = new ComposePanel(content, userName, { isReply: true });
      break;
    case 'settings':
      currentPanel = new SettingsPanel(content, () => {
        // After saving settings, re-render to show main panels
        renderApp();
      });
      break;
  }
}

// ── Office.js Initialization ──

// Sweep auto au démarrage + à chaque changement de mail. Le throttle 2min
// évite le spam API. Affiche un toast discret si N mails archivés.
async function triggerAutoSweep(): Promise<void> {
  // Sur mobile, pas de balayage depuis le téléphone : le rangement est fait par l'agent serveur.
  if (isMobile()) return;
  try {
    const r = await maybeAutoSweep();
    if (r && r.archived > 0) {
      showToast(`🧹 ${r.archived} mail${r.archived > 1 ? 's' : ''} archivé${r.archived > 1 ? 's' : ''} automatiquement`, 'success');
    }
  } catch (e) {
    console.warn('[ATLAS] auto-sweep failed:', e);
  }
}

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Outlook) {
    console.log('[ATLAS] Outlook Add-in loaded');
    // Hydrate localStorage depuis roamingSettings AVANT le 1er renderApp, et efface les
    // anciens secrets (clé Anthropic, jeton Airtable) du poste et de la boîte.
    await initRoamingStorage();
    renderApp();
    // Sweep auto en background (silencieux si rien à faire)
    triggerAutoSweep();

    // Quand le task-pane est ÉPINGLÉ par l'utilisateur (icône 📌 en haut),
    // il reste ouvert pendant qu'il change de mail. Sans handler, le contenu
    // reste figé sur l'ancien mail. On écoute ItemChanged et on re-render
    // pour rafraîchir tous les panels (tags IA, projet lié, etc.).
    try {
      Office.context.mailbox.addHandlerAsync(
        Office.EventType.ItemChanged,
        () => {
          console.log('[ATLAS] ItemChanged → re-render');
          currentTab = '';
          renderApp();
          // Sweep auto à chaque changement de mail (throttled 2min)
          triggerAutoSweep();
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
