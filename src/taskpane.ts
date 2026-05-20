/**
 * ATLAS Outlook Add-in — Main Taskpane Entry Point
 *
 * Detects context (read/compose) and renders the appropriate panel.
 * Manages navigation between tabs and handles Office.js initialization.
 * Supports ?mode=read|compose and ?tab=link|info|create|compose|reply|settings URL params.
 */

import { LinkPanel } from './components/link-panel';
import { ComposePanel } from './components/compose-panel';
import { ProjectInfoPanel } from './components/project-info';
import { CreateProjectPanel } from './components/create-project';
import { SettingsPanel } from './components/settings';
import { IAPanel } from './components/ia-panel';
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

function isConfigured(): boolean {
  return !!localStorage.getItem('atlas_addin_airtable_token');
}

function getUserName(): string {
  return localStorage.getItem('atlas_addin_user_name') || 'Utilisateur';
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

  if (!isConfigured()) {
    renderSetup(app);
    return;
  }

  // Detect project in subject before building tabs
  const isCompose = mode === 'compose';
  const hasDetectedProject = !isCompose ? detectProjectInSubject() : false;

  // Determine available tabs based on mode
  let tabs: Array<{ id: string; label: string; icon: string }>;

  if (isCompose) {
    tabs = [
      { id: 'compose', label: '\uD83D\uDCDD Templates', icon: '' },
      { id: 'settings', label: '\u2699\uFE0F', icon: '' },
    ];
  } else if (hasDetectedProject) {
    // Project detected in subject: remove "Créer" tab, default to link
    tabs = [
      { id: 'ia', label: '\u2728 IA', icon: '' },
      { id: 'link', label: '\uD83D\uDD17 Lier', icon: '' },
      { id: 'info', label: '\uD83D\uDCC1 Projet', icon: '' },
      { id: 'reply', label: '\uD83D\uDCAC R\u00e9pondre', icon: '' },
      { id: 'settings', label: '\u2699\uFE0F', icon: '' },
    ];
  } else {
    tabs = [
      { id: 'ia', label: '\u2728 IA', icon: '' },
      { id: 'link', label: '\uD83D\uDD17 Lier', icon: '' },
      { id: 'info', label: '\uD83D\uDCC1 Projet', icon: '' },
      { id: 'create', label: '\u2795 Cr\u00e9er', icon: '' },
      { id: 'reply', label: '\uD83D\uDCAC R\u00e9pondre', icon: '' },
      { id: 'settings', label: '\u2699\uFE0F', icon: '' },
    ];
  }

  // Determine which tab to activate: URL param > IA (mode lecture) > compose (mode compose).
  let defaultTab: string;
  if (tab && tabs.some(t => t.id === tab)) {
    defaultTab = tab;
  } else {
    defaultTab = isCompose ? 'compose' : 'ia';
  }

  app.innerHTML = `
    <div class="header">
      <span class="header-logo">ATLAS</span>
      <span class="header-subtitle">GOOD VIBES</span>
    </div>
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

function renderSetup(app: HTMLElement): void {
  app.innerHTML = `
    <div class="header">
      <span class="header-logo">ATLAS</span>
      <span class="header-subtitle">Configuration initiale</span>
    </div>
    <div class="content">
      <div class="setup-card">
        <h3>Bienvenue dans ATLAS</h3>
        <p>Configurez votre acc\u00e8s pour lier vos emails aux projets GOOD VIBES.</p>
      </div>
    </div>
    <div id="panel-content" class="content"></div>
  `;

  const content = document.getElementById('panel-content')!;
  currentPanel = new SettingsPanel(content, () => renderApp());
}

// ── Office.js Initialization ──

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    console.log('[ATLAS] Outlook Add-in loaded');
    renderApp();

    // Quand le task-pane est ÉPINGLÉ par l'utilisateur (icône 📌 en haut),
    // il reste ouvert pendant qu'il change de mail. Sans handler, le contenu
    // reste figé sur l'ancien mail. On écoute ItemChanged et on re-render
    // pour rafraîchir tous les panels (tags IA, projet lié, etc.).
    try {
      Office.context.mailbox.addHandlerAsync(
        Office.EventType.ItemChanged,
        () => {
          console.log('[ATLAS] ItemChanged → re-render');
          // Reset tab pour repartir sur l'IA par défaut au changement de mail
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
    renderApp();
  }
});
