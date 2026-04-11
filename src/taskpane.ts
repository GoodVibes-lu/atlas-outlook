/**
 * ATLAS Outlook Add-in — Main Taskpane Entry Point
 *
 * Detects context (read/compose) and renders the appropriate panel.
 * Manages navigation between tabs and handles Office.js initialization.
 */

import { LinkPanel } from './components/link-panel';
import { ComposePanel } from './components/compose-panel';
import { ProjectInfoPanel } from './components/project-info';
import { CreateProjectPanel } from './components/create-project';
import { SettingsPanel } from './components/settings';
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

function getMode(): AddinMode {
  const params = new URLSearchParams(window.location.search);
  return (params.get('mode') as AddinMode) || 'read';
}

// ── Rendering ──

function renderApp(): void {
  const app = document.getElementById('app')!;
  const mode = getMode();

  if (!isConfigured()) {
    renderSetup(app);
    return;
  }

  // Determine available tabs based on mode
  const isCompose = mode === 'compose';
  const tabs = isCompose
    ? [
        { id: 'compose', label: '📝 Templates', icon: '' },
        { id: 'settings', label: '⚙️', icon: '' },
      ]
    : [
        { id: 'link', label: '🔗 Lier', icon: '' },
        { id: 'info', label: '📁 Projet', icon: '' },
        { id: 'create', label: '➕ Créer', icon: '' },
        { id: 'settings', label: '⚙️', icon: '' },
      ];

  const defaultTab = isCompose ? 'compose' : 'link';

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
  app.querySelectorAll('.nav-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      const tabId = tab.getAttribute('data-tab')!;
      switchTab(tabId);
    });
  });

  switchTab(defaultTab);
}

function switchTab(tabId: string): void {
  if (tabId === currentTab) return;
  currentTab = tabId;

  // Update active tab styling
  document.querySelectorAll('.nav-tab').forEach(tab => {
    tab.classList.toggle('active', tab.getAttribute('data-tab') === tabId);
  });

  // Destroy current panel
  currentPanel?.destroy();
  currentPanel = null;

  const content = document.getElementById('panel-content')!;
  const userName = getUserName();

  switch (tabId) {
    case 'link':
      currentPanel = new LinkPanel(content, userName);
      break;
    case 'info':
      currentPanel = new ProjectInfoPanel(content);
      break;
    case 'create':
      currentPanel = new CreateProjectPanel(content);
      break;
    case 'compose':
      currentPanel = new ComposePanel(content, userName);
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
        <p>Configurez votre accès pour lier vos emails aux projets GOOD VIBES.</p>
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
  } else {
    // Running outside Office (dev mode)
    console.log('[ATLAS] Running outside Office.js — dev mode');
    renderApp();
  }
});
