/**
 * tableau-de-bord.ts — point d'entrée de la page pleine largeur de l'application ATLAS dans la
 * BARRE DE GAUCHE d'Outlook (application Microsoft 365, onglet personnel : teams-app/manifest.json).
 * Bureau (nouvel Outlook Windows, Mac à confirmer) et web ; Outlook classique et mobile gardent le
 * bandeau « Ma journée » du panneau.
 *
 * Depuis le 06/10/2026 (section B du cadrage `.claude/CADRAGE-OUTLOOK-DASHBOARD.md`) : tableau de
 * bord complet (src/tableau/app.ts) : boîte triée sur toutes les boîtes permises, priorités ARGO,
 * en attente, mis de côté, envoi programmé, réponses ARGO et modèles déposés dans le brouillon
 * Outlook, fil d'équipe, engagements, prochains RDV préparés, réactivité aux clients ; clavier
 * d'abord, palette ⌘K, glisser-déposer, mise à jour en temps réel.
 *
 * Connexion : même `getWorkerToken` que le complément (worker.ts), voie « nested app
 * authentication » activée hors Office.js par `enableHostedNaa()` après l'initialisation de
 * TeamsJS. Si l'hôte ne fournit pas la connexion automatique, MSAL tente une fenêtre ; sinon la
 * page affiche l'erreur et renvoie vers ATLAS (docs/agent-inbox-essai-complement.md §5).
 *
 * DEUXIÈME HÔTE (07/10/2026) : fenêtre de dialogue Office ouverte par le complément
 * (`?hote=office`, src/api/dialogue-tableau.ts), parce que le nouvel Outlook pour Mac n'affiche pas
 * les applications de la barre de gauche. Là : Office.js chargé à la demande (pas TeamsJS), jeton
 * fourni par la page parente (panneau ou commande du ruban), sinon connexion par redirection MSAL
 * dans la fenêtre ; les mails s'ouvrent dans Outlook par la page parente (`displayMessageForm`).
 *
 * OUTLOOK MOBILE (07/10/2026, retour de Charles sur iPhone : « Ta session Outlook doit être rouverte »,
 * compteurs à 0) : l'application ATLAS apparaît dans l'onglet « Applications » d'Outlook iOS, mais
 * cet hôte ne fournit pas toujours le pont de la connexion automatique (`nestedAppAuthBridge`) ; MSAL
 * retombait alors sans le dire sur un client classique. Désormais, dans l'hôte Teams / Microsoft 365 :
 *   1. connexion automatique (NAA) seulement si le pont est là ;
 *   2. sinon authentification unique Teams (`authentication.getAuthToken`, bloc webApplicationInfo,
 *      applications Microsoft 365 pré-autorisées dans Entra) ;
 *   3. sur « Se connecter » : fenêtre de connexion de l'hôte (`authentication.authenticate`) qui
 *      ouvre cette page en mode `?auth=debut` (connexion MSAL par redirection, URI SPA de la page) ;
 *   4. chaque jeton est contrôlé avant usage (audience, étendue, tenant, compte) et, si le worker le
 *      refuse quand même, sa raison (`GET /api/plugin/agent/jeton`) est affichée : la page dit QUOI
 *      corriger (src/api/jeton-diagnostic.ts).
 */

import {
  autoriserConnexionInteractive, enableDialogRedirectLogin, enableHostedNaa, finishDialogRedirect, getWorkerToken,
  setContexteConnexion, setExternalTokenProvider, setHoteTeams,
} from './api/worker';
import { brancherSurParent } from './api/dialogue-tableau';
import { ATLAS_BASE } from './api/platform';
import { AtlasError, getDiag, humanError } from './api/net';
import { demarrerTableau } from './tableau/app';
import { h } from './tableau/ui';

/** TeamsJS (chargé par tableau-de-bord.html depuis le CDN Microsoft), facultatif. */
const teams: any = (window as any).microsoftTeams;

/** Hôte Microsoft 365 détecté par TeamsJS : nom (« Outlook », « Teams »…) et plateforme (« ios », « android », « desktop », « web »). */
interface InfoHote { ok: boolean; appli: string; plateforme: string }

async function initHost(): Promise<InfoHote> {
  const info: InfoHote = { ok: false, appli: '', plateforme: '' };
  if (!teams?.app?.initialize) return info;
  try {
    await Promise.race([
      teams.app.initialize(),
      new Promise((_, reject) => setTimeout(() => reject(new Error('délai dépassé')), 5000)),
    ]);
    info.ok = true;
    try {
      const ctx = await teams.app.getContext();
      applyTheme(String(ctx?.app?.theme || 'default'));
      teams.app.registerOnThemeChangeHandler?.((t: string) => applyTheme(t));
      info.appli = String(ctx?.app?.host?.name || '');
      info.plateforme = String(ctx?.app?.host?.clientType || '').toLowerCase();
    } catch { /* thème du système */ }
    teams.app.notifySuccess?.();
  } catch (e) {
    // Page ouverte hors d'un hôte Microsoft 365 (navigateur seul) : on continue sans TeamsJS.
    console.warn('[tableau-de-bord] TeamsJS indisponible :', (e as Error)?.message || e);
  }
  return info;
}

/** Adresse de la page de connexion ouverte par la fenêtre de l'hôte (même page, mode `auth=debut`). */
function urlFenetreConnexion(): string {
  return `${window.location.origin}${window.location.pathname}?auth=debut`;
}

/**
 * Branche les voies de connexion propres à l'hôte Teams / Microsoft 365 (worker.ts) : SSO Teams,
 * puis, sur geste, la fenêtre de connexion de l'hôte.
 */
function brancherConnexionTeams(info: InfoHote): void {
  setContexteConnexion({ hote: 'teams', appli: info.appli, plateforme: info.plateforme });
  const auth = teams?.authentication;
  if (!info.ok || !auth) return;
  setHoteTeams({
    sso: typeof auth.getAuthToken === 'function' ? () => auth.getAuthToken() : undefined,
    fenetre: typeof auth.authenticate === 'function'
      ? () => auth.authenticate({ url: urlFenetreConnexion(), width: 600, height: 640 })
      : undefined,
  });
}

// ── Fenêtre de connexion de l'hôte Teams (`?auth=debut`) ──

const CLE_AUTH = 'teams-auth';

/** Vrai si la page est la fenêtre de connexion ouverte par `authentication.authenticate`, ou son retour de Microsoft. */
function enFenetreConnexion(): boolean {
  const params = new URLSearchParams(window.location.search);
  if (params.get('auth') === 'debut') {
    try { sessionStorage.setItem(CLE_HOTE, CLE_AUTH); } catch { /* stockage indisponible */ }
    return true;
  }
  try {
    return /[#&](code|error)=/.test(window.location.hash) && sessionStorage.getItem(CLE_HOTE) === CLE_AUTH;
  } catch { return false; }
}

/**
 * Connexion Microsoft par redirection (application Entra du complément, URI SPA = cette page), puis
 * jeton rendu à l'onglet par `authentication.notifySuccess`. La page part vers Microsoft puis revient.
 */
async function fenetreConnexion(): Promise<void> {
  const app = document.getElementById('app')!;
  app.innerHTML = '<div class="loading"><div class="spinner"></div><p>Connexion à ton compte Microsoft…</p></div>';
  await initHost();
  enableDialogRedirectLogin();
  if (/[#&](code|error)=/.test(window.location.hash)) await finishDialogRedirect();
  try {
    const token = await getWorkerToken();
    try { sessionStorage.removeItem(CLE_HOTE); } catch { /* rien */ }
    teams?.authentication?.notifySuccess?.(token);
  } catch (e) {
    const detail = e instanceof AtlasError ? e.detail : String((e as Error)?.message || e);
    // Redirection lancée : la page est en train de partir vers Microsoft, rien à signaler.
    if (/redirection\)|connexion Microsoft en cours/.test(detail)) return;
    try { sessionStorage.removeItem(CLE_HOTE); } catch { /* rien */ }
    if (typeof teams?.authentication?.notifyFailure === 'function') teams.authentication.notifyFailure(detail.slice(0, 400));
    else renderError(humanError(e));
  }
}

/** Thème d'Outlook ('default' = clair, 'dark' / 'contrast' = sombre) ; sans hôte : celui du système. */
function applyTheme(theme: string | null): void {
  if (!theme) { delete document.documentElement.dataset.theme; return; }
  const t = theme === 'dark' || theme === 'contrast' ? 'dark' : 'light';
  document.body.dataset.theme = t;
  document.documentElement.dataset.theme = t;
}

// ── Hôte « dialogue Office » (Outlook Mac) ──

const CLE_HOTE = 'atlas_tdb_hote';
const OFFICE_JS = 'https://appsforoffice.microsoft.com/lib/1/hosted/office.js';

/**
 * Vrai si la page est ouverte dans la fenêtre de dialogue du complément : paramètre `hote=office`,
 * ou retour de la connexion Microsoft (réponse dans l'adresse) commencée dans ce mode.
 */
function enDialogueOffice(): boolean {
  const params = new URLSearchParams(window.location.search);
  if (params.get('hote') === 'office') {
    try { sessionStorage.setItem(CLE_HOTE, 'office'); } catch { /* stockage indisponible */ }
    return true;
  }
  try {
    return /[#&](code|error)=/.test(window.location.hash) && sessionStorage.getItem(CLE_HOTE) === 'office';
  } catch { return false; }
}

/**
 * Charge Office.js à la demande (seulement dans le dialogue : dans l'onglet Teams il est inutile).
 * Office.js efface history.pushState / replaceState, dont MSAL a besoin : on les remet ensuite.
 */
async function chargerOfficeJs(): Promise<boolean> {
  const push = window.history.pushState;
  const replace = window.history.replaceState;
  try {
    if (typeof Office === 'undefined') {
      await new Promise<void>((resolve, reject) => {
        const s = document.createElement('script');
        s.src = OFFICE_JS;
        s.onload = () => resolve();
        s.onerror = () => reject(new Error('Office.js introuvable'));
        document.head.appendChild(s);
      });
    }
    await Promise.race([
      Office.onReady(),
      new Promise((_, reject) => setTimeout(() => reject(new Error('délai dépassé')), 8000)),
    ]);
    return true;
  } catch (e) {
    console.warn('[tableau-de-bord] Office.js indisponible :', (e as Error)?.message || e);
    return false;
  } finally {
    if (!window.history.pushState) window.history.pushState = push;
    if (!window.history.replaceState) window.history.replaceState = replace;
  }
}

/** Ouverture des liens propre à l'hôte (remplacée dans le dialogue Office). */
let openLinkHote: ((url: string) => void) | null = null;

/** Ouvre une adresse depuis l'onglet (Outlook sur le web pour un mail, ATLAS pour une fiche). */
function openLink(url: string): void {
  if (openLinkHote) { openLinkHote(url); return; }
  try {
    if (teams?.app?.openLink) { teams.app.openLink(url); return; }
  } catch { /* repli ci-dessous */ }
  window.open(url, '_blank', 'noopener');
}

function renderError(message: string): void {
  const app = document.getElementById('app')!;
  const diag = getDiag() as any;
  const detail = String(diag?.token?.lastError || '');
  app.innerHTML = `
    <div class="tb-fatal tb-panel is-raised">
      <div class="tb-h">Connexion impossible</div>
      <p>${h(message)}</p>
      <p class="tb-note">Le panneau ATLAS d'un mail (bandeau « Ma journée ») reste utilisable.</p>
      ${detail ? `<details class="tb-note"><summary>Détail technique</summary><p style="word-break:break-word">${h(detail.slice(0, 900))}</p></details>` : ''}
      <div class="tb-actions">
        <button type="button" class="tb-btn is-primary" id="tdb-retry">Se connecter</button>
        <button type="button" class="tb-btn" id="tdb-open-atlas">Ouvrir ATLAS</button>
      </div>
    </div>`;
  // Le clic autorise la fenêtre de connexion de l'hôte (jamais ouverte sans geste).
  document.getElementById('tdb-retry')?.addEventListener('click', () => { autoriserConnexionInteractive(); void start(); });
  document.getElementById('tdb-open-atlas')?.addEventListener('click', () => openLink(ATLAS_BASE));
}

async function start(): Promise<void> {
  try {
    // Connexion d'abord : son erreur exacte est plus parlante qu'un tableau vide.
    await getWorkerToken();
  } catch (e) {
    renderError(humanError(e));
    return;
  }
  demarrerTableau(document.getElementById('app')!, { openLink });
}

async function initDialogueOffice(): Promise<void> {
  setContexteConnexion({ hote: 'dialogue', appli: 'Outlook' });
  const params = new URLSearchParams(window.location.search);
  const theme = params.get('theme');
  if (theme === 'dark' || theme === 'light') applyTheme(theme);
  enableDialogRedirectLogin();
  if (await chargerOfficeJs()) {
    const lien = brancherSurParent(params.get('parent') === '1');
    setExternalTokenProvider(lien.jeton);
    openLinkHote = lien.openLink;
  }
  // Retour de la connexion par redirection : MSAL termine et revient à la page de départ.
  if (/[#&](code|error)=/.test(window.location.hash)) await finishDialogRedirect();
}

(async () => {
  document.body.classList.add('tb');
  applyTheme(null);
  if (enFenetreConnexion()) {
    await fenetreConnexion();
    return;
  }
  if (enDialogueOffice()) {
    await initDialogueOffice();
  } else {
    const info = await initHost();
    enableHostedNaa();
    brancherConnexionTeams(info);
  }
  await start();
})();
