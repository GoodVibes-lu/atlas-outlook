/**
 * worker.ts — Accès du complément aux données ATLAS et à l'IA via le worker.
 *
 * Le complément ne détient AUCUN secret (03/10/2026) : plus de clé Anthropic ni de jeton Airtable
 * dans le localStorage / roamingSettings. Toutes les lectures/écritures ATLAS et tous les appels IA
 * passent par les routes `/api/plugin/*` du worker (liste blanche d'opérations, IA par
 * l'Agency Brain), avec le jeton Microsoft de l'utilisateur : connexion automatique (nested app
 * authentication), sinon SSO Office. Le worker vérifie le jeton (signature Entra ID, tenant
 * GOOD VIBES, audience) : worker/portal-api/plugin-auth.ts.
 *
 * Robustesse (05/10/2026, retour « Load failed » de Charles) :
 *   - un seul appel de jeton à la fois (avant : bandeau, onglet et fiche demandaient chacun un jeton
 *     en parallèle, d'où plusieurs fenêtres MSAL concurrentes « interaction_in_progress ») ;
 *   - chaque appel est borné dans le temps et ses erreurs sont lisibles (net.ts) ;
 *   - les LECTURES sont retentées sur une coupure réseau : le worker ferme la connexion au bout de
 *     10 s (server.timeout) alors que la première lecture d'une liste froide (projets, clients…)
 *     peut prendre plus longtemps ; le serveur termine quand même et garde la liste en cache
 *     5 min, la nouvelle tentative la trouve. Les écritures ne sont JAMAIS retentées.
 */

import {
  AtlasError, humanMessage, isHumanText, kindForStatus, netFetch, setDiagWorkerUrl, setTokenDiag,
} from './net';
import {
  expliquerEchecConnexion, expliquerRaisonWorker, libelleProblemes, problemesJeton, resumeJeton, type ContexteConnexion,
} from './jeton-diagnostic';

const DEFAULT_WORKER_URL = 'https://worker.vibes.lu';

/** URL du worker (surchargeable au build : VITE_ATLAS_WORKER_URL). */
export const WORKER_BASE: string = String(
  ((import.meta as any).env?.VITE_ATLAS_WORKER_URL as string | undefined) || DEFAULT_WORKER_URL,
).trim().replace(/\/$/, '') || DEFAULT_WORKER_URL;
setDiagWorkerUrl(WORKER_BASE);

/** Anciennes clés de secrets stockées par les versions précédentes du complément. */
const LEGACY_SECRET_KEYS = ['atlas_addin_anthropic_key', 'atlas_addin_airtable_token'] as const;

/**
 * Efface les secrets laissés par les anciennes versions (localStorage ET roamingSettings, qui
 * les synchronisait dans la boîte Exchange). À appeler au démarrage, avant tout le reste.
 */
export async function purgeLegacySecrets(): Promise<void> {
  for (const k of LEGACY_SECRET_KEYS) {
    try { localStorage.removeItem(k); } catch { /* stockage indisponible */ }
  }
  try {
    const rs = typeof Office !== 'undefined' ? Office.context?.roamingSettings : undefined;
    if (!rs) return;
    let changed = false;
    for (const k of LEGACY_SECRET_KEYS) {
      if (rs.get(k) != null) { rs.remove(k); changed = true; }
    }
    if (!changed) return;
    await new Promise<void>((resolve) => {
      rs.saveAsync((res) => {
        if (res.status !== Office.AsyncResultStatus.Succeeded) {
          console.warn('[worker] purge roamingSettings échouée :', res.error);
        }
        resolve();
      });
    });
    console.info('[worker] anciens secrets effacés de roamingSettings');
  } catch (e) {
    console.warn('[worker] purge roamingSettings échouée :', e);
  }
}

// ── Jeton Microsoft (phase 1.1, 03/10/2026) ──
//
// Ordre d'obtention, du plus fiable au repli :
//   1. Nested app authentication (MSAL `createNestablePublicClientApplication`) : connexion
//      automatique partout, mobile compris (comptes Microsoft 365).
//   2. SSO Office (`Office.auth.getAccessToken`, bloc WebApplicationInfo du manifeste).
// Les deux voies donnent un jeton émis POUR l'application du complément (audience
// `api://<domaine>/<client id>`), vérifié localement par le worker (worker/portal-api/plugin-auth.ts).
// Aucun échange « on-behalf-of » : le worker n'a besoin que de l'identité de l'appelant.
//
// Configuration au build (outlook-addin/.env.production.local, cf. docs/agent-inbox-entra-id.md) :
//   VITE_ATLAS_ADDIN_CLIENT_ID  client id de l'application Entra du complément (sans lui : voie 1 coupée)
//   VITE_ATLAS_ADDIN_RESOURCE   facultatif, URI d'ID d'application (défaut api://goodvibes-lu.github.io/<client id>)
//   VITE_MS_TENANT_ID           facultatif, tenant GOOD VIBES (défaut ci-dessous, identifiant public)

const env = ((import.meta as any).env || {}) as Record<string, string | undefined>;
const GV_TENANT_ID = (env.VITE_MS_TENANT_ID || '50200505-c9df-4b9a-b565-1562456aaefa').trim();
// Client id public (pas un secret) de l'application Entra « ATLAS Outlook (complément) », créée le 04/10/2026.
const ADDIN_CLIENT_ID = (env.VITE_ATLAS_ADDIN_CLIENT_ID || 'fc36080c-dcf7-4784-a5f8-7fa994799371').trim();
const ADDIN_RESOURCE = (env.VITE_ATLAS_ADDIN_RESOURCE || (ADDIN_CLIENT_ID ? `api://goodvibes-lu.github.io/${ADDIN_CLIENT_ID}` : '')).trim().replace(/\/$/, '');
const ADDIN_SCOPE = ADDIN_RESOURCE ? `${ADDIN_RESOURCE}/access_as_user` : '';

let cachedToken: string | null = null;
let cachedUntil = 0;
/** Durée de vie par défaut d'un jeton dont l'expiration est illisible. */
const TOKEN_TTL = 45 * 60 * 1000;
/** Marge avant expiration : on renouvelle 5 min avant. */
const TOKEN_MARGIN = 5 * 60 * 1000;
/** Une voie de connexion qui ne répond pas (fenêtre MSAL bloquée) est abandonnée après ce délai. */
const TOKEN_STEP_TIMEOUT = 20_000;

/** Expiration (ms) lue dans le jeton, sans vérification (le worker vérifie). */
function tokenExpiry(token: string): number {
  try {
    const part = token.split('.')[1];
    if (!part) return 0;
    const json = JSON.parse(atob(part.replace(/-/g, '+').replace(/_/g, '/')));
    return typeof json?.exp === 'number' ? json.exp * 1000 : 0;
  } catch { return 0; }
}

function withTimeout<T>(p: Promise<T>, ms: number, what: string): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const t = setTimeout(() => reject(new Error(`${what} : délai dépassé (${Math.round(ms / 1000)} s)`)), ms);
    p.then(v => { clearTimeout(t); resolve(v); }, e => { clearTimeout(t); reject(e); });
  });
}

// Voie 1 : nested app authentication
let naaClient: Promise<any | null> | null = null;

/**
 * Hôte sans Office.js qui fournit lui-même la connexion automatique : le tableau de bord de la
 * barre de gauche d'Outlook (application Microsoft 365, onglet personnel, TeamsJS initialisé).
 * Là, `Office.context.requirements` n'existe pas : on tente directement MSAL (nested app
 * authentication) ; si l'hôte ne la fournit pas, MSAL échoue et le tableau de bord l'indique.
 */
let hostedNaa = false;
export function enableHostedNaa(): void {
  hostedNaa = true;
}

/**
 * Pont de la connexion automatique posé par l'hôte (TeamsJS le pose après `app.initialize()` si
 * l'hôte déclare `supports.nestedAppAuth` ; Office.js le pose dans les compléments). Sans lui,
 * `createNestablePublicClientApplication` de MSAL retombe SANS LE DIRE sur un client MSAL classique
 * (fenêtre surgissante), inutilisable dans une application d'Outlook mobile : cause du 07/10/2026
 * (Outlook iPhone, onglet « Applications »). On ne tente donc la voie NAA hébergée que si le pont existe.
 */
function naaBridgePresent(): boolean {
  try { return !!(window as any).nestedAppAuthBridge; } catch { return false; }
}

function naaSupported(): boolean {
  let ok = false;
  if (hostedNaa) ok = !!ADDIN_SCOPE && naaBridgePresent();
  else {
    try {
      ok = !!ADDIN_SCOPE && typeof Office !== 'undefined'
        && !!Office.context?.requirements?.isSetSupported('NestedAppAuth', '1.1');
    } catch { ok = false; }
  }
  setTokenDiag({ naaSupported: ok });
  return ok;
}

async function getNaaClient(): Promise<any | null> {
  if (!naaClient) {
    naaClient = (async () => {
      try {
        const msal: any = await import('@azure/msal-browser');
        if (typeof msal?.createNestablePublicClientApplication !== 'function') return null;
        return await msal.createNestablePublicClientApplication({
          auth: { clientId: ADDIN_CLIENT_ID, authority: `https://login.microsoftonline.com/${GV_TENANT_ID}` },
        });
      } catch (e) {
        console.warn('[worker] MSAL (nested app authentication) indisponible :', (e as Error)?.message || e);
        naaClient = null; // nouvel essai au prochain appel (chargement du module raté)
        return null;
      }
    })();
  }
  return naaClient;
}

async function getNaaToken(forceRefresh: boolean, errors: string[]): Promise<string> {
  if (!naaSupported()) {
    errors.push(hostedNaa ? 'NAA absente (nestedAppAuthBridge non fourni par l\'hôte)' : 'connexion automatique non proposée par Outlook');
    return '';
  }
  const pca = await getNaaClient();
  if (!pca) { errors.push('MSAL indisponible'); return ''; }
  const account = (pca.getActiveAccount?.() || pca.getAllAccounts?.()?.[0]) ?? undefined;
  const request: Record<string, unknown> = { scopes: [ADDIN_SCOPE], ...(account ? { account } : {}), ...(forceRefresh ? { forceRefresh: true } : {}) };
  try {
    const r = await withTimeout<any>(pca.acquireTokenSilent(request), TOKEN_STEP_TIMEOUT, 'NAA silencieux');
    if (r?.accessToken) return String(r.accessToken);
  } catch (e) {
    errors.push(`NAA silencieux : ${(e as any)?.errorCode || (e as Error)?.message || e}`);
  }
  try {
    const r = await withTimeout<any>(pca.acquireTokenPopup({ scopes: [ADDIN_SCOPE] }), TOKEN_STEP_TIMEOUT * 3, 'NAA fenêtre');
    return r?.accessToken ? String(r.accessToken) : '';
  } catch (e) {
    errors.push(`NAA fenêtre : ${(e as any)?.errorCode || (e as Error)?.message || e}`);
    return '';
  }
}

// Voie 2 : SSO Office
async function getSsoToken(errors: string[]): Promise<string> {
  if (typeof Office === 'undefined' || !Office.auth) { errors.push('SSO Office absent'); return ''; }
  try {
    // Pas de forMSGraphAccess : le worker n'échange pas ce jeton contre un jeton Graph.
    return await withTimeout(
      Office.auth.getAccessToken({ allowSignInPrompt: true, allowConsentPrompt: true }),
      TOKEN_STEP_TIMEOUT * 3, 'SSO Office',
    );
  } catch (e) {
    const code = (e as any)?.code;
    errors.push(`SSO ${code ?? ''} : ${(e as Error)?.message || e}`.trim());
    return '';
  }
}

// Voies propres à l'hôte Teams / Microsoft 365 (application ATLAS de la barre de gauche d'Outlook,
// onglet « Applications » d'Outlook mobile), branchées par tableau-de-bord.ts après TeamsJS :
//   - `sso` : authentification unique Teams (`authentication.getAuthToken`), jeton émis pour la
//     ressource du bloc webApplicationInfo (api://goodvibes-lu.github.io/<client id>), étendue
//     access_as_user. Exige que les applications Microsoft 365 (Teams, Outlook, Microsoft 365, web,
//     bureau, mobile) soient pré-autorisées sur cette étendue dans Entra (jeton-diagnostic.ts) ;
//   - `fenetre` : fenêtre de connexion de l'hôte (`authentication.authenticate`) qui ouvre la page en
//     mode `?auth=debut` (connexion MSAL par redirection, renvoie le jeton). Seulement sur un geste de
//     la personne (« Se connecter ») : jamais ouverte toute seule.
export interface HoteTeams {
  sso?: () => Promise<string>;
  fenetre?: () => Promise<string>;
}
let hoteTeams: HoteTeams | null = null;
export function setHoteTeams(h: HoteTeams | null): void {
  hoteTeams = h;
}
/** Autorise UNE tentative interactive (fenêtre de connexion de l'hôte) au prochain calcul du jeton. */
let interactifAutorise = false;
export function autoriserConnexionInteractive(): void {
  interactifAutorise = true;
  lastTokenFailure = null;
}

/** Hôte de la page, pour des messages d'erreur qui disent quoi corriger (null : panneau, messages historiques). */
let contexteConnexion: ContexteConnexion | null = null;
export function setContexteConnexion(c: ContexteConnexion | null): void {
  contexteConnexion = c;
}

function messageErreur(e: unknown): string {
  return String((e as any)?.errorCode || (e as Error)?.message || e);
}

async function getTeamsSsoToken(errors: string[]): Promise<string> {
  if (!hoteTeams?.sso) return '';
  try {
    return String(await withTimeout(hoteTeams.sso(), TOKEN_STEP_TIMEOUT, 'SSO Teams') || '');
  } catch (e) {
    errors.push(`SSO Teams : ${messageErreur(e)}`);
    return '';
  }
}

async function getTeamsFenetreToken(errors: string[]): Promise<string> {
  if (!hoteTeams?.fenetre || !interactifAutorise) return '';
  interactifAutorise = false;
  try {
    return String(await withTimeout(hoteTeams.fenetre(), TOKEN_STEP_TIMEOUT * 9, 'fenêtre de connexion') || '');
  } catch (e) {
    errors.push(`fenêtre de connexion : ${messageErreur(e)}`);
    return '';
  }
}

/**
 * Jeton accepté s'il passerait les contrôles du worker (hors signature) ; sinon noté et jeté, pour
 * essayer la voie suivante au lieu d'un 401 opaque (« Ta session Outlook doit être rouverte »).
 */
function jetonUtilisable(token: string, source: string, errors: string[]): string {
  if (!token || !ADDIN_CLIENT_ID) return token;
  const pb = problemesJeton(token, { clientId: ADDIN_CLIENT_ID, resource: ADDIN_RESOURCE, tenantId: GV_TENANT_ID, nowMs: Date.now() });
  if (!pb.length) return token;
  errors.push(`${source} : jeton refusé : ${libelleProblemes(pb)} (${resumeJeton(token)})`);
  return '';
}

// Voie 0 (07/10/2026) : tableau de bord ouvert dans une FENÊTRE de dialogue Office (Outlook Mac,
// qui n'affiche pas les applications de la barre de gauche). Ni la connexion automatique ni le SSO
// Office n'existent dans un dialogue : le jeton est fourni par la page qui l'a ouvert (panneau ou
// commande du ruban, src/tableau-dialogue.ts), par messageParent / messageChild (DialogApi 1.2).
let externalTokenProvider: ((forceRefresh: boolean) => Promise<string>) | null = null;
export function setExternalTokenProvider(fn: ((forceRefresh: boolean) => Promise<string>) | null): void {
  externalTokenProvider = fn;
}

// Voie 3 (dialogue seulement, repli) : connexion MSAL classique par redirection DANS la fenêtre
// de dialogue (application Entra du complément, URI de redirection SPA = la page du tableau de bord,
// à déclarer dans Entra). Utilisée seulement si la page parente ne répond pas.
let redirectLogin = false;
let redirectClient: Promise<any | null> | null = null;
export function enableDialogRedirectLogin(): void {
  redirectLogin = true;
}

/** Adresse de retour de la connexion par redirection : la page elle-même, sans paramètres. */
function redirectUri(): string {
  return `${window.location.origin}${window.location.pathname}`;
}

async function getRedirectClient(): Promise<any | null> {
  if (!redirectClient) {
    redirectClient = (async () => {
      try {
        const msal: any = await import('@azure/msal-browser');
        const pca = new msal.PublicClientApplication({
          auth: { clientId: ADDIN_CLIENT_ID, authority: `https://login.microsoftonline.com/${GV_TENANT_ID}`, redirectUri: redirectUri() },
          cache: { cacheLocation: 'localStorage' },
        });
        await pca.initialize();
        const r = await pca.handleRedirectPromise().catch(() => null);
        if (r?.account) pca.setActiveAccount(r.account);
        return pca;
      } catch (e) {
        console.warn('[worker] MSAL (redirection) indisponible :', (e as Error)?.message || e);
        redirectClient = null;
        return null;
      }
    })();
  }
  return redirectClient;
}

/** Retour de la connexion Microsoft (réponse dans l'adresse) : MSAL la traite et revient à la page de départ. */
export async function finishDialogRedirect(): Promise<void> {
  if (redirectLogin) await getRedirectClient();
}

/** Sans jeton silencieux, lance la redirection vers la connexion Microsoft (la page est quittée). */
async function getRedirectToken(forceRefresh: boolean, errors: string[]): Promise<string> {
  if (!redirectLogin || !ADDIN_SCOPE) return '';
  const pca = await getRedirectClient();
  if (!pca) { errors.push('MSAL redirection indisponible'); return ''; }
  const account = (pca.getActiveAccount?.() || pca.getAllAccounts?.()?.[0]) ?? undefined;
  if (account) {
    try {
      const r = await withTimeout<any>(pca.acquireTokenSilent({ scopes: [ADDIN_SCOPE], account, ...(forceRefresh ? { forceRefresh: true } : {}) }), TOKEN_STEP_TIMEOUT, 'redirection silencieuse');
      if (r?.accessToken) return String(r.accessToken);
    } catch (e) {
      errors.push(`redirection silencieuse : ${(e as any)?.errorCode || (e as Error)?.message || e}`);
    }
  }
  try {
    await pca.acquireTokenRedirect({ scopes: [ADDIN_SCOPE], redirectStartPage: window.location.href });
    errors.push('connexion Microsoft en cours (redirection)');
  } catch (e) {
    errors.push(`redirection : ${(e as any)?.errorCode || (e as Error)?.message || e}`);
  }
  return '';
}

let tokenInFlight: Promise<string> | null = null;
/** Dernier échec (évite d'ouvrir une fenêtre de connexion par appel en échec). */
let lastTokenFailure: { at: number; err: AtlasError } | null = null;
const TOKEN_FAILURE_COOLDOWN = 4000;

/**
 * Jeton Microsoft à présenter au worker (`Authorization: Bearer …`), mis en cache jusqu'à
 * 5 min avant son expiration. `forceRefresh` après un 401. Un seul calcul à la fois : les appels
 * simultanés attendent le même résultat.
 */
export async function getWorkerToken(forceRefresh = false): Promise<string> {
  if (!forceRefresh && cachedToken && Date.now() < cachedUntil) return cachedToken;
  if (tokenInFlight) return tokenInFlight;
  if (!forceRefresh && lastTokenFailure && Date.now() - lastTokenFailure.at < TOKEN_FAILURE_COOLDOWN) throw lastTokenFailure.err;
  tokenInFlight = (async () => {
    const errors: string[] = [];
    let source: 'naa' | 'sso' | 'parent' | 'redirection' | '' = '';
    let token = '';
    if (externalTokenProvider) {
      try { token = await externalTokenProvider(forceRefresh); } catch (e) { errors.push(`page parente : ${(e as Error)?.message || e}`); }
      token = jetonUtilisable(token, 'page parente', errors);
      if (token) source = 'parent';
    }
    // Dans un dialogue Office (redirectLogin), ni NAA ni SSO Office : on passe à la redirection.
    if (!token && !redirectLogin) {
      token = jetonUtilisable(await getNaaToken(forceRefresh, errors), 'NAA', errors);
      if (token) source = 'naa';
    }
    // Hôte Teams / Microsoft 365 : SSO Teams, puis (sur geste) la fenêtre de connexion de l'hôte.
    if (!token && hoteTeams) {
      token = jetonUtilisable(await getTeamsSsoToken(errors), 'SSO Teams', errors);
      if (token) source = 'sso';
      if (!token) {
        token = jetonUtilisable(await getTeamsFenetreToken(errors), 'fenêtre de connexion', errors);
        if (token) source = 'redirection';
      }
    }
    if (!token && !redirectLogin && !hoteTeams) {
      token = jetonUtilisable(await getSsoToken(errors), 'SSO Office', errors);
      if (token) source = 'sso';
    }
    if (!token) {
      token = jetonUtilisable(await getRedirectToken(forceRefresh, errors), 'redirection', errors);
      if (token) source = 'redirection';
    }
    if (!token) {
      const detail = errors.join(' | ') || 'aucune voie de connexion';
      setTokenDiag({ source: '', until: 0, lastError: detail, lastErrorAt: Date.now() });
      // 13002 / user_cancelled : la personne a fermé la fenêtre de connexion.
      const cancelled = /13002|user_cancel/i.test(detail);
      const err = new AtlasError('session', contexteConnexion
        ? expliquerEchecConnexion(detail, contexteConnexion)
        : cancelled
          ? 'Connexion à ton compte Microsoft annulée : clique sur Réessayer pour te connecter.'
          : humanMessage('session'), { route: 'jeton', detail });
      lastTokenFailure = { at: Date.now(), err };
      throw err;
    }
    const exp = tokenExpiry(token);
    cachedToken = token;
    cachedUntil = exp ? exp - TOKEN_MARGIN : Date.now() + TOKEN_TTL;
    lastTokenFailure = null;
    setTokenDiag({ source, until: cachedUntil });
    return token;
  })();
  try {
    return await tokenInFlight;
  } finally {
    tokenInFlight = null;
  }
}

/** Oublie le jeton en cache (bouton « Réessayer » après une erreur de session). */
export function resetWorkerToken(): void {
  cachedToken = null;
  cachedUntil = 0;
  lastTokenFailure = null;
  setTokenDiag({ source: '', until: 0 });
}

// ── Appels au worker ──

const sleep = (ms: number) => new Promise(r => setTimeout(r, ms));

export interface WorkerRequestOptions {
  /** Délai maximal d'une tentative (défaut 45 s). */
  timeoutMs?: number;
  /** Lecture sans effet de bord : retentée sur une coupure réseau (jamais une écriture). */
  retry?: boolean;
  /** Prévenu avant chaque nouvelle tentative (message « encore un instant »). */
  onRetry?: (attempt: number) => void;
}

/** Délais entre tentatives d'une lecture (cf. en-tête : liste froide plus longue que 10 s). */
const RETRY_DELAYS = [3000, 8000];

/**
 * Appelle `/api/plugin/<path>` du worker (GET, POST ou DELETE JSON) avec le jeton Microsoft.
 * Lève une AtlasError (message lisible) si la réponse n'est pas OK. Nouvelle tentative sur 401
 * (jeton renouvelé) et, pour les lectures, sur une coupure réseau.
 */
export async function workerRequest<T = Record<string, unknown>>(
  method: 'GET' | 'POST' | 'DELETE', path: string, payload?: unknown, opts: WorkerRequestOptions = {},
): Promise<T> {
  const url = `${WORKER_BASE}/api/plugin/${path.replace(/^\//, '')}`;
  const label = path.split('?')[0];
  const withBody = method !== 'GET';
  const send = async (token: string) => netFetch(url, {
    method,
    headers: { Authorization: `Bearer ${token}`, ...(withBody ? { 'Content-Type': 'application/json' } : {}) },
    body: withBody ? JSON.stringify(payload ?? {}) : undefined,
    cache: 'no-store',
  }, { service: 'atlas', label, timeoutMs: opts.timeoutMs ?? 45_000 });

  const attempts = opts.retry ? RETRY_DELAYS.length + 1 : 1;
  let lastErr: unknown = null;
  for (let i = 0; i < attempts; i++) {
    if (i > 0) { opts.onRetry?.(i); await sleep(RETRY_DELAYS[i - 1]); }
    try {
      let res = await send(await getWorkerToken());
      if (res.status === 401) {
        cachedToken = null;
        const frais = await getWorkerToken(true);
        res = await send(frais);
        if (res.status === 401 && contexteConnexion) {
          // Jeton obtenu mais refusé deux fois : on demande au worker POURQUOI (raison courte).
          const raison = await raisonRefusWorker(frais);
          setTokenDiag({ lastError: `worker 401 : ${raison || '?'} (${resumeJeton(frais)})`, lastErrorAt: Date.now() });
          const texte = expliquerRaisonWorker(raison) + (raison === 'employe_inactif' && dernierCompteRefuse ? ` Compte reconnu : ${dernierCompteRefuse}.` : '');
          if (texte) throw new AtlasError('session', texte, { status: 401, route: label, detail: `HTTP 401 · raison ${raison}` });
        }
      }
      const data: any = await res.json().catch(() => ({}));
      if (!res.ok || data?.ok === false) {
        const status = res.ok ? 400 : res.status;
        const kind = kindForStatus(status);
        const serverText = data?.error ?? data?.message;
        throw new AtlasError(kind, isHumanText(serverText) && kind !== 'serveur' ? String(serverText).trim() : humanMessage(kind), {
          status, route: label, data, detail: `HTTP ${res.status}${serverText ? ` · ${String(serverText).slice(0, 160)}` : ''}`,
        });
      }
      return data as T;
    } catch (e) {
      lastErr = e;
      const retryable = e instanceof AtlasError && (e.kind === 'reseau' || e.kind === 'delai' || (e.kind === 'serveur' && e.status >= 502));
      if (!retryable) throw e;
    }
  }
  throw lastErr;
}

/** Raison du refus d'un jeton par le worker (`GET /api/plugin/agent/jeton`) ; '' si indisponible. */
async function raisonRefusWorker(token: string): Promise<string> {
  try {
    const r = await netFetch(`${WORKER_BASE}/api/plugin/agent/jeton`, {
      method: 'GET', headers: { Authorization: `Bearer ${token}` }, cache: 'no-store',
    }, { service: 'atlas', label: 'agent/jeton', timeoutMs: 15_000 });
    if (r.ok) return '';
    const data: any = await r.json().catch(() => ({}));
    if (typeof data?.compte === 'string' && data.compte) dernierCompteRefuse = data.compte;
    return typeof data?.raison === 'string' ? data.raison : '';
  } catch { return ''; }
}
/** Compte reconnu par le worker lors du dernier refus (diagnostic affiché). */
let dernierCompteRefuse = '';

/** Lectures de la liste blanche `/api/plugin/atlas/*` (retentées sur coupure réseau). */
const ATLAS_READS = new Set([
  'projets/list', 'projets/extra', 'tiers/list', 'tiers/resolve', 'contacts/list', 'contacts/argo-profile',
  'employes/list', 'templates/list', 'emails/linked-ids', 'emails/count', 'folder-mapping/get',
  'tags/by-email', 'tags/by-conversation', 'reunion/resoudre', 'reunion/reunion',
]);

/** Appelle une route `/api/plugin/atlas/<route>` du worker. Lève une AtlasError lisible en cas d'échec. */
export async function callAtlasWorker<T = Record<string, unknown>>(route: string, payload: unknown = {}, opts: WorkerRequestOptions = {}): Promise<T> {
  const ai = route.startsWith('ai/');
  return workerRequest<T>('POST', `atlas/${route}`, payload, {
    retry: ATLAS_READS.has(route),
    // IA : le worker accorde 120 s à ces routes.
    timeoutMs: ai ? 125_000 : 45_000,
    ...opts,
  });
}
