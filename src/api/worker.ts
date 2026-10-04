/**
 * worker.ts — Accès du complément aux données ATLAS et à l'IA via le worker.
 *
 * Le complément ne détient AUCUN secret (03/10/2026) : plus de clé Anthropic ni de jeton Airtable
 * dans le localStorage / roamingSettings. Toutes les lectures/écritures ATLAS et tous les appels IA
 * passent par les routes `/api/plugin/atlas/*` du worker (liste blanche d'opérations, IA par
 * l'Agency Brain), avec le jeton Microsoft de l'utilisateur : connexion automatique (nested app
 * authentication), sinon SSO Office, sinon jeton Graph collé dans Réglages. Le worker vérifie le
 * jeton (signature Entra ID, tenant GOOD VIBES, audience) : worker/portal-api/plugin-auth.ts.
 */

const DEFAULT_WORKER_URL = 'https://worker.vibes.lu';

/** URL du worker (surchargeable au build : VITE_ATLAS_WORKER_URL). */
export const WORKER_BASE: string = String(
  ((import.meta as any).env?.VITE_ATLAS_WORKER_URL as string | undefined) || DEFAULT_WORKER_URL,
).replace(/\/$/, '');

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
//   3. Jeton Graph collé dans les Réglages (repli de transition, courte durée).
// Les voies 1 et 2 donnent un jeton émis POUR l'application du complément (audience
// `api://<domaine>/<client id>`), vérifié localement par le worker (worker/portal-api/plugin-auth.ts).
// Aucun échange « on-behalf-of » : le worker n'a besoin que de l'identité de l'appelant.
//
// Configuration au build (outlook-addin/.env.production.local, cf. docs/agent-inbox-entra-id.md) :
//   VITE_ATLAS_ADDIN_CLIENT_ID  client id de l'application Entra du complément (sans lui : voie 1 coupée)
//   VITE_ATLAS_ADDIN_RESOURCE   facultatif, URI d'ID d'application (défaut api://goodvibes-lu.github.io/<client id>)
//   VITE_MS_TENANT_ID           facultatif, tenant GOOD VIBES (défaut ci-dessous, identifiant public)
// Dépendance de la voie 1 : `@azure/msal-browser` (^5) à ajouter aux dépendances de
// outlook-addin/package.json (aujourd'hui résolue depuis le node_modules racine d'ATLAS) ; import
// protégé : si MSAL est absent ou échoue, on passe à la voie 2.

const env = ((import.meta as any).env || {}) as Record<string, string | undefined>;
const GV_TENANT_ID = (env.VITE_MS_TENANT_ID || '50200505-c9df-4b9a-b565-1562456aaefa').trim();
// Client id public (pas un secret) de l'application Entra « ATLAS Outlook (complément) », créée le 04/10/2026.
const ADDIN_CLIENT_ID = (env.VITE_ATLAS_ADDIN_CLIENT_ID || 'fc36080c-dcf7-4784-a5f8-7fa994799371').trim();
const ADDIN_RESOURCE = (env.VITE_ATLAS_ADDIN_RESOURCE || (ADDIN_CLIENT_ID ? `api://goodvibes-lu.github.io/${ADDIN_CLIENT_ID}` : '')).trim().replace(/\/$/, '');
const ADDIN_SCOPE = ADDIN_RESOURCE ? `${ADDIN_RESOURCE}/access_as_user` : '';

let cachedToken: string | null = null;
let cachedUntil = 0;
/** Durée de vie par défaut d'un jeton dont l'expiration est illisible (jeton Graph opaque). */
const TOKEN_TTL = 45 * 60 * 1000;
/** Marge avant expiration : on renouvelle 5 min avant. */
const TOKEN_MARGIN = 5 * 60 * 1000;

/** Expiration (ms) lue dans le jeton, sans vérification (le worker vérifie). */
function tokenExpiry(token: string): number {
  try {
    const part = token.split('.')[1];
    if (!part) return 0;
    const json = JSON.parse(atob(part.replace(/-/g, '+').replace(/_/g, '/')));
    return typeof json?.exp === 'number' ? json.exp * 1000 : 0;
  } catch { return 0; }
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

function naaSupported(): boolean {
  if (hostedNaa) return !!ADDIN_SCOPE;
  try {
    return !!ADDIN_SCOPE && typeof Office !== 'undefined'
      && !!Office.context?.requirements?.isSetSupported('NestedAppAuth', '1.1');
  } catch { return false; }
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
        return null;
      }
    })();
  }
  return naaClient;
}

async function getNaaToken(forceRefresh: boolean): Promise<string> {
  if (!naaSupported()) return '';
  const pca = await getNaaClient();
  if (!pca) return '';
  const account = (pca.getActiveAccount?.() || pca.getAllAccounts?.()?.[0]) ?? undefined;
  const request: Record<string, unknown> = { scopes: [ADDIN_SCOPE], ...(account ? { account } : {}), ...(forceRefresh ? { forceRefresh: true } : {}) };
  try {
    const r = await pca.acquireTokenSilent(request);
    if (r?.accessToken) return String(r.accessToken);
  } catch { /* interaction nécessaire : fenêtre de connexion ci-dessous */ }
  try {
    const r = await pca.acquireTokenPopup({ scopes: [ADDIN_SCOPE] });
    return r?.accessToken ? String(r.accessToken) : '';
  } catch (e) {
    console.warn('[worker] nested app authentication échouée :', (e as Error)?.message || e);
    return '';
  }
}

// Voie 2 : SSO Office
async function getSsoToken(): Promise<string> {
  if (typeof Office === 'undefined' || !Office.auth) return '';
  try {
    // Pas de forMSGraphAccess : le worker n'échange pas ce jeton contre un jeton Graph.
    return await Office.auth.getAccessToken({ allowSignInPrompt: true, allowConsentPrompt: true });
  } catch (e) {
    console.warn('[worker] Office SSO indisponible :', (e as Error)?.message || e);
    return '';
  }
}

/**
 * Jeton Microsoft à présenter au worker (`Authorization: Bearer …`), mis en cache jusqu'à
 * 5 min avant son expiration. `forceRefresh` après un 401.
 */
export async function getWorkerToken(forceRefresh = false): Promise<string> {
  if (!forceRefresh && cachedToken && Date.now() < cachedUntil) return cachedToken;
  let token = await getNaaToken(forceRefresh);
  if (!token) token = await getSsoToken();
  // Plus de repli « jeton Graph collé » (audit M8) : le worker n'accepte que le jeton du complément.
  if (!token) throw new Error('Connexion Microsoft impossible (connexion automatique et SSO indisponibles).');
  const exp = tokenExpiry(token);
  cachedToken = token;
  cachedUntil = exp ? exp - TOKEN_MARGIN : Date.now() + TOKEN_TTL;
  return token;
}

/** Appelle une route `/api/plugin/atlas/<route>` du worker. Lève une erreur si la réponse n'est pas OK. */
export async function callAtlasWorker<T = Record<string, unknown>>(route: string, payload: unknown = {}): Promise<T> {
  const send = async (token: string) => fetch(`${WORKER_BASE}/api/plugin/atlas/${route}`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(payload ?? {}),
  });
  let res = await send(await getWorkerToken());
  if (res.status === 401) {
    cachedToken = null;
    res = await send(await getWorkerToken(true));
  }
  const data: any = await res.json().catch(() => ({}));
  if (!res.ok || data?.ok === false) {
    throw new Error(`ATLAS ${res.status}: ${String(data?.error || 'erreur').slice(0, 120)}`);
  }
  return data as T;
}
