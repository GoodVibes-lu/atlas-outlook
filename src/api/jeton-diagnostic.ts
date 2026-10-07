/**
 * jeton-diagnostic.ts — logique PURE (sans DOM ni réseau) de la connexion du complément et du
 * tableau de bord (07/10/2026, retour de Charles : sur Outlook iPhone, l'application ATLAS de
 * l'onglet « Applications » affichait « Ta session Outlook doit être rouverte » sans dire quoi faire).
 *
 *   - `problemesJeton` : contrôle LOCAL (non vérifié, le worker vérifie la signature) des claims d'un
 *     jeton avant de l'utiliser : un jeton qui viserait une autre application, sans l'étendue
 *     `access_as_user`, d'un autre tenant ou hors @vibes.lu serait refusé par le worker (401) ;
 *     on le jette et on passe à la voie de connexion suivante.
 *   - `expliquerEchecConnexion` : à partir du détail technique des voies tentées (NAA, SSO Teams,
 *     fenêtre de connexion…), une phrase qui dit exactement ce qui manque et où le corriger.
 *   - `expliquerRaisonWorker` : idem pour la raison renvoyée par `GET /api/plugin/agent/jeton`.
 *
 * Tests : scripts/test-jeton-diagnostic.mjs.
 */

/** Applications Microsoft 365 à pré-autoriser sur l'étendue `access_as_user` (Teams, Outlook, Microsoft 365). */
export const CLIENTS_M365_A_PREAUTORISER: ReadonlyArray<{ id: string; nom: string }> = [
  { id: '1fec8e78-bce4-4aaf-ab1b-5451cc387264', nom: 'Teams bureau et mobile' },
  { id: '5e3ce6c0-2b1f-4285-8d4b-75ee78787346', nom: 'Teams web' },
  { id: '4765445b-32c6-49b0-83e6-1d93765276ca', nom: 'Microsoft 365 web' },
  { id: '0ec893e0-5785-4de6-99da-4ed124e5296c', nom: 'Microsoft 365 bureau' },
  { id: 'd3590ed6-52b3-4102-aeff-aad2292ab01c', nom: 'Microsoft 365 mobile et Outlook bureau' },
  { id: 'bc59ab01-8403-45c6-8796-ac3ef710b3e3', nom: 'Outlook web' },
  { id: '27922004-5251-4030-b22d-91ecd9a37ea4', nom: 'Outlook mobile' },
];

export interface ClaimsJeton {
  aud?: unknown; scp?: unknown; tid?: unknown; exp?: unknown; nbf?: unknown;
  preferred_username?: unknown; upn?: unknown; email?: unknown; idtyp?: unknown; roles?: unknown; ver?: unknown;
}

/** Claims d'un JWT (sans vérification). null si illisible. Fonctionne sous Node et dans le navigateur. */
export function lireClaims(token: string): ClaimsJeton | null {
  try {
    const part = String(token || '').split('.')[1];
    if (!part) return null;
    const b64 = part.replace(/-/g, '+').replace(/_/g, '/');
    const pad = b64 + '='.repeat((4 - (b64.length % 4)) % 4);
    const bin = atob(pad); // navigateur et Node 16+
    const json = decodeURIComponent(Array.from(bin, (c: string) => `%${c.charCodeAt(0).toString(16).padStart(2, '0')}`).join(''));
    const o = JSON.parse(json);
    return o && typeof o === 'object' ? o as ClaimsJeton : null;
  } catch { return null; }
}

export interface JetonAttendu {
  clientId: string;
  /** URI d'ID d'application, ex. api://goodvibes-lu.github.io/<client id>. */
  resource: string;
  tenantId: string;
  scope?: string;
  domaine?: string;
  nowMs: number;
}

export type ProblemeJeton = 'illisible' | 'audience' | 'scope' | 'tenant' | 'expire' | 'compte' | 'application';

/**
 * Ce qui ferait refuser le jeton par le worker (même règles que worker/portal-api/plugin-auth.ts,
 * sauf la signature). Liste vide : le jeton peut être présenté.
 */
export function problemesJeton(token: string, att: JetonAttendu): ProblemeJeton[] {
  const c = lireClaims(token);
  if (!c) return ['illisible'];
  const out: ProblemeJeton[] = [];
  const aud = String(c.aud ?? '').trim().toLowerCase();
  const id = att.clientId.trim().toLowerCase();
  const res = att.resource.trim().toLowerCase().replace(/\/$/, '');
  if (!(aud === id || aud === res || aud === `api://${id}` || (aud.startsWith('api://') && aud.endsWith(`/${id}`)))) out.push('audience');
  if (c.idtyp === 'app' || (!c.scp && c.roles)) out.push('application');
  const scopes = String(c.scp ?? '').split(/\s+/).filter(Boolean);
  if (!scopes.includes(att.scope || 'access_as_user')) out.push('scope');
  if (String(c.tid ?? '').toLowerCase() !== att.tenantId.trim().toLowerCase()) out.push('tenant');
  if (typeof c.exp !== 'number' || c.exp * 1000 + 120_000 < att.nowMs) out.push('expire');
  const email = String(c.preferred_username || c.upn || c.email || '').trim().toLowerCase();
  if (!email.endsWith(att.domaine || '@vibes.lu')) out.push('compte');
  return out;
}

/** Résumé court d'un jeton pour le diagnostic (jamais le jeton lui-même ni l'adresse complète). */
export function resumeJeton(token: string): string {
  const c = lireClaims(token);
  if (!c) return 'jeton illisible';
  const email = String(c.preferred_username || c.upn || c.email || '');
  const dom = email.includes('@') ? `@${email.split('@')[1]}` : '?';
  return `aud=${String(c.aud ?? '?')} scp=${String(c.scp ?? '?')} ver=${String(c.ver ?? '?')} compte=${dom}`;
}

const LIBELLE_PROBLEME: Record<ProblemeJeton, string> = {
  illisible: 'jeton illisible',
  audience: 'jeton émis pour une autre application que « ATLAS Outlook (complément) »',
  scope: 'étendue access_as_user absente du jeton',
  tenant: 'compte d\'un autre tenant que GOOD VIBES',
  expire: 'jeton expiré (heure du téléphone ?)',
  compte: 'compte hors @vibes.lu',
  application: 'jeton d\'application (pas de personne connectée)',
};

export function libelleProblemes(p: ProblemeJeton[]): string {
  return p.map(x => LIBELLE_PROBLEME[x]).join(', ');
}

export interface ContexteConnexion {
  /**
   * 'teams' : application de la barre de gauche / onglet Applications ; 'dialogue' : grande fenêtre du
   * complément ; 'office' : panneau ; 'navigateur' : page ouverte seule dans un navigateur (repli mobile).
   */
  hote: 'teams' | 'dialogue' | 'office' | 'navigateur';
  /** Nom de l'application hôte (« Outlook », « Teams »…) et plateforme (« ios », « android », « desktop », « web »), si connus. */
  appli?: string;
  plateforme?: string;
}

const ENTRA = 'Entra › Inscriptions d\'applications › ATLAS Outlook (complément)';
const PAGE = 'https://goodvibes-lu.github.io/atlas-outlook/tableau-de-bord.html';

/**
 * Phrase destinée à l'utilisateur (et à Charles) à partir du détail des voies tentées
 * (`getWorkerToken`, entrées séparées par « | »). Dit QUOI manque et OÙ le corriger.
 */
export function expliquerEchecConnexion(detail: string, ctx: ContexteConnexion): string {
  const d = String(detail || '');
  const ou = [ctx.appli, ctx.plateforme === 'ios' ? 'iPhone / iPad' : ctx.plateforme === 'android' ? 'Android' : ''].filter(Boolean).join(' ');
  const ici = ou ? ` dans ${ou}` : '';
  if (/13002|user_cancel|CancelledByUser|annulée/i.test(d)) {
    return 'Connexion à ton compte Microsoft annulée : touche « Se connecter » pour recommencer.';
  }
  if (/AADSTS500011|invalid_resource|resource principal/i.test(d)) {
    return `Connexion refusée par Microsoft : l'URI d'ID d'application est introuvable. À corriger dans ${ENTRA} › Exposer une API : « api://goodvibes-lu.github.io/<ID d'application> », identique au bloc webApplicationInfo de l'application Teams.`;
  }
  if (/AADSTS65005|resourceDisabled|App resource defined in manifest/i.test(d)) {
    return `Connexion refusée${ici} : la ressource déclarée par l'application ATLAS (webApplicationInfo) ne correspond pas à l'application Entra, ou le domaine de la page n'est pas celui de l'URI d'ID d'application. À vérifier dans ${ENTRA} › Exposer une API.`;
  }
  if (/AADSTS65001|consent_required|resourceRequiresConsent|interaction_required.*consent|AADSTS90094|AADSTS90008/i.test(d)) {
    return `Connexion refusée${ici} : ${ctx.hote === 'teams' ? 'cette application Microsoft 365 n\'est pas pré-autorisée' : 'consentement manquant'} sur l'étendue access_as_user. À faire une fois dans ${ENTRA} › Exposer une API › Applications clientes autorisées : ajouter ${CLIENTS_M365_A_PREAUTORISER.map(c => `${c.id} (${c.nom})`).join(', ')}, puis API autorisées › Accorder un consentement d'administrateur.`;
  }
  if (/AADSTS50011|redirect_uri|redirect uri/i.test(d)) {
    return `Connexion refusée : adresse de retour non déclarée. À ajouter dans ${ENTRA} › Authentification › Application monopage (SPA) : « ${PAGE} » et « brk-multihub://goodvibes-lu.github.io ».`;
  }
  if (/AADSTS700016|unauthorized_client|invalid_client/i.test(d)) {
    return `Connexion refusée : l'application Entra « ATLAS Outlook (complément) » est introuvable ou désactivée dans le tenant GOOD VIBES (${ENTRA}).`;
  }
  if (/AADSTS53003|AADSTS53000|AADSTS53001|conditional access|accès conditionnel/i.test(d)) {
    return 'Connexion bloquée par une règle d\'accès conditionnel Microsoft 365 (appareil non conforme ou non géré) : à vérifier dans Entra › Sécurité › Accès conditionnel.';
  }
  if (/AADSTS50076|AADSTS50079|mfa/i.test(d)) {
    return 'Microsoft demande une vérification en deux étapes : touche « Se connecter » pour la faire.';
  }
  if (/jeton refusé|refusé par le worker/i.test(d)) {
    return `Connexion obtenue${ici} mais le jeton ne convient pas (${d.split('jeton refusé : ')[1]?.split(' |')[0] || 'voir le détail'}).`;
  }
  if (ctx.hote === 'teams') {
    const naaAbsent = /NAA absente|nestedAppAuthBridge/i.test(d);
    const ssoTente = /SSO Teams/i.test(d);
    const fenetreBloquee = /fenêtre|FailedToOpenWindow|popup|window/i.test(d);
    if (naaAbsent && ssoTente) {
      return `Connexion impossible${ici} : la connexion automatique (nested app authentication) n'est pas fournie ici et l'authentification unique (SSO) Teams a échoué. Le plus souvent, il manque la pré-autorisation des applications Microsoft 365 sur l'étendue access_as_user (${ENTRA} › Exposer une API › Applications clientes autorisées, dont Outlook mobile 27922004-5251-4030-b22d-91ecd9a37ea4)${fenetreBloquee ? ' ; la fenêtre de connexion de secours n\'a pas pu s\'ouvrir' : ''}. Touche « Se connecter » pour essayer la fenêtre de connexion Microsoft.`;
    }
  }
  return 'Connexion à ton compte Microsoft impossible : touche « Se connecter ». Si ça persiste, quitte et relance Outlook.';
}

/** Raison courte de `GET /api/plugin/agent/jeton` (worker) → phrase. Chaîne vide si inconnue. */
export function expliquerRaisonWorker(raison: string): string {
  switch (String(raison || '')) {
    case 'audience':
    case 'graph_refuse':
      return 'Le worker refuse le jeton : il vise une autre application que « ATLAS Outlook (complément) » (MS_ADDIN_CLIENT_ID / MS_ADDIN_APP_ID_URI du worker à comparer avec l\'application Entra).';
    case 'scope':
      return `Le worker refuse le jeton : l'étendue access_as_user est absente (${ENTRA} › Exposer une API : étendue access_as_user « Activée »).`;
    case 'tenant':
    case 'issuer':
      return 'Le worker refuse le jeton : compte d\'un autre tenant. Dans Outlook, utilise ton compte @vibes.lu.';
    case 'identity':
      return 'Le worker refuse le jeton : compte invité ou hors @vibes.lu. Dans Outlook, utilise ton compte @vibes.lu.';
    case 'employe_inactif':
      return 'Le worker refuse l\'accès : ton compte n\'a pas de fiche Employés active dans ATLAS.';
    case 'expired':
    case 'not_yet_valid':
      return 'Le worker refuse le jeton : heure décalée (téléphone ou NAS). Vérifie la mise à l\'heure automatique.';
    case 'unknown_key':
      return 'Le worker n\'atteint pas les clés Microsoft (login.microsoftonline.com) : accès Internet sortant du NAS à vérifier.';
    case 'signature':
    case 'alg':
    case 'malformed':
      return 'Le worker refuse le jeton : signature invalide.';
    case 'not_configured':
      return 'Le worker n\'est pas configuré pour le complément (MS_ADDIN_CLIENT_ID absent).';
    case 'app_token':
      return 'Le worker refuse le jeton : jeton d\'application, sans personne connectée.';
    default:
      return '';
  }
}
