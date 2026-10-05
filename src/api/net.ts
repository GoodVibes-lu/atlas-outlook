/**
 * net.ts · Appels réseau du complément, erreurs lisibles et journal de diagnostic (05/10/2026).
 *
 * Pourquoi : sur Outlook Mac (WKWebView), un `fetch` qui n'obtient AUCUNE réponse HTTP exploitable
 * (connexion coupée par le serveur, réponse sans en-têtes CORS, proxy qui renvoie une page d'erreur,
 * délai dépassé) lève un TypeError dont le message brut est « Load failed ». Ce texte remontait tel
 * quel dans le panneau. Ici :
 *   - chaque erreur devient une AtlasError avec une catégorie (`kind`) et un message humain en
 *     français ; le détail technique reste dans `detail` et dans le journal de diagnostic ;
 *   - chaque appel est borné dans le temps (AbortController) ;
 *   - le journal (`getDiag`) garde l'origine de la page, l'URL du worker, l'état du jeton et les
 *     derniers appels, affichés dans Réglages › Diagnostic pour diagnostiquer avec Charles.
 * Aucune donnée personnelle n'est journalisée : chemin d'URL sans paramètres, code HTTP, durée.
 */

export type ErrorKind =
  | 'hors-ligne'   // le poste n'a pas de réseau
  | 'reseau'       // aucune réponse HTTP exploitable (« Load failed », « Failed to fetch »)
  | 'delai'        // délai dépassé
  | 'session'      // jeton Microsoft impossible à obtenir, ou refusé (401)
  | 'acces'        // 403
  | 'debit'        // 429
  | 'introuvable'  // 404
  | 'requete'      // autre 4xx
  | 'serveur'      // 5xx
  | 'outlook'      // Office.js a refusé l'opération
  | 'inconnu';

export type Service = 'atlas' | 'outlook';

const HUMAN: Record<ErrorKind, Record<Service, string>> = {
  'hors-ligne': {
    atlas: 'Pas de connexion Internet : vérifie ton réseau puis réessaie.',
    outlook: 'Pas de connexion Internet : vérifie ton réseau puis réessaie.',
  },
  reseau: {
    atlas: 'Impossible de joindre ATLAS : vérifie ta connexion puis réessaie.',
    outlook: 'Impossible de joindre ta boîte Outlook pour le moment : réessaie dans un instant.',
  },
  delai: {
    atlas: 'ATLAS met trop de temps à répondre : réessaie dans un instant.',
    outlook: 'Ta boîte Outlook met trop de temps à répondre : réessaie dans un instant.',
  },
  session: {
    atlas: 'Ta session Outlook doit être rouverte : ferme puis rouvre le panneau ATLAS. Si ça persiste, quitte et relance Outlook.',
    outlook: 'Ta session Outlook doit être rouverte : quitte et relance Outlook, puis réessaie.',
  },
  acces: {
    atlas: 'Ton compte n\'a pas accès à cette fonction d\'ATLAS.',
    outlook: 'Outlook n\'autorise pas encore ATLAS à faire cela dans ta boîte.',
  },
  debit: {
    atlas: 'Beaucoup de demandes d\'un coup : patiente une minute puis réessaie.',
    outlook: 'Outlook limite les demandes pour le moment : patiente une minute puis réessaie.',
  },
  introuvable: {
    atlas: 'Cette fonction n\'est pas encore disponible sur le serveur ATLAS.',
    outlook: 'Élément introuvable dans ta boîte (déplacé ou supprimé ?).',
  },
  requete: {
    atlas: 'ATLAS n\'a pas pu traiter cette demande.',
    outlook: 'Outlook n\'a pas pu traiter cette demande.',
  },
  serveur: {
    atlas: 'ATLAS a rencontré un problème : réessaie dans un instant. Si ça persiste, préviens Charles.',
    outlook: 'Outlook a rencontré un problème : réessaie dans un instant.',
  },
  outlook: {
    atlas: 'Outlook n\'a pas permis cette action ici.',
    outlook: 'Outlook n\'a pas permis cette action ici.',
  },
  inconnu: {
    atlas: 'Un problème est survenu : réessaie dans un instant.',
    outlook: 'Un problème est survenu : réessaie dans un instant.',
  },
};

/** Message humain d'une catégorie d'erreur. */
export function humanMessage(kind: ErrorKind, service: Service = 'atlas'): string {
  return (HUMAN[kind] || HUMAN.inconnu)[service];
}

/** Erreur du complément : `message` est TOUJOURS une phrase lisible, `detail` le texte technique. */
export class AtlasError extends Error {
  readonly kind: ErrorKind;
  /** Code HTTP (0 si aucune réponse). */
  readonly status: number;
  readonly service: Service;
  readonly route: string;
  readonly detail: string;
  /** Corps JSON de la réponse (si lisible). */
  readonly data: any;
  /** L'action peut être retentée telle quelle (réseau, délai, serveur, débit). */
  readonly retryable: boolean;

  constructor(kind: ErrorKind, message: string, opts: { status?: number; service?: Service; route?: string; detail?: string; data?: any } = {}) {
    super(message);
    this.name = 'AtlasError';
    this.kind = kind;
    this.status = opts.status ?? 0;
    this.service = opts.service ?? 'atlas';
    this.route = opts.route ?? '';
    this.detail = opts.detail ?? '';
    this.data = opts.data;
    this.retryable = ['hors-ligne', 'reseau', 'delai', 'serveur', 'debit', 'session'].includes(kind);
  }
}

/** Catégorie d'un code HTTP en échec. */
export function kindForStatus(status: number): ErrorKind {
  if (status === 401) return 'session';
  if (status === 403) return 'acces';
  if (status === 404) return 'introuvable';
  if (status === 408 || status === 504) return 'delai';
  if (status === 429) return 'debit';
  if (status >= 500) return 'serveur';
  if (status >= 400) return 'requete';
  return 'inconnu';
}

/**
 * Le serveur renvoie parfois une phrase destinée à l'utilisateur (« Le point a déjà été arbitré »),
 * parfois un code (« invalid_projetId »). On ne garde que la première.
 */
export function isHumanText(s: unknown): s is string {
  if (typeof s !== 'string') return false;
  const t = s.trim();
  return t.length >= 8 && t.length <= 240 && /\s/.test(t) && !/^[a-z0-9_.:-]+$/i.test(t) && !/^(atlas \d{3}|https?:|error|typeerror)/i.test(t);
}

// ── Journal de diagnostic ──

export interface DiagEvent {
  at: number;
  service: Service;
  /** Chemin appelé, sans paramètres (ex. « atlas/projets/list », « outlook /me/mailFolders »). */
  label: string;
  status: number;
  ms: number;
  kind?: ErrorKind;
  detail?: string;
}

export interface TokenDiag {
  source: 'naa' | 'sso' | '';
  /** Fin de validité du jeton en cache (ms epoch), 0 si aucun. */
  until: number;
  /** Dernière erreur d'obtention (texte technique). */
  lastError: string;
  lastErrorAt: number;
  /** Connexion automatique (nested app authentication) annoncée par Outlook. */
  naaSupported: boolean | null;
}

const recent: DiagEvent[] = [];
let lastError: DiagEvent | null = null;
let workerUrl = '';
const token: TokenDiag = { source: '', until: 0, lastError: '', lastErrorAt: 0, naaSupported: null };

export function setDiagWorkerUrl(url: string): void { workerUrl = url; }
export function setTokenDiag(patch: Partial<TokenDiag>): void { Object.assign(token, patch); }

export function noteDiag(ev: DiagEvent): void {
  recent.unshift(ev);
  if (recent.length > 20) recent.length = 20;
  if (ev.kind) lastError = ev;
}

export function getDiag() {
  let origin = '';
  try { origin = window.location.origin; } catch { /* hors navigateur */ }
  let host = '';
  let version = '';
  let platform = '';
  try {
    const d: any = (Office as any)?.context?.diagnostics || (Office as any)?.context?.mailbox?.diagnostics;
    host = String(d?.host || d?.hostName || '');
    version = String(d?.version || d?.hostVersion || '');
    platform = String(d?.platform || '');
  } catch { /* hors Outlook */ }
  return {
    origin,
    workerUrl,
    online: typeof navigator !== 'undefined' ? navigator.onLine : true,
    host, version, platform,
    userAgent: typeof navigator !== 'undefined' ? navigator.userAgent : '',
    token: { ...token },
    lastError: lastError ? { ...lastError } : null,
    recent: recent.slice(),
  };
}

/** Chemin lisible d'une URL, sans paramètres ni identifiants longs. */
function labelFor(url: string, service: Service): string {
  try {
    const u = new URL(url);
    const path = u.pathname
      .replace(/^\/api\/plugin\//, '')
      .replace(/\/(messages|mailFolders)\/[^/]{20,}/g, '/$1/…');
    return service === 'atlas' ? path : `${u.host}${path}`;
  } catch { return url.split('?')[0].slice(0, 80); }
}

/**
 * `fetch` borné dans le temps, journalisé, dont l'échec réseau devient une AtlasError lisible.
 * Ne lève PAS sur une réponse HTTP en erreur : l'appelant décide (voir worker.ts).
 */
export async function netFetch(url: string, init: RequestInit = {}, opts: { service?: Service; timeoutMs?: number; label?: string } = {}): Promise<Response> {
  const service = opts.service ?? 'atlas';
  const label = opts.label ?? labelFor(url, service);
  const timeoutMs = opts.timeoutMs ?? 45_000;
  const started = Date.now();
  if (typeof navigator !== 'undefined' && navigator.onLine === false) {
    const err = new AtlasError('hors-ligne', humanMessage('hors-ligne', service), { service, route: label, detail: 'navigator.onLine = false' });
    noteDiag({ at: started, service, label, status: 0, ms: 0, kind: err.kind, detail: err.detail });
    throw err;
  }
  const ctrl = typeof AbortController !== 'undefined' ? new AbortController() : null;
  let timedOut = false;
  const timer = ctrl ? setTimeout(() => { timedOut = true; ctrl.abort(); }, timeoutMs) : null;
  try {
    const res = await fetch(url, ctrl ? { ...init, signal: ctrl.signal } : init);
    noteDiag({ at: started, service, label, status: res.status, ms: Date.now() - started, ...(res.ok ? {} : { kind: kindForStatus(res.status) }) });
    return res;
  } catch (e) {
    const kind: ErrorKind = timedOut ? 'delai' : (typeof navigator !== 'undefined' && navigator.onLine === false) ? 'hors-ligne' : 'reseau';
    const detail = `${(e as Error)?.name || 'Error'}: ${(e as Error)?.message || String(e)}${timedOut ? ` (délai ${Math.round(timeoutMs / 1000)} s)` : ''}`;
    noteDiag({ at: started, service, label, status: 0, ms: Date.now() - started, kind, detail });
    throw new AtlasError(kind, humanMessage(kind, service), { service, route: label, detail });
  } finally {
    if (timer) clearTimeout(timer);
  }
}

/** Appel direct à la boîte (Graph ou API REST d'Outlook) : mêmes garanties que netFetch. */
export function outlookFetch(url: string, init: RequestInit = {}): Promise<Response> {
  return netFetch(url, init, { service: 'outlook', timeoutMs: 30_000 });
}

/** Erreur lisible pour une réponse en échec de la boîte (le corps technique va dans `detail`). */
export function outlookHttpError(status: number, what: string, body = ''): AtlasError {
  const kind = kindForStatus(status);
  return new AtlasError(kind, humanMessage(kind, 'outlook'), { status, service: 'outlook', route: what, detail: `${what} HTTP ${status}${body ? ` · ${body.slice(0, 160)}` : ''}` });
}

/**
 * Message lisible pour N'IMPORTE quelle erreur (AtlasError, erreur Office.js, TypeError brut…).
 * À utiliser partout où une erreur est montrée à l'utilisateur : jamais de texte technique brut.
 */
export function humanError(e: unknown, fallback = 'Un problème est survenu : réessaie dans un instant.'): string {
  if (e instanceof AtlasError) return e.message;
  const msg = String((e as Error)?.message ?? e ?? '');
  if (/load failed|failed to fetch|networkerror|network request failed|réseau/i.test(msg)) return humanMessage('reseau');
  if (/abort|timeout|délai/i.test(msg)) return humanMessage('delai');
  if (/\b401\b|token|jeton|getAccessToken|sso|13001|13006/i.test(msg)) return humanMessage('session');
  // Messages déjà rédigés pour l'utilisateur dans le code du complément (« Choisis une date. »).
  if (isHumanText(msg) && !/\b\d{3}\b.*:|[{}<>]|exception|undefined|null/i.test(msg)) return msg;
  return fallback;
}
