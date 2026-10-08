/**
 * dialogue-tableau.ts : tableau de bord ATLAS dans une GRANDE FENÊTRE de dialogue Office (07/10/2026).
 *
 * Pourquoi : le nouvel Outlook pour Mac n'affiche pas les applications Microsoft 365 / Teams dans la
 * barre de gauche (onglet personnel de teams-app/manifest.json). Pour avoir le tableau de bord DANS
 * Outlook sur Mac, le complément l'ouvre avec `Office.context.ui.displayDialogAsync` (fenêtre de
 * 95 % × 90 % de l'écran), depuis le menu ATLAS du ruban (commande `atlasTableauCommand`,
 * src/commands.ts) ou depuis le bouton en tête du panneau (src/taskpane.ts). Windows et le web
 * gardent aussi l'onglet de la barre de gauche (même page, hôte TeamsJS).
 *
 * Deux côtés, un seul protocole (messages JSON marqués `atlas: 'tableau'`) :
 *   - PAGE PARENTE (panneau ou fichier de commandes, Office.js complet) : `ouvrirTableauDialogue`.
 *     Elle fournit le jeton Microsoft au dialogue (sa propre connexion, `getWorkerToken`) et ouvre
 *     les mails demandés par le tableau de bord (`displayMessageForm`, le dialogue ne pouvant pas
 *     piloter la fenêtre principale d'Outlook).
 *   - DIALOGUE (tableau-de-bord.html?hote=office) : `brancherSurParent`.
 *     Demande le jeton par `messageParent`, le reçoit par `messageChild` (DialogApi 1.2) ; sans
 *     réponse, la page passe à sa propre connexion Microsoft (redirection MSAL, worker.ts).
 * Aucun secret dans la page : le jeton est celui de la personne connectée, de courte durée.
 */

import { getWorkerToken } from './worker';
import { openExternal, supportsSet } from './platform';
import { humanError } from './net';

const MARQUE = 'tableau';
const PAGE = 'tableau-de-bord.html';

type Message =
  | { atlas: typeof MARQUE; type: 'jeton'; id: number; force?: boolean }
  | { atlas: typeof MARQUE; type: 'jeton-reponse'; id: number; token?: string; erreur?: string }
  | { atlas: typeof MARQUE; type: 'ouvrir'; url: string }
  /** Séance de tri (07/10/2026) : formulaire de réponse d'Outlook sur le mail affiché par la fenêtre parente. */
  | { atlas: typeof MARQUE; type: 'repondre'; id: number; messageId: string; html: string }
  | { atlas: typeof MARQUE; type: 'repondre-reponse'; id: number; ok: boolean }
  | { atlas: typeof MARQUE; type: 'fermer' };

function lire(raw: unknown): Message | null {
  try {
    const m = typeof raw === 'string' ? JSON.parse(raw) : raw;
    return m && m.atlas === MARQUE && typeof m.type === 'string' ? (m as Message) : null;
  } catch { return null; }
}

// ── Côté PAGE PARENTE ──

let dialogue: Office.Dialog | null = null;

/** Thème d'Outlook lu par Office.js (fond sombre ⇒ sombre), transmis au dialogue par l'adresse. */
function themeOutlook(): 'dark' | 'light' | '' {
  try {
    const bg = (Office.context as any)?.officeTheme?.bodyBackgroundColor as string | undefined;
    const m = bg && /^#?([0-9a-f]{6})$/i.exec(bg.trim());
    if (!m) return '';
    const n = parseInt(m[1], 16);
    const lum = (0.2126 * ((n >> 16) & 255) + 0.7152 * ((n >> 8) & 255) + 0.0722 * (n & 255)) / 255;
    return lum < 0.5 ? 'dark' : 'light';
  } catch { return ''; }
}

/** Adresse du tableau de bord en mode dialogue (même dossier que la page parente, même domaine). */
export function urlTableauDialogue(): string {
  const u = new URL(PAGE, window.location.href);
  u.search = '';
  u.hash = '';
  u.searchParams.set('hote', 'office');
  // Le parent sait répondre (messageChild) : sinon le dialogue se connecte seul tout de suite.
  if (supportsSet('DialogApi', '1.2')) u.searchParams.set('parent', '1');
  const t = themeOutlook();
  if (t) u.searchParams.set('theme', t);
  return u.toString();
}

/**
 * Identifiant EWS d'un mail à partir de son lien Outlook sur le web (`webLink` de Graph :
 * `…/owa/?ItemID=…`, ou `…/mail/deeplink/read/<id>`), pour `displayMessageForm`.
 */
function idDuLienOutlook(url: string): string {
  let u: URL;
  try { u = new URL(url); } catch { return ''; }
  if (!/(^|\.)outlook\.(office|office365|live)\.com$/i.test(u.hostname)) return '';
  let id = u.searchParams.get('ItemID') || u.searchParams.get('itemid') || '';
  if (!id) {
    const m = /\/read\/([^/?#]+)/.exec(u.pathname);
    if (m) { try { id = decodeURIComponent(m[1]); } catch { id = m[1]; } }
  }
  if (!id) return '';
  // Forme REST (base64 « url », - et _) : conversion en EWS (Mailbox 1.3).
  if (/[-_]/.test(id) && !/[+/]/.test(id)) {
    try { id = Office.context.mailbox.convertToEwsId(id, Office.MailboxEnums.RestVersion.v2_0); } catch { /* tel quel */ }
  }
  return id;
}

/**
 * Séance de tri : ouvre le formulaire de réponse d'Outlook (Mailbox 1.1 `displayReplyForm`, nouvel
 * Outlook pour Mac compris) avec le texte préparé, SEULEMENT si la fenêtre parente montre ce mail-là
 * (même internetMessageId). Office.js ne sait pas répondre à un autre mail que l'élément courant :
 * sinon false, et le tableau copie le texte puis ouvre le mail. Rien n'est envoyé.
 */
function repondreSurElementCourant(messageId: string, html: string): boolean {
  try {
    const item: any = Office.context?.mailbox?.item;
    if (!item || typeof item.displayReplyForm !== 'function') return false;
    if (String(item.internetMessageId || '').trim() !== String(messageId || '').trim()) return false;
    item.displayReplyForm(html);
    return true;
  } catch (e) {
    console.warn('[dialogue-tableau] displayReplyForm impossible :', e);
    return false;
  }
}

/** Ouvre, depuis la fenêtre principale, l'adresse demandée par le tableau de bord. */
function ouvrirDepuisTableau(url: string): void {
  const id = idDuLienOutlook(url);
  const mailbox = (() => { try { return Office.context?.mailbox; } catch { return undefined; } })();
  if (id && mailbox && typeof mailbox.displayMessageForm === 'function') {
    try { mailbox.displayMessageForm(id); return; } catch (e) {
      console.warn('[dialogue-tableau] displayMessageForm impossible, repli sur le lien :', e);
    }
  }
  openExternal(url);
}

export interface OuvrirTableauOptions {
  /** Ouvre directement la séance de tri (bouton « Séance de tri » du panneau). */
  seance?: boolean;
  /** Appelé quand la fenêtre se ferme (ou n'a pas pu s'ouvrir). */
  onFerme?: () => void;
  /** Message lisible si la fenêtre ne s'ouvre pas. */
  onErreur?: (message: string) => void;
}

/** Message lisible pour les codes d'erreur de displayDialogAsync. */
function messageOuverture(code: number | undefined): string {
  switch (code) {
    case 12007: return 'Le tableau de bord est déjà ouvert dans une autre fenêtre.';
    case 12009: return 'Ouverture du tableau de bord refusée : autorise la fenêtre ATLAS puis réessaie.';
    case 12011: return 'Le navigateur bloque la fenêtre du tableau de bord : autorise les fenêtres pour Outlook.';
    case 12004: case 12005: return 'Adresse du tableau de bord refusée par Outlook (complément à republier).';
    default: return 'Impossible d\'ouvrir le tableau de bord dans cette version d\'Outlook.';
  }
}

/**
 * Ouvre le tableau de bord dans une grande fenêtre de dialogue Office. Une seule fenêtre à la fois
 * (limite d'Office) : un second appel signale qu'elle est déjà ouverte.
 */
export function ouvrirTableauDialogue(opts: OuvrirTableauOptions = {}): void {
  let fini = false;
  const fin = () => {
    if (fini) return;
    fini = true;
    dialogue = null;
    opts.onFerme?.();
  };
  if (typeof Office === 'undefined' || !Office.context?.ui?.displayDialogAsync) {
    opts.onErreur?.(messageOuverture(undefined));
    fin();
    return;
  }
  if (dialogue) {
    opts.onErreur?.(messageOuverture(12007));
    // Pas de fin() : la fenêtre ouverte garde son propre suivi.
    return;
  }
  const url = new URL(urlTableauDialogue());
  if (opts.seance) url.searchParams.set('seance', '1');
  Office.context.ui.displayDialogAsync(url.toString(), {
    width: 95,
    height: 90,
    // Web : vraie fenêtre (pas un cadre), sans demande de confirmation. Ignoré sur Mac / Windows.
    displayInIframe: false,
    promptBeforeOpen: false,
  } as Office.DialogOptions, (res) => {
    if (res.status !== Office.AsyncResultStatus.Succeeded) {
      console.warn('[dialogue-tableau] ouverture impossible :', res.error);
      opts.onErreur?.(messageOuverture(res.error?.code));
      fin();
      return;
    }
    const d = res.value;
    dialogue = d;
    const repondre = (m: Message) => {
      try { (d as any).messageChild?.(JSON.stringify(m)); } catch (e) {
        console.warn('[dialogue-tableau] messageChild impossible :', e);
      }
    };
    d.addEventHandler(Office.EventType.DialogMessageReceived, (arg: any) => {
      if (!arg || 'error' in arg) return;
      // DialogApi 1.2 donne l'origine : seule la page du complément (même domaine) est écoutée.
      if (arg.origin && arg.origin !== window.location.origin) return;
      const m = lire(arg.message);
      if (!m) return;
      if (m.type === 'jeton') {
        getWorkerToken(!!m.force)
          .then(token => repondre({ atlas: MARQUE, type: 'jeton-reponse', id: m.id, token }))
          .catch(e => repondre({ atlas: MARQUE, type: 'jeton-reponse', id: m.id, erreur: humanError(e) }));
      } else if (m.type === 'repondre' && typeof m.messageId === 'string' && typeof m.html === 'string') {
        repondre({ atlas: MARQUE, type: 'repondre-reponse', id: m.id, ok: repondreSurElementCourant(m.messageId, m.html) });
      } else if (m.type === 'ouvrir' && typeof m.url === 'string') {
        // Un mail ouvert dans Outlook s'affichait DERRIÈRE la fenêtre du tableau de bord (Office la garde
        // au premier plan). Pour un mail, la fenêtre se ferme donc ; on la rouvre d'un clic (« Tableau de bord »).
        const mail = !!idDuLienOutlook(m.url);
        ouvrirDepuisTableau(m.url);
        if (mail) { try { d.close(); } catch { /* déjà fermée */ } fin(); }
      } else if (m.type === 'fermer') {
        try { d.close(); } catch { /* déjà fermée */ }
        fin();
      }
    });
    // Fenêtre fermée par la personne (12006) ou page du dialogue en erreur.
    d.addEventHandler(Office.EventType.DialogEventReceived, () => fin());
  });
}

// ── Côté DIALOGUE (tableau-de-bord.html?hote=office) ──

/** Délai d'attente du jeton envoyé par la page parente. */
const JETON_PARENT_DELAI = 15_000;

export interface LienParent {
  /** Jeton Microsoft obtenu par la page parente (lève une erreur si elle ne répond pas). */
  jeton: (forceRefresh: boolean) => Promise<string>;
  /** Ouvre une adresse (un mail : dans Outlook, par la page parente ; sinon le navigateur). */
  openLink: (url: string) => void;
  /** Séance de tri : réponse dans Outlook par la page parente (false : pas le mail affiché, ou pas de réponse). */
  repondre: (messageId: string, html: string) => Promise<boolean>;
}

/**
 * Branche le dialogue sur la page qui l'a ouvert. `avecJeton` : la page parente sait répondre
 * (paramètre `parent=1`, DialogApi 1.2) ; sinon seule l'ouverture des liens passe par elle.
 */
export function brancherSurParent(avecJeton: boolean): LienParent {
  const ui: any = (Office.context as any)?.ui;
  const attente = new Map<number, { ok: (t: string) => void; ko: (e: Error) => void }>();
  const reponses = new Map<number, (ok: boolean) => void>();
  let suivant = 1;
  let ecoute = false;

  const envoyer = (m: Message): boolean => {
    try { ui.messageParent(JSON.stringify(m)); return true; } catch (e) {
      console.warn('[dialogue-tableau] messageParent impossible :', e);
      return false;
    }
  };

  if (avecJeton && typeof ui?.addHandlerAsync === 'function') {
    try {
      ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg: any) => {
        const m = lire(arg?.message);
        if (m?.type === 'repondre-reponse') { const r = reponses.get(m.id); reponses.delete(m.id); r?.(m.ok === true); return; }
        if (!m || m.type !== 'jeton-reponse') return;
        const p = attente.get(m.id);
        if (!p) return;
        attente.delete(m.id);
        if (m.token) p.ok(m.token);
        else p.ko(new Error(m.erreur || 'jeton refusé par la page parente'));
      }, (r: any) => { ecoute = r?.status === Office.AsyncResultStatus.Succeeded; });
      ecoute = true; // l'accusé arrive après coup ; une erreur se verra au délai dépassé
    } catch (e) {
      console.warn('[dialogue-tableau] écoute du parent impossible :', e);
    }
  }

  const jeton = (forceRefresh: boolean): Promise<string> => {
    if (!avecJeton || !ecoute || typeof ui?.messageParent !== 'function') {
      return Promise.reject(new Error('page parente sans échange de messages (DialogApi 1.2)'));
    }
    const id = suivant++;
    return new Promise<string>((ok, ko) => {
      const t = setTimeout(() => { attente.delete(id); ko(new Error('la page parente ne répond pas')); }, JETON_PARENT_DELAI);
      attente.set(id, {
        ok: (v) => { clearTimeout(t); ok(v); },
        ko: (e) => { clearTimeout(t); ko(e); },
      });
      if (!envoyer({ atlas: MARQUE, type: 'jeton', id, force: forceRefresh })) {
        clearTimeout(t); attente.delete(id); ko(new Error('messageParent indisponible'));
      }
    });
  };

  const openLink = (url: string): void => {
    if (!/^(https:\/\/|atlas-app:\/\/)/i.test(String(url || '').trim())) return;
    if (typeof ui?.messageParent === 'function' && envoyer({ atlas: MARQUE, type: 'ouvrir', url })) return;
    window.open(url, '_blank', 'noopener');
  };

  const repondreParent = (messageId: string, html: string): Promise<boolean> => {
    if (!avecJeton || !ecoute || typeof ui?.messageParent !== 'function') return Promise.resolve(false);
    const id = suivant++;
    return new Promise<boolean>(ok => {
      const t = setTimeout(() => { reponses.delete(id); ok(false); }, 5000);
      reponses.set(id, v => { clearTimeout(t); ok(v); });
      if (!envoyer({ atlas: MARQUE, type: 'repondre', id, messageId, html })) { clearTimeout(t); reponses.delete(id); ok(false); }
    });
  };

  return { jeton, openLink, repondre: repondreParent };
}
