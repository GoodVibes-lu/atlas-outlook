/**
 * dialogue-tableau.ts : liaison du tableau de bord ATLAS avec Outlook.
 *
 * Depuis le 10/10/2026, le tableau de bord s'affiche DANS LE VOLET du complément (src/tableau-volet.ts),
 * plus dans une fenêtre de dialogue Office : sur le nouvel Outlook pour Mac, cette fenêtre restait
 * au-dessus de toutes les applications. Ce fichier garde :
 *   - `ouvrirDepuisTableau` : ouvre un mail demandé par le tableau de bord dans l'application Outlook
 *     (`displayMessageForm`, Mailbox 1.1, pris en charge sur Mac), sinon l'adresse telle quelle ;
 *   - `repondreSurElementCourant` : formulaire de réponse d'Outlook sur le mail affiché ;
 *   - `brancherSurParent` : côté page tableau-de-bord.html ouverte en dialogue (ancien hôte, gardé
 *     pour une page encore ouverte par un vieux manifeste).
 */

import { openExternal } from './platform';

const MARQUE = 'tableau';

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

/**
 * Identifiant EWS d'un mail à partir de son lien Outlook sur le web (`webLink` de Graph :
 * `…/owa/?ItemID=…`, ou `…/mail/deeplink/read/<id>`), pour `displayMessageForm`.
 */
export function idDuLienOutlook(url: string): string {
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
export function repondreSurElementCourant(messageId: string, html: string): boolean {
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

/** Lien hébergé par Outlook (mail, brouillon, calendrier, compose…) : Outlook l'ouvrira devant, la fenêtre doit se fermer. */
export function ouvrirDepuisTableau(url: string): void {
  const id = idDuLienOutlook(url);
  const mailbox = (() => { try { return Office.context?.mailbox; } catch { return undefined; } })();
  if (id && mailbox && typeof mailbox.displayMessageForm === 'function') {
    try { mailbox.displayMessageForm(id); return; } catch (e) {
      console.warn('[dialogue-tableau] displayMessageForm impossible, repli sur le lien :', e);
    }
  }
  openExternal(url);
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
