/**
 * envoi-direct.ts · « Envoyer » depuis le tableau de bord (08/10/2026, Charles : le nouvel Outlook pour
 * Mac n'ouvre pas un brouillon déposé, la réponse restait dans les Brouillons).
 *
 * Deux clics, sans autre éditeur : le worker prépare la réponse dans la boîte PERSONNELLE de la personne
 * (createReply / createReplyAll : destinataires et citation posés par Outlook, notre texte au-dessus) et
 * renvoie les destinataires ; la personne confirme « Envoyer à X (et Y en copie) ? » ; le mail part
 * par remise différée d'Exchange de quelques secondes, un toast « Envoyé » avec « Annuler » le retire
 * tant qu'il n'est pas parti. Exclaimer ajoute la signature côté serveur. Une fois parti, le mail
 * reçu passe dans « Traités à ranger ».
 */
import { abandonnerEnvoi, confirmerEnvoi, preparerEnvoi, type EnvoiPret } from '../api/tableau';
import { humanError } from '../api/net';
import { h, ouvrirCouche, toast } from './ui';

/** Durée pendant laquelle le toast « Envoyé · Annuler » reste affiché (ms). */
export const TOAST_ANNULER_MS = 10_000;

const nomDe = (x: { email: string; nom: string }) => x.nom || x.email;

/** « Envoyer à Marie Muller (et Paul Weber en copie) ? » */
export function questionEnvoi(p: Pick<EnvoiPret, 'a' | 'cc'>): string {
  const a = p.a.map(nomDe).join(', ');
  const cc = p.cc.map(nomDe).join(', ');
  if (!a) return `Envoyer à ${cc} (en copie) ?`;
  return `Envoyer à ${a}${cc ? ` (et ${cc} en copie)` : ''} ?`;
}

export interface EnvoiDirectOptions {
  messageId: string;
  mailbox: string;
  texte: string;
  tous?: boolean;
  /** Confirmation (dans la page ou en modale) : true = envoyer, false = abandonner (le brouillon est retiré). */
  confirmer: (p: EnvoiPret) => Promise<boolean>;
  /** Après l'envoi confirmé (le mail est dans la Boîte d'envoi, annulable quelques secondes). */
  apresEnvoi?: (p: EnvoiPret) => void;
  /** Après « Annuler » pendant la fenêtre d'annulation (le texte est rendu à la personne). */
  apresAnnulation?: () => void;
}

export type IssueEnvoi = 'envoye' | 'abandonne' | 'erreur';

/** Prépare, fait confirmer, envoie. Toute erreur est affichée (toast) ; rien ne part sans le clic de confirmation. */
export async function envoyerDirect(o: EnvoiDirectOptions): Promise<IssueEnvoi> {
  let p: EnvoiPret;
  try { p = await preparerEnvoi(o.messageId, o.mailbox, o.texte, o.tous === true); }
  catch (e) { toast(humanError(e), 'error'); return 'erreur'; }
  if (!p?.pret || !p.brouillonId) { toast('Envoi indisponible : réponds depuis Outlook', 'error'); return 'erreur'; }
  let ok = false;
  try { ok = await o.confirmer(p); } catch { ok = false; }
  if (!ok) {
    try { await abandonnerEnvoi(p.brouillonId); } catch { /* brouillon orphelin : sans gravité */ }
    return 'abandonne';
  }
  let r: { envoye: boolean; partLe: string; annulableS: number };
  try { r = await confirmerEnvoi(p.brouillonId, o.messageId); }
  catch (e) {
    toast(humanError(e), 'error');
    try { await abandonnerEnvoi(p.brouillonId); } catch { /* idem */ }
    return 'erreur';
  }
  if (!r?.envoye) { toast('Envoi refusé par le worker', 'error'); return 'erreur'; }
  const fenetre = Math.min(TOAST_ANNULER_MS, Math.max(3000, (r.annulableS || 0) * 1000 - 2000));
  toast(`Envoyé à ${p.a.map(nomDe).join(', ') || p.cc.map(nomDe).join(', ')}`, 'success', () => {
    void abandonnerEnvoi(p.brouillonId)
      .then(() => { toast('Envoi annulé : le mail n\'est pas parti', 'info'); o.apresAnnulation?.(); })
      .catch(e => toast(/déjà parti/i.test(humanError(e)) ? 'Trop tard : le mail est déjà parti' : humanError(e), 'error'));
  }, fenetre);
  o.apresEnvoi?.(p);
  return 'envoye';
}

/** Confirmation en modale (séance de tri, proposition de RDV) : « Envoyer à X ? » Confirmer / Annuler. */
export function confirmerParModale(p: EnvoiPret): Promise<boolean> {
  return new Promise(resolve => {
    let v = false;
    const el = document.createElement('div');
    el.className = 'tb-modal';
    el.innerHTML = `<h3>${h(questionEnvoi(p))}</h3>${p.sujet ? `<p class="tb-note">${h(p.sujet)}</p>` : ''}
      <div class="tb-actions"><button type="button" class="tb-btn is-primary" data-ok>Envoyer</button><button type="button" class="tb-btn is-ghost" data-non>Annuler</button></div>
      <p class="tb-note">Part dans quelques secondes : « Annuler » le retient tant qu'il n'est pas parti. Signature ajoutée à l'envoi.</p>`;
    const fermer = ouvrirCouche(el, () => resolve(v));
    el.querySelector('[data-ok]')?.addEventListener('click', () => { v = true; fermer(); });
    el.querySelector('[data-non]')?.addEventListener('click', () => fermer());
    (el.querySelector('[data-ok]') as HTMLButtonElement | null)?.focus();
  });
}

/** Confirmation dans la page (zone de réponse du détail) : même question, boutons sous le texte. */
export function confirmerDansLaPage(host: HTMLElement, p: EnvoiPret): Promise<boolean> {
  return new Promise(resolve => {
    host.innerHTML = `<div class="tb-sep"></div><p><b>${h(questionEnvoi(p))}</b></p>
      <div class="tb-actions"><button type="button" class="tb-btn is-primary" data-ok>Envoyer</button><button type="button" class="tb-btn is-ghost" data-non>Annuler</button></div>
      <p class="tb-note">Part dans quelques secondes : « Annuler » le retient tant qu'il n'est pas parti. Signature ajoutée à l'envoi.</p>`;
    const fin = (v: boolean) => { host.innerHTML = ''; resolve(v); };
    host.querySelector('[data-ok]')?.addEventListener('click', () => fin(true));
    host.querySelector('[data-non]')?.addEventListener('click', () => fin(false));
    (host.querySelector('[data-ok]') as HTMLButtonElement | null)?.focus();
  });
}
