/**
 * agent-outils.ts — Blocs « à la demande » de l'agent d'inbox (phases 3 et 4), communs au panneau
 * « Agent » (mail lu) et au tableau de bord de la barre de gauche (tableau-de-bord.ts).
 *
 *  • « Que faire de ce mail ? » : jusqu'à 3 cartes d'actions métier (GET /api/plugin/agent/actions),
 *    exécution au clic (POST …/actions/executer), choix du projet quand il est requis (jamais
 *    deviné), bandeau « Fait » + Annuler (POST /api/plugin/agent/annuler) ; actions déjà faites en
 *    grisé, jamais reproposées.
 *  • « Répondre plus tard » (mail reçu) et « Relancer si pas de réponse » / « Pas de relance »
 *    (mail envoyé) : POST | DELETE /api/plugin/agent/plus-tard | relancer.
 *  • « Résumer le fil » (GET …/resume-fil), « Poser une question à ma boîte » (POST …/question),
 *    « Rattraper depuis… » (GET …/rattrapage).
 *  • Phase 6 : « Traduire » (FR / EN / DE / LB, POST …/traduire), « Modèles de réponse » (GET
 *    …/modeles, POST …/modeles/remplir → texte à copier, à insérer ou à mettre dans une réponse),
 *    fiche mémoire du contact (GET …/fiche-contact, la personne seule), encart des détections
 *    (rebond, absence avec date de retour, candidature), pièces jointes lourdes → lien court
 *    (POST …/piece-jointe-lourde, …/deposer-partager : dépôt sur le CLOUD, transit 30 jours, jamais
 *    le NAS ; …/lien-court : lien court d'un lien de partage collé).
 *
 * Droit à la déconnexion : tout est déclenché par un clic, rien n'est notifié ni rafraîchi seul.
 * Office.js : aucune API ici (les écritures passent par le worker) ; ouverture des mails et des
 * liens par agent-lists.ts / platform.ts, qui testent la disponibilité des ensembles d'API.
 * Erreurs : toujours affichées (message court dans le bloc), jamais avalées.
 */

import {
  fetchActions, executerAction, annulerAction, mettrePlusTard, retirerPlusTard, demanderRelance,
  annulerRelance, poserQuestion, fetchResumeFil, fetchRattrapage,
  traduire, fetchModeles, remplirModele, fetchFicheContact, verifierPiecesLourdes, creerLienCourt, deposerPiecesLourdes, fetchDossiersNas,
  type DepotPiecesLourdes,
  type LangueTraduction, type ModeleReponse, type ModeleRempli, type DetectionsMail,
} from '../api/agent';
import type { ActionProposee, ActionFaite, CandidatChoix, SourceMail } from '../api/inbox-actions.types';
import { escapeHtml, dateCourte, openAgentMail } from './agent-lists';
import { openExternal, ATLAS_BASE } from '../api/platform';
import { WORKER_BASE } from '../api/worker';
import { humanError } from '../api/net';
import { icon } from '../ui/icons';

export type InfoFn = (message: string, type?: 'success' | 'error' | 'info') => void;

const MAX_CARTES = 3;

function messageErreur(e: unknown): string {
  return humanError(e);
}

function erreurHtml(quoi: string, e: unknown): string {
  return `<p class="agent-error" role="alert">${escapeHtml(quoi)} : ${escapeHtml(messageErreur(e))}</p>`;
}

/**
 * Lien ATLAS ouvrable depuis Outlook : https tel quel ; `atlas://go/<écran>` → page relais du
 * worker (`/go?to=…`, comme les liens des mails internes) ; chemin relatif → ATLAS web.
 */
export function lienAtlasOuvrable(lien: string | undefined | null, repli?: { table?: string; id?: string }): string {
  const l = String(lien || '').trim();
  if (/^https:\/\//i.test(l)) return l;
  const m = /^atlas:\/\/go\/(.+)$/i.exec(l);
  if (m) return `${WORKER_BASE}/go?to=${encodeURIComponent(m[1])}`;
  if (l.startsWith('/') && !l.startsWith('//')) return `${ATLAS_BASE}${l}`;
  if (repli?.id && /projet/i.test(repli.table || '')) return `${ATLAS_BASE}/projet/${encodeURIComponent(repli.id)}`;
  return '';
}

/** Lien ATLAS d'un compte rendu de réunion (écran Réunion ponctuelle, fiche ouverte par id). */
export function lienReunion(id: string): string {
  return `${ATLAS_BASE}/reunion-adhoc?id=${encodeURIComponent(id)}`;
}

function nomSource(s: SourceMail): string {
  const f = s.from;
  return typeof f === 'string' ? f : (f?.name || f?.email || '');
}

/** Liste cliquable de mails (sources d'une réponse, rattrapage) : ouverture par webLink. */
function sourcesHtml(sources: Array<SourceMail & { note?: string }>, prefixe: string): string {
  return `<ul class="mail-list">${sources.map((s, i) => `
    <li class="mail-item">
      <button type="button" class="mail-open" data-${prefixe}="${i}" ${s.webLink ? '' : 'disabled'} title="${s.webLink ? 'Ouvrir le mail' : 'Lien indisponible'}">
        <span class="mail-line1">
          <span class="mail-from">${escapeHtml(nomSource(s) || 'Expéditeur inconnu')}</span>
          <span class="mail-date">${escapeHtml(dateCourte(s.date))}</span>
        </span>
        <span class="mail-subject">${escapeHtml(s.subject || '(sans objet)')}</span>
        ${s.note ? `<span class="mail-resume">${escapeHtml(s.note)}</span>` : ''}
      </button>
    </li>`).join('')}</ul>`;
}

function bindSources(host: HTMLElement, sources: SourceMail[], prefixe: string, openLink?: (url: string) => void, onInfo?: InfoFn): void {
  host.querySelectorAll<HTMLButtonElement>(`button[data-${prefixe}]`).forEach(btn => {
    btn.addEventListener('click', () => {
      const s = sources[Number(btn.getAttribute(`data-${prefixe}`))];
      if (s && !openAgentMail({ graphId: '', webLink: s.webLink }, openLink)) {
        onInfo?.('Ce mail ne peut pas être ouvert d\'ici : retrouve-le dans ta boîte.', 'info');
      }
    });
  });
}

// ── « Que faire de ce mail ? » ──

function candidatsDe(a: ActionProposee): CandidatChoix[] {
  const c = (a.donnees as any)?.candidats;
  if (!Array.isArray(c)) return [];
  return c
    .filter((x: any) => x && typeof x.id === 'string' && x.id)
    .map((x: any) => ({ id: x.id, libelle: String(x.libelle || x.id), score: x.score, signaux: x.signaux }));
}

/** Choix de projet requis : avertissement « Projet à choisir… » ou candidats sans projet retenu. */
function choixProjetRequis(a: ActionProposee): boolean {
  const projetId = typeof (a.donnees as any)?.projetId === 'string' ? (a.donnees as any).projetId : '';
  if (projetId) return false;
  const av = (a.avertissements || []).join(' ');
  return /projet à choisir|aucun projet identifié|projet existant possible/i.test(av) || candidatsDe(a).length > 1;
}

function selectProjetHtml(i: number, candidats: CandidatChoix[], courant = ''): string {
  return `
    <label class="agent-muted" for="act-choix-${i}">Projet</label>
    <select class="agent-select" id="act-choix-${i}" data-choix="${i}">
      <option value="">Choisir un projet…</option>
      ${candidats.map(c => `<option value="${escapeHtml(c.id)}" ${c.id === courant ? 'selected' : ''}>${escapeHtml(c.libelle)}${c.signaux?.length ? ` (${escapeHtml(c.signaux[0])})` : ''}</option>`).join('')}
    </select>`;
}

/** Facture hors projet sans règle apprise : le dossier du NAS est à choisir (cadrage du 06/10/2026). */
function choixDossierRequis(a: ActionProposee): boolean {
  return a.type === 'facture-hors-projet' && (a.donnees as any)?.choixDossier === true;
}

/**
 * Sélecteur de dossier du NAS (GET /api/plugin/agent/nas/dossiers) : navigation dans les
 * sous-dossiers existants de la racine des factures, « Ranger ici » fixe le choix. Jamais de création.
 */
function brancherSelecteurDossier(host: HTMLElement, i: number, choisis: Map<number, string>): void {
  const ouvrir = host.querySelector<HTMLButtonElement>(`[data-dossier-ouvrir="${i}"]`);
  const liste = host.querySelector<HTMLElement>(`[data-dossier-liste="${i}"]`);
  const choisi = host.querySelector<HTMLElement>(`[data-dossier-choisi="${i}"]`);
  if (!ouvrir || !liste || !choisi) return;
  const afficher = async (chemin?: string) => {
    liste.hidden = false;
    liste.innerHTML = '<div class="agent-loading"><div class="spinner"></div><span>Dossiers du NAS…</span></div>';
    try {
      const d = await fetchDossiersNas(chemin);
      const parent = d.chemin && d.chemin !== d.racine ? d.chemin.slice(0, d.chemin.lastIndexOf('/')) : '';
      liste.innerHTML = `
        <div class="agent-muted">${escapeHtml(d.chemin)}</div>
        <ul class="agent-faites">
          ${parent ? `<li><button type="button" class="agent-link" data-nas-aller="${escapeHtml(parent)}">Dossier parent</button></li>` : ''}
          ${d.dossiers.map(x => `<li><button type="button" class="agent-link" data-nas-aller="${escapeHtml(x.chemin)}">${escapeHtml(x.nom)}</button></li>`).join('')}
        </ul>
        <button type="button" class="btn btn-secondary agent-btn" data-nas-ici="${escapeHtml(d.chemin)}">Ranger ici</button>`;
      liste.querySelectorAll<HTMLButtonElement>('[data-nas-aller]').forEach(b => b.addEventListener('click', () => { void afficher(b.dataset.nasAller); }));
      liste.querySelector<HTMLButtonElement>('[data-nas-ici]')?.addEventListener('click', (e) => {
        const c = (e.currentTarget as HTMLButtonElement).dataset.nasIci || '';
        choisis.set(i, c);
        choisi.textContent = c;
        liste.hidden = true;
      });
    } catch (e) {
      liste.innerHTML = erreurHtml('Dossiers indisponibles', e);
    }
  };
  ouvrir.addEventListener('click', () => { void afficher(choisis.get(i) || undefined); });
}

function faitesHtml(faites: ActionFaite[]): string {
  if (!faites.length) return '';
  return `<ul class="agent-faites">${faites.map(f => {
    const lien = lienAtlasOuvrable(f.lien);
    const etat = f.annuleLe
      ? ' <span class="agent-muted">· annulé</span>'
      : f.annulable && f.actionId ? ` <button type="button" class="agent-link" data-annuler-faite="${escapeHtml(f.actionId)}">Annuler</button>` : '';
    return `<li>${icon('check', 12)} ${escapeHtml(f.resume || f.type)}${f.at ? ` <span class="agent-muted">· ${escapeHtml(dateCourte(f.at))}</span>` : ''}${lien ? ` <a href="#" class="agent-faite-lien" data-href="${escapeHtml(lien)}">Voir</a>` : ''}${etat}</li>`;
  }).join('')}</ul>`;
}

/**
 * Cartes « Que faire de ce mail ? » dans `host` (masqué s'il n'y a rien à proposer ni rien de fait).
 * `observation` : agent en mode à blanc (les exécutions reviennent `simule: true`, on le dit).
 */
export async function renderActionsMetier(host: HTMLElement, ctx: { messageId: string; mailbox: string; onInfo?: InfoFn }): Promise<void> {
  if (!ctx.messageId) { host.hidden = true; return; }
  host.hidden = false;
  host.innerHTML = `
    <div class="agent-section-title">Que faire de ce mail ?</div>
    <div class="agent-loading"><div class="spinner"></div><span>Recherche des actions…</span></div>`;
  let actions: ActionProposee[];
  let faites: ActionFaite[];
  try {
    ({ actions, faites } = await fetchActions(ctx.messageId, ctx.mailbox));
  } catch (e) {
    if (!host.isConnected) return;
    host.innerHTML = `<div class="agent-section-title">Que faire de ce mail ?</div>${erreurHtml('Actions indisponibles', e)}`;
    return;
  }
  if (!host.isConnected) return;
  const typesFaits = new Set(faites.map(f => String(f.type)));
  const cartes = actions.filter(a => a && a.type && !typesFaits.has(a.type)).slice(0, MAX_CARTES);
  if (!cartes.length && !faites.length) { host.hidden = true; host.innerHTML = ''; return; }

  host.innerHTML = `
    <div class="agent-section-title">Que faire de ce mail ?</div>
    ${cartes.map((a, i) => {
      const candidats = candidatsDe(a);
      const choix = choixProjetRequis(a) && candidats.length > 0;
      const courant = typeof (a.donnees as any)?.projetId === 'string' ? (a.donnees as any).projetId : '';
      return `
      <div class="agent-carte" data-carte="${i}">
        <div class="agent-carte-titre">${escapeHtml(a.libelle)}</div>
        ${a.apercu ? `<div class="agent-carte-apercu">${escapeHtml(a.apercu)}</div>` : ''}
        ${a.avertissements?.length ? `<ul class="agent-avertissements">${a.avertissements.map(w => `<li>${escapeHtml(w)}</li>`).join('')}</ul>` : ''}
        <div class="agent-carte-choix" data-choix-wrap="${i}" ${choix ? '' : 'hidden'}>${choix ? selectProjetHtml(i, candidats, courant) : ''}</div>
        ${choixDossierRequis(a) ? `<div class="agent-carte-choix" data-dossier-wrap="${i}">
          <div class="agent-muted">Dossier du NAS : <span data-dossier-choisi="${i}">à choisir</span></div>
          <button type="button" class="btn btn-secondary agent-btn" data-dossier-ouvrir="${i}">Choisir le dossier…</button>
          <div data-dossier-liste="${i}" hidden></div>
        </div>` : ''}
        <button type="button" class="btn btn-primary btn-block agent-btn" data-exec="${i}">${escapeHtml(a.libelle)}</button>
        <div class="agent-carte-resultat" data-res="${i}" hidden></div>
      </div>`;
    }).join('')}
    ${faites.length ? `<div class="agent-muted">Déjà fait sur ce mail :</div>${faitesHtml(faites)}` : ''}
  `;

  host.querySelectorAll<HTMLAnchorElement>('a[data-href]').forEach(a => {
    a.addEventListener('click', (e) => { e.preventDefault(); openExternal(a.dataset.href!); });
  });

  host.querySelectorAll<HTMLButtonElement>('button[data-annuler-faite]').forEach(btn => {
    btn.addEventListener('click', async () => {
      btn.disabled = true;
      try {
        await annulerAction(btn.dataset.annulerFaite!);
        btn.replaceWith(Object.assign(document.createElement('span'), { className: 'agent-muted', textContent: '· annulé' }));
        ctx.onInfo?.('Action annulée', 'success');
      } catch (e) {
        btn.disabled = false;
        ctx.onInfo?.(`Annulation impossible : ${messageErreur(e)}`, 'error');
      }
    });
  });

  const dossiersChoisis = new Map<number, string>();
  cartes.forEach((a, i) => { if (choixDossierRequis(a)) brancherSelecteurDossier(host, i, dossiersChoisis); });

  host.querySelectorAll<HTMLButtonElement>('button[data-exec]').forEach(btn => {
    btn.addEventListener('click', async () => {
      const i = Number(btn.dataset.exec);
      const a = cartes[i];
      const carte = host.querySelector<HTMLElement>(`[data-carte="${i}"]`);
      const res = host.querySelector<HTMLElement>(`[data-res="${i}"]`);
      const choixWrap = host.querySelector<HTMLElement>(`[data-choix-wrap="${i}"]`);
      if (!a || !carte || !res || !choixWrap) return;
      const select = choixWrap.querySelector<HTMLSelectElement>('select');
      const projetId = select?.value || '';
      if (!choixWrap.hidden && select && !projetId) {
        res.hidden = false;
        res.innerHTML = '<p class="agent-error" role="alert">Choisis d\'abord le projet.</p>';
        select.focus();
        return;
      }
      const dossierNas = dossiersChoisis.get(i) || '';
      if (choixDossierRequis(a) && !dossierNas) {
        res.hidden = false;
        res.innerHTML = '<p class="agent-error" role="alert">Choisis d\'abord le dossier du NAS.</p>';
        return;
      }
      btn.disabled = true;
      const libelle = btn.textContent || '';
      btn.textContent = 'En cours…';
      res.hidden = true;
      try {
        const r = await executerAction({
          messageId: ctx.messageId,
          mailbox: ctx.mailbox || undefined,
          type: a.type,
          donnees: projetId ? { ...a.donnees, projetId } : a.donnees,
          ...(projetId || dossierNas ? { choix: { ...(projetId ? { projetId } : {}), ...(dossierNas ? { dossierNas } : {}) } } : {}),
        });
        if (!host.isConnected) return;
        if (r.ok) {
          const lien = lienAtlasOuvrable(r.cible?.lien, r.cible);
          carte.classList.add('is-fait');
          carte.innerHTML = `
            <div class="agent-fait" role="status">
              <span>${r.simule ? 'Simulé (agent en observation, rien n\'a été écrit)' : 'Fait'} : ${escapeHtml(r.resume || a.libelle)}</span>
              <span class="agent-fait-btns">
                ${lien ? '<button type="button" class="btn btn-secondary agent-btn" data-voir>Voir dans ATLAS</button>' : ''}
                ${r.actionId && !r.simule ? '<button type="button" class="btn btn-secondary agent-btn" data-annuler>Annuler</button>' : ''}
              </span>
            </div>`;
          carte.querySelector('[data-voir]')?.addEventListener('click', () => openExternal(lien));
          const undo = carte.querySelector<HTMLButtonElement>('[data-annuler]');
          undo?.addEventListener('click', async () => {
            undo.disabled = true;
            undo.textContent = 'Annulation…';
            try {
              await annulerAction(r.actionId!);
              undo.replaceWith(Object.assign(document.createElement('span'), { className: 'agent-muted', textContent: 'Annulé' }));
              ctx.onInfo?.('Action annulée', 'success');
            } catch (e) {
              undo.disabled = false;
              undo.textContent = 'Annuler';
              ctx.onInfo?.(`Annulation impossible : ${messageErreur(e)}`, 'error');
            }
          });
          ctx.onInfo?.(r.simule ? 'Action simulée (mode à blanc)' : 'Action faite', 'success');
          return;
        }
        btn.disabled = false;
        btn.textContent = libelle;
        res.hidden = false;
        if (r.erreur === 'choix-requis') {
          const candidats = candidatsDe(a);
          if (candidats.length && choixWrap.hidden) {
            choixWrap.innerHTML = selectProjetHtml(i, candidats);
            choixWrap.hidden = false;
          }
          res.innerHTML = `<p class="agent-error" role="alert">${escapeHtml(r.resume || 'Un choix est nécessaire.')}${candidats.length ? '' : ' À faire dans ATLAS.'}</p>`;
          return;
        }
        res.innerHTML = `<p class="agent-error" role="alert">${escapeHtml(r.resume || 'Action impossible')}${r.erreur ? ` (${escapeHtml(r.erreur)})` : ''}</p>`;
      } catch (e) {
        if (!host.isConnected) return;
        btn.disabled = false;
        btn.textContent = libelle;
        res.hidden = false;
        res.innerHTML = erreurHtml('Action impossible', e);
      }
    });
  });
}

// ── « Répondre plus tard » / « Relancer si pas de réponse » ──

function demainIso(): string {
  const d = new Date(Date.now() + 86_400_000);
  return d.toLocaleDateString('sv-SE', { timeZone: 'Europe/Luxembourg' }); // AAAA-MM-JJ
}

function jourLisible(iso: string | undefined): string {
  if (!iso) return '';
  const jourSeul = /^\d{4}-\d{2}-\d{2}$/.test(iso);
  const d = new Date(jourSeul ? `${iso}T12:00:00Z` : iso);
  if (Number.isNaN(d.getTime())) return iso;
  return d.toLocaleDateString('fr-FR', { timeZone: jourSeul ? 'UTC' : 'Europe/Luxembourg', weekday: 'long', day: 'numeric', month: 'long' });
}

/**
 * Rappels d'un mail : reçu → « Répondre plus tard » (demain, lundi, date) ; envoyé → « Relancer si
 * pas de réponse » (1 à 30 jours ouvrés) et « Pas de relance ». Écritures par le worker seulement.
 */
export function renderRappels(host: HTMLElement, ctx: { messageId: string; envoye: boolean; onInfo?: InfoFn; onChange?: () => void }): void {
  if (!ctx.messageId) { host.hidden = true; return; }
  host.hidden = false;
  if (ctx.envoye) renderRelance(host, ctx);
  else renderPlusTard(host, ctx);
}

function renderPlusTard(host: HTMLElement, ctx: { messageId: string; onInfo?: InfoFn; onChange?: () => void }): void {
  host.innerHTML = `
    <div class="agent-section-title">Répondre plus tard</div>
    <div class="agent-rappel-btns">
      <button type="button" class="btn btn-secondary agent-btn" data-quand="demain">Demain</button>
      <button type="button" class="btn btn-secondary agent-btn" data-quand="lundi">Lundi</button>
    </div>
    <div class="agent-rappel-date">
      <input type="date" class="agent-input" id="agent-plustard-date" min="${demainIso()}" aria-label="Date du retour" />
      <button type="button" class="btn btn-secondary agent-btn" data-quand="date">Ce jour-là</button>
    </div>
    <div class="agent-muted">Le mail sort de « à traiter » et revient en tête ce jour-là à 8 h (jours ouvrés).</div>
    <div id="agent-plustard-res" hidden></div>
  `;
  const res = host.querySelector<HTMLElement>('#agent-plustard-res')!;
  const boutons = () => host.querySelectorAll<HTMLButtonElement>('button[data-quand]');
  boutons().forEach(btn => {
    btn.addEventListener('click', async () => {
      let quand = btn.dataset.quand!;
      if (quand === 'date') {
        quand = host.querySelector<HTMLInputElement>('#agent-plustard-date')?.value || '';
        if (!quand) { res.hidden = false; res.innerHTML = '<p class="agent-error" role="alert">Choisis une date.</p>'; return; }
      }
      boutons().forEach(b => { b.disabled = true; });
      res.hidden = true;
      try {
        const r = await mettrePlusTard(ctx.messageId, quand);
        if (!host.isConnected) return;
        host.innerHTML = `
          <div class="agent-section-title">Répondre plus tard</div>
          <div class="agent-fait" role="status">
            <span>Revient ${escapeHtml(jourLisible(r.jusqua))}${r.mode !== 'actif' ? ' (agent en observation : rien n\'est écrit dans ta boîte)' : ''}.</span>
            <span class="agent-fait-btns"><button type="button" class="btn btn-secondary agent-btn" data-retirer>Le remettre maintenant</button></span>
          </div>`;
        host.querySelector<HTMLButtonElement>('[data-retirer]')?.addEventListener('click', async (ev) => {
          const b = ev.currentTarget as HTMLButtonElement;
          b.disabled = true;
          try {
            await retirerPlusTard(ctx.messageId);
            ctx.onInfo?.('Mail remis dans « à traiter »', 'success');
            ctx.onChange?.();
            if (host.isConnected) renderPlusTard(host, ctx);
          } catch (e) {
            b.disabled = false;
            ctx.onInfo?.(`Impossible : ${messageErreur(e)}`, 'error');
          }
        });
        ctx.onInfo?.('Rappel posé', 'success');
        ctx.onChange?.();
      } catch (e) {
        if (!host.isConnected) return;
        boutons().forEach(b => { b.disabled = false; });
        res.hidden = false;
        res.innerHTML = erreurHtml('Rappel impossible', e);
      }
    });
  });
}

function renderRelance(host: HTMLElement, ctx: { messageId: string; onInfo?: InfoFn; onChange?: () => void }): void {
  const options = Array.from({ length: 30 }, (_, k) => k + 1)
    .map(n => `<option value="${n}" ${n === 3 ? 'selected' : ''}>${n} jour${n > 1 ? 's' : ''} ouvré${n > 1 ? 's' : ''}</option>`).join('');
  host.innerHTML = `
    <div class="agent-section-title">Relancer si pas de réponse</div>
    <div class="agent-rappel-date">
      <select class="agent-select" id="agent-relance-delai" aria-label="Délai avant relance">${options}</select>
      <button type="button" class="btn btn-secondary agent-btn" data-relance="oui">Relancer</button>
    </div>
    <button type="button" class="btn btn-secondary btn-block agent-btn" data-relance="non">Pas de relance</button>
    <div class="agent-muted">Sans réponse dans le délai, l'agent prépare un brouillon de relance (jamais envoyé).</div>
    <div id="agent-relance-res" hidden></div>
  `;
  const res = host.querySelector<HTMLElement>('#agent-relance-res')!;
  host.querySelectorAll<HTMLButtonElement>('button[data-relance]').forEach(btn => {
    btn.addEventListener('click', async () => {
      const oui = btn.dataset.relance === 'oui';
      const delai = Number(host.querySelector<HTMLSelectElement>('#agent-relance-delai')?.value || 3);
      host.querySelectorAll<HTMLButtonElement>('button').forEach(b => { b.disabled = true; });
      res.hidden = true;
      try {
        const r = oui ? await demanderRelance(ctx.messageId, delai) : await annulerRelance(ctx.messageId);
        if (!host.isConnected) return;
        host.querySelectorAll<HTMLButtonElement>('button').forEach(b => { b.disabled = false; });
        res.hidden = false;
        res.innerHTML = `<div class="agent-fait" role="status"><span>${oui
          ? `Relance prévue ${escapeHtml(jourLisible(r.relanceLe || r.jusqua))} sans réponse d'ici là.`
          : 'Pas de relance pour ce mail (ni manuelle ni automatique).'}</span></div>`;
        ctx.onInfo?.(oui ? 'Relance programmée' : 'Relance retirée', 'success');
        ctx.onChange?.();
      } catch (e) {
        if (!host.isConnected) return;
        host.querySelectorAll<HTMLButtonElement>('button').forEach(b => { b.disabled = false; });
        res.hidden = false;
        res.innerHTML = erreurHtml(oui ? 'Relance impossible' : 'Impossible', e);
      }
    });
  });
}

// ── Résumé du fil ──

export function renderResumeFil(host: HTMLElement, ctx: { conversationId: string; mailbox: string }): void {
  if (!ctx.conversationId) { host.hidden = true; return; }
  host.hidden = false;
  host.innerHTML = `
    <button type="button" class="btn btn-secondary btn-block agent-btn" id="agent-resume-fil-btn">Résumer le fil</button>
    <div id="agent-resume-fil" hidden></div>`;
  const btn = host.querySelector<HTMLButtonElement>('#agent-resume-fil-btn')!;
  const out = host.querySelector<HTMLElement>('#agent-resume-fil')!;
  btn.addEventListener('click', async () => {
    btn.disabled = true;
    out.hidden = false;
    out.innerHTML = '<div class="agent-loading"><div class="spinner"></div><span>Résumé du fil…</span></div>';
    try {
      const r = await fetchResumeFil(ctx.conversationId, ctx.mailbox);
      if (!host.isConnected) return;
      out.innerHTML = r.lignes.length
        ? `<div class="agent-section-title">Le fil en bref${r.nbMails ? ` (${r.nbMails} mail${r.nbMails > 1 ? 's' : ''})` : ''}</div>
           <ul class="agent-resume-fil">${r.lignes.map(l => `<li>${escapeHtml(l)}</li>`).join('')}</ul>
           ${r.misAJour ? `<div class="agent-muted">Mis à jour ${escapeHtml(dateCourte(r.misAJour))}</div>` : ''}`
        : '<p class="agent-muted">Rien à résumer pour ce fil.</p>';
      btn.textContent = 'Actualiser le résumé';
    } catch (e) {
      if (!host.isConnected) return;
      out.innerHTML = erreurHtml('Résumé indisponible', e);
    } finally {
      btn.disabled = false;
    }
  });
}

// ── Poser une question à ma boîte ──

export function renderQuestion(host: HTMLElement, ctx: { mailbox?: string; openLink?: (url: string) => void; onInfo?: InfoFn }): void {
  host.innerHTML = `
    <div class="agent-section-title">Poser une question à ma boîte</div>
    <form class="agent-question" novalidate>
      <input type="text" class="agent-input" name="q" maxlength="500" autocomplete="off"
        placeholder="Qu'a dit le client sur le budget ?" aria-label="Question" />
      <button type="submit" class="btn btn-primary agent-btn">Demander</button>
    </form>
    <div class="agent-question-res" hidden></div>`;
  const form = host.querySelector<HTMLFormElement>('form')!;
  const input = form.querySelector<HTMLInputElement>('input')!;
  const submit = form.querySelector<HTMLButtonElement>('button')!;
  const out = host.querySelector<HTMLElement>('.agent-question-res')!;
  form.addEventListener('submit', async (ev) => {
    ev.preventDefault();
    const q = input.value.trim();
    if (q.length < 3) { out.hidden = false; out.innerHTML = '<p class="agent-error" role="alert">Écris ta question.</p>'; return; }
    submit.disabled = true;
    out.hidden = false;
    out.innerHTML = '<div class="agent-loading"><div class="spinner"></div><span>Recherche dans ta boîte…</span></div>';
    try {
      const r = await poserQuestion(q, ctx.mailbox);
      if (!host.isConnected) return;
      out.innerHTML = `
        <p class="agent-reponse">${escapeHtml(r.reponse || 'Pas de réponse trouvée dans tes mails.')}</p>
        ${r.sources.length ? `<div class="agent-muted">Mails cités :</div>${sourcesHtml(r.sources, 'src')}` : ''}`;
      bindSources(out, r.sources, 'src', ctx.openLink, ctx.onInfo);
    } catch (e) {
      if (!host.isConnected) return;
      out.innerHTML = erreurHtml('Question impossible', e);
    } finally {
      submit.disabled = false;
    }
  });
}

// ── Rattrapage (« ce qui s'est passé pendant ton absence ») ──

function ilYaJours(n: number): string {
  return new Date(Date.now() - n * 86_400_000).toLocaleDateString('sv-SE', { timeZone: 'Europe/Luxembourg' });
}

export function renderRattrapage(host: HTMLElement, ctx: { mailbox?: string; openLink?: (url: string) => void; onInfo?: InfoFn }): void {
  host.innerHTML = `
    <form class="agent-rattrapage" novalidate>
      <label class="agent-muted" for="tdb-rattrapage-date">Rattraper depuis le</label>
      <input type="date" class="agent-input" id="tdb-rattrapage-date" value="${ilYaJours(7)}" max="${ilYaJours(0)}" min="${ilYaJours(90)}" />
      <button type="submit" class="btn btn-secondary agent-btn">Rattraper</button>
    </form>
    <div class="agent-rattrapage-res" hidden></div>`;
  const form = host.querySelector<HTMLFormElement>('form')!;
  const input = form.querySelector<HTMLInputElement>('input')!;
  const submit = form.querySelector<HTMLButtonElement>('button')!;
  const out = host.querySelector<HTMLElement>('.agent-rattrapage-res')!;
  form.addEventListener('submit', async (ev) => {
    ev.preventDefault();
    const jour = input.value;
    if (!/^\d{4}-\d{2}-\d{2}$/.test(jour)) { out.hidden = false; out.innerHTML = '<p class="agent-error" role="alert">Choisis une date.</p>'; return; }
    // Minuit à Luxembourg ≈ 22 h / 23 h UTC la veille : on envoie le début de journée UTC, le worker borne.
    const depuis = new Date(`${jour}T00:00:00Z`).toISOString();
    submit.disabled = true;
    out.hidden = false;
    out.innerHTML = '<div class="agent-loading"><div class="spinner"></div><span>Rattrapage…</span></div>';
    try {
      const r = await fetchRattrapage(depuis, ctx.mailbox);
      if (!host.isConnected) return;
      const tous: SourceMail[] = [];
      const sections = r.sections.filter(s => s.elements.length).map(s => {
        const debut = tous.length;
        tous.push(...s.elements);
        const html = sourcesHtml(s.elements, 'rat').replace(/data-rat="(\d+)"/g, (_m, k) => `data-rat="${debut + Number(k)}"`);
        return `<div class="agent-section-title">${escapeHtml(s.titre)} (${s.elements.length})</div>${html}`;
      });
      out.innerHTML = `
        ${r.intro ? `<p class="agent-reponse">${escapeHtml(r.intro)}</p>` : ''}
        ${sections.length ? sections.join('') : '<p class="agent-muted">Rien de notable depuis cette date.</p>'}`;
      bindSources(out, tous, 'rat', ctx.openLink, ctx.onInfo);
    } catch (e) {
      if (!host.isConnected) return;
      out.innerHTML = erreurHtml('Rattrapage indisponible', e);
    } finally {
      submit.disabled = false;
    }
  });
}

// ── Phase 6 : copier un texte ──

/** Copie un texte (presse-papiers ; repli : sélection de la zone + execCommand). */
export async function copierTexte(texte: string, zone?: HTMLTextAreaElement | null): Promise<boolean> {
  try {
    if (navigator.clipboard?.writeText) { await navigator.clipboard.writeText(texte); return true; }
  } catch { /* presse-papiers refusé dans la fenêtre du complément : repli */ }
  try {
    if (zone) {
      zone.focus();
      zone.select();
      return document.execCommand('copy');
    }
  } catch { /* rien */ }
  return false;
}

function zoneTexteHtml(id: string, texte: string, lignes = 8): string {
  return `<textarea class="agent-input agent-texte" id="${id}" rows="${lignes}" readonly>${escapeHtml(texte)}</textarea>`;
}

function bindCopier(host: HTMLElement, bouton: string, zone: string, texte: () => string, onInfo?: InfoFn): void {
  host.querySelector<HTMLButtonElement>(bouton)?.addEventListener('click', async () => {
    const ok = await copierTexte(texte(), host.querySelector<HTMLTextAreaElement>(zone));
    onInfo?.(ok ? 'Texte copié' : 'Copie impossible : sélectionne le texte et copie-le', ok ? 'success' : 'error');
  });
}

// ── Phase 6 : « Traduire » (mail lu) ──

const LANGUES: Array<{ code: LangueTraduction; libelle: string }> = [
  { code: 'FR', libelle: 'Français' }, { code: 'EN', libelle: 'English' }, { code: 'DE', libelle: 'Deutsch' }, { code: 'LB', libelle: 'Lëtzebuergesch' },
];

/** Traduction du mail lu, à la demande (IA côté worker, sous le plafond du jour). */
export function renderTraduction(host: HTMLElement, ctx: { messageId: string; mailbox?: string; onInfo?: InfoFn }): void {
  if (!ctx.messageId) { host.hidden = true; return; }
  host.hidden = false;
  host.innerHTML = `
    <div class="agent-section-title">Traduire ce mail</div>
    <div class="agent-rappel-btns agent-langues">
      ${LANGUES.map(l => `<button type="button" class="btn btn-secondary agent-btn" data-vers="${l.code}" title="${escapeHtml(l.libelle)}">${l.code}</button>`).join('')}
    </div>
    <div class="agent-traduction-res" hidden></div>`;
  const out = host.querySelector<HTMLElement>('.agent-traduction-res')!;
  const boutons = () => host.querySelectorAll<HTMLButtonElement>('button[data-vers]');
  let texte = '';
  boutons().forEach(btn => {
    btn.addEventListener('click', async () => {
      const vers = btn.dataset.vers as LangueTraduction;
      boutons().forEach(b => { b.disabled = true; });
      out.hidden = false;
      out.innerHTML = '<div class="agent-loading"><div class="spinner"></div><span>Traduction…</span></div>';
      try {
        const r = await traduire({ vers, messageId: ctx.messageId, ...(ctx.mailbox ? { mailbox: ctx.mailbox } : {}) });
        if (!host.isConnected) return;
        texte = r.texte;
        out.innerHTML = `
          <div class="agent-muted">Traduction (${escapeHtml(vers)}), à relire :</div>
          ${zoneTexteHtml('agent-traduction-texte', texte, 10)}
          <button type="button" class="btn btn-secondary btn-block agent-btn" data-copier>Copier</button>`;
        bindCopier(out, '[data-copier]', '#agent-traduction-texte', () => texte, ctx.onInfo);
      } catch (e) {
        if (!host.isConnected) return;
        out.innerHTML = erreurHtml('Traduction indisponible', e);
      } finally {
        boutons().forEach(b => { b.disabled = false; });
      }
    });
  });
}

// ── Phase 6 : « Modèles de réponse » ──

export interface CtxModeles {
  /** Mail lu (sinon `conversationId` : rédaction d'une réponse). */
  messageId?: string;
  conversationId?: string;
  mailbox?: string;
  onInfo?: InfoFn;
  /** Rédaction : insère le texte dans le mail en cours (absent si l'API Office manque). */
  inserer?: (texte: string, objet: string) => Promise<boolean>;
  /** Lecture (bureau / web) : ouvre une réponse avec ce texte (absent sur mobile). */
  repondre?: (texte: string) => void;
}

/** Modèles de réponse à variables (fiches Communications), remplis pour ce mail ; jamais de signature. */
export function renderModeles(host: HTMLElement, ctx: CtxModeles): void {
  if (!ctx.messageId && !ctx.conversationId) { host.hidden = true; return; }
  host.hidden = false;
  host.innerHTML = `
    <button type="button" class="btn btn-secondary btn-block agent-btn" data-modeles>Modèles de réponse</button>
    <div class="agent-modeles" hidden></div>`;
  const ouvrir = host.querySelector<HTMLButtonElement>('[data-modeles]')!;
  const zone = host.querySelector<HTMLElement>('.agent-modeles')!;
  let modeles: ModeleReponse[] | null = null;
  ouvrir.addEventListener('click', async () => {
    if (!zone.hidden && modeles) { zone.hidden = true; return; }
    zone.hidden = false;
    if (modeles) return;
    ouvrir.disabled = true;
    zone.innerHTML = '<div class="agent-loading"><div class="spinner"></div><span>Modèles…</span></div>';
    try {
      modeles = await fetchModeles();
    } catch (e) {
      if (!host.isConnected) return;
      zone.innerHTML = erreurHtml('Modèles indisponibles', e);
      ouvrir.disabled = false;
      return;
    }
    ouvrir.disabled = false;
    if (!host.isConnected) return;
    if (!modeles.length) { zone.innerHTML = '<p class="agent-muted">Aucun modèle de réponse dans Communications.</p>'; return; }
    zone.innerHTML = `
      <div class="agent-section-title">Modèles de réponse</div>
      <div class="agent-rappel-date">
        <select class="agent-select" aria-label="Modèle" data-choix-modele>
          ${modeles.map(m => `<option value="${escapeHtml(m.id)}">${escapeHtml(m.nom)}${m.categorie ? ` · ${escapeHtml(m.categorie)}` : ''}</option>`).join('')}
        </select>
        <button type="button" class="btn btn-primary agent-btn" data-remplir>Remplir</button>
      </div>
      <div class="agent-modele-res" hidden></div>`;
    const res = zone.querySelector<HTMLElement>('.agent-modele-res')!;
    const remplir = zone.querySelector<HTMLButtonElement>('[data-remplir]')!;
    remplir.addEventListener('click', async () => {
      const modeleId = zone.querySelector<HTMLSelectElement>('[data-choix-modele]')?.value || '';
      if (!modeleId) return;
      remplir.disabled = true;
      res.hidden = false;
      res.innerHTML = '<div class="agent-loading"><div class="spinner"></div><span>Remplissage…</span></div>';
      try {
        const r: ModeleRempli = await remplirModele({
          modeleId, ...(ctx.messageId ? { messageId: ctx.messageId } : { conversationId: ctx.conversationId }), ...(ctx.mailbox ? { mailbox: ctx.mailbox } : {}),
        });
        if (!host.isConnected) return;
        res.innerHTML = `
          ${r.objet ? `<div class="agent-muted">Objet : <strong>${escapeHtml(r.objet)}</strong></div>` : ''}
          ${zoneTexteHtml('agent-modele-texte', r.corps, 10)}
          ${r.manquantes.length ? `<ul class="agent-avertissements"><li>À compléter à la main : ${r.manquantes.map(v => escapeHtml(v)).join(', ')}</li></ul>` : ''}
          <div class="agent-muted">Sans signature : Exclaimer l'ajoute à l'envoi.</div>
          <div class="agent-rappel-btns">
            <button type="button" class="btn btn-secondary agent-btn" data-copier>Copier</button>
            ${ctx.inserer ? '<button type="button" class="btn btn-primary agent-btn" data-inserer>Insérer dans le mail</button>' : ''}
            ${ctx.repondre ? '<button type="button" class="btn btn-primary agent-btn" data-repondre>Répondre avec ce texte</button>' : ''}
          </div>`;
        bindCopier(res, '[data-copier]', '#agent-modele-texte', () => r.corps, ctx.onInfo);
        res.querySelector<HTMLButtonElement>('[data-inserer]')?.addEventListener('click', async () => {
          const ok = await ctx.inserer!(r.corps, r.objet).catch(() => false);
          ctx.onInfo?.(ok ? 'Modèle inséré dans le mail' : 'Insertion impossible : copie le texte', ok ? 'success' : 'error');
        });
        res.querySelector<HTMLButtonElement>('[data-repondre]')?.addEventListener('click', () => ctx.repondre!(r.corps));
      } catch (e) {
        if (!host.isConnected) return;
        res.innerHTML = erreurHtml('Modèle impossible à remplir', e);
      } finally {
        remplir.disabled = false;
      }
    });
  });
}

// ── Phase 6 : fiche mémoire du contact (la personne seule) ──

const SOURCE_FICHE: Record<string, string> = { correction: 'ta correction', profil: 'profil de conversation', observation: 'observé dans tes échanges' };

/** Ce que l'agent retient de ce contact POUR TOI (mémoire personnelle) ; masqué s'il n'y a rien. */
export async function renderFicheContact(host: HTMLElement, ctx: { email: string }): Promise<void> {
  host.hidden = true;
  if (!ctx.email) return;
  let f;
  try { f = await fetchFicheContact(ctx.email); } catch { return; /* fiche facultative */ }
  if (!f || !host.isConnected) return;
  const lignes: string[] = [];
  if (typeof f.tutoiement === 'boolean') lignes.push(`${f.tutoiement ? 'Tutoiement' : 'Vouvoiement'}${f.sources?.tutoiement ? ` <span class="agent-muted">(${escapeHtml(SOURCE_FICHE[f.sources.tutoiement] || f.sources.tutoiement)})</span>` : ''}`);
  if (f.langue) lignes.push(`Langue : ${escapeHtml(f.langue)}${f.sources?.langue ? ` <span class="agent-muted">(${escapeHtml(SOURCE_FICHE[f.sources.langue] || f.sources.langue)})</span>` : ''}`);
  if (f.ton) lignes.push(`Ton : ${escapeHtml(f.ton)}`);
  if (f.sujets.length) lignes.push(`Sujets en cours : ${f.sujets.slice(0, 5).map(x => escapeHtml(x)).join(' · ')}`);
  if (!lignes.length && !f.echanges) return;
  host.hidden = false;
  host.innerHTML = `
    <div class="agent-muted">Ta mémoire de ce contact (visible par toi seul${f.echanges ? `, ${f.echanges} échange${f.echanges > 1 ? 's' : ''}` : ''}) :</div>
    <ul class="agent-facts">${lignes.map(l => `<li>${l}</li>`).join('')}</ul>
    <div class="agent-muted">À corriger ou oublier dans ATLAS › Mon agent › Mes contacts.</div>`;
}

// ── Phase 6 : détections (rebond, absence, candidature) ──

const TYPES_REBOND: Record<string, string> = {
  'adresse-invalide': 'Adresse en échec (mail non remis)',
  depart: 'Départ annoncé : la personne ne reçoit plus à cette adresse',
  'boite-pleine': 'Boîte du destinataire pleine',
  temporaire: 'Non-remise temporaire (nouvel essai du serveur)',
};
const SIGNALEMENTS: Record<string, string> = {
  fait: 'fiche ATLAS signalée',
  simule: 'fiche ATLAS à signaler (agent en observation)',
  'sans-fiche': 'aucune fiche ATLAS pour cette adresse',
  'non-autorise': 'fiche non modifiée (tes réglages d’autonomie)',
};

function jourCourt(iso: string): string {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(iso)) return iso;
  return new Date(`${iso}T12:00:00Z`).toLocaleDateString('fr-FR', { timeZone: 'UTC', weekday: 'long', day: 'numeric', month: 'long' });
}

/** Encart des détections du mail ; masqué s'il n'y a rien. */
export function renderDetections(host: HTMLElement, d: DetectionsMail | null | undefined): void {
  const items: string[] = [];
  if (d?.rebond) {
    const r = d.rebond;
    items.push(`<li><strong>${escapeHtml(TYPES_REBOND[r.type] || 'Mail non remis')}</strong>${r.adresse ? ` : ${escapeHtml(r.adresse)}` : ''}${r.signale && SIGNALEMENTS[r.signale] ? ` <span class="agent-muted">(${escapeHtml(SIGNALEMENTS[r.signale])})</span>` : ''}</li>`);
    if (r.nouveauContact?.email) items.push(`<li>Nouveau contact suggéré : ${escapeHtml(r.nouveauContact.nom ? `${r.nouveauContact.nom} ` : '')}&lt;${escapeHtml(r.nouveauContact.email)}&gt; <span class="agent-muted">(à créer d'un clic dans « Que faire de ce mail ? » ou dans ATLAS)</span></li>`);
  }
  if (d?.absence?.retourLe) {
    items.push(`<li><strong>Absent</strong> : de retour ${escapeHtml(jourCourt(d.absence.retourLe))}. Les relances prévues pendant son absence sont décalées au lendemain ouvré de son retour.</li>`);
  }
  if (d?.candidat) {
    items.push('<li><strong>Candidature</strong> : relevée pour le RGPD (effacement à prévoir 12 mois après la fin du processus si le candidat n\'est pas retenu ; jamais automatique).</li>');
  }
  if (!items.length) { host.hidden = true; host.innerHTML = ''; return; }
  host.hidden = false;
  host.innerHTML = `<div class="agent-section-title">À savoir</div><ul class="agent-avertissements">${items.join('')}</ul>`;
}

// ── Phase 6 : pièces jointes lourdes → lien court ──

/** Pièce jointe vue par Office (identifiant Office : EWS ou REST selon l'hôte). */
export interface PieceOffice { id?: string; nom: string; octets: number; isInline?: boolean }

/** Contexte du dépôt : mail source (identifiant REST) et, en rédaction, insertion / retrait dans le brouillon. */
export interface DepotCtx {
  /** Mail dont le worker lit les pièces (reçu, ou brouillon enregistré juste avant). */
  source: () => Promise<{ messageId: string; mailbox?: string }>;
  /** Identifiant REST d'une pièce (convertToRestId) ; identité si indisponible. */
  idRest?: (id: string) => string;
  /** Rédaction : insère le HTML des liens dans le brouillon (true si inséré). */
  inserer?: (html: string) => Promise<boolean>;
  /** Rédaction : retire du brouillon les pièces déposées (identifiants Office) ; nombre retiré. */
  retirer?: (ids: string[]) => Promise<number>;
}

const tailleMo = (o: number) => `${(o / (1024 * 1024)).toFixed(1).replace('.', ',')} Mo`;
const attendre = (ms: number) => new Promise(ok => setTimeout(ok, ms));

/**
 * Pièces jointes au-delà du seuil (worker) : « Déposer et partager » (cloud, transit 30 jours,
 * jamais le NAS) → liens courts vibes.lu/p/… insérés dans le brouillon (rédaction) ou à copier
 * (lecture) ; parcours conservé : lien court à partir d'un lien de partage collé. Masqué sinon.
 */
export async function renderPiecesLourdes(host: HTMLElement, ctx: { pieces: PieceOffice[]; onInfo?: InfoFn; depot?: DepotCtx }): Promise<void> {
  host.hidden = true;
  const pieces = ctx.pieces.filter(p => p && p.octets > 0 && !p.isInline);
  if (!pieces.length) return;
  let lourde;
  try { lourde = await verifierPiecesLourdes(pieces); } catch { return; /* suggestion facultative */ }
  if (!lourde || !host.isConnected) return;
  const lourdesNoms = new Set(lourde.lourdes.map(l => l.nom));
  const depot = ctx.depot;
  const deposables = depot ? pieces.filter(p => !!p.id) : [];
  host.hidden = false;
  host.innerHTML = `
    <div class="agent-section-title">Pièces jointes lourdes</div>
    <p class="agent-muted">${escapeHtml(lourde.suggestion)}</p>
    ${deposables.length ? `
      <div class="agent-pj-depot">
        ${deposables.map((p, i) => `<label class="agent-check"><input type="checkbox" data-pj="${i}" ${lourdesNoms.has(p.nom) ? 'checked' : ''}/> ${escapeHtml(p.nom)} <span class="agent-muted">(${tailleMo(p.octets)})</span></label>`).join('')}
        ${depot?.retirer ? '<label class="agent-check"><input type="checkbox" data-retirer checked/> Retirer ensuite ces pièces du brouillon</label>' : ''}
        <button type="button" class="btn btn-primary btn-block agent-btn" data-deposer>${depot?.inserer ? 'Déposer sur le cloud et insérer les liens' : 'Déposer sur le cloud et créer les liens'}</button>
        <p class="agent-muted">Cloud de l'agence (dossier de transit), lien en lecture seule valable 30 jours, puis fichier effacé. 150 Mo au plus par fichier ; exécutables et scripts refusés.</p>
        <div class="agent-pj-depot-res" hidden></div>
      </div>
      <div class="agent-muted">Ou, si le fichier est déjà partagé :</div>` : ''}
    <form class="agent-rappel-date agent-lien-court" novalidate>
      <input type="url" class="agent-input" name="url" placeholder="Lien de partage (https://…)" aria-label="Lien de partage existant" />
      <button type="submit" class="btn btn-secondary agent-btn">Lien court</button>
    </form>
    <div class="agent-lien-court-res" hidden></div>`;

  // ── Déposer et partager ──
  const btnDepot = host.querySelector<HTMLButtonElement>('[data-deposer]');
  const resDepot = host.querySelector<HTMLElement>('.agent-pj-depot-res');
  if (btnDepot && resDepot && depot) {
    btnDepot.addEventListener('click', async () => {
      const choisies = deposables.filter((_, i) => host.querySelector<HTMLInputElement>(`[data-pj="${i}"]`)?.checked);
      if (!choisies.length) { resDepot.hidden = false; resDepot.innerHTML = '<p class="agent-error" role="alert">Coche au moins une pièce.</p>'; return; }
      btnDepot.disabled = true;
      resDepot.hidden = false;
      resDepot.innerHTML = '<div class="agent-loading"><div class="spinner"></div><span>Dépôt sur le cloud… (quelques minutes pour les gros fichiers)</span></div>';
      try {
        const src = await depot.source();
        const idRest = depot.idRest || ((x: string) => x);
        const demande = { ...src, attachmentIds: choisies.map(p => idRest(String(p.id))), pieces: choisies.map(p => ({ nom: p.nom, octets: p.octets })) };
        let r: DepotPiecesLourdes | null = null;
        // Brouillon tout juste enregistré : le serveur peut mettre quelques secondes à le voir (404).
        for (let essai = 0; ; essai++) {
          try { r = await deposerPiecesLourdes(demande); break; } catch (e) {
            if ((e as any)?.status === 404 && essai < 3) { await attendre(3000); continue; }
            throw e;
          }
        }
        if (!host.isConnected || !r) return;
        const refus = r.refus.length ? `<ul class="agent-avertissements">${r.refus.map(x => `<li>${escapeHtml(x.nom)} : ${escapeHtml(x.motif)}</li>`).join('')}</ul>` : '';
        // Le worker refuse une pièce sans lien court (jamais de lien brut du cloud) : rien à signaler ici.
        const nonCourt = '';
        const texte = r.liens.map(l => `${l.nom} : ${l.lienCourt} (lien valable 30 jours)`).join('\n');
        let insere = false;
        if (depot.inserer && r.html) insere = await depot.inserer(r.html);
        let retirees = 0;
        if (insere && depot.retirer && host.querySelector<HTMLInputElement>('[data-retirer]')?.checked) {
          // `rang` = position dans la demande : seules les pièces réellement déposées sont retirées.
          const aRetirer = r.liens.map(l => choisies[l.rang]?.id).filter((x): x is string => !!x);
          retirees = await depot.retirer(aRetirer);
        }
        resDepot.innerHTML = `
          ${insere ? `<p class="agent-muted">${r.liens.length} lien(s) inséré(s) dans le brouillon${retirees ? `, ${retirees} pièce(s) retirée(s)` : ''}. Rien n'est envoyé.</p>` : ''}
          ${zoneTexteHtml('agent-pj-depot-liens', texte, Math.min(6, Math.max(1, r.liens.length)))}
          <button type="button" class="btn btn-secondary btn-block agent-btn" data-copier-depot>Copier les liens</button>
          ${nonCourt}${refus}`;
        bindCopier(resDepot, '[data-copier-depot]', '#agent-pj-depot-liens', () => texte, ctx.onInfo);
        if (insere) ctx.onInfo?.('Liens insérés dans le brouillon', 'success');
      } catch (e) {
        if (!host.isConnected) return;
        const refus: Array<{ nom: string; motif: string }> = Array.isArray((e as any)?.data?.refus) ? (e as any).data.refus : [];
        resDepot.innerHTML = erreurHtml('Dépôt impossible', e) + (refus.length ? `<ul class="agent-avertissements">${refus.map(x => `<li>${escapeHtml(x.nom)} : ${escapeHtml(x.motif)}</li>`).join('')}</ul>` : '');
      } finally {
        btnDepot.disabled = false;
      }
    });
  }

  // ── Lien court d'un lien de partage collé (parcours existant) ──
  const form = host.querySelector<HTMLFormElement>('form')!;
  const out = host.querySelector<HTMLElement>('.agent-lien-court-res')!;
  form.addEventListener('submit', async (ev) => {
    ev.preventDefault();
    const url = (form.querySelector<HTMLInputElement>('input')?.value || '').trim();
    if (!/^https:\/\//i.test(url)) { out.hidden = false; out.innerHTML = '<p class="agent-error" role="alert">Colle un lien de partage https.</p>'; return; }
    const btn = form.querySelector<HTMLButtonElement>('button')!;
    btn.disabled = true;
    out.hidden = false;
    out.innerHTML = '<div class="agent-loading"><div class="spinner"></div><span>Lien court…</span></div>';
    try {
      const lien = await creerLienCourt(url);
      if (!host.isConnected) return;
      out.innerHTML = `${zoneTexteHtml('agent-lien-court', lien, 1)}<button type="button" class="btn btn-secondary btn-block agent-btn" data-copier>Copier le lien</button>`;
      bindCopier(out, '[data-copier]', '#agent-lien-court', () => lien, ctx.onInfo);
    } catch (e) {
      if (!host.isConnected) return;
      out.innerHTML = erreurHtml('Lien court indisponible (garde le lien d’origine)', e);
    } finally {
      btn.disabled = false;
    }
  });
}
