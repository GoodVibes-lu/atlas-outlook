/**
 * journee-banner.ts — Bandeau « Ma journée » en tête du panneau ATLAS.
 *
 * Tient lieu de tableau de bord sur mobile et dans Outlook classique Windows (décision du
 * 03/10/2026, backlog agent d'inbox §7) : « 12 à traiter · 5 en attente · 3 relances dues ·
 * 2 clients > 48 h · 4 à filtrer ». Données : GET /api/plugin/agent/journee (aucun appel IA),
 * rafraîchies à l'ouverture et au changement de mail, au plus une fois par minute (cf. fetchJournee).
 *
 * Phase 2 : les compteurs de pile (à traiter, en attente, à filtrer) sont cliquables : clic = liste
 * des mails de la pile sous le bandeau (GET /api/plugin/agent/liste), clic sur un mail = ouverture ;
 * dans « à filtrer », Accepter / Refuser l'expéditeur. Second clic sur le même compteur = fermer.
 * Phase 3 : compteur « N plus tard » (s'il y en a) → pile « Plus tard » avec « Remettre maintenant ».
 * Phase 5 (membres de good@ seulement) : « N à attribuer (good@) » et « N pour toi sur good@ »,
 * cliquables → liste des mails de la boîte partagée (assigné / pris par).
 */

import { fetchJournee, journeeMemorisee, type AgentJournee, type AgentListePile } from '../api/agent';
import { renderMailList, type MailListOptions } from './agent-lists';
import { isMobile } from '../api/platform';

/** Séance de tri (07/10/2026) : ouverte dans la grande fenêtre du tableau de bord (dialogue Office, Mac compris). */
function seancePossible(): boolean {
  try { return !isMobile() && typeof Office !== 'undefined' && typeof Office.context?.ui?.displayDialogAsync === 'function'; } catch { return false; }
}

interface Compteur { texte: string; n: number; libelle: string; pile?: AgentListePile; alerte?: boolean }

function c(n: number, libelle: string, extra: Partial<Compteur> = {}): Compteur {
  return { n, libelle, texte: `${n} ${libelle}`, ...extra };
}

function compteurs(j: AgentJournee): Compteur[] {
  const out: Compteur[] = [
    c(j.aTraiter, 'à traiter', { pile: 'a_traiter', alerte: j.aTraiter > 0 }),
    c(j.enAttente, 'en attente', { pile: 'en_attente' }),
    c(j.relancesDues, j.relancesDues > 1 ? 'relances dues' : 'relance due'),
    c(j.clientsPlus48h, j.clientsPlus48h > 1 ? 'clients > 48 h' : 'client > 48 h', { alerte: j.clientsPlus48h > 0 }),
  ];
  if (j.aFiltrer > 0) out.push(c(j.aFiltrer, 'à filtrer', { pile: 'a_filtrer' }));
  // Phase 3 : mails « répondre plus tard » encore masqués (affiché seulement s'il y en a).
  if (j.plusTard > 0) out.push(c(j.plusTard, 'plus tard', { pile: 'plus_tard' }));
  // Phase 5 : good@ (présent seulement pour ses membres autorisés).
  if (j.equipe) {
    out.push(c(j.equipe.aAttribuer, 'à attribuer (good@)', { pile: 'equipe_a_attribuer' }));
    out.push(c(j.equipe.pourMoi, 'pour toi sur good@', { pile: 'equipe_pour_moi' }));
  }
  return out;
}

function chipInner(x: Compteur): string {
  return `<span class="journee-n">${x.n}</span><span class="journee-l">${x.libelle}</span>`;
}

/** Texte du bandeau (exporté pour le plan d'essai et la cohérence avec le brief). */
export function formatJournee(j: AgentJournee): string {
  return compteurs(j).map(c => c.texte).join(' · ');
}

/**
 * Affiche / rafraîchit le bandeau dans `host`. Silencieux si la route n'est pas disponible.
 * `onInfo` : message court (toast) après une décision ou un mail impossible à ouvrir.
 */
export async function refreshJourneeBanner(host: HTMLElement | null, opts: Pick<MailListOptions, 'onInfo'> = {}): Promise<void> {
  if (!host) return;
  // Derniers compteurs connus affichés tout de suite (avant la connexion), puis mis à jour sur place.
  const pre = !host.querySelector('.journee-banner') ? journeeMemorisee() : null;
  if (pre) dessiner(host, pre, opts);
  const j = await fetchJournee();
  if (!j) {
    // Bandeau facultatif : absent si la route ne répond pas (jamais d'erreur ici).
    if (!host.querySelector('.journee-banner')) { host.innerHTML = ''; host.hidden = true; }
    return;
  }
  // Liste d'une pile ouverte : compteurs mis à jour sans la refermer ; sinon bandeau redessiné.
  const liste = host.querySelector<HTMLElement>('#journee-list');
  if (liste && !liste.hidden) { majCompteurs(host, j); return; }
  dessiner(host, j, opts);
}

function dessiner(host: HTMLElement, j: AgentJournee, opts: Pick<MailListOptions, 'onInfo'>): void {
  host.hidden = false;
  const cs = compteurs(j);
  host.innerHTML = `
    <section class="journee-banner" aria-label="Ma journée">
      <span class="journee-title">Ma journée</span>
      ${seancePossible() ? '<button type="button" class="journee-seance" id="journee-seance" title="Séance de tri : une carte par mail, l\'action est déjà préparée">Séance de tri</button>' : ''}
      <div class="journee-counts">${cs.map(x => x.pile
        ? `<button type="button" class="journee-count${x.alerte ? ' is-alert' : ''}" data-pile="${x.pile}" aria-expanded="false" aria-controls="journee-list" title="Voir la liste : ${x.libelle}">${chipInner(x)}</button>`
        : `<span class="journee-count is-static${x.alerte ? ' is-alert' : ''}">${chipInner(x)}</span>`).join('')}</div>
    </section>
    <div class="journee-list" id="journee-list" hidden></div>
  `;
  const list = host.querySelector<HTMLElement>('#journee-list')!;
  host.querySelector('#journee-seance')?.addEventListener('click', () => {
    // Tableau de bord dans ce volet, séance ouverte (src/taskpane.ts › ouvrirTableau).
    window.dispatchEvent(new CustomEvent('atlas:tableau', { detail: { seance: true } }));
  });
  let ouverte: AgentListePile | null = null;
  host.querySelectorAll<HTMLButtonElement>('button[data-pile]').forEach(btn => {
    btn.addEventListener('click', () => {
      const pile = btn.dataset.pile as AgentListePile;
      host.querySelectorAll<HTMLButtonElement>('button[data-pile]').forEach(b => {
        b.classList.remove('active');
        b.setAttribute('aria-expanded', 'false');
      });
      if (ouverte === pile) {
        ouverte = null;
        list.hidden = true;
        list.innerHTML = '';
        return;
      }
      ouverte = pile;
      btn.classList.add('active');
      btn.setAttribute('aria-expanded', 'true');
      list.hidden = false;
      renderMailList(list, pile, {
        limite: 50,
        onInfo: opts.onInfo,
        // Après Accepter / Refuser : compteurs relus (le cache a été invalidé), liste laissée ouverte.
        onChange: () => { updateCounts(host).catch(() => { /* bandeau facultatif */ }); },
      });
    });
  });
}

/** Met à jour le texte des compteurs cliquables sans refermer la liste ouverte. */
async function updateCounts(host: HTMLElement): Promise<void> {
  const j = await fetchJournee();
  if (j) majCompteurs(host, j);
}

function majCompteurs(host: HTMLElement, j: AgentJournee): void {
  for (const x of compteurs(j)) {
    const btn = x.pile ? host.querySelector<HTMLButtonElement>(`button[data-pile="${x.pile}"]`) : null;
    if (btn) { btn.innerHTML = chipInner(x); btn.classList.toggle('is-alert', !!x.alerte); }
  }
}
