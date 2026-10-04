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

import { fetchJournee, type AgentJournee, type AgentListePile } from '../api/agent';
import { renderMailList, type MailListOptions } from './agent-lists';

function plural(n: number, one: string, many: string): string {
  return `${n} ${n > 1 ? many : one}`;
}

interface Compteur { texte: string; pile?: AgentListePile }

function compteurs(j: AgentJournee): Compteur[] {
  const out: Compteur[] = [
    { texte: `${j.aTraiter} à traiter`, pile: 'a_traiter' },
    { texte: `${j.enAttente} en attente`, pile: 'en_attente' },
    { texte: plural(j.relancesDues, 'relance due', 'relances dues') },
    { texte: plural(j.clientsPlus48h, 'client > 48 h', 'clients > 48 h') },
  ];
  if (j.aFiltrer > 0) out.push({ texte: `${j.aFiltrer} à filtrer`, pile: 'a_filtrer' });
  // Phase 3 : mails « répondre plus tard » encore masqués (affiché seulement s'il y en a).
  if (j.plusTard > 0) out.push({ texte: `${j.plusTard} plus tard`, pile: 'plus_tard' });
  // Phase 5 : good@ (présent seulement pour ses membres autorisés).
  if (j.equipe) {
    out.push({ texte: `${j.equipe.aAttribuer} à attribuer (good@)`, pile: 'equipe_a_attribuer' });
    out.push({ texte: `${j.equipe.pourMoi} pour toi sur good@`, pile: 'equipe_pour_moi' });
  }
  return out;
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
  const j = await fetchJournee();
  if (!j) {
    host.innerHTML = '';
    host.style.display = 'none';
    return;
  }
  host.style.display = '';
  const cs = compteurs(j);
  host.innerHTML = `
    <div class="journee-banner" role="status" aria-label="Ma journée">
      <span class="journee-title">Ma journée</span>
      <span class="journee-counts">${cs.map(c => c.pile
        ? `<button type="button" class="journee-count" data-pile="${c.pile}" aria-expanded="false" title="Voir la liste">${c.texte}</button>`
        : `<span class="journee-count-static">${c.texte}</span>`).join('<span class="journee-sep"> · </span>')}</span>
    </div>
    <div class="journee-list" id="journee-list" hidden></div>
  `;
  const list = host.querySelector<HTMLElement>('#journee-list')!;
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
  if (!j) return;
  for (const c of compteurs(j)) {
    const btn = c.pile ? host.querySelector<HTMLButtonElement>(`button[data-pile="${c.pile}"]`) : null;
    if (btn) btn.textContent = c.texte;
  }
}
