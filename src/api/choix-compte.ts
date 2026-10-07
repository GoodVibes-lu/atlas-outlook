/**
 * choix-compte.ts · logique PURE (sans DOM ni réseau) du compte Microsoft utilisé par le tableau de
 * bord ATLAS (07/10/2026, retour de Charles sur Outlook iPhone : l'onglet « Applications › ATLAS »
 * recevait le jeton de la boîte d'équipe good@vibes.lu, ajoutée elle aussi dans Outlook, au lieu de
 * charles@vibes.lu ; le worker refusait « pas de fiche Employés active » et « Se connecter » ne
 * faisait rien, parce que l'authentification unique redonnait aussitôt le même compte).
 *
 * Règles :
 *   - un compte CHOISI par la personne (connexion interactive, « Changer de compte ») est retenu sur
 *     l'appareil (localStorage) ; un jeton automatique (connexion automatique, SSO, page parente)
 *     d'un AUTRE compte ne le remplace jamais ;
 *   - un compte refusé par le worker faute de fiche Employés active (boîte partagée, compte d'équipe)
 *     est mis de côté : ses jetons automatiques sont écartés, et la connexion interactive propose
 *     le choix du compte (prompt=select_account) ;
 *   - sans compte retenu, la connexion interactive propose toujours le choix du compte ; avec un
 *     compte retenu, elle le pré-remplit (login_hint) pour aller vite.
 *
 * Tests : scripts/test-choix-compte.mjs.
 */

/** Clé localStorage (par appareil) de l'état des comptes. */
export const CLE_COMPTES = 'atlas_tdb_compte';

/** Nombre maximal de comptes refusés mémorisés. */
const MAX_REFUSES = 6;

export interface EtatComptes {
  /** Compte choisi par la personne sur cet appareil ('' : aucun). */
  retenu: string;
  /** Comptes refusés par le worker (pas de fiche Employés active). */
  refuses: string[];
}

export const ETAT_VIDE: EtatComptes = { retenu: '', refuses: [] };

/** Adresse en minuscules, '' si ce n'est pas une adresse. */
export function normaliserCompte(compte: unknown): string {
  const c = String(compte ?? '').trim().toLowerCase();
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(c) ? c : '';
}

/** Lecture tolérante de l'état stocké (JSON abîmé ou absent : état vide). */
export function lireEtatComptes(brut: string | null | undefined): EtatComptes {
  if (!brut) return { ...ETAT_VIDE, refuses: [] };
  try {
    const o = JSON.parse(brut);
    const retenu = normaliserCompte(o?.retenu);
    const refuses = Array.isArray(o?.refuses)
      ? [...new Set(o.refuses.map(normaliserCompte).filter(Boolean) as string[])].filter(c => c !== retenu).slice(-MAX_REFUSES)
      : [];
    return { retenu, refuses };
  } catch {
    return { ...ETAT_VIDE, refuses: [] };
  }
}

export function ecrireEtatComptes(e: EtatComptes): string {
  return JSON.stringify({ retenu: normaliserCompte(e.retenu), refuses: e.refuses.map(normaliserCompte).filter(Boolean) });
}

export type DecisionJeton =
  | { ok: true }
  | { ok: false; raison: 'autre_compte_retenu' | 'compte_refuse'; compte: string };

/**
 * Un jeton du compte `compte` peut-il être utilisé ? `interactif` : il vient d'une connexion où la
 * personne a choisi son compte (toujours accepté : c'est son choix).
 */
export function deciderJeton(compte: string, etat: EtatComptes, interactif: boolean): DecisionJeton {
  const c = normaliserCompte(compte);
  if (interactif || !c) return { ok: true };
  if (etat.retenu && c !== etat.retenu) return { ok: false, raison: 'autre_compte_retenu', compte: c };
  if (etat.refuses.includes(c)) return { ok: false, raison: 'compte_refuse', compte: c };
  return { ok: true };
}

/** Après une connexion interactive réussie : le compte choisi est retenu (et n'est plus « refusé »). */
export function apresConnexionInteractive(etat: EtatComptes, compte: string): EtatComptes {
  const c = normaliserCompte(compte);
  if (!c) return etat;
  return { retenu: c, refuses: etat.refuses.filter(x => x !== c) };
}

/** Après un refus du worker (pas de fiche Employés active) : le compte est mis de côté. */
export function apresRefusWorker(etat: EtatComptes, compte: string): EtatComptes {
  const c = normaliserCompte(compte);
  if (!c) return etat;
  return {
    retenu: etat.retenu === c ? '' : etat.retenu,
    refuses: [...etat.refuses.filter(x => x !== c), c].slice(-MAX_REFUSES),
  };
}

/** Options de la connexion interactive (fenêtre de l'hôte, redirection). */
export interface OptionsConnexion {
  /** Proposer le choix du compte (prompt=select_account). */
  choix: boolean;
  /** Compte pré-rempli (login_hint), '' sinon. */
  indice: string;
}

/**
 * `changement` : la personne a touché « Changer de compte » (choix forcé, sans pré-remplissage).
 * Sinon : compte retenu pré-rempli ; sans compte retenu, choix proposé (sur un appareil où plusieurs
 * comptes sont ouverts, la session par défaut peut être celle d'une boîte d'équipe).
 */
export function optionsConnexion(etat: EtatComptes, changement: boolean): OptionsConnexion {
  if (changement) return { choix: true, indice: '' };
  if (etat.retenu) return { choix: false, indice: etat.retenu };
  return { choix: true, indice: '' };
}

/** Adresse de la fenêtre de connexion (`?auth=debut`, plus le choix et l'indice). */
export function urlFenetreConnexion(base: string, o: OptionsConnexion): string {
  const p = new URLSearchParams({ auth: 'debut' });
  if (o.choix) p.set('choix', '1');
  const indice = normaliserCompte(o.indice);
  if (indice) p.set('indice', indice);
  return `${base}?${p.toString()}`;
}

/** Relit les options passées à la fenêtre de connexion. */
export function lireOptionsFenetre(search: string): OptionsConnexion {
  const p = new URLSearchParams(search);
  return { choix: p.get('choix') === '1', indice: normaliserCompte(p.get('indice')) };
}

/** Message affiché quand le seul compte proposé automatiquement n'est pas utilisable. */
export function messageCompteEcarte(d: { compte: string; raison: 'autre_compte_retenu' | 'compte_refuse' | 'employe_inactif'; retenu?: string; mobile?: boolean }): string {
  const c = normaliserCompte(d.compte) || 'ce compte';
  const base = d.raison === 'autre_compte_retenu' && d.retenu
    ? `Outlook propose le compte ${c}, mais ce tableau de bord est relié à ${d.retenu} sur cet appareil. Touche « Se connecter » pour reprendre ${d.retenu}, ou « Changer de compte ».`
    : `Le compte ${c} n'a pas de fiche Employés active dans ATLAS (boîte d'équipe ou partagée ?). Touche « Choisir mon compte » et choisis ton compte personnel @vibes.lu : il sera retenu sur cet appareil.`;
  return d.mobile
    ? `${base} Si la fenêtre de connexion ne s'ouvre pas, touche « Ouvrir dans le navigateur » : la même page s'ouvre dans Safari, où tu choisis ton compte.`
    : base;
}
