// COPIE de src/utils/mail-folder-tree.ts (le complément est publié seul : il ne peut pas importer hors de
// son dossier). Ne pas modifier ici : modifier l'original puis recopier (cp). Identité vérifiée par
// scripts/test-inbox-dossiers.mjs.
/**
 * mail-folder-tree.ts · ARBORESCENCE COMPLÈTE des dossiers Outlook d'une boîte et rangement par
 * projet / mandat (PUR : aucun réseau ici, la lecture d'une page Graph est passée en paramètre).
 *
 * Demande de Charles du 07/10/2026 : ses dossiers vont bien au-delà de deux niveaux (ex.
 * « Projets/2026/755 Gala ») ; toutes les lectures de dossiers (worker, complément Outlook, ATLAS)
 * lisent désormais l'arbre ENTIER : tous les niveaux, pages Graph suivies (@odata.nextLink), sans
 * plafond arbitraire. Seul garde-fou : PLAFOND_DOSSIERS (5 000, arbre emballé), signalé quand il est
 * atteint (`tronque` + rappel `surPlafond`).
 *
 * Utilisé par :
 *  - worker/inbox-agent/dossiers-outlook.ts (jeton d'application, cache, création sous double verrou) ;
 *  - src/services/mail-folders-tree.service.ts (ATLAS, jeton délégué) ;
 *  - outlook-addin/src/api/graph.ts (complément, jeton délégué).
 *
 * Nommage d'un dossier créé : MÊME convention que l'Inbox ATLAS (InboxWidget, zone STABLE non
 * modifiée ; logique recopiée ici à l'identique) :
 *  - « Classer dans Outlook » (FileEmailModal › handleCreateSuggested) : client connu →
 *    « Clients/<Client>/#<n°> <Nom du projet> » (niveaux existants repris, casse comprise) ;
 *  - rangement par projet sans client (move-mail-to-project-folder.ts) : « <n°> - <Nom> » sous le
 *    dossier racine « Projets » s'il existe, sinon à la racine ;
 *  - mandat / association : « Mandats/<Nom> » (forme reconnue par inbox-mandats › dossierMandatPour).
 */

/** Garde-fou contre un arbre emballé (jamais atteint par une boîte normale) : signalé, jamais muet. */
export const PLAFOND_DOSSIERS = 5000;

export interface DossierArbre {
  id: string;
  /** Nom affiché du dossier. */
  nom: string;
  /** Chemin complet « Parent/Enfant/Sous-enfant ». */
  chemin: string;
  parentId?: string;
  /** 0 = dossier racine (enfant de la racine de la boîte). */
  profondeur: number;
}

export interface PageDossiers {
  value?: Array<{ id?: string; Id?: string; displayName?: string; DisplayName?: string; childFolderCount?: number; ChildFolderCount?: number }>;
  '@odata.nextLink'?: string;
}

export interface OptionsParcours {
  /** URL (ou chemin) de la première page des dossiers racine. */
  racine: string;
  /** URL (ou chemin) de la première page des sous-dossiers d'un dossier. */
  enfants: (id: string) => string;
  /** Lecture d'une page (URL complète d'un nextLink comprise). Une erreur sur la racine remonte. */
  lirePage: (url: string) => Promise<PageDossiers>;
  plafond?: number;
  /** Appelé une fois si le plafond est atteint (journal de l'appelant). */
  surPlafond?: (n: number) => void;
  /** Appelé quand un sous-arbre est illisible (le reste est lu). */
  surErreur?: (chemin: string, e: unknown) => void;
}

/**
 * Parcourt TOUTE l'arborescence (profondeur quelconque, pages suivies). Ordre : parent avant ses
 * enfants (profondeur d'abord). Un dossier déjà vu (cycle, doublon de page) n'est jamais relu.
 * `childFolderCount` absent : les enfants sont demandés quand même (jamais un niveau oublié).
 */
export async function parcourirArborescence(o: OptionsParcours): Promise<{ dossiers: DossierArbre[]; tronque: boolean }> {
  const plafond = o.plafond && o.plafond > 0 ? o.plafond : PLAFOND_DOSSIERS;
  const out: DossierArbre[] = [];
  const vus = new Set<string>();
  let tronque = false;
  const signalerPlafond = () => { if (!tronque) { tronque = true; o.surPlafond?.(plafond); } };

  type Noeud = DossierArbre & { aEnfants: boolean };
  const lireNiveau = async (premiere: string, parent: DossierArbre | null): Promise<Noeud[]> => {
    const niveau: Noeud[] = [];
    const pagesVues = new Set<string>();
    let url: string | undefined = premiere;
    while (url && !pagesVues.has(url)) {
      pagesVues.add(url);
      const page: PageDossiers = await o.lirePage(url);
      for (const d of page.value || []) {
        const id = String(d.id || d.Id || '');
        if (!id || vus.has(id)) continue;
        if (out.length >= plafond) { signalerPlafond(); return niveau; }
        vus.add(id);
        const nom = String(d.displayName ?? d.DisplayName ?? '');
        const n = d.childFolderCount ?? d.ChildFolderCount;
        const x: Noeud = { id, nom, chemin: parent ? `${parent.chemin}/${nom}` : nom, ...(parent ? { parentId: parent.id } : {}), profondeur: parent ? parent.profondeur + 1 : 0, aEnfants: n === undefined || n === null || Number(n) > 0 };
        out.push({ id: x.id, nom: x.nom, chemin: x.chemin, ...(x.parentId ? { parentId: x.parentId } : {}), profondeur: x.profondeur });
        niveau.push(x);
      }
      url = page['@odata.nextLink'];
    }
    return niveau;
  };

  const descendre = async (niveau: Noeud[]): Promise<void> => {
    for (const d of niveau) {
      if (tronque) return;
      if (d.aEnfants === false) continue;
      let enfants: Noeud[] = [];
      try { enfants = await lireNiveau(o.enfants(d.id), d); }
      catch (e) { o.surErreur?.(d.chemin, e); continue; } // sous-arbre illisible : le reste est lu
      await descendre(enfants);
    }
  };

  const racines = await lireNiveau(o.racine, null);
  await descendre(racines);
  return { dossiers: out, tronque };
}

// ── Normalisation ─────────────────────────────────────────────────────────────────

export const normNom = (s: string | null | undefined): string => String(s || '')
  .normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();

/** Noms de la boîte de réception selon la langue de la boîte. */
const NOMS_BOITE_RECEPTION = new Set(['inbox', 'boite de reception', 'posteingang', 'postvak in']);
export const estBoiteDeReception = (nom: string): boolean => NOMS_BOITE_RECEPTION.has(normNom(nom));

/** Dossiers système (pas un rangement choisi par la personne). */
const SYSTEME = /^(boite de reception|inbox|posteingang|postvak in|elements envoyes|sent items|gesendete elemente|elements supprimes|deleted items|geloschte elemente|brouillons|drafts|entwurfe|courrier indesirable|junk e-?mail|junk-e-mail|boite d'envoi|outbox|postausgang|historique des conversations|conversation history|flux rss|rss feeds|rss|archive|archiv|archives|sync issues|problemes de synchronisation|notes|taches|tasks|calendrier|calendar|contacts|journal)$/;
/** Dossiers système dont les SOUS-dossiers sont aussi exclus (corbeille, indésirables, brouillons…). */
const SYSTEME_FERME = /^(elements envoyes|sent items|gesendete elemente|elements supprimes|deleted items|geloschte elemente|brouillons|drafts|entwurfe|courrier indesirable|junk e-?mail|junk-e-mail|boite d'envoi|outbox|postausgang|historique des conversations|conversation history|flux rss|rss feeds|rss|sync issues|problemes de synchronisation)$/;

export const estDossierSysteme = (nom: string): boolean => SYSTEME.test(normNom(nom));

/**
 * Dossiers où une personne RANGE ses mails : tout l'arbre, sauf les dossiers système eux-mêmes et
 * les sous-arbres fermés (corbeille, indésirables, brouillons, envoyés…). Les sous-dossiers de la
 * boîte de réception et des archives sont gardés (beaucoup de rangements y vivent).
 */
export function dossiersDeRangement(dossiers: DossierArbre[]): DossierArbre[] {
  const exclus = new Set<string>();
  const parId = new Map(dossiers.map(d => [d.id, d]));
  const fermeParAncetre = (d: DossierArbre): boolean => {
    let p = d.parentId ? parId.get(d.parentId) : undefined;
    while (p) {
      if (exclus.has(p.id) || (p.profondeur === 0 && SYSTEME_FERME.test(normNom(p.nom)))) return true;
      p = p.parentId ? parId.get(p.parentId) : undefined;
    }
    return false;
  };
  const out: DossierArbre[] = [];
  for (const d of dossiers) {
    if (fermeParAncetre(d)) { exclus.add(d.id); continue; }
    if (d.profondeur === 0 && estDossierSysteme(d.nom)) continue;
    out.push(d);
  }
  return out;
}

/** Recherche par mots (tous présents dans le chemin), tri par profondeur puis chemin. */
export function rechercherDossiers(dossiers: DossierArbre[], recherche?: string | null, max = 50): DossierArbre[] {
  const mots = normNom(recherche).split(' ').filter(Boolean);
  const hits = mots.length ? dossiers.filter(d => { const c = normNom(d.chemin); return mots.every(m => c.includes(m)); }) : dossiers;
  return [...hits].sort((a, b) => a.profondeur - b.profondeur || a.chemin.localeCompare(b.chemin, 'fr')).slice(0, max);
}

// ── Dossier d'un projet (où qu'il soit dans l'arbre) ──────────────────────────────

export interface ProjetPourDossier { numero?: number | string | null; nom?: string | null; clientNom?: string | null }

const numeroDe = (n: ProjetPourDossier['numero']): string => {
  const s = String(n ?? '').replace(/[^\d]/g, '').replace(/^0+/, '');
  return /^\d{1,6}$/.test(s) ? s : '';
};

/** Le nom (dernier segment) porte-t-il le numéro de projet (« #755 Gala », « 755 - Gala », « Gala 755 ») ? */
export function nomPorteNumero(nom: string, numero: number | string): boolean {
  const n = numeroDe(numero);
  if (!n) return false;
  return new RegExp(`(^|[^0-9])#?\\s?0*${n}([^0-9]|$)`).test(nom);
}

/** Nom du projet sans le numéro ni la ponctuation de tête (« #755 - Gala ESA » → « gala esa »). */
const nomSansNumero = (nom: string): string => normNom(nom).replace(/^#?\s*\d{1,6}\s*[-–—:._]*\s*/, '').trim();

const motsUtiles = (s: string): string[] => normNom(s).split(/[^a-z0-9]+/).filter(m => m.length >= 3);

export interface DossierProjetTrouve {
  dossier: DossierArbre;
  /** 'numero' : n° de projet dans le nom (le plus sûr) ; 'nom' : nom du projet seul. */
  par: 'numero' | 'nom';
}

/**
 * Dossier EXISTANT du projet, n'importe où dans l'arbre. Le numéro de projet prime (dossier dont le
 * NOM porte le n°) ; plusieurs dossiers portent le n° : celui dont le nom reprend aussi le nom du
 * projet, puis celui dont le nom COMMENCE par le n° ; égalité restante = ambigu (null, jamais deviné).
 * Sans dossier au n° : nom du projet (≥ 5 lettres), identique au nom du dossier sans n° ; unique.
 */
export function trouverDossierProjet(dossiers: DossierArbre[], p: ProjetPourDossier): DossierProjetTrouve | null {
  const n = numeroDe(p.numero);
  const nomProjet = nomSansNumero(String(p.nom || ''));
  const mots = motsUtiles(nomProjet).filter(m => !/^\d+$/.test(m));
  if (n) {
    const hits = dossiers.filter(d => nomPorteNumero(d.nom, n) && !estDossierSysteme(d.nom));
    if (hits.length === 1) return { dossier: hits[0], par: 'numero' };
    if (hits.length > 1) {
      const score = (d: DossierArbre): number => {
        const nm = normNom(d.nom);
        let s = 0;
        if (mots.length && mots.some(m => nm.includes(m))) s += 4;
        if (new RegExp(`^#?\\s?0*${n}([^0-9]|$)`).test(nm)) s += 2;
        if (nm.replace(/[#\s]/g, '') === n) s -= 1; // dossier qui ne porte QUE le n° (ex. année) : moins sûr
        return s;
      };
      const tries = hits.map(d => ({ d, s: score(d) })).sort((a, b) => b.s - a.s);
      if (tries[0].s > tries[1].s) return { dossier: tries[0].d, par: 'numero' };
      return null;
    }
  }
  if (nomProjet.length >= 5) {
    const hits = dossiers.filter(d => !estDossierSysteme(d.nom) && (normNom(d.nom) === nomProjet || nomSansNumero(d.nom) === nomProjet));
    if (hits.length === 1) return { dossier: hits[0], par: 'nom' };
  }
  return null;
}

// ── Nommage (convention de l'Inbox ATLAS) ─────────────────────────────────────────

/** Nom de dossier sûr : sans « / » (séparateur de chemin), sans caractères de contrôle, 120 caractères. */
export function nomDossierSur(s: string | null | undefined): string {
  return String(s || '').replace(/[\x00-\x1f\x7f]/g, ' ').replace(/[\\/]/g, '-').replace(/\s+/g, ' ').trim().slice(0, 120);
}

/** « #755 Gala ESA » (FileEmailModal › handleCreateSuggested). */
export function nomDossierProjet(p: ProjetPourDossier): string {
  const n = numeroDe(p.numero);
  const nom = nomDossierSur(p.nom) || 'Projet';
  return n ? `#${n} ${nom}` : nom;
}

/** Racine existante (profondeur 0) qui porte ce nom (casse et accents ignorés), sinon le nom donné. */
const racineExistante = (dossiers: DossierArbre[], noms: RegExp, defaut: string): string =>
  dossiers.find(d => d.profondeur === 0 && noms.test(normNom(d.nom)))?.nom || defaut;

/**
 * Chemin du dossier à CRÉER pour un projet, à la même place et avec le même nom que l'Inbox ATLAS :
 *  - client connu : « Clients/<Client>/#<n°> <Nom> » ; si le dossier du client existe déjà sous
 *    « Clients », sa casse est reprise, et un dossier de projet déjà là sous l'ancien nom (sans n°)
 *    est réutilisé (comme handleCreateSuggested) ;
 *  - sans client : « Projets/<n°> - <Nom> » si la racine « Projets » existe, sinon « <n°> - <Nom> ».
 */
export function cheminDossierProjetPropose(p: ProjetPourDossier, dossiers: DossierArbre[] = []): string {
  const client = nomDossierSur(p.clientNom);
  const n = numeroDe(p.numero);
  if (client) {
    const racine = racineExistante(dossiers, /^clients$/, 'Clients');
    const rac = dossiers.find(d => d.profondeur === 0 && normNom(d.nom) === normNom(racine));
    const dossierClient = rac ? dossiers.find(d => d.parentId === rac.id && normNom(d.nom) === normNom(client)) : undefined;
    const nomClient = dossierClient?.nom || client;
    const nomProjet = nomDossierProjet(p);
    if (dossierClient) {
      const existant = dossiers.find(d => d.parentId === dossierClient.id && (normNom(d.nom) === normNom(nomProjet) || normNom(d.nom) === normNom(p.nom)));
      if (existant) return existant.chemin;
    }
    return `${racine}/${nomClient}/${nomProjet}`;
  }
  const nom = nomDossierSur(p.nom) || 'Projet';
  const feuille = n ? `${n} - ${nom}` : nom;
  const projets = dossiers.find(d => d.profondeur === 0 && /^projets?$/.test(normNom(d.nom)));
  return projets ? `${projets.nom}/${feuille}` : feuille;
}

/** « Mandats/<Nom> » (racine « Mandats » / « Mandat » existante reprise). */
export function cheminDossierMandatPropose(nom: string, dossiers: DossierArbre[] = []): string {
  return `${racineExistante(dossiers, /^mandats?$/, 'Mandats')}/${nomDossierSur(nom) || 'Mandat'}`;
}

/** Segments d'un chemin (vides retirés, noms sûrs). */
export const segmentsChemin = (chemin: string): string[] => String(chemin || '').split('/').map(s => nomDossierSur(s)).filter(Boolean);

export interface PlanCreation {
  /** Dossier final s'il existe déjà (rien à créer). */
  existant: DossierArbre | null;
  /** Dernier niveau existant du chemin (parent des créations), null = racine de la boîte. */
  parent: DossierArbre | null;
  /** Noms à créer dans l'ordre, sous `parent`. */
  aCreer: string[];
}

/**
 * Plan de création d'un chemin dans l'arbre (pur) : niveaux existants repris (casse et accents
 * ignorés ; 1er segment « Inbox » = boîte de réception quelle que soit sa langue), le reste à créer.
 */
export function planCreationChemin(dossiers: DossierArbre[], chemin: string): PlanCreation {
  const segs = segmentsChemin(chemin);
  let parent: DossierArbre | null = null;
  for (let i = 0; i < segs.length; i++) {
    const seg = segs[i];
    const freres = dossiers.filter(d => (parent ? d.parentId === parent.id : d.profondeur === 0));
    let hit = freres.find(d => normNom(d.nom) === normNom(seg));
    if (!hit && i === 0 && estBoiteDeReception(seg)) hit = freres.find(d => estBoiteDeReception(d.nom));
    if (!hit) return { existant: null, parent, aCreer: segs.slice(i) };
    parent = hit;
  }
  return { existant: parent, parent, aCreer: [] };
}

/** Même recherche sur une liste « { id, chemin } » (forme de l'agent) : renvoie { id, chemin } ou null. */
export function dossierProjetDansListe(dossiers: Array<{ id: string; chemin: string }>, p: ProjetPourDossier): { id: string; chemin: string } | null {
  const arbre: DossierArbre[] = dossiers.map(d => { const segs = d.chemin.split('/'); return { id: d.id, nom: segs[segs.length - 1] || d.chemin, chemin: d.chemin, profondeur: segs.length - 1 }; });
  const t = trouverDossierProjet(arbre, p);
  return t ? { id: t.dossier.id, chemin: t.dossier.chemin } : null;
}
