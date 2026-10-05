/**
 * agent.ts — Routes worker de l'agent d'inbox et du complément fusionné (phase 1.5).
 *
 *  • `/api/plugin/agent/*` : état calculé par l'agent serveur (aucun appel IA à l'ouverture d'un
 *    mail : le serveur a déjà lu le mail, le panneau ne fait qu'afficher) ; phase 2 : listes par
 *    pile (`liste`), filtre des nouveaux expéditeurs (`expediteur`), journal et Annuler
 *    (`journal`, `annuler`).
 *    Phase 6 : traduction, modèles de réponse, fiche contact personnelle, détections, pièces
 *    jointes lourdes et lien court (routes-memoire.ts côté worker), toujours à la demande.
 *  • `/api/plugin/email/*` : fonctions reprises du complément « ATLAS Assistant » (outlook-plugin,
 *    remplacé) : actions suggérées (sans IA), réponse rapide (modèle fixe, sans IA), brouillon
 *    automatique (IA côté worker, Agency Brain, sur clic seulement).
 *
 * Même worker et même jeton Microsoft que `callAtlasWorker` (worker.ts, dont l'obtention du jeton
 * relève de la tâche « auth Entra ») : ce module ne détient aucun secret et ne gère pas le jeton
 * lui-même, il le demande à worker.ts (`getWorkerToken`).
 */

import { workerRequest } from './worker';
import { AtlasError, humanError } from './net';
import type {
  InboxMessageState, InboxJournee, InboxPile, InboxRappelReponse, InboxEquipe, InboxEquipeVue, InboxEquipeCommentaire,
} from './inbox-agent.types';
import type {
  ActionProposee, ActionFaite, ReponseActions, ResultatExecution, ReponseQuestion, ResumeFil, Rattrapage,
} from './inbox-actions.types';

/** Erreur d'appel au worker : message lisible, `status` HTTP (0 sans réponse) et corps `data`. */
const PluginHttpError = AtlasError;

/**
 * Appelle `/api/plugin/<path>` du worker (GET, POST ou DELETE JSON) : jeton, délai, erreurs
 * lisibles et nouvelles tentatives des lectures (GET) sont gérés par workerRequest (worker.ts).
 * Les routes qui font travailler l'IA côté worker (brouillon, question, traduction…) ont 90 s.
 */
async function pluginFetch<T>(method: 'GET' | 'POST' | 'DELETE', path: string, payload?: unknown): Promise<T> {
  return workerRequest<T>(method, path, payload, { retry: method === 'GET', timeoutMs: 90_000 });
}

// ── État d'un message ──

export type AgentStateResult =
  | { kind: 'state'; state: InboxMessageState }
  | { kind: 'pending' }
  | { kind: 'absent' }
  | { kind: 'error'; message: string };

/**
 * Lit l'état calculé par l'agent pour un message (clé : internetMessageId). Lecture seule, sans IA.
 * Le worker répond HTTP 404 `{ status: 'pending' }` tant que le message n'a pas été lu.
 * `absent` = route pas encore déployée (404 sans corps « pending ») ou réponse inattendue.
 */
export async function fetchAgentState(messageId: string, mailbox: string): Promise<AgentStateResult> {
  if (!messageId) return { kind: 'absent' };
  const qs = `messageId=${encodeURIComponent(messageId)}&mailbox=${encodeURIComponent(mailbox || '')}`;
  try {
    const data = await pluginFetch<any>('GET', `agent/state?${qs}`);
    if (data?.status === 'pending' || data?.state?.status === 'pending') return { kind: 'pending' };
    const state = (data?.state && typeof data.state === 'object' ? data.state : data) as InboxMessageState;
    if (state && typeof state.pile === 'string') return { kind: 'state', state };
    return { kind: 'absent' };
  } catch (e) {
    if (e instanceof PluginHttpError && e.status === 404) {
      return e.data?.status === 'pending' ? { kind: 'pending' } : { kind: 'absent' };
    }
    return { kind: 'error', message: humanError(e) };
  }
}

// ── Bandeau « Ma journée » ──

const JOURNEE_MIN_INTERVAL = 60 * 1000; // au plus une fois par minute
let journeeCache: { data: AgentJournee | null; ts: number } | null = null;
let journeeInFlight: Promise<AgentJournee | null> | null = null;

/**
 * Compteurs du bandeau, avec la pile « À filtrer » (phase 2), « Plus tard » (phase 3, 0 si absent)
 * et, pour les membres de good@ (phase 5), « à attribuer » / « pour toi sur good@ » (absent sinon).
 */
export type AgentJournee = InboxJournee & { aFiltrer: number; plusTard: number };

/** Oublie les compteurs mémorisés (après une décision qui les change : accepter / refuser, annuler). */
export function invalidateJournee(): void {
  journeeCache = null;
}

/**
 * Compteurs de la journée. Rafraîchis au plus une fois par minute (ouverture du panneau et
 * changement de mail) : entre deux, renvoie la dernière valeur connue.
 */
export async function fetchJournee(): Promise<AgentJournee | null> {
  if (journeeCache && Date.now() - journeeCache.ts < JOURNEE_MIN_INTERVAL) return journeeCache.data;
  if (journeeInFlight) return journeeInFlight;
  journeeInFlight = (async () => {
    try {
      const data = await pluginFetch<any>('GET', 'agent/journee');
      const src = data?.journee && typeof data.journee === 'object' ? data.journee : data;
      const n = (v: unknown) => (typeof v === 'number' && Number.isFinite(v) ? v : 0);
      const j: AgentJournee = {
        aTraiter: n(src?.aTraiter),
        enAttente: n(src?.enAttente),
        relancesDues: n(src?.relancesDues),
        clientsPlus48h: n(src?.clientsPlus48h),
        pourInfo: n(src?.pourInfo),
        bruit: n(src?.bruit),
        aFiltrer: n(src?.aFiltrer),
        plusTard: n(src?.plusTard),
        ...(src?.equipe && typeof src.equipe === 'object'
          ? { equipe: { aAttribuer: n(src.equipe.aAttribuer), pourMoi: n(src.equipe.pourMoi), boites: Array.isArray(src.equipe.boites) ? src.equipe.boites.map(String) : [] } }
          : {}),
      };
      journeeCache = { data: j, ts: Date.now() };
      return j;
    } catch (e) {
      console.warn('[agent] journée indisponible :', (e as Error).message);
      // On mémorise aussi l'échec pour ne pas réessayer à chaque mail.
      journeeCache = { data: journeeCache?.data ?? null, ts: Date.now() };
      return journeeCache.data;
    } finally {
      journeeInFlight = null;
    }
  })();
  return journeeInFlight;
}

// ── Phase 2 : listes par pile, nouveaux expéditeurs, journal et Annuler ──

/**
 * Piles listables : les 4 piles du tri + « À filtrer » (premier mail d'un expéditeur inconnu)
 * + « Plus tard » (phase 3 : mails masqués jusqu'à une date, liste seulement).
 */
export type AgentPile = InboxPile | 'a_filtrer' | 'plus_tard';

/** Listes demandables : les piles + (phase 5, membres de good@) « à attribuer » et « pour toi sur good@ ». */
export type AgentListePile = AgentPile | 'equipe_a_attribuer' | 'equipe_pour_moi';

/** Un mail d'une pile (GET /api/plugin/agent/liste). */
export interface AgentListeElement {
  messageId: string;
  graphId: string;
  from: { email: string; name: string };
  subject: string;
  receivedAt: string;
  pile: AgentPile;
  resume: string;
  urgence: number;
  categorie: string;
  webLink?: string;
  /** Pile « plus tard » : remis en tête de « à traiter » à cette date (ISO). */
  plusTardJusqua?: string;
  /** Un brouillon de l'agent est prêt dans le fil. */
  brouillonPret?: boolean;
  /** « En attente » : relance prévue ce jour-là ('YYYY-MM-DD') sans réponse. */
  relanceLe?: string;
  /** Phase 5 (listes de good@) : boîte du mail, assigné / preneur. */
  mailbox?: string;
  equipe?: Pick<InboxEquipe, 'assigneA' | 'assigneNom' | 'prisPar' | 'prisParNom'>;
}

/** Une action de l'agent (GET /api/plugin/agent/journal). */
export interface AgentJournalAction {
  id: string;
  type: string;
  libelle: string;
  mail?: { subject?: string; from?: string | { email?: string; name?: string } };
  date: string;
  annulable: boolean;
}

/** Mails d'une pile, du plus récent au plus ancien (lecture seule, sans IA). */
export async function fetchListe(pile: AgentListePile, limite = 50): Promise<AgentListeElement[]> {
  const qs = `pile=${encodeURIComponent(pile)}&limite=${Math.max(1, Math.min(200, Math.round(limite)))}`;
  const r = await pluginFetch<{ elements?: AgentListeElement[] }>('GET', `agent/liste?${qs}`);
  return Array.isArray(r.elements) ? r.elements.filter(e => e && typeof e === 'object') : [];
}

/**
 * Filtre des nouveaux expéditeurs : accepter (les mails suivants arrivent normalement) ou refuser
 * (les mails suivants vont au bruit), une fois pour toutes. Les compteurs sont relus ensuite.
 */
export async function decideExpediteur(email: string, decision: 'accepter' | 'refuser'): Promise<void> {
  await pluginFetch('POST', 'agent/expediteur', { email: String(email || '').trim().toLowerCase(), decision });
  invalidateJournee();
}

/** Ce que l'agent a fait récemment (plus récent d'abord). */
export async function fetchJournal(limite = 30): Promise<AgentJournalAction[]> {
  const r = await pluginFetch<{ actions?: AgentJournalAction[]; journal?: AgentJournalAction[] }>('GET', `agent/journal?limite=${Math.max(1, Math.min(100, Math.round(limite)))}`);
  const liste = Array.isArray(r.actions) ? r.actions : Array.isArray(r.journal) ? r.journal : [];
  return liste.filter(a => a && typeof a === 'object' && a.id);
}

/** Annule une action de l'agent (déplacement, catégorie, filtre…). */
export async function annulerAction(actionId: string): Promise<void> {
  await pluginFetch('POST', 'agent/annuler', { actionId });
  invalidateJournee();
}

// ── Phase 3 : « répondre plus tard », « relancer si pas de réponse » ──

/** `quand` : 'demain' | 'lundi' | date 'AAAA-MM-JJ' ou ISO (jours ouvrés, 8 h, 120 jours au plus). */
export async function mettrePlusTard(messageId: string, quand: string): Promise<InboxRappelReponse> {
  const r = await pluginFetch<InboxRappelReponse>('POST', 'agent/plus-tard', { messageId, quand });
  invalidateJournee();
  return r;
}

/** Le mail revient tout de suite dans « à traiter ». */
export async function retirerPlusTard(messageId: string): Promise<InboxRappelReponse> {
  const r = await pluginFetch<InboxRappelReponse>('DELETE', 'agent/plus-tard', { messageId });
  invalidateJournee();
  return r;
}

/** Mail ENVOYÉ : relance préparée après `delaiJours` jours ouvrés (1 à 30) sans réponse. */
export async function demanderRelance(messageId: string, delaiJours: number): Promise<InboxRappelReponse> {
  const d = Math.max(1, Math.min(30, Math.round(delaiJours)));
  const r = await pluginFetch<InboxRappelReponse>('POST', 'agent/relancer', { messageId, delaiJours: d });
  invalidateJournee();
  return r;
}

/** Mail ENVOYÉ : « pas de relance » (ni manuelle ni automatique J+3 / J+5). */
export async function annulerRelance(messageId: string): Promise<InboxRappelReponse> {
  const r = await pluginFetch<InboxRappelReponse>('DELETE', 'agent/relancer', { messageId });
  invalidateJournee();
  return r;
}

// ── Phase 4 : « Que faire de ce mail ? » ──

/** Actions proposées (détection sans IA côté worker) et actions déjà faites sur ce mail. */
export async function fetchActions(messageId: string, mailbox: string): Promise<ReponseActions> {
  const qs = `messageId=${encodeURIComponent(messageId)}&mailbox=${encodeURIComponent(mailbox || '')}`;
  const r = await pluginFetch<Partial<ReponseActions>>('GET', `agent/actions?${qs}`);
  const ok = (a: unknown) => !!a && typeof a === 'object';
  return {
    actions: Array.isArray(r.actions) ? (r.actions.filter(ok) as ActionProposee[]) : [],
    faites: Array.isArray(r.faites) ? (r.faites.filter(ok) as ActionFaite[]) : [],
  };
}

/**
 * Exécute une action proposée. Un refus métier (`ok: false`, par ex. `erreur: 'choix-requis'`)
 * est RENVOYÉ (pas levé) pour que le panneau puisse demander le choix ; les erreurs techniques
 * (réseau, 5xx, jeton) sont levées.
 */
export async function executerAction(args: {
  messageId: string; mailbox?: string; type: string; donnees?: Record<string, unknown>; choix?: Record<string, unknown>;
}): Promise<ResultatExecution> {
  try {
    const r = await pluginFetch<ResultatExecution>('POST', 'agent/actions/executer', args);
    invalidateJournee();
    return r;
  } catch (e) {
    if (e instanceof PluginHttpError && e.status < 500 && e.data && typeof e.data.resume === 'string' && typeof e.data.erreur === 'string') {
      return { ...e.data, ok: false } as ResultatExecution;
    }
    throw e;
  }
}

// ── Phase 5 : good@ et travail d'équipe (membres autorisés de la boîte partagée) ──

/**
 * Équipe d'un mail de good@ (assigné, pris par, commentaires internes, collision, membres).
 * null = mail hors boîte partagée, ou personne non membre (le bloc Équipe reste masqué).
 * `mailbox` facultatif : sans lui, le worker cherche dans les boîtes partagées de la personne.
 */
export async function fetchEquipe(messageId: string, mailbox?: string): Promise<InboxEquipeVue | null> {
  if (!messageId) return null;
  const qs = `messageId=${encodeURIComponent(messageId)}${mailbox ? `&mailbox=${encodeURIComponent(mailbox)}` : ''}`;
  try {
    const r = await pluginFetch<InboxEquipeVue>('GET', `agent/equipe?${qs}`);
    return r && typeof r === 'object' && r.equipe ? r : null;
  } catch (e) {
    if (e instanceof PluginHttpError && (e.status === 404 || e.status === 403)) return null;
    throw e;
  }
}

export interface ReponseEquipe { ok: boolean; equipe?: InboxEquipe; error?: string }

/**
 * « Je prends » / « Relâcher » / « Attribuer à… ». Un refus (409 : déjà pris par quelqu'un d'autre)
 * est RENVOYÉ avec l'équipe actuelle (pour afficher « Pris par X »), pas levé.
 */
async function actionEquipe(method: 'POST' | 'DELETE', path: string, payload: Record<string, unknown>): Promise<ReponseEquipe> {
  try {
    const r = await pluginFetch<ReponseEquipe>(method, path, payload);
    invalidateJournee();
    return r;
  } catch (e) {
    if (e instanceof PluginHttpError && e.status === 409 && e.data?.equipe) return { ok: false, equipe: e.data.equipe, error: String(e.data.error || '') };
    throw e;
  }
}

export const prendreMail = (messageId: string, mailbox: string) => actionEquipe('POST', 'agent/equipe/prendre', { messageId, mailbox });
export const relacherMail = (messageId: string, mailbox: string) => actionEquipe('DELETE', 'agent/equipe/prendre', { messageId, mailbox });
export const attribuerMail = (messageId: string, mailbox: string, a: string) => actionEquipe('POST', 'agent/equipe/attribuer', { messageId, mailbox, a });

/** Commentaire interne (jamais dans le mail, invisible du client) : renvoie la liste à jour. */
export async function ajouterCommentaire(messageId: string, mailbox: string, texte: string): Promise<InboxEquipeCommentaire[]> {
  const r = await pluginFetch<{ commentaires?: InboxEquipeCommentaire[] }>('POST', 'agent/equipe/commentaires', { messageId, mailbox, texte: texte.trim().slice(0, 1000) });
  return Array.isArray(r.commentaires) ? r.commentaires.filter(c => c && typeof c === 'object') : [];
}

// ── Phase 3 : question à sa boîte, résumé de fil, rattrapage (à la demande, sans notification) ──

export async function poserQuestion(question: string, mailbox?: string): Promise<ReponseQuestion> {
  const r = await pluginFetch<Partial<ReponseQuestion>>('POST', 'agent/question', { question: question.trim().slice(0, 500), ...(mailbox ? { mailbox } : {}) });
  return { reponse: String(r.reponse || ''), sources: Array.isArray(r.sources) ? r.sources.filter(x => x && typeof x === 'object') : [] };
}

export async function fetchResumeFil(conversationId: string, mailbox: string): Promise<ResumeFil> {
  const qs = `conversationId=${encodeURIComponent(conversationId)}&mailbox=${encodeURIComponent(mailbox || '')}`;
  const r = await pluginFetch<Partial<ResumeFil>>('GET', `agent/resume-fil?${qs}`);
  return {
    lignes: Array.isArray(r.lignes) ? r.lignes.map(String).filter(Boolean) : [],
    nbMails: typeof r.nbMails === 'number' ? r.nbMails : 0,
    misAJour: String(r.misAJour || ''),
  };
}

export async function fetchRattrapage(depuisIso: string, mailbox?: string): Promise<Rattrapage> {
  const qs = `depuis=${encodeURIComponent(depuisIso)}${mailbox ? `&mailbox=${encodeURIComponent(mailbox)}` : ''}`;
  const r = await pluginFetch<Partial<Rattrapage>>('GET', `agent/rattrapage?${qs}`);
  return {
    depuis: String(r.depuis || depuisIso),
    ...(r.intro ? { intro: String(r.intro) } : {}),
    sections: Array.isArray(r.sections)
      ? r.sections.filter(s => s && typeof s === 'object').map(s => ({
        titre: String(s.titre || ''),
        elements: Array.isArray(s.elements) ? s.elements.filter(x => x && typeof x === 'object') : [],
      }))
      : [],
  };
}

// ── Phase 6 : traduction, modèles de réponse, fiche contact, détections, pièces lourdes ──
// (routes de worker/inbox-agent/routes-memoire.ts ; tout à la demande, rien n'est notifié)

export type LangueTraduction = 'FR' | 'EN' | 'DE' | 'LB';

/**
 * Traduit le mail lu (`messageId`) ou un texte (`texte`, prioritaire : brouillon) vers FR / EN / DE /
 * LB. IA côté worker (Haiku), sous le plafond IA du jour : 429 si atteint (erreur levée).
 */
export async function traduire(args: { vers: LangueTraduction; messageId?: string; mailbox?: string; texte?: string; brouillon?: boolean }): Promise<{ texte: string; vers: string; source: 'mail' | 'texte' }> {
  const r = await pluginFetch<{ texte?: string; vers?: string; source?: string }>('POST', 'agent/traduire', args);
  return { texte: String(r.texte || ''), vers: String(r.vers || args.vers), source: r.source === 'mail' ? 'mail' : 'texte' };
}

/** Modèle de réponse (fiche Communications « Réponse type »). */
export interface ModeleReponse {
  id: string;
  nom: string;
  categorie: string;
  destinataires: string[];
  objet: string;
  apercu: string;
  variables: string[];
  statut: string;
}

/** Modèle rempli pour un mail : texte à copier ou à insérer, SANS signature (Exclaimer l'ajoute). */
export interface ModeleRempli {
  modeleId: string;
  nom: string;
  langue: string;
  objet: string;
  corps: string;
  /** Variables à compléter à la main (restent visibles en {{X}} dans le texte). */
  manquantes: string[];
  variables: string[];
}

export async function fetchModeles(filtre: { public?: string; q?: string } = {}): Promise<ModeleReponse[]> {
  const qs = new URLSearchParams();
  if (filtre.public) qs.set('public', filtre.public);
  if (filtre.q) qs.set('q', filtre.q);
  const r = await pluginFetch<{ modeles?: ModeleReponse[] }>('GET', `agent/modeles${qs.toString() ? `?${qs}` : ''}`);
  return Array.isArray(r.modeles) ? r.modeles.filter(m => m && typeof m === 'object' && m.id) : [];
}

/**
 * Remplit un modèle pour le mail lu (`messageId`) ou, en rédaction d'une réponse, pour le dernier
 * mail reçu du fil (`conversationId`).
 */
export async function remplirModele(args: { modeleId: string; messageId?: string; conversationId?: string; mailbox?: string }): Promise<ModeleRempli> {
  const r = await pluginFetch<Partial<ModeleRempli>>('POST', 'agent/modeles/remplir', args);
  return {
    modeleId: String(r.modeleId || args.modeleId), nom: String(r.nom || ''), langue: String(r.langue || ''),
    objet: String(r.objet || ''), corps: String(r.corps || ''),
    manquantes: Array.isArray(r.manquantes) ? r.manquantes.map(String) : [],
    variables: Array.isArray(r.variables) ? r.variables.map(String) : [],
  };
}

/** Fiche mémoire d'un contact : la MIENNE (mémoire personnelle, visible par moi seul). */
export interface FicheContact {
  email: string;
  nom?: string;
  tutoiement?: boolean;
  ton?: string;
  langue?: string;
  sujets: string[];
  echanges: number;
  dernierLe?: string;
  sources?: { tutoiement?: string; langue?: string; ton?: string };
}

/** null = aucune fiche pour ce contact (404). */
export async function fetchFicheContact(email: string): Promise<FicheContact | null> {
  const e = String(email || '').trim().toLowerCase();
  if (!e) return null;
  try {
    const r = await pluginFetch<FicheContact>('GET', `agent/fiche-contact?email=${encodeURIComponent(e)}`);
    return r && typeof r === 'object' && r.email ? { ...r, sujets: Array.isArray(r.sujets) ? r.sujets.map(String) : [] } : null;
  } catch (err) {
    if (err instanceof PluginHttpError && err.status === 404) return null;
    throw err;
  }
}

/** Détections d'un mail (rebond / départ, absence avec date de retour, candidature). */
export interface DetectionsMail {
  rebond?: { type: string; adresse?: string; nouveauContact?: { email: string; nom?: string }; definitif?: boolean; signale?: string };
  absence?: { retourLe: string; extrait?: string };
  candidat?: boolean;
}

/** Détections relues côté worker (repli quand l'état du mail ne les porte pas encore). */
export async function fetchDetections(messageId: string, mailbox: string): Promise<DetectionsMail | null> {
  if (!messageId) return null;
  const qs = `messageId=${encodeURIComponent(messageId)}&mailbox=${encodeURIComponent(mailbox || '')}`;
  try {
    const r = await pluginFetch<DetectionsMail>('GET', `agent/detections?${qs}`);
    return r && typeof r === 'object' && (r.rebond || r.absence || r.candidat) ? r : null;
  } catch (err) {
    if (err instanceof PluginHttpError && (err.status === 404 || err.status === 403)) return null;
    throw err;
  }
}

export interface PiecesLourdes { totalOctets: number; lourdes: Array<{ nom: string; octets: number }>; suggestion: string }

/** Pièces jointes trop lourdes ? (seuil côté worker, sans IA) ; null sinon. */
export async function verifierPiecesLourdes(pieces: Array<{ nom: string; octets: number; isInline?: boolean }>): Promise<PiecesLourdes | null> {
  if (!pieces.length) return null;
  const r = await pluginFetch<{ lourde?: PiecesLourdes | null }>('POST', 'agent/piece-jointe-lourde', { pieces: pieces.slice(0, 50) });
  return r.lourde && typeof r.lourde === 'object' ? r.lourde : null;
}

/** Lien court vibes.lu/p/… d'une URL de partage DÉJÀ créée (cloud du projet). Lève une erreur sinon. */
export async function creerLienCourt(url: string, libelle?: string): Promise<string> {
  const r = await pluginFetch<{ lien?: string }>('POST', 'agent/lien-court', { url: url.trim(), ...(libelle ? { libelle } : {}) });
  if (!r.lien) throw new Error('lien court indisponible');
  return String(r.lien);
}

/** `rang` : position de la pièce dans `attachmentIds`. */
export interface LienPieceLourde { nom: string; taille: number; lienCourt: string; court: boolean; expireLe: string; rang: number }
export interface DepotPiecesLourdes { liens: LienPieceLourde[]; refus: Array<{ nom: string; motif: string }>; html: string }

/**
 * « Déposer et partager » : le worker lit les pièces du mail par Graph (lecture seule), les dépose
 * dans le transit du CLOUD (jamais le NAS ; purgé à 30 jours), crée un lien public lecture seule
 * (30 jours) et son lien court vibes.lu/p/…. `messageId` = identifiant REST du mail (reçu, ou
 * brouillon enregistré). `pieces` : nom et taille vus par Office (repli si les identifiants diffèrent).
 */
export async function deposerPiecesLourdes(p: { mailbox?: string; messageId: string; attachmentIds: string[]; pieces?: Array<{ nom: string; octets: number }> }): Promise<DepotPiecesLourdes> {
  const r = await pluginFetch<Partial<DepotPiecesLourdes>>('POST', 'agent/deposer-partager', p);
  return { liens: Array.isArray(r.liens) ? r.liens : [], refus: Array.isArray(r.refus) ? r.refus : [], html: typeof r.html === 'string' ? r.html : '' };
}

// ── Fonctions reprises d'« ATLAS Assistant » (outlook-plugin) ──

export interface EmailPayload {
  emailId: string;
  conversationId?: string;
  subject: string;
  fromEmail: string;
  fromName: string;
  receivedAt: string;
  bodyPreview: string;
}

export interface EmailActionSuggestion {
  id: string;
  type: string;
  label: string;
  icon?: string;
  color?: string;
  priority: number;
  payload?: Record<string, unknown>;
  tooltip?: string;
}

/** Actions ATLAS suggérées pour un mail (règles + tag existant, SANS appel IA). */
export async function fetchEmailSuggestions(payload: EmailPayload): Promise<EmailActionSuggestion[]> {
  const r = await pluginFetch<{ suggestions?: EmailActionSuggestion[] }>('POST', 'email/suggestions', payload);
  return Array.isArray(r.suggestions) ? r.suggestions : [];
}

export type QuickReplyType = 'acknowledge' | 'quote_pending' | 'polite_refusal' | 'reschedule';

/** Réponse rapide : modèle fixe FR / EN / DE × tu / vous (sans IA). Renvoie du texte brut. */
export async function fetchQuickReply(args: {
  type: QuickReplyType;
  language: 'FR' | 'EN' | 'DE';
  address: 'tu' | 'vous';
  recipientName: string;
}): Promise<string> {
  const r = await pluginFetch<{ body?: string }>('POST', 'email/quick-reply', args);
  return String(r.body || '');
}

/** Brouillon automatique (IA côté worker, Agency Brain), sur action volontaire uniquement. */
export async function fetchAutoDraft(args: {
  email: { id: string; subject: string; from: { name: string; email: string }; receivedAt: string; bodyPreview: string };
  fullBody?: string;
}): Promise<{ subject?: string; body?: string } | null> {
  const r = await pluginFetch<{ draft?: { subject?: string; body?: string } }>('POST', 'email/auto-draft', args);
  return r.draft || null;
}
