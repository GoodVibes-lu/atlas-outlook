/**
 * tableau.ts · routes du worker pour le tableau de bord pleine largeur (application ATLAS de la
 * barre de gauche) : /api/plugin/agent/tableau* (worker/inbox-agent/routes-tableau.ts). Même jeton
 * que le panneau (getWorkerToken, aucun secret ici). Les actions déjà existantes (mettre de côté,
 * relance, « Je prends », commentaires, actions proposées, modèles) restent dans agent.ts.
 * Types : miroir de worker/inbox-agent/tableau.ts (à garder alignés).
 */
import { workerRequest } from './worker';
import type { AgentListeElement } from './agent';

export type FamilleTableau = 'personnes' | 'mandats' | 'notifications' | 'newsletters' | 'factures';

export type TableauMail = AgentListeElement & {
  mailbox: string;
  conversationId: string;
  famille: FamilleTableau;
  correspondant?: 'direction' | 'client' | 'partenaire';
  priorite: number;
  raisons: string[];
  repondu?: boolean;
  actions?: Array<{ type: string; libelle: string }>;
  langue?: string;
  tonalite?: string;
  /** « En attente » : envoi de la personne sans réponse (et non mail reçu). */
  envoye?: boolean;
};

export interface DigestNewsletter {
  expediteur: string; nom: string; nb: number; dernier: string; sujets: string[]; desinscription?: string; mails: TableauMail[];
}

export interface EnvoiProgramme {
  id: string; mailbox: string; quand: string; sujet: string; a: string[]; programmeLe: string; par: string;
  enReponseA?: string; webLink?: string; relanceJours?: number; annuleLe?: string;
}

export interface Tableau {
  boites: string[];
  personnes: TableauMail[];
  notifications: TableauMail[];
  newsletters: DigestNewsletter[];
  factures: TableauMail[];
  /** Mandats / associations (classement de l'agent) : section masquée si vide. */
  mandats?: TableauMail[];
  priorites: TableauMail[];
  /**
   * « En attente » : mails REÇUS encore dans la boîte de réception (un par fil), puis envois de la
   * personne sans réponse depuis 14 jours au plus (`envoye`, hors boîte de réception par nature).
   */
  enAttente: Array<AgentListeElement & { mailbox?: string; conversationId?: string; envoye?: boolean }>;
  deCote: TableauMail[];
  compteurs: Record<'personnes' | 'clients' | 'mandats' | 'notifications' | 'newsletters' | 'factures' | 'priorites' | 'enAttente' | 'deCote', number>;
  version: string;
  misAJour: string;
  programmes: EnvoiProgramme[];
  quiTraite: Array<{ personne: string; nom: string; mails: number }>;
  ecrituresActives: boolean;
  mode: string;
  moi: string;
  /** Boîtes dont la boîte de réception n'a pas pu être relue (Graph) : rien n'y est masqué. */
  inboxNonVerifiee?: string[];
  /** Traités à ranger (08/10/2026) : mails répondus encore dans la boîte, dossier proposé (boîte personnelle). */
  traites?: TraiteARanger[];
  /** Rangés seuls par l'agent après la réponse (3 jours), annulables. */
  rangesSeuls?: TraiteARanger[];
}

// ── Traités à ranger (08/10/2026) : miroir de src/utils/inbox-traites-a-ranger.ts ──

export interface TraiteDestination {
  type: 'projet' | 'expediteur' | 'regle' | 'aucun';
  dossierId?: string;
  chemin?: string;
  aCreer?: boolean;
  projetId?: string;
  libelle?: string;
}

export interface TraiteARanger {
  mailbox: string;
  messageId: string;
  graphId: string;
  conversationId: string;
  subject: string;
  from: { email: string; name: string };
  receivedAt: string;
  webLink?: string;
  repondu: { at: string; messageId: string };
  destination: TraiteDestination;
  statut: 'a_ranger' | 'range' | 'ignore';
  calculeLe: string;
  actionId?: string;
  rangeLe?: string;
  auto?: boolean;
}

/** Phrase de la suggestion (même texte que le worker). */
export function phraseSuggestion(d: TraiteDestination): string {
  if (d.type === 'aucun' || !d.chemin) return 'Répondu : le ranger dans un dossier ?';
  return `Répondu : le ranger dans « ${d.chemin} » ?${d.aCreer ? ' (dossier à créer)' : ''}`;
}
export function sourceSuggestion(d: TraiteDestination): string {
  return d.type === 'projet' ? (d.aCreer ? 'dossier du projet, à créer' : 'dossier du projet')
    : d.type === 'expediteur' ? 'dossier habituel de cet expéditeur'
    : d.type === 'regle' ? 'ta règle' : '';
}

export interface ReponseRedigee {
  texte: string;
  objet?: string;
  source: 'agent' | 'argo' | 'modele';
  resumeIntention?: string;
  aCompleter?: string[];
  manquantes?: string[];
  /** Engagement : mail sur lequel la réponse a été rédigée. */
  messageId?: string;
  mailbox?: string;
}

export interface Engagement {
  recordId: string; direction: 'made' | 'received'; counterpart: string; counterpartName: string; promiseText: string;
  dueDate: string; status: string; sourceConversationId: string; sourceSubject: string; createdAt: string;
  enRetard: boolean; jours: number;
}

export interface RdvPrep {
  id: string; jour: string; sujet: string; debut: string; fin: string; lieu?: string; organisateur: boolean; avec: string[]; utile: boolean;
  prep: null | {
    sections: Array<{ key: string; titre: string; lignes: Array<{ texte: string; meta?: string }> }>;
    notesRdv: Array<{ titre: string; contenu: string }>;
    projets: Array<{ id: string; nom: string }>;
    tiers: Array<{ id: string; nom: string }>;
    viaIntitule?: string;
  };
}

export interface ReactiviteVue {
  seuilHeures: number;
  personnes: Array<{
    personne: string; moyenneHeures: number | null; repondus: number;
    enRetard: Array<{ mailbox: string; messageId: string; subject: string; from: string; client: string; recuLe: string; webLink?: string; responsable: string; heures: number }>;
  }>;
  clients: Array<{ client: string; moyenneHeures: number | null; mails: number; enRetard: number }>;
  calculeLe: string;
}

const lecture = <T>(path: string) => workerRequest<T>('GET', `agent/${path}`, undefined, { retry: true, timeoutMs: 60_000 });
const ecriture = <T>(method: 'POST' | 'DELETE', path: string, body: unknown) => workerRequest<T>(method, `agent/${path}`, body, { timeoutMs: 90_000 });
const qMailbox = (mailbox: string) => (mailbox && mailbox !== 'toutes' ? `?mailbox=${encodeURIComponent(mailbox)}` : '');

/** `frais` (« Actualiser », après une action) : le worker relit la boîte de réception sans attendre son cache. */
export const fetchTableau = (mailbox = 'toutes', frais = false) => {
  const q = qMailbox(mailbox);
  return lecture<Tableau>(`tableau${q}${frais ? `${q ? '&' : '?'}frais=1` : ''}`);
};
export const fetchVersion = (mailbox = 'toutes') => workerRequest<{ version: string; misAJour: string }>('GET', `agent/tableau/version${qMailbox(mailbox)}`, undefined, { timeoutMs: 20_000 });

export const redigerArgo = (messageId: string, mailbox: string) => ecriture<ReponseRedigee>('POST', 'tableau/reponse', { messageId, mailbox, source: 'argo' });
export const remplirModeleTableau = (messageId: string, mailbox: string, modeleId: string) => ecriture<ReponseRedigee>('POST', 'tableau/reponse', { messageId, mailbox, source: 'modele', modeleId });

export interface Depot { depose: boolean; brouillonId?: string; webLink?: string; raison?: string }
export const deposerBrouillon = (messageId: string, mailbox: string, texte: string, tous = false) => ecriture<Depot>('POST', 'tableau/brouillon', { messageId, mailbox, texte, tous });

/** Envoi direct (08/10/2026) : réponse préparée dans la boîte personnelle, destinataires à confirmer ; rien ne part. */
export interface EnvoiPret { pret: boolean; brouillonId: string; a: Array<{ email: string; nom: string }>; cc: Array<{ email: string; nom: string }>; sujet?: string }
export const preparerEnvoi = (messageId: string, mailbox: string, texte: string, tous = false) => ecriture<EnvoiPret>('POST', 'tableau/envoyer', { messageId, mailbox, texte, tous });
/** Confirmation : remise différée de quelques secondes (`annulableS`), « Annuler » le retire tant qu'il n'est pas parti. */
export const confirmerEnvoi = (brouillonId: string, enReponseA?: string) => ecriture<{ envoye: boolean; partLe: string; annulableS: number }>('POST', 'tableau/envoyer/confirmer', { brouillonId, enReponseA });
export const abandonnerEnvoi = (brouillonId: string) => ecriture<{ ok: boolean }>('POST', 'tableau/envoyer/abandonner', { brouillonId });

export const programmerEnvoi = (p: { brouillonId: string; quand: string; relanceJours?: number; enReponseA?: string }) =>
  ecriture<{ ok: boolean; envoi: EnvoiProgramme }>('POST', 'tableau/envois', p);
export const annulerEnvoi = (brouillonId: string) => ecriture<{ ok: boolean; envoi: EnvoiProgramme }>('DELETE', 'tableau/envois', { brouillonId });

/** promis / attendus : échus depuis 30 jours au plus, ou à venir ; anciens : au-delà (repliés). */
export const fetchEngagements = () => lecture<{ promis: Engagement[]; attendus: Engagement[]; anciens?: Engagement[]; verificationEnCours: boolean }>('tableau/engagements');
/** « Marquer abandonné » en lot (statut existant `cancelled` de la table Email Promises). */
export const abandonnerEngagements = (recordIds: string[]) => ecriture<{ ok: boolean; faits: number; echecs: number; inconnus: number }>('POST', 'tableau/engagements/abandonner', { recordIds });
export const marquerTenu = (recordId: string) => ecriture<{ ok: boolean }>('POST', 'tableau/engagements/tenu', { recordId });
export const redigerEngagement = (recordId: string) => ecriture<ReponseRedigee>('POST', 'tableau/engagements/rediger', { recordId });

export const fetchRdv = () => lecture<{ rdv: RdvPrep[]; jours: { aujourdhui: string; lendemain: string } }>('tableau/rdv');
export const fetchReactivite = () => lecture<ReactiviteVue>('tableau/reactivite');

// ── Traités à ranger (08/10/2026) ──

/** Mails répondus à ranger (sa boîte) ; `messageId` : ce mail seulement (panneau « Ce mail »). */
export const fetchTraites = (messageId?: string) =>
  lecture<{ mailbox: string; traites: TraiteARanger[]; rangesSeuls: TraiteARanger[] }>(`tableau/traites${messageId ? `?messageId=${encodeURIComponent(messageId)}` : ''}`);
/** « Ranger » : dossier proposé par défaut, ou un autre dossier (existant, ou chemin à créer). */
export const rangerTraite = (p: { messageId: string; dossierId?: string; projetId?: string; creer?: { chemin: string; projetId?: string } }) =>
  ecriture<{ ok: boolean; messageId: string; dossier: { id: string; chemin: string }; cree: boolean; actionId?: string }>('POST', 'tableau/traites/ranger', p);
export const ignorerTraite = (messageId: string) => ecriture<{ ok: boolean }>('POST', 'tableau/traites/ignorer', { messageId });
/** « Annuler » : le mail revient dans la boîte de réception et redevient à ranger. */
export const annulerTraite = (messageId: string) => ecriture<{ ok: boolean }>('POST', 'tableau/traites/annuler', { messageId });
