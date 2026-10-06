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
  enAttente: AgentListeElement[];
  deCote: TableauMail[];
  compteurs: Record<'personnes' | 'clients' | 'mandats' | 'notifications' | 'newsletters' | 'factures' | 'priorites' | 'enAttente' | 'deCote', number>;
  version: string;
  misAJour: string;
  programmes: EnvoiProgramme[];
  quiTraite: Array<{ personne: string; nom: string; mails: number }>;
  ecrituresActives: boolean;
  mode: string;
  moi: string;
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

export const fetchTableau = (mailbox = 'toutes') => lecture<Tableau>(`tableau${qMailbox(mailbox)}`);
export const fetchVersion = (mailbox = 'toutes') => workerRequest<{ version: string; misAJour: string }>('GET', `agent/tableau/version${qMailbox(mailbox)}`, undefined, { timeoutMs: 20_000 });

export const redigerArgo = (messageId: string, mailbox: string) => ecriture<ReponseRedigee>('POST', 'tableau/reponse', { messageId, mailbox, source: 'argo' });
export const remplirModeleTableau = (messageId: string, mailbox: string, modeleId: string) => ecriture<ReponseRedigee>('POST', 'tableau/reponse', { messageId, mailbox, source: 'modele', modeleId });

export interface Depot { depose: boolean; brouillonId?: string; webLink?: string; raison?: string }
export const deposerBrouillon = (messageId: string, mailbox: string, texte: string, tous = false) => ecriture<Depot>('POST', 'tableau/brouillon', { messageId, mailbox, texte, tous });

export const programmerEnvoi = (p: { brouillonId: string; quand: string; relanceJours?: number; enReponseA?: string }) =>
  ecriture<{ ok: boolean; envoi: EnvoiProgramme }>('POST', 'tableau/envois', p);
export const annulerEnvoi = (brouillonId: string) => ecriture<{ ok: boolean; envoi: EnvoiProgramme }>('DELETE', 'tableau/envois', { brouillonId });

export const fetchEngagements = () => lecture<{ promis: Engagement[]; attendus: Engagement[]; verificationEnCours: boolean }>('tableau/engagements');
export const marquerTenu = (recordId: string) => ecriture<{ ok: boolean }>('POST', 'tableau/engagements/tenu', { recordId });
export const redigerEngagement = (recordId: string) => ecriture<ReponseRedigee>('POST', 'tableau/engagements/rediger', { recordId });

export const fetchRdv = () => lecture<{ rdv: RdvPrep[]; jours: { aujourdhui: string; lendemain: string } }>('tableau/rdv');
export const fetchReactivite = () => lecture<ReactiviteVue>('tableau/reactivite');
