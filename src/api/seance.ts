/**
 * seance.ts · routes du worker pour la « séance de tri » (backlog recz1ck93h164mv9x, 07/10/2026) :
 * /api/plugin/agent/seance* (worker/inbox-agent/routes-seance.ts). Même jeton que le reste du
 * complément (getWorkerToken, aucun secret ici). Les actions des cartes passent par les routes
 * existantes (agent.ts : rangerDansDossier, executerAction, mettrePlusTard, annulerAction ;
 * tableau.ts : redigerArgo, deposerBrouillon).
 * Types : miroir de src/utils/inbox-seance.ts (le complément n'importe rien hors de son dossier).
 */
import { workerRequest } from './worker';

export type SeanceActionType = 'repondre' | 'classer-projet' | 'metier' | 'confier' | 'plus-tard' | 'archiver' | 'ouvrir';

export interface SeanceAction {
  cle: string;
  type: SeanceActionType;
  libelle: string;
  detail?: string;
  donnees?: Record<string, unknown>;
}

export interface SeanceContexte {
  tiers?: string;
  projet?: { id: string; libelle: string; statut?: string; dates?: string; echeance?: string };
  offre?: { libelle: string; montant?: string; envoyeeLe?: string; statut?: string; enAttente: boolean };
  impayes: null | Array<{ libelle: string; montant?: string }>;
  dernierEchange?: { sens: 'recu' | 'envoye'; at: string; sujet?: string };
  confie?: { a: string; nom: string; le: string };
}

export interface SeanceCarte {
  mailbox: string;
  messageId: string;
  graphId: string;
  conversationId: string;
  webLink?: string;
  from: { email: string; name: string };
  subject: string;
  receivedAt: string;
  resume: string;
  langue?: string;
  pile: string;
  urgence: number;
  correspondant?: string;
  valeur: number;
  raisons: string[];
  important: boolean;
  risque?: { niveau: 'a_verifier' | 'eleve'; raisons: string[] };
  recommandee: SeanceAction | null;
  alternatives: SeanceAction[];
  appris?: string;
  brouillon?: { texte: string; resumeIntention: string; alertes?: Array<{ message: string; gravite: string }> };
  contexte: SeanceContexte;
}

export interface SeanceSemaine {
  semaine: string;
  jours: Array<{ jour: string; restants?: number; minutes: number; traites: number }>;
  minutes: number;
  traites: number;
  objectifAtteint: number;
  joursReleves: number;
  moyenneRestants: number | null;
}

export interface Seance {
  mailbox: string;
  dansLaBoite: number | null;
  objectif: number;
  aTrier: number;
  dureeMinutes: number;
  cartes: SeanceCarte[];
  collegues: Array<{ email: string; nom: string }>;
  stats: SeanceSemaine[];
  ecrituresActives: boolean;
  mode: string;
  genereLe: string;
}

export const fetchSeance = (frais = false) =>
  workerRequest<Seance>('GET', `agent/seance${frais ? '?frais=1' : ''}`, undefined, { retry: true, timeoutMs: 90_000 });
export const fetchStatsSeance = () =>
  workerRequest<{ stats: SeanceSemaine[]; dansLaBoite: number | null; objectif: number }>('GET', 'agent/seance/stats', undefined, { retry: true, timeoutMs: 30_000 });

/** Apprentissage : action recommandée acceptée ou remplacée (`annule` : retour en arrière). Jamais bloquant. */
export function noterDecisionSeance(recommandee: string | null, faite: string, annule = false): void {
  void workerRequest('POST', 'agent/seance/decision', { recommandee, faite, ...(annule ? { annule: true } : {}) }, { timeoutMs: 15_000 }).catch(() => { /* facultatif */ });
}

export const confierMail = (p: { messageId: string; mailbox: string; a: string; quand: string }) =>
  workerRequest<{ ok: boolean; nom?: string; jusqua?: string; prevenu?: boolean }>('POST', 'agent/seance/confier', p, { timeoutMs: 30_000 });
export const annulerConfier = (messageId: string, mailbox: string) =>
  workerRequest<{ ok: boolean }>('DELETE', 'agent/seance/confier', { messageId, mailbox }, { timeoutMs: 30_000 });

export const finSeance = (p: { dureeMs: number; traites: number }) =>
  workerRequest<{ ok: boolean; restants: number | null; stats: SeanceSemaine[] }>('POST', 'agent/seance/fin', p, { timeoutMs: 30_000 });
