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
  devis?: { id: string; libelle: string; expireLe?: string };
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
  patron?: string;
  regle?: { id: string; interpretation: string };
  /** Trois réponses courtes (touches 5, 6, 7), jamais envoyées. */
  suggestions?: string[];
  /** Mise en relation probable (touche 8 : introducteur en copie cachée). */
  miseEnRelation?: boolean;
}

// « Ne rien laisser filer » (07/10/2026, miroir de src/utils/inbox-seance.ts).
export interface SeancePromesse { messageId: string; index: number; texte: string; echeance: string; sujet: string; a: string; sentAt: string; enRetard: boolean; webLink?: string }
export interface SeanceOffreConsultee { messageId: string; projetId: string; offre: string; montant?: string; a: string; sentAt: string; libelle: string; webLink?: string }
export interface SeanceCrASuivre { id: string; titre: string; valideLe: string; lien: string }

// Assistant inbox (07/10/2026) : cartes spéciales (miroir de src/utils/inbox-seance.ts).
export interface SeanceGroupe {
  cle: string;
  libelle: string;
  messageIds: string[];
  apercu: Array<{ messageId: string; subject: string; receivedAt: string }>;
  actions: Array<'archiver' | 'marquer-lu' | 'classer'>;
  dossier?: { id: string; chemin: string };
}
export interface SeanceDesabonnement { expediteur: string; nom?: string; recus30j: number; mode: 'un-clic' | 'brouillon' | 'lien'; dansLaBoite: number }
export interface SeanceSollicitations { semaine: string; mails: Array<{ messageId: string; from: { email: string; name: string }; subject: string; receivedAt: string; raisons: string[] }> }
export interface SeancePropositionAutonomie { patron: string; libelle: string; acceptees: number }

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
  groupes?: SeanceGroupe[];
  desabonnements?: SeanceDesabonnement[];
  sollicitations?: SeanceSollicitations | null;
  sollicitationsHorsSeance?: number;
  autonomie?: SeancePropositionAutonomie[];
  concentration?: { actif: boolean; creneaux: string[]; retenus: number; prochain?: string; libelle?: string };
  promesses?: SeancePromesse[];
  offresConsultees?: SeanceOffreConsultee[];
  crsASuivre?: SeanceCrASuivre[];
}

export const fetchSeance = (frais = false) =>
  workerRequest<Seance>('GET', `agent/seance${frais ? '?frais=1' : ''}`, undefined, { retry: true, timeoutMs: 90_000 });
export const fetchStatsSeance = () =>
  workerRequest<{ stats: SeanceSemaine[]; dansLaBoite: number | null; objectif: number }>('GET', 'agent/seance/stats', undefined, { retry: true, timeoutMs: 30_000 });

/** Apprentissage : action recommandée acceptée ou remplacée (`annule` : retour en arrière). Jamais bloquant. */
export function noterDecisionSeance(recommandee: string | null, faite: string, annule = false, patron?: string, temps?: { dureeMs: number; client?: string }): void {
  void workerRequest('POST', 'agent/seance/decision', {
    recommandee, faite, ...(annule ? { annule: true } : {}), ...(patron ? { patron } : {}),
    ...(temps && temps.client && !annule ? { dureeMs: Math.round(temps.dureeMs), client: temps.client } : {}),
  }, { timeoutMs: 15_000 }).catch(() => { /* facultatif */ });
}

// ── Répondre plus vite, ne rien laisser filer (worker/inbox-agent/reponse-suivi.ts) ──
export const fetchSuggestions = (messageId: string, mailbox: string) =>
  workerRequest<{ textes: string[] }>('POST', 'agent/seance/suggestions', { messageId, mailbox }, { timeoutMs: 45_000 });
export const preparerMiseEnRelation = (messageId: string, mailbox: string) =>
  workerRequest<{ a: string[]; cci: string[]; sujet: string; texte: string; introducteur: { email: string; name: string } }>('POST', 'agent/seance/mise-en-relation', { messageId, mailbox }, { timeoutMs: 30_000 });
export const promesseTache = (messageId: string, index: number) =>
  workerRequest<{ ok: boolean; tacheId?: string; echeance?: string; projet?: string; lien?: string }>('POST', 'agent/seance/promesse', { messageId, index, action: 'tache' }, { timeoutMs: 45_000 });
export const promesseEcarter = (messageId: string, index: number) =>
  workerRequest<{ ok: boolean }>('POST', 'agent/seance/promesse', { messageId, index, action: 'ecarter' }, { timeoutMs: 20_000 });
export const promesseAnnuler = (messageId: string, index: number, tacheId?: string) =>
  workerRequest<{ ok: boolean }>('DELETE', 'agent/seance/promesse', { messageId, index, ...(tacheId ? { tacheId } : {}) }, { timeoutMs: 30_000 });
export const relanceOffre = (messageId: string) =>
  workerRequest<{ texte: string; resumeIntention: string; a: string[]; sujet: string; webLink: string }>('POST', 'agent/seance/relance-offre', { messageId }, { timeoutMs: 90_000 });
export interface StatistiquesBoite {
  topExpediteurs: Array<{ email: string; nom?: string; n: number }>;
  horsHoraires: { n: number; part: number };
  total: number;
  tempsParClient: Array<{ client: string; minutes: number }>;
  courbe: Array<{ jour: string; restants: number }>;
  objectif: number;
}
export const fetchStatistiques = () => workerRequest<StatistiquesBoite>('GET', 'agent/seance/statistiques', undefined, { retry: true, timeoutMs: 30_000 });
export const enregistrerRemplacant = (remplacant: string | null) =>
  workerRequest<{ ok: boolean; remplacant: string | null }>('POST', 'agent/assistant/remplacant', { remplacant: remplacant || '' }, { timeoutMs: 20_000 });

export const confierMail = (p: { messageId: string; mailbox: string; a: string; quand: string }) =>
  workerRequest<{ ok: boolean; nom?: string; jusqua?: string; prevenu?: boolean }>('POST', 'agent/seance/confier', p, { timeoutMs: 30_000 });
export const annulerConfier = (messageId: string, mailbox: string) =>
  workerRequest<{ ok: boolean }>('DELETE', 'agent/seance/confier', { messageId, mailbox }, { timeoutMs: 30_000 });

export const finSeance = (p: { dureeMs: number; traites: number }) =>
  workerRequest<{ ok: boolean; restants: number | null; stats: SeanceSemaine[] }>('POST', 'agent/seance/fin', p, { timeoutMs: 30_000 });

// ── Assistant inbox (worker/inbox-agent/routes-assistant.ts) ──────────────────────

export interface RegleAssistant { id: string; texte: string; interpretation: string; etat: 'a-confirmer' | 'active' | 'pause'; creeLe: string; appliquee: number }
export interface ReglagesAssistant {
  concentration: { actif: boolean; creneaux: string[] };
  regles: RegleAssistant[];
  delegations: Array<{ patron: string; libelle: string; accordeLe: string; appliquees: number; etat: 'appliquer' | 'prete' | null }>;
  ecrituresAgent: boolean;
  autonomieRanger: boolean;
  boiteSuivie: boolean;
  /** Remplaçant pendant les congés (mails clients proposés), et collègues possibles. */
  remplacant?: string | null;
  collegues?: string[];
}
export type ConditionRetour =
  | { type: 'date'; at: string }
  | { type: 'evenement-j7'; projetId: string; libelle?: string }
  | { type: 'etape-projet'; projetId: string; statutInitial: string; libelle?: string }
  | { type: 'expiration-devis'; devisId: string; libelle?: string };

export const fetchReglagesAssistant = () => workerRequest<ReglagesAssistant>('GET', 'agent/assistant/reglages', undefined, { retry: true, timeoutMs: 30_000 });
export const enregistrerConcentration = (p: { actif: boolean; creneaux: string[] }) =>
  workerRequest<{ ok: boolean; concentration: { actif: boolean; creneaux: string[] } }>('POST', 'agent/assistant/concentration', p, { timeoutMs: 20_000 });
export const creerRegle = (texte: string) => workerRequest<{ ok: boolean; regle: RegleAssistant }>('POST', 'agent/assistant/regles', { texte }, { timeoutMs: 60_000 });
export const etatRegle = (id: string, etat: 'active' | 'pause') => workerRequest<{ ok: boolean }>('POST', 'agent/assistant/regles/etat', { id, etat }, { timeoutMs: 20_000 });
export const supprimerRegle = (id: string) => workerRequest<{ ok: boolean }>('DELETE', 'agent/assistant/regles', { id }, { timeoutMs: 20_000 });
export const repondreAutonomie = (patron: string, accorder: boolean) =>
  workerRequest<{ ok: boolean; etat?: string; message?: string }>('POST', 'agent/assistant/autonomie', { patron, accorder }, { timeoutMs: 20_000 });
export const retirerDelegation = (patron: string) => workerRequest<{ ok: boolean }>('DELETE', 'agent/assistant/autonomie', { patron }, { timeoutMs: 20_000 });
export const desabonner = (expediteur: string, opts: { archiver?: boolean; ignorer?: boolean } = {}) =>
  workerRequest<{ ok: boolean; statut: string; mode?: string; archives?: number; actionIds?: string[]; brouillon?: { a: string; sujet: string; corps: string } | null; lien?: string; erreur?: string }>('POST', 'agent/assistant/desabonner', { expediteur, ...opts }, { timeoutMs: 60_000 });
export const actionGroupe = (p: { action: 'archiver' | 'marquer-lu' | 'classer'; messageIds: string[]; dossierId?: string }) =>
  workerRequest<{ ok: boolean; faits: number; echecs: number; actionIds: string[]; lus: Array<{ messageId: string; graphId: string }> }>('POST', 'agent/assistant/groupe', p, { timeoutMs: 90_000 });
export const annulerGroupe = (p: { actionIds: string[]; lus: Array<{ messageId: string; graphId: string }> }) =>
  workerRequest<{ ok: boolean; annules: number; echecs: number }>('POST', 'agent/assistant/groupe/annuler', p, { timeoutMs: 90_000 });
export const poserRetour = (messageId: string, mailbox: string, condition: ConditionRetour) =>
  workerRequest<{ ok: boolean; libelle?: string; attente?: string; cible?: string; jusqua?: string }>('POST', 'agent/assistant/retour', { messageId, mailbox, condition }, { timeoutMs: 30_000 });
export const retirerRetour = (messageId: string, mailbox: string) => workerRequest<{ ok: boolean }>('DELETE', 'agent/assistant/retour', { messageId, mailbox }, { timeoutMs: 30_000 });
export const sollicitationsVues = (semaine: string) => workerRequest<{ ok: boolean }>('POST', 'agent/assistant/sollicitations/vues', { semaine }, { timeoutMs: 15_000 });
export interface DelaisClients { cibleHeures: number; mois: Array<{ mois: string; clients: Array<{ client: string; repondus: number; moyenneHeures: number | null; medianeHeures: number | null; horsDelai: number; sansReponse: number }> }> }
export const fetchDelaisClients = () => workerRequest<DelaisClients>('GET', 'agent/assistant/delais-clients', undefined, { retry: true, timeoutMs: 60_000 });
