/**
 * COPIE LOCALE IDENTIQUE de `src/types/inbox-agent.types.ts` (racine du dépôt, contrat de l'agent
 * serveur) : le complément est publié depuis son propre dossier (GitHub Pages) et n'importe rien
 * hors de `outlook-addin/`. Toute évolution du fichier racine doit être recopiée ici telle quelle
 * (le contenu ci-dessous, à partir du commentaire d'origine, ne doit pas diverger).
 */

/**
 * inbox-agent.types.ts · contrat COMMUN de l'agent d'inbox (backlog `.claude/BACKLOG-AGENT-INBOX.md`,
 * phase 1.3). Le worker (`worker/inbox-agent/`) écrit cet état, le complément Outlook unifié et ATLAS
 * le lisent (routes `/api/plugin/agent/*`). Ne pas changer la forme de `InboxMessageState` sans
 * prévenir les consommateurs (complément, page « Mon agent », brief).
 */

/**
 * Les piles du tri (« À traiter par toi », « En attente d'une réponse », « Pour info », « Bruit »),
 * plus « À filtrer » (phase 2) : premier mail d'un expéditeur externe inconnu, en attente de la
 * décision de la personne (accepter / refuser une fois pour toutes ; reste en boîte de réception).
 */
export type InboxPile = 'a_traiter' | 'en_attente' | 'pour_info' | 'bruit' | 'a_filtrer';
export type InboxLangue = 'FR' | 'EN' | 'DE' | 'LU' | 'AUTRE';
export type InboxUrgence = 0 | 1 | 2 | 3;
/** 'blanc' = l'agent lit et classe sans rien écrire dans les boîtes (phase 1.7, défaut). */
export type InboxAgentMode = 'blanc' | 'actif';

export interface InboxEngagement {
  /** 'recu' = promesse faite PAR le correspondant ; 'fait' = promesse faite par la personne de la boîte. */
  sens: 'recu' | 'fait';
  texte: string;
  /** 'YYYY-MM-DD' si une échéance est connue. */
  echeance?: string;
}

export interface InboxActionProbable {
  type: string;
  libelle: string;
  cible?: { table: string; id?: string };
}

/** État calculé par le serveur pour UN message d'UNE boîte (une seule lecture, partagée entre les écrans). */
export interface InboxMessageState {
  mailbox: string;
  /** internetMessageId (stable d'une boîte à l'autre). */
  messageId: string;
  graphId: string;
  conversationId: string;
  receivedAt: string;
  from: { email: string; name: string };
  subject: string;
  pile: InboxPile;
  categorie: string;
  urgence: InboxUrgence;
  resume: string;
  langue: InboxLangue;
  besoinReponse: boolean;
  engagements: InboxEngagement[];
  actionProbable?: InboxActionProbable;
  source: 'regles' | 'ia';
  /**
   * Règle qui a décidé (identifiant de `LIBELLES_REGLES`, 'ia:lot' / 'ia:direct'), ou libellé
   * lisible d'une règle PERSONNELLE de la personne (« Règle personnelle : Les mails de … ») :
   * voir `PREFIXE_REGLE_PERSO` (src/utils/inbox-rules.ts). Le rapport admin ne montre jamais ce
   * libellé personnel.
   */
  regle?: string;
  /** Dossier où la personne range ces mails (règle personnelle) : proposé, rien n'est déplacé en phase 1. */
  dossier?: string;
  /** Actions que la personne a refusées pour ces mails (règles personnelles) : 'brouillon', 'ranger'… */
  actionsInterdites?: string[];
  mode: InboxAgentMode;
  traiteLe: string;
  coutEur?: number;
  /** Lien Outlook sur le web du message (Graph `webLink`), pour ouvrir le mail depuis une liste. */
  webLink?: string;
  /**
   * Correspondant important (phase 2) : jamais rangé ni mis en bruit, urgence minimale 1.
   * 'direction' = INBOX_AGENT_VIP ; 'client' = tiers lié à un projet non clôturé ; 'partenaire'.
   */
  correspondant?: InboxCorrespondant;
  /** Préférence de la personne qui a décidé de la pile (09/10/2026, miroir de src/types/inbox-agent.types.ts). */
  preference?: { effet: 'montrer' | 'bruit'; source: 'explicite' | 'appris' | 'defaut'; raison: string };
  /** Newsletter : lien de désinscription (en-tête List-Unsubscribe, https de préférence). */
  desinscription?: string;
  /**
   * Tonalité du mail (lecture IA, phase 3). 'negatif' / 'escalade' d'un client actif = client
   * mécontent : urgence 3 et responsable du projet prévenu (dans ses horaires).
   */
  tonalite?: InboxTonalite;
  /**
   * Brouillon de réponse préparé par l'agent (phase 3) : déposé dans le fil (dossier Brouillons,
   * jamais envoyé, jamais de signature) sous double verrou, sinon seulement simulé (`cree: false`).
   */
  brouillon?: InboxBrouillon;
  /** Comptes rendus de réunion ATLAS liés au même tiers / projet (lecture seule, phase 3). */
  reunionsLiees?: InboxReunionLiee[];
  /**
   * Remis dans « à traiter » par le suivi des engagements (phase 3) : promesse du correspondant
   * échue sans réponse de sa part (échéance en jours ouvrés).
   */
  relance?: { motif: 'engagement'; echeance: string; texte: string };
  /**
   * Actions métier proposées (phase 4, « Que faire de ce mail ? », `src/services/inbox-actions`),
   * calculées sans IA (arbitrage Haiku unique en cas de doute) et mises en cache ici ; jamais une
   * action déjà faite pour ce mail. Absent = pas encore calculé.
   */
  actionsProposees?: InboxActionMetierProposee[];
  /** Date du calcul des propositions ; `actionsProposeesSansIa` : doute laissé sans arbitrage IA (hors horaires, budget). */
  actionsProposeesLe?: string;
  actionsProposeesSansIa?: boolean;
  /** Actions métier faites sur ce mail (jamais reproposées ; annulables 30 jours par `actionId`). */
  actionsFaites?: InboxActionMetierFaite[];
  /**
   * Phase 5 (boîtes PARTAGÉES seulement, good@) : assignation, « pris par », commentaires internes,
   * escalade. Tenu par le worker à part (worker/inbox-agent/equipe-store.ts, jamais dans le mail) et
   * joint à l'état quand un MEMBRE autorisé de la boîte le lit. Absent sur une boîte personnelle.
   */
  equipe?: InboxEquipe;
  /**
   * Phase 6 : détections SANS IA après lecture (worker/inbox-agent/detections.ts) : avis de non-remise
   * ou départ annoncé (`rebond`), réponse automatique avec date de retour (`absence` : relances
   * décalées au lendemain ouvré du retour), mail de candidature relevé pour l'effacement RGPD
   * (`candidat`). Absent = rien de détecté.
   */
  detections?: InboxDetections;
}

// ── Phase 6 : détections après lecture ────────────────────────────────────────────

/** Avis de non-remise ou départ annoncé (réponse automatique). */
export interface InboxDetectionRebond {
  type: 'adresse-invalide' | 'depart' | 'boite-pleine' | 'temporaire';
  /** Adresse qui ne reçoit plus (absente si illisible). */
  adresse?: string;
  /** Remplaçant indiqué par le message (jamais créé sans validation). */
  nouveauContact?: { email: string; nom?: string };
  /** Changement durable (adresse invalide, départ) : fiche contact à revoir. */
  definitif: boolean;
  /** Fiche contact ATLAS : 'fait' (signalée), 'simule' (double verrou fermé), 'sans-fiche', 'non-autorise' (autonomie). */
  signale?: 'fait' | 'simule' | 'sans-fiche' | 'non-autorise';
}

export interface InboxDetections {
  rebond?: InboxDetectionRebond;
  /** Absence du correspondant : premier jour de retour ('YYYY-MM-DD'). */
  absence?: { retourLe: string };
  /** Mail de candidature (RGPD : effacement au terme si le candidat n'est pas retenu). */
  candidat?: boolean;
}

// ── Phase 5 : good@ et travail d'équipe ───────────────────────────────────────────

/** Commentaire interne sur un mail de good@ : invisible du client, visible des seuls membres autorisés. */
export interface InboxEquipeCommentaire {
  id: string;
  /** Adresse de l'auteur (identité du jeton). */
  auteur: string;
  auteurNom?: string;
  texte: string;
  at: string;
}

/** Un niveau d'escalade (1 = responsable / assigné, 2 = direction), une fois par mail et par niveau. */
export interface InboxEquipeEscaladeNiveau {
  niveau: 1 | 2;
  /** Instant où le niveau est devenu dû. */
  at: string;
  dest: string[];
  /** Cloche livrée (dans les horaires du destinataire) ; false = en attente de ses horaires, ou simulée. */
  notifie: boolean;
  /** Double verrou fermé : rien n'est envoyé, compté dans le rapport à blanc. */
  simule: boolean;
}

export interface InboxEquipe {
  /** Collègue à qui le mail revient (responsable du projet / du client), ou attribué par un membre. */
  assigneA?: string;
  assigneNom?: string;
  /** 'agent' (règles, sans IA) ou adresse du membre qui a attribué. */
  assignePar: 'agent' | string;
  assigneLe?: string;
  /** Pourquoi (« Responsable du projet #530 Lunex », « À attribuer : aucun projet lié »). */
  motif?: string;
  /** « Je prends » : même principe que TACTIK (jamais écrasé par quelqu'un d'autre ; « Relâcher » le vide). */
  prisPar?: string;
  prisParNom?: string;
  prisLe?: string;
  commentaires?: InboxEquipeCommentaire[];
  /** Délai de réponse et escalade si personne ne prend ni ne répond. */
  escalade?: {
    /** Délai cible en heures ouvrées de l'agence (client actif : 4 h ; autres : 9 h = 1 jour ouvré). */
    delaiHeures: number;
    /** Échéance de la réponse (ISO) ; niveau 2 à deux fois le délai. */
    echeance: string;
    niveaux?: InboxEquipeEscaladeNiveau[];
  };
  /** Catégorie « ATLAS · Pour <Prénom> » posée dans good@ (double verrou ouvert), sinon absente. */
  categorie?: string;
  majLe?: string;
}

/** Avertissement de collision (jamais bloquant) : quelqu'un d'autre répond déjà au même mail. */
export interface InboxEquipeCollision {
  niveau: 'info' | 'avertissement';
  message: string;
  /** Adresses repérées comme répondant (envoi ou brouillon), si connues. */
  auteurs: string[];
}

/** Réponse de `GET /api/plugin/agent/equipe?messageId=&mailbox=` (membres de good@ seulement). */
export interface InboxEquipeVue {
  mailbox: string;
  messageId: string;
  equipe: InboxEquipe;
  collision: InboxEquipeCollision | null;
  /** Membres autorisés de la boîte (choix « Attribuer à… »). */
  membres: Array<{ email: string; nom: string }>;
  moi: string;
  mode: InboxAgentMode;
}

/**
 * Action métier proposée (phase 4) : même forme que `ActionProposee` de
 * `src/services/inbox-actions/types.ts` (le type est un `TypeAction` de cette bibliothèque).
 */
export interface InboxActionMetierProposee {
  type: string;
  libelle: string;
  confiance: number;
  donnees: Record<string, unknown>;
  cible?: { table: string; id?: string; lien?: string };
  apercu: string;
  avertissements?: string[];
}

/** Action métier faite sur un mail (phase 4) : catégorie Outlook « ATLAS : … fait » sous double verrou. */
export interface InboxActionMetierFaite {
  type: string;
  /** Identifiant de l'action (journal de l'agent, « Annuler ») : `a_…`. */
  actionId: string;
  resume: string;
  lien?: string;
  at: string;
  /** Réversible (« Annuler » 30 jours) ? */
  annulable?: boolean;
  annuleLe?: string;
}

/** Tonalité d'un mail (lecture IA, phase 3). */
export type InboxTonalite = 'neutre' | 'positif' | 'negatif' | 'escalade';

/** Pourquoi un brouillon : réponse à un mail, relance d'une promesse reçue échue, relance d'un envoi sans réponse. */
export type InboxMotifBrouillon = 'reponse' | 'relance-engagement' | 'relance-envoi';
/** Sort d'un brouillon déposé (mesure, phase 3) ; 'non_utilise' = resté dans Brouillons plus de 30 jours. */
export type InboxIssueBrouillon = 'tel_quel' | 'corrige' | 'jete' | 'non_utilise';

export interface InboxBrouillon {
  /** Identifiant Graph du brouillon déposé ('' tant que rien n'est déposé : mode à blanc). */
  graphId: string;
  /** true = déposé dans la boîte (double verrou ouvert) ; false = simulé. */
  cree: boolean;
  mode: InboxAgentMode;
  /** Intention de la réponse en une ligne (« Confirmer le RDV de jeudi 14 h ») : jamais le texte. */
  resumeIntention: string;
  motif?: InboxMotifBrouillon;
  prepareLe?: string;
  /** Préparé en lot (Batch API) hors horaires de la personne : jamais notifié. */
  enLot?: boolean;
  /** Alertes du contrôle avant envoi (src/utils/client-send-check.ts) sur le brouillon. */
  alertes?: Array<{ code: string; message: string; gravite: 'bloquant' | 'avertissement' }>;
  /** Mesure : envoyé tel quel, corrigé, jeté (absent tant que rien n'est constaté). */
  issue?: InboxIssueBrouillon;
  /** Phase 6 : non déposé parce que l'autonomie « brouillon » de la personne s'arrête à « préparer ». */
  simuleMotif?: 'autonomie';
}

export interface InboxReunionLiee { id: string; titre: string; date: string }

/**
 * Vue PAR CONVERSATION (phase 1.3 / 3) : dernier message entrant et sortant de chaque fil
 * (conversationId), tenue par le worker à partir des mails lus et du relevé des éléments envoyés.
 * « La personne a écrit en dernier » = elle attend une réponse (pile « en attente »).
 */
export interface InboxConversation {
  conversationId: string;
  sujet?: string;
  dernierEntrant?: { at: string; de: string; messageId: string };
  dernierSortant?: { at: string; a: string[]; messageId: string };
  /** Dernier mail reçu de chaque correspondant du fil (adresse → date ISO). */
  parExpediteur?: Record<string, string>;
}

export type InboxCorrespondant = 'direction' | 'client' | 'partenaire';

/** Réponse de `GET /api/plugin/agent/state` quand le message n'a pas encore été lu (HTTP 404). */
export interface InboxPendingState { status: 'pending' }

/** Réponse de `GET /api/plugin/agent/journee` (bandeau « Ma journée »). */
export interface InboxJournee {
  aTraiter: number;
  enAttente: number;
  relancesDues: number;
  clientsPlus48h: number;
  pourInfo: number;
  bruit: number;
  /** Premiers mails d'expéditeurs inconnus en attente de décision (pile « À filtrer », 14 jours). */
  aFiltrer: number;
  /** Phase 3 : mails « répondre plus tard » encore masqués. */
  plusTard?: number;
  /** Phase 5 (membres de good@) : mails de la boîte partagée « à attribuer » et « pour toi ». */
  equipe?: { aAttribuer: number; pourMoi: number; boites: string[] };
}

/** Élément de `GET /api/plugin/agent/liste?pile=…` (piles visibles dans le complément et ATLAS). */
export type InboxListeElement = Pick<InboxMessageState, 'messageId' | 'graphId' | 'from' | 'subject' | 'receivedAt' | 'pile' | 'resume' | 'urgence' | 'categorie'> & {
  webLink?: string;
  /** « Répondre plus tard » : masqué de « à traiter » jusqu'à cette date (pile `plus_tard`), puis remis en tête. */
  plusTardJusqua?: string;
  /** Un brouillon de l'agent est prêt (déposé, ou simulé à blanc). */
  brouillonPret?: boolean;
  /** « En attente » : relance prévue ce jour-là ('YYYY-MM-DD') si aucune réponse n'arrive. */
  relanceLe?: string;
  /** « En attente » : mail d'un flux ATLAS (offre, portail Créas…) dont ATLAS gère déjà les rappels : « Relance gérée par ATLAS (…) », jamais de relance de l'agent. */
  relanceGeree?: string;
  /** Phase 5 (listes de good@) : boîte du mail, assigné / preneur. */
  mailbox?: string;
  equipe?: Pick<InboxEquipe, 'assigneA' | 'assigneNom' | 'prisPar' | 'prisParNom'>;
};

/**
 * Piles demandables à `GET /api/plugin/agent/liste` : les piles du tri + « plus tard » (phase 3) ;
 * phase 5 : mails de good@ « à attribuer » et « pour toi » (membres autorisés seulement).
 */
export type InboxListePile = InboxPile | 'plus_tard' | 'equipe_a_attribuer' | 'equipe_pour_moi';

/** Réponse de `POST /api/plugin/agent/plus-tard` et `POST /api/plugin/agent/relancer` (phase 3). */
export interface InboxRappelReponse {
  ok: boolean;
  messageId: string;
  /** Plus tard : ISO de la remise en tête ; relancer : jour de la relance ('YYYY-MM-DD'). */
  jusqua?: string;
  relanceLe?: string;
  /** Catégorie « ATLAS · Plus tard » posée dans la boîte (mode actif seulement). */
  ecrit: boolean;
  mode: InboxAgentMode;
}

/** Réponse de `POST /api/plugin/agent/expediteur` (filtre des nouveaux expéditeurs). */
export interface InboxDecisionExpediteur {
  ok: boolean;
  email: string;
  decision: 'accepter' | 'refuser';
  /** Mails « À filtrer » de cet expéditeur concernés par la décision. */
  mailsConcernes: number;
  /** Mails réellement déplacés / recatégorisés (mode actif seulement ; 0 à blanc). */
  ecrits: number;
  mode: InboxAgentMode;
}

/**
 * Action d'écriture de l'agent dans une boîte (phase 2) : exécutée (mode actif ET boîte dans
 * INBOX_AGENT_ACTIVE_MAILBOXES) ou seulement simulée (« ce que l'agent aurait fait »).
 */
/** 'metier' (phase 4) : action métier écrite dans ATLAS (pas dans la boîte), voir `metier`. */
export type InboxActionType = 'categorie' | 'ranger' | 'metier';
export type InboxMotifRangement = 'copie-lue' | 'bruit' | 'projet' | 'expediteur-refuse';

export interface InboxAgentAction {
  /** Identifiant (aussi celui de l'entrée du journal de la personne) : `a_…`. */
  id: string;
  mailbox: string;
  messageId: string;
  type: InboxActionType;
  /** Rangement : pourquoi. */
  motif?: InboxMotifRangement;
  /** Phrase lisible, sans contenu de mail (« Rangé dans ATLAS · Bruit »). */
  libelle: string;
  /** Avant / après : dossier (id + nom) pour un rangement, catégories pour une catégorie. */
  avant: { dossierId?: string; dossierNom?: string; categories?: string[] };
  apres: { dossierId?: string; dossierNom?: string; categories?: string[] };
  /** Identifiant Graph du message avant / après l'action (il change à chaque déplacement). */
  graphIdAvant: string;
  graphIdApres?: string;
  /** true = écrit dans la boîte ; false = simulé (mode à blanc ou boîte hors liste active). */
  executee: boolean;
  /** Simulée parce que l'AUTONOMIE de la personne ne le permet pas (phase 6), et pas seulement le double verrou. */
  simuleMotif?: 'autonomie';
  at: string;
  annuleLe?: string;
  /** Action métier (type 'metier', phase 4) : ce qui a été fait dans ATLAS et de quoi l'annuler. */
  metier?: InboxActionMetierTrace;
}

/** Trace d'une action métier (phase 4) dans le journal des actions de l'agent. */
export interface InboxActionMetierTrace {
  /** `TypeAction` de `src/services/inbox-actions`. */
  type: string;
  resume: string;
  cible?: { table: string; id?: string; lien?: string };
  /** Données d'annulation renvoyées par l'action (identifiants ATLAS, valeurs d'avant). */
  annulation?: Record<string, unknown>;
  annulable: boolean;
  /** 'clic' = décision explicite de la personne dans le complément ; 'agent' = déclenchée par l'agent. */
  declencheur: 'clic' | 'agent';
  /** Personne qui a déclenché l'action (identité du jeton) : son journal porte l'action. */
  personne?: string;
  /** Catégorie « ATLAS : … fait » réellement posée dans la boîte (double verrou ouvert). */
  categorieEcrite: boolean;
  categorie: string;
}

/** Résumé quotidien des newsletters d'une personne (section du brief du matin, sans IA par défaut). */
export interface InboxNewsletterDigest {
  mailbox: string;
  /** Jour du brief ('YYYY-MM-DD') ; couvre les newsletters reçues depuis le jour ouvré précédent. */
  jour: string;
  periode: { du: string; au: string };
  elements: Array<{ titre: string; expediteur: string; nom?: string; receivedAt: string; webLink?: string }>;
  /** Résumé IA en lot (INBOX_AGENT_NEWSLETTER_IA=1), préparé hors horaires ; absent sinon. */
  resumeIa?: string;
  /** Newsletters jamais ouvertes depuis 30 jours : désinscription PROPOSÉE (jamais automatique). */
  desinscriptions: Array<{ expediteur: string; nom?: string; lien?: string; recus30j: number }>;
}

/** « Ce que l'agent aurait fait » (rapport à blanc) ou a fait (mode actif), par type d'action. */
export interface InboxRapportActions {
  categories: { total: number; parPile: Record<InboxPile, number> };
  rangements: { total: number; parMotif: Record<InboxMotifRangement, number>; dossiers: Array<{ dossier: string; n: number }> };
  filtrage: { aFiltrer: number; acceptes: number; refuses: number };
  executees: number;
  simulees: number;
  annulees: number;
  /** Phase 4 : actions métier (faites sur clic, simulées sinon hors double verrou), propositions calculées. */
  metier?: {
    proposees: number;
    mailsAvecPropositions: number;
    faites: number;
    simulees: number;
    annulees: number;
    categoriesEcrites: number;
    parType: Record<string, number>;
  };
}

/** Ligne de l'échantillon du rapport à blanc (validation du classement par Charles). */
export interface InboxRapportEchantillon {
  receivedAt: string;
  mailbox: string;
  expediteur: string;
  objet: string;
  pile: InboxPile;
  categorie: string;
  source: 'regles' | 'ia';
  raison: string;
}

/** Réponse de `GET /api/atlas/worker/inbox-agent/report` (mode à blanc, phase 1.7). */
export interface InboxAgentRapport {
  genereLe: string;
  mode: InboxAgentMode;
  periode: { du: string; au: string; joursOuvres: number };
  total: number;
  sansIa: { n: number; part: number };
  parPile: Record<InboxPile, number>;
  parBoite: Array<{ mailbox: string; total: number; piles: Record<InboxPile, number>; sansIa: number; coutEur: number }>;
  parRegle: Array<{ regle: string; libelle: string; n: number }>;
  cout: {
    reelEur: number;
    reelParJourOuvreEur: number;
    projeteParJourOuvreEur: number;
    objectifParJourOuvreEur: number;
    lecturesIa: number;
    coutMoyenLectureEur: number;
    partEnLot: number;
    hypotheses: string;
  };
  echantillon: InboxRapportEchantillon[];
  /** Phase 2 : écritures de l'agent (simulées à blanc : « ce que l'agent aurait fait »). */
  actions?: InboxRapportActions;
  /** Boîtes où l'agent écrit vraiment (INBOX_AGENT_MODE=actif ET INBOX_AGENT_ACTIVE_MAILBOXES). */
  boitesActives?: string[];
  /** Phase 3 : brouillons (simulés à blanc), engagements, relances, plus tard, clients mécontents. */
  phase3?: InboxRapportPhase3;
  /** Phase 5 : good@ et équipe (assignation, prises, commentaires, escalades). */
  phase5?: InboxRapportPhase5;
  /**
   * Phase 6 : écritures automatiques simulées parce que l'autonomie réglée par la personne ne les
   * permet pas (motif « autonomie »), par action ('trier', 'ranger', 'brouillon', 'tache', 'rappel',
   * 'action-metier'). Agrégé : aucun réglage individuel n'y figure.
   */
  autonomie?: { motif: 'autonomie'; refus: number; parAction: Record<string, number> };
}

/** Rapport de la phase 3 (« ce que l'agent aurait fait » à blanc, mesure en mode actif). */
export interface InboxRapportPhase3 {
  brouillons: {
    total: number;
    deposes: number;
    simules: number;
    enLot: number;
    parMotif: Record<InboxMotifBrouillon, number>;
    /** Brouillons avec au moins une alerte du contrôle avant envoi. */
    avecAlertes: number;
    /** Mesure (brouillons déposés) : envoyés tels quels, corrigés, jetés, non utilisés, en attente. */
    issues: Record<InboxIssueBrouillon | 'en_attente', number>;
    coutEur: number;
    /** Derniers brouillons (intention seulement, jamais le texte). */
    exemples: Array<{ mailbox: string; objet: string; motif: InboxMotifBrouillon; resumeIntention: string; cree: boolean }>;
  };
  engagements: { recus: number; recusEchus: number; faits: number; taches: { creees: number; simulees: number } };
  relances: { prevues: number; dues: number; preparees: number; manuelles: number; pasDeRelance: number };
  plusTard: number;
  clientsMecontents: { detectes: number; notifies: number; simules: number };
  reunionsLiees: number;
}

/** Rapport de la phase 5 (good@ et équipe) : agrégé, sans contenu de mail ni commentaire. */
export interface InboxRapportPhase5 {
  mails: number;
  assignesParAgent: number;
  aAttribuer: number;
  attribuesParMembre: number;
  pris: number;
  commentaires: number;
  escalades: { niveau1: number; niveau2: number; notifiees: number; simulees: number; enAttenteHoraires: number };
  categories: { ecrites: number };
}
