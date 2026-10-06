/**
 * inbox-actions.types.ts — COPIE LOCALE du contrat « Que faire de ce mail ? » (phase 4 de l'agent
 * d'inbox) : `TypeAction` et `ActionProposee` repris tels quels de
 * `src/services/inbox-actions/types.ts` (racine du dépôt). Le complément est publié depuis son
 * propre dossier et n'importe rien hors de `outlook-addin/` : toute évolution du fichier racine
 * doit être recopiée ici.
 *
 * Plus les réponses des routes du worker (contrat fixé le 04/10/2026) :
 *   GET  /api/plugin/agent/actions?messageId=&mailbox=        → ReponseActions
 *   POST /api/plugin/agent/actions/executer { messageId, mailbox?, type, donnees?, choix? } → ResultatExecution
 *   POST /api/plugin/agent/annuler { actionId }                 (route existante du journal)
 *   POST /api/plugin/agent/question { question, mailbox? }      → ReponseQuestion
 *   GET  /api/plugin/agent/resume-fil?conversationId=&mailbox=  → ResumeFil
 *   GET  /api/plugin/agent/rattrapage?depuis=ISO&mailbox=       → Rattrapage
 */

// ── Copie de src/services/inbox-actions/types.ts (contrat fixé) ──

export type TypeAction =
  | 'facture'
  | 'facture-projet'
  | 'demande-client'
  | 'rdv'
  | 'signature-contact'
  | 'reponse-prospect'
  | 'candidature'
  | 'devis-fournisseur'
  | 'accord-client'
  | 'bat'
  | 'changement-projet'
  | 'pieces-jointes'
  | 'nouveau-contact-copie'
  | 'confirmation-fournisseur'
  | 'infos-lieu'
  | 'reservation-studio'
  | 'hotesse'
  | 'impaye-fournisseur'
  | 'mention-presse'
  | 'chrono'
  // Cadrage Outlook du 06/10/2026 (section A).
  | 'ranger-projet'
  | 'facture-hors-projet'
  | 'newsletter-interessante'
  | 'newsletter-sans-interet'
  | 'mandat-associatif'
  | 'modele-appris';

export interface ActionProposee {
  type: TypeAction;
  libelle: string;
  confiance: number;
  donnees: Record<string, unknown>;
  cible?: { table: string; id?: string; lien?: string };
  apercu: string;
  avertissements?: string[];
}

// ── Réponses des routes du worker ──

/**
 * Action déjà exécutée sur ce mail (affichée en grisé, jamais reproposée). Même forme que
 * `InboxActionMetierFaite` (inbox-agent.types.ts, `etat.actionsFaites`).
 */
export interface ActionFaite {
  type: TypeAction | string;
  actionId: string;
  resume: string;
  lien?: string;
  at: string;
  /** Réversible (« Annuler » 30 jours) ? */
  annulable?: boolean;
  annuleLe?: string;
}

export interface ReponseActions {
  actions: ActionProposee[];
  faites: ActionFaite[];
}

export interface ResultatExecution {
  ok: boolean;
  resume: string;
  cible?: { table: string; id?: string; lien?: string };
  actionId?: string;
  /** 'choix-requis' : un choix humain (projet…) est nécessaire avant d'exécuter (jamais deviné). */
  erreur?: string;
  /** true = mode à blanc : rien n'a été écrit, l'action est seulement simulée. */
  simule?: boolean;
  /** 07/10/2026 : dossier Outlook absent (projet, mandat) : chemin proposé à la création. */
  dossierACreer?: { chemin: string; projetId?: string; mandatId?: string };
}

/** Projet candidat proposé par la détection (`donnees.candidats`, inbox-actions/commun.ts). */
export interface CandidatChoix {
  id: string;
  libelle: string;
  score?: number;
  signaux?: string[];
}

export interface SourceMail {
  messageId: string;
  subject: string;
  from: string | { email?: string; name?: string };
  date: string;
  webLink?: string;
}

export interface ReponseQuestion {
  reponse: string;
  sources: SourceMail[];
}

export interface ResumeFil {
  lignes: string[];
  nbMails: number;
  misAJour: string;
}

export interface Rattrapage {
  depuis: string;
  intro?: string;
  sections: Array<{ titre: string; elements: Array<SourceMail & { note?: string }> }>;
}
