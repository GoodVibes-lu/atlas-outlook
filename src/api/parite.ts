/**
 * parite.ts · Routes worker de la fin de PARITÉ avec l'Inbox ATLAS (07/10/2026), côté complément :
 * `/api/plugin/agent/parite/*` (worker/inbox-agent/routes-parite.ts). Même jeton et mêmes erreurs
 * lisibles que les autres routes de l'agent (workerRequest). Aucune règle métier ici : le worker
 * applique les services d'ATLAS.
 */
import { workerRequest } from './worker';
import { AtlasError } from './net';

const lire = <T>(path: string) => workerRequest<T>('GET', `agent/parite/${path}`, undefined, { retry: true, timeoutMs: 60_000 });
const ecrire = <T>(path: string, corps: unknown, timeoutMs = 60_000) => workerRequest<T>('POST', `agent/parite/${path}`, corps, { timeoutMs });
const qs = (o: Record<string, string | number | undefined>) => {
  const q = new URLSearchParams();
  for (const [k, v] of Object.entries(o)) if (v !== undefined && v !== '') q.set(k, String(v));
  return q.toString();
};

/** Données jointes à une erreur du worker (ex. doublons d'un tiers), ou null. */
export function donneesErreur(e: unknown): Record<string, any> | null {
  return e instanceof AtlasError && e.data && typeof e.data === 'object' ? e.data as Record<string, any> : null;
}

// ── Contact de la demande de devis ──

export interface EcartContact { senderEmail: string; originalEmail: string; senderName?: string; newContact?: { id: string; name: string; societe?: string; email?: string } }

export async function fetchEcartContactDevis(messageId: string, mailbox: string | undefined, devisId: string): Promise<EcartContact | null> {
  const r = await lire<{ ecart?: EcartContact | null }>(`devis-contact?${qs({ messageId, mailbox, devisId })}`);
  return r.ecart || null;
}
export const rattacherContactDevis = (messageId: string, mailbox: string | undefined, devisId: string, contactId: string) =>
  ecrire<{ ok: boolean; contact: { id: string; name: string } }>('devis-contact', { messageId, mailbox, devisId, contactId });
export const annulerContactDevis = (messageId: string, mailbox: string | undefined, devisId: string) =>
  ecrire<{ ok: boolean }>('devis-contact', { messageId, mailbox, devisId, annuler: true });

// ── Import des mails d'un dossier ──

export interface LotImport { importes: number; deja: number; traites: number; total: number | null; suivant: number | null; dossier: { id: string; chemin: string } }
export const importerLotDossier = (p: { mailbox?: string; projetId: string; dossierId: string; curseur: number }) =>
  ecrire<LotImport>('import-dossier', p, 120_000);

// ── Reclasser ──

export async function fetchProjetDuMail(messageId: string, mailbox?: string): Promise<{ id: string; libelle: string } | null> {
  const r = await lire<{ projet?: { id: string; libelle: string } | null }>(`reclasser?${qs({ messageId, mailbox })}`);
  return r.projet || null;
}
export interface ReponseReclassement { delies: number; deplacer: 'non' | 'ranger' | 'creer'; dossierVers?: { id: string; chemin: string }; dossierACreer?: string }
export const reclasserMail = (p: { messageId: string; mailbox?: string; deProjetId: string; versProjetId: string; deplacer?: boolean }) =>
  ecrire<ReponseReclassement>('reclasser', p);

// ── Pièces jointes vers le projet ──

export interface ContextePiecesProjet { destinations: Array<{ cle: string; libelle: string }>; pieces: Array<{ id: string; nom: string; taille: number }>; projetId: string | null }
export const fetchPiecesProjet = (messageId: string, mailbox?: string) => lire<ContextePiecesProjet>(`pieces-projet?${qs({ messageId, mailbox })}`);
export const joindreAuProjet = (p: { messageId: string; mailbox?: string; projetId: string; champ: string; pieceJointeIds: string[] }) =>
  ecrire<{ joints: number; champ: string; refus: Array<{ id: string; motif: string }> }>('pieces-projet', p, 120_000);

// ── Proposer un RDV ──

export type LieuRdv = 'onsite_gv' | 'onsite_client' | 'teams' | 'phone';
export const LIBELLES_LIEUX: Record<LieuRdv, string> = { onsite_gv: 'Chez nous', onsite_client: 'Chez le client', teams: 'Teams (visio)', phone: 'Téléphone' };
export interface CreneauRdv { debut: number; fin: number; libelle: string; propose: boolean; conflit: boolean }
export const fetchCreneauxRdv = (messageId: string, mailbox: string | undefined, duree: number) =>
  lire<{ creneaux: CreneauRdv[]; agendaLu: boolean; duree: number; durees: number[] }>(`rdv?${qs({ messageId, mailbox, duree })}`);
export const propositionRdv = (p: { messageId: string; mailbox?: string; lieu: LieuRdv; duree: number; creneaux: number[] }) =>
  ecrire<{ texte: string; html: string }>('rdv', p);

// ── Nouveau tiers ──

export interface SaisieTiersOutlook {
  nom: string; prenom?: string; categories: string[]; particulier?: boolean; pays: string; email?: string; tel?: string; tva?: string;
  adresse?: string; cp?: string; ville?: string; web?: string; matricule?: string; nbEmployes?: string; anneeCreation?: string;
  contact?: { prenom?: string; nom?: string; genre?: string; fonction?: string; langue?: string; email?: string; tel?: string; gsm?: string };
}
export interface PreparationTiers {
  erreurs: string[];
  doublons: Array<{ id: string; nom: string; score: number; categories: string[] }>;
  candidatsBob: Array<{ refBOB: string; nom: string; reason: string; score: number }>;
  codePropose: string;
  options: { categories: string[]; nbEmployes: string[]; genres: string[]; langues: string[] };
}
export const preparerTiers = (saisie: SaisieTiersOutlook) => ecrire<PreparationTiers>('tiers/preparer', { saisie });
export const creerTiers = (saisie: SaisieTiersOutlook, o: { refBob?: string; doublonsVus?: boolean }) =>
  ecrire<{ id: string; nom: string; codeBob: string; contactCree: boolean; lieuCree: boolean; aCompleter?: string[]; lien: string }>('tiers/creer', { saisie, ...o });

// ── Prospection ──

export const lierProspection = (messageId: string, mailbox?: string) =>
  ecrire<{ contact: { id: string; nom: string; tiers: string } | null; chemin?: string }>('prospection', { messageId, mailbox });
export const CHEMIN_PROSPECTION = 'Inbox/Prospection';
