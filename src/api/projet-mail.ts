/**
 * projet-mail.ts · projet du mail sur clic (lots 2 et 3 du complément, 10/10/2026) : « Lier et classer »,
 * « Changer de projet », « Délier », recherche de projets et autonomie de la personne. Tout passe par
 * le worker (aucun jeton de boîte) ; rien n'est envoyé chez un client.
 */
import { workerRequest } from './worker';

export type NiveauAutonomie = 'suggerer' | 'preparer' | 'faire';

/** Vue préparée par l'agent à l'arrivée du mail (GET agent/vue › pret). */
export interface VuePrete {
  projet: { id: string; libelle: string; numero?: number; source: string; confiance: 'forte' | 'moyenne' | 'faible'; raison: string } | null;
  candidats?: Array<{ id: string; libelle: string; numero?: number }>;
  raison?: string;
  delie?: boolean;
  lie?: boolean;
  dossier?: { existant: { id: string; chemin: string } | null; propose: string | null } | null;
  auto?: { lie?: boolean; range?: boolean; actionRangementId?: string; categorie?: boolean; motif?: string; at: string };
  autonomie?: { lier: NiveauAutonomie; ranger: NiveauAutonomie; trier: NiveauAutonomie } | null;
  prepareLe: string;
}

export interface ResultatLiaison {
  projet: { id: string; libelle: string };
  rangement: { dossier?: { id: string; chemin: string }; cree?: boolean; actionId?: string } | null;
  rangementErreur?: string;
  prochainsSeuls: boolean;
}

const lireResultat = (r: any): ResultatLiaison => ({
  projet: { id: String(r?.projet?.id || ''), libelle: String(r?.projet?.libelle || '') },
  rangement: r?.rangement && typeof r.rangement === 'object' ? r.rangement : null,
  ...(r?.rangementErreur ? { rangementErreur: String(r.rangementErreur) } : {}),
  prochainsSeuls: r?.prochainsSeuls === true,
});

export async function lierEtClasser(p: { messageId: string; projetId: string; mailbox?: string; chemin?: string; ranger?: boolean }): Promise<ResultatLiaison> {
  return lireResultat(await workerRequest<any>('POST', 'agent/projet/lier-classer', p, { timeoutMs: 45_000 }));
}

export async function changerProjet(p: { messageId: string; projetId: string; ancienProjetId?: string; mailbox?: string; ranger?: boolean }): Promise<ResultatLiaison> {
  return lireResultat(await workerRequest<any>('POST', 'agent/projet/changer', p, { timeoutMs: 45_000 }));
}

export async function delierProjet(p: { messageId: string; projetId?: string; conversationId?: string; actionRangementId?: string; mailbox?: string }): Promise<{ rangementAnnule: boolean }> {
  const r = await workerRequest<any>('POST', 'agent/projet/delier', p, { timeoutMs: 30_000 });
  return { rangementAnnule: r?.rangementAnnule === true };
}

export async function chercherProjets(q: string): Promise<Array<{ id: string; libelle: string; client: string }>> {
  const r = await workerRequest<any>('GET', `agent/projet/projets?q=${encodeURIComponent(q.slice(0, 80))}`, undefined, { retry: true, timeoutMs: 20_000 });
  return Array.isArray(r?.projets) ? r.projets : [];
}

export interface Autonomie {
  reglages: Record<string, NiveauAutonomie>;
  plafonds: Record<string, NiveauAutonomie>;
  libelles: Record<string, { titre: string; niveaux: Partial<Record<NiveauAutonomie, string>> }>;
}

export async function lireAutonomie(): Promise<Autonomie> {
  return workerRequest<Autonomie>('GET', 'agent/projet/autonomie', undefined, { retry: true, timeoutMs: 20_000 });
}

export async function ecrireAutonomie(r: Partial<Record<string, NiveauAutonomie>>): Promise<Autonomie> {
  return workerRequest<Autonomie>('PUT', 'agent/projet/autonomie', r, { timeoutMs: 20_000 });
}
