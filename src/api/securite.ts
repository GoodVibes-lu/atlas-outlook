/**
 * securite.ts · Sécurité d'un mail (backlog reczJ0zLhPXqkWF9S, 07/10/2026), côté complément :
 * routes `/api/plugin/agent/securite*` (worker/inbox-agent/routes-securite.ts). Le complément n'a
 * plus de jeton de boîte : tout passe par le worker (lecture Graph, rangement sur clic de la
 * propriétaire). Aucune règle ici : le worker calcule le risque.
 */
import { workerRequest } from './worker';

export type NiveauSecurite = 'faible' | 'a_verifier' | 'eleve';

export interface SecuriteMail {
  niveau: NiveauSecurite;
  score: number;
  raisons: string[];
  codes: string[];
  lienOuPiece: boolean;
  protege: boolean;
  calculeLe: string;
  autresBoites?: string[];
  iban?: { iban: string; statut: 'conforme' | 'different' | 'inconnu' };
  surveillanceJusqua?: string;
  sur?: { par: string; le: string };
  quarantaine?: { par: string; le: string; actionId?: string };
}

export interface ReponseSecurite {
  mailbox: string;
  securite: SecuriteMail | null;
  protege: boolean;
  message: { from: string; fromName: string; subject: string; receivedAt: string } | null;
  quarantaineDisponible: boolean;
}

export interface CollegueAPrevenir { nom: string; email: string; langue: string; objet: string; texte: string; html: string }
export interface Prevenir {
  mailbox: string;
  langue: string;
  contact: { nom: string; email: string; societe: string | null };
  telephones: Array<{ libelle: string; numero: string }>;
  texteTelephone: string;
  sms: boolean;
  collegues: CollegueAPrevenir[];
}

const qs = (o: Record<string, string | undefined>) => {
  const q = new URLSearchParams();
  for (const [k, v] of Object.entries(o)) if (v) q.set(k, v);
  return q.toString();
};

export const fetchSecurite = (messageId: string, mailbox?: string) =>
  workerRequest<ReponseSecurite>('GET', `agent/securite?${qs({ messageId, mailbox })}`, undefined, { retry: true, timeoutMs: 45_000 });
export const marquerSur = (messageId: string, mailbox?: string) =>
  workerRequest<{ ok: boolean; securite: SecuriteMail }>('POST', 'agent/securite/sur', { messageId, mailbox });
export const mettreEnQuarantaine = (messageId: string, mailbox?: string) =>
  workerRequest<{ ok: boolean; dossier: { id: string; chemin: string }; actionId: string; surveillanceJusqua: string | null }>('POST', 'agent/securite/quarantaine', { messageId, mailbox }, { timeoutMs: 60_000 });
export const fetchPrevenir = (messageId: string, mailbox?: string) =>
  workerRequest<Prevenir>('GET', `agent/securite/prevenir?${qs({ messageId, mailbox })}`, undefined, { retry: true, timeoutMs: 60_000 });
