/**
 * vue-mail.ts · « Vue du mail » en un appel (GET /api/plugin/agent/vue, lot 1 « vitesse », 09/10/2026)
 * et dernières vues gardées sur le poste.
 *
 * Le volet faisait une quinzaine d'appels à chaque mail ouvert. Le worker rend maintenant en une
 * réponse : état de l'agent, sécurité (verdict déjà calculé), projet du mail, correspondant et mail
 * répondu à ranger. Ce qui n'était pas prêt à temps est listé dans `aCharger` : le volet le demande
 * alors par sa route habituelle.
 *
 * Dernières vues (30 mails, 3 jours) : affichées tout de suite à la réouverture d'un mail, avant même
 * la connexion, puis remplacées par la réponse fraîche. Rien de plus que ce que le volet affiche.
 */
import { workerRequest } from './worker';
import type { ReponseSecurite } from './securite';
import type { TraiteARanger } from './tableau';

export interface FicheCorrespondant {
  email: string;
  nom: string;
  societe: string;
  categorie: string;
  fonction: string;
  connu: boolean;
  interne: boolean;
  projetsEnCours: Array<{ id: string; libelle: string }>;
  ton: string;
  langue: string;
}

export interface VueMail {
  etat: { status: number; data: any };
  securite: ReponseSecurite | null;
  projet: { id: string; libelle: string; source?: 'lie' | 'appris' | 'agent' } | null;
  correspondant: FicheCorrespondant | null;
  traite: TraiteARanger | null;
  aCharger: string[];
  ms?: number;
}

export async function fetchVueMail(p: { messageId: string; conversationId?: string; from?: string; nom?: string }): Promise<VueMail> {
  const q = new URLSearchParams({ messageId: p.messageId });
  if (p.conversationId) q.set('conversationId', p.conversationId);
  if (p.from) q.set('from', p.from);
  if (p.nom) q.set('nom', p.nom.slice(0, 200));
  const r = await workerRequest<any>('GET', `agent/vue?${q.toString()}`, undefined, { retry: true, timeoutMs: 30_000 });
  return {
    etat: r?.etat && typeof r.etat === 'object' ? r.etat : { status: 404, data: { status: 'pending' } },
    securite: r?.securite || null,
    projet: r?.projet || null,
    correspondant: r?.correspondant || null,
    traite: r?.traite || null,
    aCharger: Array.isArray(r?.aCharger) ? r.aCharger.map(String) : [],
    ...(typeof r?.ms === 'number' ? { ms: r.ms } : {}),
  };
}

/** Fiche du correspondant seule (quand la vue n'a pas pu la préparer à temps). */
export async function fetchCorrespondant(email: string, nom = ''): Promise<FicheCorrespondant | null> {
  const r = await workerRequest<{ fiche?: FicheCorrespondant | null }>('POST', 'atlas/correspondant', { email, nom }, { retry: true, timeoutMs: 30_000 });
  return r?.fiche || null;
}

// ── Dernières vues gardées sur le poste ──

const CLE = 'atlas_addin_vues';
const MAX = 30;
const DUREE_MS = 3 * 86_400_000;

type Stock = Record<string, { at: number; v: VueMail }>;

function lire(): Stock {
  try { const x = JSON.parse(localStorage.getItem(CLE) || '{}'); return x && typeof x === 'object' ? x : {}; } catch { return {}; }
}

export function vueMemorisee(messageId: string): VueMail | null {
  if (!messageId) return null;
  const e = lire()[messageId];
  return e && Date.now() - e.at < DUREE_MS ? e.v : null;
}

export function memoriserVue(messageId: string, v: VueMail): void {
  if (!messageId || v.etat?.status !== 200) return; // seulement une vue complète (état de l'agent présent)
  try {
    const s = lire();
    s[messageId] = { at: Date.now(), v: { ...v, aCharger: [] } };
    const cles = Object.keys(s).sort((a, b) => s[b].at - s[a].at);
    for (const k of cles.slice(MAX)) delete s[k];
    for (const k of Object.keys(s)) if (Date.now() - s[k].at >= DUREE_MS) delete s[k];
    localStorage.setItem(CLE, JSON.stringify(s));
  } catch { /* stockage plein ou indisponible : sans mémoire */ }
}
