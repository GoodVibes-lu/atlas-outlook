/**
 * reprise.ts · REPRISE du tableau de bord (demande de Charles du 08/10/2026). La fenêtre du tableau
 * (dialogue Office) se ferme chaque fois qu'Outlook doit montrer quelque chose devant (mail ouvert,
 * réponse) ; rouverte dans les 30 minutes, elle revient là où la personne était : même boîte, même
 * section, même mail choisi (ou le suivant si celui-là a été traité ou rangé), séance de tri à la
 * même carte, et la carte « Traités à ranger » du mail auquel elle vient de répondre.
 * État gardé dans le localStorage de la page (même origine), petit JSON daté, toujours sous try/catch.
 * Fonctions de choix pures (testées par scripts/test-inbox-traites.mjs) ; le DOM reste dans app.ts.
 */

export const REPRISE_MAX_MS = 30 * 60_000;
const CLE = 'atlas.tdb.reprise';

export interface RepriseSeance {
  /** Carte affichée (messageId), null = carte spéciale ou fin de séance. */
  messageId: string | null;
  /** Cartes qui suivaient (messageIds) : si la carte a disparu, la première encore présente prend le relais. */
  suivants: string[];
}

export interface Reprise {
  at: number;
  boite: string;
  section: string;
  /** Clé du mail choisi (`${mailbox}|${messageId}`), null = « Aujourd'hui ». */
  sel: string | null;
  /** Mails qui suivaient le mail choisi dans sa section (clés), pour reprendre au suivant. */
  suivants: string[];
  /** Séance de tri ouverte à la fermeture. */
  seance?: RepriseSeance | null;
}

/** Reprise encore valable (moins de 30 minutes) ? Pur. */
export function repriseValide(r: Reprise | null | undefined, now = Date.now(), maxMs = REPRISE_MAX_MS): r is Reprise {
  return !!r && typeof r.at === 'number' && now - r.at >= 0 && now - r.at <= maxMs && typeof r.section === 'string';
}

/**
 * Mail à rechoisir (pur) : le même s'il est encore listé, sinon le premier de ceux qui le suivaient
 * encore présent, sinon null (« Aujourd'hui »).
 */
export function selectionReprise(sel: string | null, suivants: readonly string[], presents: readonly string[]): string | null {
  if (!sel) return null;
  const p = new Set(presents);
  if (p.has(sel)) return sel;
  return suivants.find(s => p.has(s)) || null;
}

/** Les N clés qui suivent `cle` dans `lignes` (pur). */
export function suivantsDe(cle: string | null, lignes: readonly string[], n = 6): string[] {
  if (!cle) return [];
  const i = lignes.indexOf(cle);
  return i < 0 ? [] : lignes.slice(i + 1, i + 1 + n);
}

/** Index de la carte à reprendre dans une séance (pur) : la même, sinon la première suivante présente, sinon 0. */
export function indexReprise(s: RepriseSeance | null | undefined, messageIds: readonly string[]): number {
  if (!s) return 0;
  const cible = selectionReprise(s.messageId, s.suivants, messageIds);
  const i = cible ? messageIds.indexOf(cible) : -1;
  return i < 0 ? 0 : i;
}

/** messageId d'une clé `${mailbox}|${messageId}` (pur). */
export function messageIdDeCle(cle: string | null): string {
  if (!cle) return '';
  const i = cle.indexOf('|');
  return i < 0 ? cle : cle.slice(i + 1);
}

// ── Stockage (localStorage de la page, jamais bloquant) ──

export function lireReprise(now = Date.now()): Reprise | null {
  try {
    const raw = localStorage.getItem(CLE);
    if (!raw) return null;
    const r = JSON.parse(raw) as Reprise;
    if (!repriseValide(r, now)) { localStorage.removeItem(CLE); return null; }
    return { ...r, suivants: Array.isArray(r.suivants) ? r.suivants.slice(0, 10) : [], seance: r.seance && typeof r.seance === 'object' ? { messageId: r.seance.messageId ?? null, suivants: Array.isArray(r.seance.suivants) ? r.seance.suivants.slice(0, 10) : [] } : null };
  } catch { return null; }
}

export function sauverReprise(r: Omit<Reprise, 'at'>, now = Date.now()): void {
  try { localStorage.setItem(CLE, JSON.stringify({ ...r, at: now, suivants: r.suivants.slice(0, 10) })); } catch { /* stockage indisponible */ }
}

export function effacerReprise(): void {
  try { localStorage.removeItem(CLE); } catch { /* rien */ }
}
