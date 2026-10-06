/**
 * logique.ts · fonctions PURES du tableau de bord (barre de gauche d'Outlook) : dates proposées
 * pour « mettre de côté » et « envoyer plus tard », durées lisibles, recherche de la palette ⌘K,
 * navigation au clavier, nouveautés entre deux rafraîchissements, animation des chiffres.
 * Aucun DOM, aucun réseau : testées par scripts/test-tableau-outlook.mjs.
 */

const H = 3_600_000;
const J = 86_400_000;
const TZ = 'Europe/Luxembourg';

/** Parties de date à Luxembourg (heure murale). */
export function partsLux(ms: number): { y: number; m: number; d: number; h: number; mi: number; wd: number } {
  const f = new Intl.DateTimeFormat('en-GB', { timeZone: TZ, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', weekday: 'short', hour12: false }).formatToParts(new Date(ms));
  const g = (t: string) => f.find(p => p.type === t)?.value || '';
  const wd = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'].indexOf(g('weekday'));
  return { y: +g('year'), m: +g('month'), d: +g('day'), h: +g('hour') % 24, mi: +g('minute'), wd };
}

/** Instant (ms) d'une heure murale de Luxembourg (correct aux changements d'heure). */
export function luxVersMs(y: number, m: number, d: number, h: number, mi = 0): number {
  const guess = Date.UTC(y, m - 1, d, h, mi);
  for (let k = 0, t = guess; k < 3; k++) {
    const p = partsLux(t);
    const ecart = Date.UTC(p.y, p.m - 1, p.d, p.h, p.mi) - guess;
    if (ecart === 0) return t;
    t -= ecart;
    if (k === 2) return t;
  }
  return guess;
}

export interface Preset { cle: string; libelle: string; iso: string }

/**
 * Moments proposés (pur) : ce soir 18 h (avant 17 h seulement), demain 8 h, lundi prochain 8 h
 * (sauf un dimanche, où « demain » suffit), dans une semaine 8 h. `envoi` : pour un envoi
 * programmé, « dans 1 heure » en plus. Jamais un moment passé.
 */
export function presetsMoments(now: number, opts: { envoi?: boolean } = {}): Preset[] {
  const p = partsLux(now);
  const out: Preset[] = [];
  if (opts.envoi) out.push({ cle: '1h', libelle: 'Dans 1 heure', iso: new Date(now + H).toISOString() });
  if (p.h < 17) out.push({ cle: 'soir', libelle: 'Ce soir 18 h', iso: new Date(luxVersMs(p.y, p.m, p.d, 18)).toISOString() });
  const demain = partsLux(now + J);
  out.push({ cle: 'demain', libelle: 'Demain 8 h', iso: new Date(luxVersMs(demain.y, demain.m, demain.d, 8)).toISOString() });
  const jusquaLundi = ((8 - p.wd) % 7) || 7;
  const lundi = partsLux(now + jusquaLundi * J);
  if (jusquaLundi > 1) out.push({ cle: 'lundi', libelle: 'Lundi 8 h', iso: new Date(luxVersMs(lundi.y, lundi.m, lundi.d, 8)).toISOString() });
  const sem = partsLux(now + 7 * J);
  out.push({ cle: 'semaine', libelle: 'Dans une semaine', iso: new Date(luxVersMs(sem.y, sem.m, sem.d, 8)).toISOString() });
  return out.filter(x => Date.parse(x.iso) > now + 60_000);
}

/** Valeur d'un champ <input type="datetime-local"> (heure murale de Luxembourg) → ISO. null si invalide. */
export function depuisChampDate(v: string): string | null {
  const m = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})$/.exec(String(v || '').trim());
  if (!m) return null;
  return new Date(luxVersMs(+m[1], +m[2], +m[3], +m[4], +m[5])).toISOString();
}

/** ISO → valeur de <input type="datetime-local"> (Luxembourg). */
export function versChampDate(iso: string): string {
  const p = partsLux(Date.parse(iso));
  const z = (n: number) => String(n).padStart(2, '0');
  return `${p.y}-${z(p.m)}-${z(p.d)}T${z(p.h)}:${z(p.mi)}`;
}

/** « il y a 5 min », « il y a 3 h », « hier », « lun. 6 oct. » (pur). */
export function ilYA(iso: string | undefined, now: number): string {
  const t = Date.parse(iso || '');
  if (!Number.isFinite(t)) return '';
  const d = now - t;
  if (d < 60_000) return 'à l\'instant';
  if (d < H) return `il y a ${Math.floor(d / 60_000)} min`;
  if (d < 24 * H && partsLux(t).d === partsLux(now).d) return `il y a ${Math.floor(d / H)} h`;
  if (d < 48 * H && partsLux(t).d === partsLux(now - J).d) return 'hier';
  return new Date(t).toLocaleDateString('fr-FR', { timeZone: TZ, weekday: 'short', day: 'numeric', month: 'short' });
}

/** « demain 8:00 », « lun. 13 oct. 8:00 » (pur). */
export function quandLisible(iso: string, now: number): string {
  const t = Date.parse(iso);
  if (!Number.isFinite(t)) return '';
  const heure = new Date(t).toLocaleTimeString('fr-FR', { timeZone: TZ, hour: '2-digit', minute: '2-digit' });
  const p = partsLux(t), a = partsLux(now), dm = partsLux(now + J);
  if (p.y === a.y && p.m === a.m && p.d === a.d) return `aujourd'hui ${heure}`;
  if (p.y === dm.y && p.m === dm.m && p.d === dm.d) return `demain ${heure}`;
  return `${new Date(t).toLocaleDateString('fr-FR', { timeZone: TZ, weekday: 'short', day: 'numeric', month: 'short' })} ${heure}`;
}

/** Heures ouvrées lisibles : « 26 h », « 3,5 j » au-delà de 3 jours de 8 h. */
export function dureeHeures(h: number | null | undefined): string {
  if (h == null || !Number.isFinite(h)) return '-';
  if (h >= 24) return `${(Math.round((h / 8) * 10) / 10).toString().replace('.', ',')} j`;
  return `${(Math.round(h * 10) / 10).toString().replace('.', ',')} h`;
}

// ── Palette ⌘K ──

const plier = (s: string) => String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');

/**
 * Score de correspondance (pur) : 0 = aucune. Tous les mots de la requête doivent se trouver ;
 * début de mot et début de libellé comptent plus ; les mots-clés (alias) aussi.
 */
export function scoreRecherche(requete: string, libelle: string, motsCles = ''): number {
  const q = plier(requete).trim();
  if (!q) return 1;
  const cible = plier(`${libelle} ${motsCles}`);
  const mots = q.split(/\s+/).filter(Boolean);
  let score = 0;
  for (const m of mots) {
    const i = cible.indexOf(m);
    if (i < 0) return 0;
    score += 10;
    if (i === 0) score += 8;
    else if (/[\s\-«(@.]/.test(cible[i - 1])) score += 4;
  }
  if (plier(libelle).startsWith(q)) score += 12;
  return score - Math.min(5, Math.floor(cible.length / 40));
}

export interface Commande { id: string; libelle: string; groupe: string; motsCles?: string; raccourci?: string }

/** Commandes triées par pertinence (pur), 12 au plus par défaut. */
export function filtrerCommandes<T extends Commande>(cmds: T[], requete: string, max = 12): T[] {
  return cmds
    .map((c, i) => ({ c, i, s: scoreRecherche(requete, c.libelle, c.motsCles) }))
    .filter(x => x.s > 0)
    .sort((a, b) => b.s - a.s || a.i - b.i)
    .slice(0, max)
    .map(x => x.c);
}

// ── Navigation et rafraîchissement ──

/** Index suivant / précédent borné (j/k, flèches). -1 = rien de sélectionné. */
export function deplacer(index: number, delta: number, total: number): number {
  if (total <= 0) return -1;
  if (index < 0) return delta > 0 ? 0 : total - 1;
  return Math.max(0, Math.min(total - 1, index + delta));
}

/** Identifiants apparus depuis le rendu précédent (pour les signaler sans tout réafficher). */
export function nouveautes(avant: string[], apres: string[]): string[] {
  const a = new Set(avant);
  return apres.filter(x => !a.has(x));
}

/** Courbe d'animation des chiffres (pur) : sortie douce. */
export function adoucir(t: number): number {
  const x = Math.max(0, Math.min(1, t));
  return 1 - Math.pow(1 - x, 3);
}

/** Valeur intermédiaire d'un chiffre animé (entier). */
export function valeurAnimee(de: number, vers: number, t: number): number {
  return Math.round(de + (vers - de) * adoucir(t));
}

/** Initiales (pastille d'expéditeur). */
export function initiales(nom: string, email = ''): string {
  const src = (nom || '').trim() || email.split('@')[0].replace(/[._-]+/g, ' ');
  const mots = src.split(/\s+/).filter(w => /[a-zà-ÿ]/i.test(w));
  return ((mots[0]?.[0] || '?') + (mots.length > 1 ? mots[mots.length - 1][0] : '')).toUpperCase();
}

/** Prénom lisible d'une adresse interne (« marie.dupont@… » → « Marie »). */
export function prenomDe(email: string): string {
  const p = String(email || '').split('@')[0].split(/[._-]/)[0] || '';
  return p ? p.charAt(0).toUpperCase() + p.slice(1) : email;
}
