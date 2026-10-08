/**
 * traites.ts · « TRAITÉS À RANGER » du tableau de bord (demande de Charles du 08/10/2026) : une fois
 * un mail traité (réponse depuis « Répondre » du tableau, ou depuis Outlook), ATLAS propose de le
 * ranger dans le bon dossier sans qu'on ait à y penser : une carte par mail, le dossier proposé
 * (projet, dossier habituel de l'expéditeur, règle), UN clic « Ranger », « Autre dossier… » (le
 * même choix que « Classer ce mail » : recherche dans tout l'arbre, chemin tapé = dossier à créer),
 * « Ignorer », et « Annuler » (le mail revient). Les écritures passent par le worker
 * (tableau/traites/ranger = même rangement que dossiers/ranger, journal, Annuler 30 jours).
 * Affiché en tête d'« Aujourd'hui » (aujourdhui.ts) et de la séance de tri (seance.ts).
 */
import { annulerTraite, ignorerTraite, phraseSuggestion, rangerTraite, sourceSuggestion, type TraiteARanger } from '../api/tableau';
import { fetchDossiers, type DossierOutlook } from '../api/agent';
import { humanError } from '../api/net';
import { icon } from '../ui/icons';
import { ilYA } from './logique';
import { h, ouvrirCouche, toast } from './ui';

export interface CibleRangement { dossier?: DossierOutlook; creer?: { chemin: string; projetId?: string }; projetId?: string }

/**
 * « Autre dossier… » : fenêtre de recherche dans tout l'arbre de la boîte ; un clic = dossier choisi ;
 * ce qui est tapé peut devenir un nouveau dossier (« A/B/C » = niveaux). Résout null si fermée.
 */
export function choisirAutreDossier(mailbox: string): Promise<CibleRangement | null> {
  return new Promise(resolve => {
    const el = document.createElement('div');
    el.className = 'tb-modal';
    el.setAttribute('role', 'dialog');
    el.setAttribute('aria-modal', 'true');
    el.setAttribute('aria-label', 'Autre dossier');
    el.innerHTML = `
      <h3>Ranger dans un autre dossier</h3>
      <input type="search" class="tb-input" id="tb-ad-q" autocomplete="off" placeholder="Tape un nom ou un n° de projet (ex. 871, DealsUp)" aria-label="Chercher un dossier Outlook">
      <div id="tb-ad-liste" class="tb-note">Tape pour chercher dans toute la boîte.</div>`;
    let fini = false;
    const finir = (c: CibleRangement | null) => { if (fini) return; fini = true; fermer(); resolve(c); };
    const fermer = ouvrirCouche(el, () => { if (!fini) { fini = true; resolve(null); } });
    const q = el.querySelector<HTMLInputElement>('#tb-ad-q')!;
    const liste = el.querySelector<HTMLElement>('#tb-ad-liste')!;
    let minuterie: ReturnType<typeof setTimeout> | undefined;
    let n = 0;
    const afficher = async (texte: string) => {
      const k = ++n;
      if (!texte) { liste.innerHTML = 'Tape pour chercher dans toute la boîte.'; return; }
      liste.innerHTML = 'Recherche…';
      try {
        const ds = (await fetchDossiers(mailbox || undefined, texte)).slice(0, 8);
        if (k !== n) return;
        const cheminTape = texte.split('/').map(x => x.trim()).filter(Boolean).join('/');
        const exact = ds.some(d => d.chemin.toLowerCase() === cheminTape.toLowerCase() || d.chemin.toLowerCase().endsWith(`/${cheminTape.toLowerCase()}`));
        liste.innerHTML = `<div class="tb-presets">
          ${ds.map((d, i) => `<button type="button" class="tb-btn" data-k="${i}">${icon('folder', 13)}${h(d.chemin)}</button>`).join('')}
          ${cheminTape && !exact ? `<button type="button" class="tb-btn is-ghost" data-nouveau>${icon('plus', 13)}Créer « ${h(cheminTape)} » et ranger</button>` : ''}
        </div>${ds.length ? '' : '<p class="tb-note">Aucun dossier ne correspond.</p>'}`;
        liste.querySelectorAll<HTMLButtonElement>('[data-k]').forEach(b => b.addEventListener('click', () => { const d = ds[Number(b.dataset.k)]; if (d) finir({ dossier: d }); }));
        liste.querySelector<HTMLButtonElement>('[data-nouveau]')?.addEventListener('click', () => finir({ creer: { chemin: cheminTape } }));
      } catch (e) { if (k === n) liste.innerHTML = `<p class="tb-note is-err">Dossiers indisponibles : ${h(humanError(e))}</p>`; }
    };
    q.addEventListener('input', () => { if (minuterie) clearTimeout(minuterie); const v = q.value.trim(); minuterie = setTimeout(() => void afficher(v), 250); });
    q.focus();
  });
}

/** Range un mail traité (dossier proposé, ou cible choisie) ; renvoie le chemin et de quoi annuler. */
export async function rangerTraiteMail(t: TraiteARanger, cible?: CibleRangement | null): Promise<{ chemin: string; cree: boolean; annuler: () => Promise<void> }> {
  const r = await rangerTraite({
    messageId: t.messageId,
    ...(cible?.dossier ? { dossierId: cible.dossier.id } : {}),
    ...(cible?.creer ? { creer: cible.creer } : {}),
    ...(cible?.projetId ? { projetId: cible.projetId } : {}),
  });
  return { chemin: r.dossier.chemin, cree: r.cree, annuler: async () => { await annulerTraite(t.messageId); } };
}

export interface OptionsTraites {
  openLink: (url: string) => void;
  /** Après un rangement, une annulation, un « ignorer » : le tableau se relit. */
  onChange: () => void;
}

/** Une carte par mail traité : suggestion, « Ranger », « Autre dossier… », « Ignorer ». */
function carteHtml(t: TraiteARanger, now: number): string {
  const d = t.destination;
  const src = sourceSuggestion(d);
  return `<div class="tb-eng tb-traite" data-id="${h(t.messageId)}">
    <p><b>${h(t.subject || '(sans objet)')}</b><br><small>${h(t.from?.name || t.from?.email || '')} · répondu ${h(ilYA(t.repondu.at, now))}</small><br>
      <span class="tb-traite-sugg">${icon('folder', 12)} ${h(phraseSuggestion(d))}</span>${src ? ` <small>(${h(src)})</small>` : ''}</p>
    <span class="tb-actions">
      ${d.type !== 'aucun' && d.chemin ? `<button type="button" class="tb-btn is-primary" data-ranger title="${h(d.chemin)}">${icon('check', 14)}Ranger</button>` : ''}
      <button type="button" class="tb-btn" data-autre>Autre dossier…</button>
      ${t.webLink ? `<button type="button" class="tb-btn is-ghost" data-ouvrir title="Ouvrir dans Outlook">${icon('external', 14)}</button>` : ''}
      <button type="button" class="tb-btn is-ghost" data-ignorer title="Laisser dans la boîte">Ignorer</button>
    </span>
  </div>`;
}

/**
 * Panneau « Traités à ranger (N) » : rien n'est affiché s'il n'y a ni mail à ranger ni rangement
 * seul récent. Les rangements seuls de l'agent (autonomie) sont dits, annulables.
 */
export function renderTraites(host: HTMLElement, liste: TraiteARanger[], rangesSeuls: TraiteARanger[], opts: OptionsTraites): void {
  const now = Date.now();
  if (!liste.length && !rangesSeuls.length) { host.hidden = true; host.innerHTML = ''; return; }
  host.hidden = false;
  host.innerHTML = `<div class="tb-h">Traités à ranger<span class="tb-h-n">${liste.length}</span></div>
    ${liste.length ? '<p class="tb-note">Tu as répondu : un clic range le mail dans le dossier proposé, le tableau reste net.</p>' : ''}
    ${liste.map(t => carteHtml(t, now)).join('')}
    ${rangesSeuls.length ? `<details class="tb-anciens"><summary>Rangés seuls par ATLAS après ta réponse (${rangesSeuls.length})</summary>
      ${rangesSeuls.map(t => `<div class="tb-eng" data-auto="${h(t.messageId)}"><p>${h(t.subject || '(sans objet)')}<br><small>dans « ${h(t.destination.chemin || '')} » · ${h(ilYA(t.rangeLe || t.repondu.at, now))}</small></p><span class="tb-actions"><button type="button" class="tb-btn is-ghost" data-annuler-auto>Annuler</button></span></div>`).join('')}
    </details>` : ''}`;
  const trouver = (el: HTMLElement) => liste.find(t => t.messageId === el.closest<HTMLElement>('[data-id]')?.dataset.id);
  const occuper = (el: HTMLElement, oui: boolean) => el.querySelectorAll<HTMLButtonElement>('button').forEach(b => { b.disabled = oui; });

  const ranger = async (btn: HTMLButtonElement, cible?: CibleRangement | null) => {
    const carte = btn.closest<HTMLElement>('[data-id]')!;
    const t = trouver(btn);
    if (!t) return;
    occuper(carte, true);
    try {
      const r = await rangerTraiteMail(t, cible);
      carte.innerHTML = `<p>${icon('check-circle', 13)} ${r.cree ? 'Dossier créé, mail rangé' : 'Rangé'} dans « ${h(r.chemin)} »</p><span class="tb-actions"><button type="button" class="tb-btn is-ghost" data-annuler>Annuler</button></span>`;
      const defaire = async () => {
        try { await r.annuler(); toast('Rangement annulé : le mail est revenu', 'success'); opts.onChange(); }
        catch (e) { toast(`Annulation impossible : ${humanError(e)}`, 'error'); }
      };
      carte.querySelector<HTMLButtonElement>('[data-annuler]')?.addEventListener('click', ev => { (ev.currentTarget as HTMLButtonElement).disabled = true; void defaire(); });
      toast(`Rangé dans ${r.chemin}`, 'success', () => void defaire());
      opts.onChange();
    } catch (e) {
      occuper(carte, false);
      toast(`Rangement impossible : ${humanError(e)}`, 'error');
    }
  };

  host.querySelectorAll<HTMLButtonElement>('[data-ranger]').forEach(b => b.addEventListener('click', () => void ranger(b)));
  host.querySelectorAll<HTMLButtonElement>('[data-autre]').forEach(b => b.addEventListener('click', async () => {
    const t = trouver(b);
    if (!t) return;
    const cible = await choisirAutreDossier(t.mailbox);
    if (cible) void ranger(b, cible);
  }));
  host.querySelectorAll<HTMLButtonElement>('[data-ouvrir]').forEach(b => b.addEventListener('click', () => { const t = trouver(b); if (t?.webLink) opts.openLink(t.webLink); }));
  host.querySelectorAll<HTMLButtonElement>('[data-ignorer]').forEach(b => b.addEventListener('click', async () => {
    const t = trouver(b);
    if (!t) return;
    const carte = b.closest<HTMLElement>('[data-id]')!;
    occuper(carte, true);
    try {
      await ignorerTraite(t.messageId);
      carte.remove();
      toast('Laissé dans la boîte', 'info', () => { void annulerTraite(t.messageId).then(() => opts.onChange()).catch(e => toast(humanError(e), 'error')); });
      opts.onChange();
    } catch (e) { occuper(carte, false); toast(humanError(e), 'error'); }
  }));
  host.querySelectorAll<HTMLButtonElement>('[data-annuler-auto]').forEach(b => b.addEventListener('click', async () => {
    const id = b.closest<HTMLElement>('[data-auto]')?.dataset.auto || '';
    b.disabled = true;
    try { await annulerTraite(id); toast('Le mail est revenu dans la boîte', 'success'); opts.onChange(); }
    catch (e) { b.disabled = false; toast(`Annulation impossible : ${humanError(e)}`, 'error'); }
  }));
}
