/**
 * carte-pret.ts · carte « Prêt » en tête du volet (lots 2 et 3 du complément, 10/10/2026).
 *
 * L'agent a préparé la vue du mail à son arrivée : projet reconnu (et pourquoi), dossier Outlook du
 * projet. La carte s'affiche tout de suite, sans appel de plus :
 *   - « Lier et classer » en un clic (liaison au projet + rangement dans son dossier, créé au chemin
 *     d'ATLAS s'il manque) ; la personne accepte ainsi ce fil et cet expéditeur pour la suite ;
 *   - ce que l'agent a fait seul (« Lié et classé seul ») avec « Annuler » ;
 *   - « Changer de projet » (recherche) et « Délier » : corrections mémorisées, elles priment ensuite.
 * Aussi : réglages d'autonomie de la personne (trier, ranger, lier : seul / me proposer / rien).
 * Rien n'est envoyé chez un client.
 */
import { annulerAction } from '../api/agent';
import { humanError } from '../api/net';
import {
  changerProjet, chercherProjets, delierProjet, ecrireAutonomie, lierEtClasser, lireAutonomie,
  type NiveauAutonomie, type ResultatLiaison, type VuePrete,
} from '../api/projet-mail';
import { escapeHtml } from '../utils/html';
import { icon } from '../ui/icons';
import { inlineLoadingHtml } from '../ui/states';

type InfoFn = (msg: string, kind?: 'success' | 'error' | 'info') => void;

export interface CtxPret {
  messageId: string;
  mailbox?: string;
  conversationId?: string;
  onInfo?: InfoFn;
  /** Projet lié ou classé : la carte « Classer ce mail » est redessinée par l'appelant. */
  onChange?: (projet: { id: string; libelle: string } | null) => void;
}

const CONFIANCE: Record<string, string> = { forte: 'sûr', moyenne: 'probable', faible: 'possible' };

function resumeLiaison(r: ResultatLiaison): string {
  const d = r.rangement?.dossier;
  const range = d ? ` et classé dans « ${d.chemin} »${r.rangement?.cree ? ' (dossier créé)' : ''}` : '';
  return `Lié à ${r.projet.libelle}${range}.`;
}

/**
 * Sélecteur « Changer de projet » (recherche parmi les projets en cours). `choisir` reçoit le projet ;
 * partagé avec la carte « Classer ce mail » (volet et tableau de bord).
 */
export function monterChoixProjet(host: HTMLElement, choisir: (p: { id: string; libelle: string }) => void | Promise<void>, annuler?: () => void): void {
  host.hidden = false;
  host.innerHTML = `
    <input type="search" class="agent-input" data-q autocomplete="off" placeholder="N°, nom du projet ou client" aria-label="Chercher un projet">
    <div data-res></div>
    ${annuler ? '<button type="button" class="agent-link" data-fermer>Fermer</button>' : ''}`;
  const q = host.querySelector<HTMLInputElement>('[data-q]')!;
  const res = host.querySelector<HTMLElement>('[data-res]')!;
  host.querySelector('[data-fermer]')?.addEventListener('click', () => { host.hidden = true; host.innerHTML = ''; annuler?.(); });
  let n = 0;
  let minuterie: ReturnType<typeof setTimeout> | undefined;
  const chercher = async () => {
    const k = ++n;
    res.innerHTML = inlineLoadingHtml('Recherche…');
    try {
      const ps = await chercherProjets(q.value.trim());
      if (k !== n || !host.isConnected) return;
      res.innerHTML = ps.length
        ? ps.map((p, i) => `<button type="button" class="agent-btn-liste" data-i="${i}">${icon('folder', 14)} ${escapeHtml(p.libelle)}${p.client ? ` <span class="agent-muted">· ${escapeHtml(p.client)}</span>` : ''}</button>`).join('')
        : '<p class="agent-muted">Aucun projet en cours ne correspond.</p>';
      res.querySelectorAll<HTMLButtonElement>('[data-i]').forEach(b => b.addEventListener('click', () => {
        const p = ps[Number(b.dataset.i)];
        res.querySelectorAll<HTMLButtonElement>('button').forEach(x => { x.disabled = true; });
        void choisir({ id: p.id, libelle: p.libelle });
      }));
    } catch (e) {
      if (k === n) res.innerHTML = `<p class="agent-error" role="alert">Projets indisponibles : ${escapeHtml(humanError(e))}</p>`;
    }
  };
  q.addEventListener('input', () => { clearTimeout(minuterie); minuterie = setTimeout(() => void chercher(), 250); });
  void chercher();
  q.focus();
}

/** Carte « Prêt » ; masquée s'il n'y a ni projet ni candidat (ou si la personne a réglé « rien »). */
export function renderCartePret(host: HTMLElement, pret: VuePrete | null, ctx: CtxPret): void {
  const auto = pret?.autonomie;
  const rien = !!auto && auto.lier === 'suggerer' && auto.ranger === 'suggerer';
  if (!pret || rien || (!pret.projet && !pret.candidats?.length)) { host.hidden = true; host.innerHTML = ''; return; }
  host.hidden = false;
  let projet = pret.projet;
  let lie = !!pret.lie || !!pret.auto?.lie;
  let occupe = false;

  const dessiner = (message = '', actionRangementId = pret.auto?.range ? pret.auto.actionRangementId : undefined) => {
    const titre = projet
      ? `<div class="agent-carte-titre">${icon('folder', 14)} ${escapeHtml(projet.libelle)}</div>
         <div class="agent-muted">${lie ? 'Lié dans ATLAS' : `Projet ${CONFIANCE[projet.confiance] || 'probable'}`} · ${escapeHtml(projet.raison)}</div>`
      : `<div class="agent-carte-titre">${icon('folder', 14)} Quel projet ?</div><div class="agent-muted">${escapeHtml(pret.raison || 'Plusieurs projets en cours pour ce contact')}</div>`;
    const dossier = pret.dossier?.existant?.chemin || pret.dossier?.propose || '';
    const autoFait = pret.auto && (pret.auto.lie || pret.auto.range)
      ? `<p class="status-linked">${icon('check-circle', 14)}${pret.auto.lie ? 'Lié' : ''}${pret.auto.lie && pret.auto.range ? ' et classé' : pret.auto.range ? 'Classé' : ''} seul, comme tu l'as demandé.</p>` : '';
    const boutons = projet && !lie
      ? `<button type="button" class="btn btn-primary btn-block agent-btn" data-lier>${icon('link', 14)} Lier et classer</button>
         ${dossier ? `<div class="agent-muted">Dans « ${escapeHtml(dossier)} »${pret.dossier?.existant ? '' : ' (dossier créé au besoin)'}</div>` : ''}`
      : '';
    const candidats = !projet && pret.candidats?.length
      ? pret.candidats.map((c, i) => `<button type="button" class="btn btn-secondary btn-block agent-btn" data-cand="${i}">Lier et classer dans ${escapeHtml(c.libelle)}</button>`).join('')
      : '';
    host.innerHTML = `
      <div class="agent-section-title">Prêt</div>
      <div class="agent-carte">
        ${titre}${autoFait}${message ? `<p class="status-linked">${icon('check-circle', 14)}${escapeHtml(message)}</p>` : ''}
        ${boutons}${candidats}
        <div class="agent-liens">
          <button type="button" class="agent-link" data-changer>Changer de projet</button>
          ${projet && lie ? '<button type="button" class="agent-link" data-delier>Délier</button>' : ''}
          ${actionRangementId ? '<button type="button" class="agent-link" data-annuler>Annuler le rangement</button>' : ''}
        </div>
        <div data-choix hidden></div>
        <div data-err hidden></div>
      </div>`;
    const err = host.querySelector<HTMLElement>('[data-err]')!;
    const erreur = (texte: string, e: unknown) => { err.hidden = false; err.innerHTML = `<p class="agent-error" role="alert">${escapeHtml(texte)} : ${escapeHtml(humanError(e))}</p>`; };
    const lier = async (p: { id: string; libelle: string }, changer: boolean, btn?: HTMLButtonElement) => {
      if (occupe) return;
      occupe = true;
      if (btn) { btn.disabled = true; btn.textContent = 'Liaison…'; }
      try {
        const base = { messageId: ctx.messageId, projetId: p.id, ...(ctx.mailbox ? { mailbox: ctx.mailbox } : {}) };
        const r = changer ? await changerProjet({ ...base, ...(projet && lie ? { ancienProjetId: projet.id } : {}) }) : await lierEtClasser(base);
        projet = { id: r.projet.id, libelle: r.projet.libelle, source: 'correction-fil', confiance: 'forte', raison: changer ? 'Projet choisi par toi' : 'Lié par toi' };
        lie = true;
        pret.auto = undefined;
        const suite = r.prochainsSeuls ? ' Les prochains mails de ce fil et de cet expéditeur le seront seuls.' : '';
        occupe = false;
        dessiner(`${resumeLiaison(r)}${suite}`, r.rangement?.actionId);
        if (r.rangementErreur) erreur('Lié, mais pas classé', new Error(r.rangementErreur));
        ctx.onInfo?.(resumeLiaison(r), 'success');
        ctx.onChange?.(r.projet);
      } catch (e) {
        occupe = false;
        if (btn) { btn.disabled = false; btn.textContent = 'Lier et classer'; }
        erreur('Liaison impossible', e);
      }
    };
    host.querySelector<HTMLButtonElement>('[data-lier]')?.addEventListener('click', ev => { if (projet) void lier(projet, false, ev.currentTarget as HTMLButtonElement); });
    host.querySelectorAll<HTMLButtonElement>('[data-cand]').forEach(b => b.addEventListener('click', () => {
      const c = pret.candidats![Number(b.dataset.cand)];
      void lier({ id: c.id, libelle: c.libelle }, true, b);
    }));
    host.querySelector('[data-changer]')?.addEventListener('click', () => {
      const zone = host.querySelector<HTMLElement>('[data-choix]')!;
      monterChoixProjet(zone, p => lier(p, true), () => undefined);
    });
    host.querySelector<HTMLButtonElement>('[data-delier]')?.addEventListener('click', async ev => {
      if (occupe) return;
      occupe = true;
      const b = ev.currentTarget as HTMLButtonElement;
      b.disabled = true;
      try {
        const r = await delierProjet({ messageId: ctx.messageId, ...(projet ? { projetId: projet.id } : {}), ...(ctx.conversationId ? { conversationId: ctx.conversationId } : {}), ...(actionRangementId ? { actionRangementId } : {}), ...(ctx.mailbox ? { mailbox: ctx.mailbox } : {}) });
        projet = null; lie = false; pret.candidats = undefined; pret.raison = 'Tu as délié ce fil : ATLAS ne le proposera plus.';
        pret.auto = undefined;
        occupe = false;
        dessiner(r.rangementAnnule ? 'Délié, mail remis à sa place.' : 'Délié. Ce choix est retenu pour ce fil.', undefined);
        ctx.onInfo?.('Délié du projet', 'success');
        ctx.onChange?.(null);
      } catch (e) { occupe = false; b.disabled = false; erreur('Impossible de délier', e); }
    });
    host.querySelector<HTMLButtonElement>('[data-annuler]')?.addEventListener('click', async ev => {
      const b = ev.currentTarget as HTMLButtonElement;
      b.disabled = true;
      try { await annulerAction(actionRangementId!); b.replaceWith(Object.assign(document.createElement('span'), { className: 'agent-muted', textContent: 'Rangement annulé : mail remis à sa place' })); ctx.onInfo?.('Rangement annulé', 'success'); }
      catch (e) { b.disabled = false; erreur('Annulation impossible', e); }
    });
  };
  dessiner();
}

// ── Autonomie de la personne (Réglages) ────────────────────────────────────────────

const ACTIONS: Array<{ cle: 'trier' | 'ranger' | 'lier'; titre: string; aide: string }> = [
  { cle: 'lier', titre: 'Lier les mails à leur projet', aide: 'Seulement pour un fil ou un expéditeur que tu as déjà lié une fois.' },
  { cle: 'ranger', titre: 'Classer dans le dossier du projet', aide: 'Annulable 30 jours. Un mail à traiter attend ta réponse.' },
  { cle: 'trier', titre: 'Étiqueter dans Outlook', aide: 'Piles « À traiter », « Pour info »… et « ATLAS · #n° » du projet.' },
];
const POSITIONS: Array<{ n: NiveauAutonomie; libelle: string }> = [
  { n: 'faire', libelle: 'Seul' }, { n: 'preparer', libelle: 'Me proposer' }, { n: 'suggerer', libelle: 'Rien' },
];

/** Trois positions par action (seul / me proposer / rien), pour la personne elle-même. */
export function monterAutonomie(host: HTMLElement, onInfo?: InfoFn): void {
  host.innerHTML = inlineLoadingHtml('Autonomie d’ATLAS…');
  const rendre = (reglages: Record<string, NiveauAutonomie>) => {
    host.innerHTML = `
      <p class="tool-row-title">Ce qu'ATLAS fait seul pour toi</p>
      ${ACTIONS.map(a => `
        <fieldset class="agent-autonomie">
          <legend>${escapeHtml(a.titre)}</legend>
          <div class="agent-segment" role="radiogroup" aria-label="${escapeHtml(a.titre)}">
            ${POSITIONS.map(p => `<label><input type="radio" name="auto-${a.cle}" value="${p.n}" ${reglages[a.cle] === p.n ? 'checked' : ''}> ${p.libelle}</label>`).join('')}
          </div>
          <div class="agent-muted">${escapeHtml(a.aide)}</div>
        </fieldset>`).join('')}
      <p class="agent-muted">Rien n'est jamais envoyé à un client. Tout se corrige en un clic.</p>`;
    host.querySelectorAll<HTMLInputElement>('input[type=radio]').forEach(r => r.addEventListener('change', async () => {
      const cle = r.name.replace('auto-', '');
      try { const v = await ecrireAutonomie({ [cle]: r.value as NiveauAutonomie }); rendre(v.reglages); onInfo?.('Réglage enregistré', 'success'); }
      catch (e) { onInfo?.(`Réglage non enregistré : ${humanError(e)}`, 'error'); void lireAutonomie().then(v => rendre(v.reglages)).catch(() => undefined); }
    }));
  };
  lireAutonomie().then(v => rendre(v.reglages)).catch(e => { host.innerHTML = `<p class="agent-error" role="alert">Réglages indisponibles : ${escapeHtml(humanError(e))}</p>`; });
}
