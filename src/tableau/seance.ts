/**
 * seance.ts · « SÉANCE DE TRI » du tableau de bord (backlog recz1ck93h164mv9x et recnkz3B90JFpzkA9,
 * programme « Assistant inbox » du 07/10/2026 : fermer Outlook avec 5 mails au plus dans la boîte).
 *
 * Plein écran dans la fenêtre du tableau de bord (dialogue Office sur le nouvel Outlook pour Mac,
 * onglet de la barre de gauche ailleurs) : compteur « N dans la boîte · objectif 5 », UNE carte à la
 * fois (les mails importants d'abord, par valeur pour l'agence, du plus ancien au plus récent), le
 * contexte à droite (projet, dernière offre, impayés si connus, dernier échange) et UNE action déjà
 * préparée par le worker (GET /api/plugin/agent/seance).
 *
 * Clavier : Entrée = action recommandée ; 1 répondre ; 2 confier ; 3 plus tard (demain / lundi) ;
 * 4 autre (menu) ; 5, 6, 7 = répondre avec l'une des trois réponses courtes proposées ; 8 = mise en
 * relation (réponse préparée, introducteur en copie cachée) ; Échap = passer ; ⌘Z / Ctrl+Z ou
 * « Annuler » = défaire la dernière action.
 * « Ne rien laisser filer » (07/10/2026) : cartes « Tu as promis… » (tâche sur clic), « offre consultée,
 * pas de réponse » (relance rédigée, copiée, jamais envoyée), « compte rendu à envoyer » (module Réunion).
 * Mail à risque (sécurité) : bandeau rouge, aucune action au clavier, aucun brouillon.
 *
 * Toutes les écritures passent par les routes EXISTANTES du worker : rangement sur clic dans SA boîte
 * (dossiers/ranger), actions métier (actions/executer), « plus tard », « annuler », réponse ARGO et
 * dépôt de brouillon (tableau/reponse, tableau/brouillon). Rien n'est jamais envoyé : « Répondre »
 * ouvre la réponse dans Outlook (brouillon déposé si le verrou du tableau est ouvert, formulaire de
 * réponse d'Outlook si la fenêtre parente montre ce mail, sinon texte copié et mail ouvert).
 * Chaque décision (recommandation acceptée ou remplacée) est notée pour l'apprentissage.
 */
import {
  fetchSeance, noterDecisionSeance, confierMail, annulerConfier, finSeance, actionGroupe, annulerGroupe, desabonner, poserRetour, retirerRetour,
  repondreAutonomie, sollicitationsVues, fetchSuggestions, preparerMiseEnRelation, promesseTache, promesseEcarter, promesseAnnuler, relanceOffre,
  type ConditionRetour, type Seance, type SeanceAction, type SeanceCarte, type SeanceDesabonnement, type SeanceGroupe,
  type SeancePropositionAutonomie, type SeanceSemaine, type SeanceSollicitations, type SeancePromesse, type SeanceOffreConsultee, type SeanceCrASuivre,
} from '../api/seance';
import { annulerAction, executerAction, fetchDossierProjet, mettrePlusTard, rangerDansDossier, retirerPlusTard } from '../api/agent';
import { deposerBrouillon, redigerArgo } from '../api/tableau';
import { callAtlasWorker } from '../api/worker';
import { humanError } from '../api/net';
import { openAgentMail } from '../components/agent-lists';
import { copierTexte } from '../components/agent-outils';
import { icon } from '../ui/icons';
import { ilYA, prenomDe } from './logique';
import { h, toast, ouvrirCouche, animerChiffre } from './ui';

export interface OptionsSeance {
  openLink: (url: string) => void;
  /** Formulaire de réponse d'Outlook par la fenêtre parente (dialogue Office), si elle montre ce mail. */
  repondreDansOutlook?: (messageId: string, html: string) => Promise<boolean>;
  /** Appelé à la fermeture (le tableau se relit). */
  onFerme?: () => void;
}

/** Cartes de l'Assistant inbox (07/10/2026) : regroupements avant les mails ; désabonnements, sollicitations et autonomie après. */
type Special =
  | { kind: 'groupe'; g: SeanceGroupe }
  | { kind: 'desabonnement'; d: SeanceDesabonnement }
  | { kind: 'sollicitations'; s: SeanceSollicitations }
  | { kind: 'autonomie'; a: SeancePropositionAutonomie }
  | { kind: 'promesse'; p: SeancePromesse }
  | { kind: 'offre'; o: SeanceOffreConsultee }
  | { kind: 'cr'; r: SeanceCrASuivre };

interface Defaire {
  carte?: SeanceCarte; index: number; recommandee: string | null; faite: string; annuler?: () => Promise<void>; sortie: boolean | number;
  patron?: string; special?: { liste: 'avant' | 'apres'; item: Special };
}

const texteVersHtml = (t: string) => `<div>${t.split(/\n/).map(l => (l.trim() ? h(l) : '<br>')).join('<br>')}</div>`;
const jourFr = (iso?: string) => (iso && /^\d{4}-\d{2}-\d{2}/.test(iso) ? iso.slice(0, 10).split('-').reverse().join('/') : '');
const minutesLisibles = (ms: number) => { const m = Math.round(ms / 60_000); return m < 1 ? 'moins d\'une minute' : `${m} min`; };

/** Petit menu numéroté (1 à 9) dans une fenêtre ATLAS ; résout l'index choisi ou null. */
function menu(titre: string, choix: Array<{ libelle: string; detail?: string }>): Promise<number | null> {
  return new Promise(resolve => {
    const el = document.createElement('div');
    el.className = 'tb-modal tb-seance-menu';
    el.setAttribute('role', 'dialog');
    el.setAttribute('aria-modal', 'true');
    el.setAttribute('aria-label', titre);
    el.innerHTML = `<h3>${h(titre)}</h3><div class="tb-presets">${choix.slice(0, 9).map((c, i) => `<button type="button" class="tb-btn" data-i="${i}"><span class="tb-kbd">${i + 1}</span><span>${h(c.libelle)}${c.detail ? `<small>${h(c.detail)}</small>` : ''}</span></button>`).join('')}</div>`;
    let fini = false;
    const finir = (i: number | null) => { if (fini) return; fini = true; fermer(); resolve(i); };
    const fermer = ouvrirCouche(el, () => { if (!fini) { fini = true; resolve(null); } });
    el.querySelectorAll<HTMLButtonElement>('[data-i]').forEach(b => b.addEventListener('click', () => finir(Number(b.dataset.i))));
    el.addEventListener('keydown', e => {
      const n = Number(e.key);
      if (n >= 1 && n <= Math.min(9, choix.length)) { e.preventDefault(); e.stopPropagation(); finir(n - 1); }
    });
    (el.querySelector('[data-i]') as HTMLButtonElement | null)?.focus();
  });
}

/** Ouvre la séance de tri en plein écran. */
export function ouvrirSeance(opts: OptionsSeance): void {
  if (document.getElementById('tb-seance')) return;
  const root = document.createElement('div');
  root.id = 'tb-seance';
  root.className = 'tb-seance';
  root.setAttribute('role', 'dialog');
  root.setAttribute('aria-modal', 'true');
  root.setAttribute('aria-label', 'Séance de tri');
  root.innerHTML = `
    <header class="tb-seance-top">
      <span class="tb-mark">Séance de tri</span>
      <span class="tb-seance-compteur" id="sc-compteur" aria-live="polite"><b data-v="0">…</b> dans la boîte · objectif <b id="sc-objectif">5</b></span>
      <span class="tb-seance-prog" id="sc-prog"></span>
      <span class="tb-spacer"></span>
      <button type="button" class="tb-btn" id="sc-annuler" disabled title="Annuler la dernière action (⌘Z)">${icon('refresh', 14)}Annuler</button>
      <button type="button" class="tb-btn" id="sc-fin">${icon('check', 14)}Terminer</button>
    </header>
    <main class="tb-seance-main" id="sc-main"><div class="loading"><div class="spinner"></div><p>ARGO prépare tes cartes…</p></div></main>
    <footer class="tb-seance-pied" id="sc-pied"></footer>`;
  document.body.appendChild(root);
  requestAnimationFrame(() => root.classList.add('is-on'));
  const $ = (id: string) => root.querySelector<HTMLElement>(`#${id}`)!;

  let s: Seance | null = null;
  let cartes: SeanceCarte[] = [];
  let avant: Special[] = [];
  let apres: Special[] = [];
  let i = 0;
  let occupe = false;
  let traites = 0;
  const debut = Date.now();
  const pile: Defaire[] = [];
  let ferme = false;
  /** Temps passé sur la carte (statistiques : temps par client). */
  let carteAfficheeLe = Date.now();
  const suggestionsDemandees = new Set<string>();

  // ── Rendu ──
  function majCompteur(): void {
    const b = $('sc-compteur').querySelector('b')!;
    if (s?.dansLaBoite == null) { b.textContent = '?'; b.title = 'Boîte de réception illisible pour le moment'; }
    else animerChiffre(b, s.dansLaBoite);
    $('sc-objectif').textContent = String(s?.objectif ?? 5);
    $('sc-compteur').classList.toggle('is-ok', s?.dansLaBoite != null && s.dansLaBoite <= (s?.objectif ?? 5));
    const restSp = avant.length + apres.length;
    $('sc-prog').textContent = cartes.length ? `Carte ${Math.min(i + 1, cartes.length)} sur ${cartes.length}${restSp ? ` · ${restSp} regroupement${restSp > 1 ? 's' : ''} et propositions` : ''}` : restSp ? `${restSp} regroupement${restSp > 1 ? 's' : ''} et propositions` : '';
    ($('sc-annuler') as HTMLButtonElement).disabled = !pile.length;
  }

  function renderStats(stats: SeanceSemaine[]): string {
    const sem = stats[0];
    if (!sem) return '<span>Cette semaine : première séance.</span>';
    const releves = sem.joursReleves ? `${sem.objectifAtteint}/${sem.joursReleves} jour${sem.joursReleves > 1 ? 's' : ''} à 5 mails ou moins à la clôture` : 'pas encore de relevé à 17 h 30';
    return `<span>Cette semaine : ${sem.traites} mail${sem.traites > 1 ? 's' : ''} trié${sem.traites > 1 ? 's' : ''} en ${sem.minutes} min · ${releves}${sem.moyenneRestants != null ? ` · ${sem.moyenneRestants} en moyenne` : ''}</span>`;
  }

  function contexteHtml(c: SeanceCarte): string {
    const x = c.contexte;
    const lignes: string[] = [];
    if (x.tiers) lignes.push(`<div class="tb-seance-ctx-l"><span>Client</span><b>${h(x.tiers)}</b></div>`);
    if (x.projet) lignes.push(`<div class="tb-seance-ctx-l"><span>Projet</span><b>${h(x.projet.libelle)}</b>${x.projet.statut ? `<em>${h(x.projet.statut)}</em>` : ''}${x.projet.dates ? `<em>${h(x.projet.dates)}</em>` : ''}</div>`);
    if (x.offre) lignes.push(`<div class="tb-seance-ctx-l"><span>Dernière offre</span><b>${h(x.offre.libelle)}</b>${x.offre.montant ? `<em>${h(x.offre.montant)}</em>` : ''}${x.offre.envoyeeLe ? `<em>envoyée le ${h(jourFr(x.offre.envoyeeLe))}</em>` : ''}${x.offre.enAttente ? '<em class="is-warn">en attente de décision</em>' : x.offre.statut ? `<em>${h(x.offre.statut)}</em>` : ''}</div>`);
    lignes.push(`<div class="tb-seance-ctx-l"><span>Impayés</span>${x.impayes == null ? '<em>non disponible dans ATLAS</em>' : x.impayes.length ? x.impayes.map(f => `<b>${h(f.libelle)}</b>${f.montant ? `<em>${h(f.montant)}</em>` : ''}`).join('') : '<em>aucun</em>'}</div>`);
    if (x.dernierEchange) lignes.push(`<div class="tb-seance-ctx-l"><span>Dernier échange</span><b>${x.dernierEchange.sens === 'envoye' ? 'Tu as écrit' : 'Reçu'} ${h(ilYA(x.dernierEchange.at, Date.now()))}</b>${x.dernierEchange.sujet ? `<em>${h(x.dernierEchange.sujet)}</em>` : ''}</div>`);
    if (x.devis) lignes.push(`<div class="tb-seance-ctx-l"><span>Devis fournisseur</span><b>${h(x.devis.libelle)}</b>${x.devis.expireLe ? `<em class="is-warn">expire le ${h(jourFr(x.devis.expireLe))}</em>` : '<em>ouvert</em>'}</div>`);
    if (x.confie) lignes.push(`<div class="tb-seance-ctx-l"><span>Confié</span><b>à ${h(x.confie.nom)}</b><em>${h(ilYA(x.confie.le, Date.now()))}, revenu sans suite</em></div>`);
    return lignes.join('');
  }

  function specialCourant(): { liste: 'avant' | 'apres'; item: Special } | null {
    if (avant.length) return { liste: 'avant', item: avant[0] };
    if (!cartes[i] && apres.length) return { liste: 'apres', item: apres[0] };
    return null;
  }

  /** Boutons d'une carte spéciale : [libellé, action] ; Entrée = le premier. */
  function boutonsSpecial(sp: Special): Array<{ libelle: string; detail?: string; faire: () => Promise<Partial<Defaire> & { message: string } | null> }> {
    switch (sp.kind) {
      case 'groupe': {
        const g = sp.g;
        const lancer = (action: 'archiver' | 'marquer-lu' | 'classer', libelle: string, detail?: string) => ({
          libelle, detail,
          faire: async () => {
            const r = await actionGroupe({ action, messageIds: g.messageIds, ...(action === 'classer' && g.dossier ? { dossierId: g.dossier.id } : {}) });
            if (!r.faits) throw new Error('Aucun mail traité (déjà rangés ?)');
            const sortis = action === 'marquer-lu' ? 0 : r.faits;
            return { message: `${r.faits} mail${r.faits > 1 ? 's' : ''} ${action === 'marquer-lu' ? 'marqué' + (r.faits > 1 ? 's' : '') + ' comme lu' + (r.faits > 1 ? 's' : '') : action === 'classer' ? 'classé' + (r.faits > 1 ? 's' : '') : 'archivé' + (r.faits > 1 ? 's' : '')}${r.echecs ? ` (${r.echecs} en échec)` : ''}`, sortie: sortis, annuler: async () => { await annulerGroupe({ actionIds: r.actionIds, lus: r.lus }); } };
          },
        });
        return [
          ...(g.actions.includes('classer') && g.dossier ? [lancer('classer', 'Classer tout', `Dans ${g.dossier.chemin}`)] : []),
          ...(g.actions.includes('archiver') ? [lancer('archiver', 'Archiver tout')] : []),
          lancer('marquer-lu', 'Marquer tout comme lu', 'Les mails restent dans la boîte'),
        ];
      }
      case 'desabonnement': {
        const d = sp.d;
        const qui = d.nom || d.expediteur;
        const principal = d.mode === 'un-clic' ? 'Me désabonner en un clic' : d.mode === 'brouillon' ? 'Préparer le mail de désabonnement' : 'Ouvrir la page de désabonnement';
        return [
          {
            libelle: principal, detail: d.dansLaBoite ? `et archiver les ${d.dansLaBoite} mail${d.dansLaBoite > 1 ? 's' : ''} de ${qui}` : undefined,
            faire: async () => {
              const r = await desabonner(d.expediteur, { archiver: true });
              if (r.brouillon) {
                const u = `mailto:${encodeURIComponent(r.brouillon.a)}?subject=${encodeURIComponent(r.brouillon.sujet)}&body=${encodeURIComponent(r.brouillon.corps)}`;
                opts.openLink(u);
              } else if (r.lien) opts.openLink(r.lien);
              const ids = r.actionIds || [];
              const msg = r.statut === 'fait' ? `Désabonné de ${qui}` : r.statut === 'brouillon' ? 'Brouillon ouvert : relis puis envoie-le toi-même' : r.statut === 'echec' ? (r.erreur || 'Désabonnement refusé : page ouverte') : 'Page de désabonnement ouverte';
              return { message: `${msg}${r.archives ? `, ${r.archives} archivé${r.archives > 1 ? 's' : ''}` : ''}`, sortie: r.archives || 0, ...(ids.length ? { annuler: async () => { await annulerGroupe({ actionIds: ids, lus: [] }); } } : {}) };
            },
          },
          { libelle: `Garder ${qui}`, detail: 'Plus proposé', faire: async () => { await desabonner(d.expediteur, { ignorer: true }); return { message: `${qui} gardé` }; } },
        ];
      }
      case 'sollicitations': {
        const so = sp.s;
        return [
          {
            libelle: `Tout archiver (${so.mails.length})`,
            faire: async () => {
              const r = await actionGroupe({ action: 'archiver', messageIds: so.mails.map(m => m.messageId) });
              await sollicitationsVues(so.semaine);
              return { message: `${r.faits} sollicitation${r.faits > 1 ? 's' : ''} archivée${r.faits > 1 ? 's' : ''}`, sortie: r.faits, annuler: async () => { await annulerGroupe({ actionIds: r.actionIds, lus: r.lus }); } };
            },
          },
          { libelle: 'Vu, je les garde', detail: 'Elles restent hors de la séance', faire: async () => { await sollicitationsVues(so.semaine); return { message: 'Sollicitations passées en revue' }; } },
        ];
      }
      case 'promesse': {
        const p = sp.p;
        return [
          {
            libelle: 'Créer la tâche', detail: `Échéance ${jourFr(p.echeance)}${p.enRetard ? ' (dépassée)' : ''}`,
            faire: async () => {
              const r = await promesseTache(p.messageId, p.index);
              const id = r.tacheId;
              return { message: `Tâche créée${r.projet ? ` dans ${r.projet}` : ''} pour le ${jourFr(r.echeance || p.echeance)}`, annuler: async () => { await promesseAnnuler(p.messageId, p.index, id); } };
            },
          },
          { libelle: 'C\'est déjà fait', detail: 'Plus proposé', faire: async () => { await promesseEcarter(p.messageId, p.index); return { message: 'Promesse tenue, notée', annuler: async () => { await promesseAnnuler(p.messageId, p.index); } }; } },
          ...(p.webLink ? [{ libelle: 'Ouvrir le mail envoyé', faire: async () => { opts.openLink(p.webLink!); return null; } }] : []),
        ];
      }
      case 'offre': {
        const o = sp.o;
        return [
          {
            libelle: 'Préparer la relance', detail: 'Texte copié, tu relis et tu envoies depuis Outlook',
            faire: async () => {
              const r = await relanceOffre(o.messageId);
              const copie = await copierTexte(r.texte);
              if (r.webLink) opts.openLink(r.webLink);
              return { message: copie ? 'Relance copiée : réponds au mail envoyé, colle, relis puis envoie' : 'Relance prête mais copie impossible ici : réessaie depuis le panneau du mail' };
            },
          },
          { libelle: 'Plus tard', detail: 'Revient à la prochaine séance', faire: async () => ({ message: 'Gardé pour plus tard' }) },
        ];
      }
      case 'cr': {
        const r = sp.r;
        return [
          { libelle: 'Envoyer le compte rendu', detail: 'ATLAS ouvre la réunion : mail des participants avec relecture', faire: async () => { opts.openLink(r.lien); return { message: 'Réunion ouverte dans ATLAS' }; } },
          { libelle: 'Pas de mail pour cette réunion', faire: async () => ({ message: 'Noté pour cette séance' }) },
        ];
      }
      case 'autonomie': {
        const a = sp.a;
        return [
          { libelle: 'Oui, fais-le seul', detail: a.libelle, faire: async () => { const r = await repondreAutonomie(a.patron, true); return { message: r.message || 'Noté' }; } },
          { libelle: 'Non, je garde la main', faire: async () => { await repondreAutonomie(a.patron, false); return { message: 'Noté : ATLAS continuera de proposer' }; } },
        ];
      }
    }
  }

  function renderSpecial(cur: { liste: 'avant' | 'apres'; item: Special }): void {
    const sp = cur.item;
    const boutons = boutonsSpecial(sp);
    let titre = '', corps = '';
    if (sp.kind === 'groupe') {
      titre = sp.g.libelle;
      corps = `<p class="tb-seance-resume">Des mails semblables : une seule décision pour tous. Chaque action s'annule (⌘Z).</p><ul class="tb-seance-liste">${sp.g.apercu.map(m => `<li>${h(m.subject || '(sans objet)')}<span>${h(ilYA(m.receivedAt, Date.now()))}</span></li>`).join('')}${sp.g.messageIds.length > sp.g.apercu.length ? `<li class="tb-note">et ${sp.g.messageIds.length - sp.g.apercu.length} autre${sp.g.messageIds.length - sp.g.apercu.length > 1 ? 's' : ''}</li>` : ''}</ul>`;
    } else if (sp.kind === 'desabonnement') {
      titre = `Te désabonner de ${sp.d.nom || sp.d.expediteur} ?`;
      corps = `<p class="tb-seance-resume">${sp.d.recus30j} envoi${sp.d.recus30j > 1 ? 's' : ''} en 30 jours, aucun ouvert.${sp.d.mode === 'brouillon' ? ' Cet expéditeur demande un mail : ATLAS le prépare, tu l\'envoies toi-même.' : sp.d.mode === 'lien' ? ' Cet expéditeur passe par une page web : elle s\'ouvre, tu confirmes.' : ' Désabonnement direct, sans ouvrir de page.'}</p>`;
    } else if (sp.kind === 'sollicitations') {
      titre = `${sp.s.mails.length} sollicitation${sp.s.mails.length > 1 ? 's' : ''} commerciale${sp.s.mails.length > 1 ? 's' : ''} cette semaine`;
      corps = `<p class="tb-seance-resume">Démarchages d'inconnus, sortis de ta séance. Un coup d'œil, puis on range.</p><ul class="tb-seance-liste">${sp.s.mails.slice(0, 8).map(m => `<li><b>${h(m.from?.name || m.from?.email)}</b> ${h(m.subject || '(sans objet)')}<span>${h(m.raisons.slice(0, 2).join(' · '))}</span></li>`).join('')}</ul>`;
    } else if (sp.kind === 'promesse') {
      titre = `Tu as promis : « ${sp.p.texte} »`;
      corps = `<p class="tb-seance-resume">Dans ton mail « ${h(sp.p.sujet || '(sans objet)')} » à ${h(sp.p.a)}, ${h(ilYA(sp.p.sentAt, Date.now()))}. Échéance ${h(jourFr(sp.p.echeance))}${sp.p.enRetard ? ' : <b class="is-warn">dépassée</b>' : ''}. Aucune tâche ATLAS ne la suit encore.</p>`;
    } else if (sp.kind === 'offre') {
      titre = `Offre ${sp.o.offre} ${sp.o.libelle}, pas de réponse`;
      corps = `<p class="tb-seance-resume">Envoyée à ${h(sp.o.a)} ${h(ilYA(sp.o.sentAt, Date.now()))}${sp.o.montant ? ` (${h(sp.o.montant)})` : ''}. Le lien de l'offre a été ouvert (clic sur le lien du portail, sans pixel de suivi) et rien n'est revenu depuis. ARGO prépare une relance courte, sans jamais dire au client qu'il a ouvert l'offre.</p>`;
    } else if (sp.kind === 'cr') {
      titre = `Compte rendu à envoyer : ${sp.r.titre}`;
      corps = `<p class="tb-seance-resume">Validé ${h(ilYA(sp.r.valideLe, Date.now()))}, pas encore partagé avec les participants. Le module Réunion compose le mail de suivi (relu avant envoi).</p>`;
    } else {
      titre = 'Je le fais seul désormais ?';
      corps = `<p class="tb-seance-resume">Tu as choisi « ${h(sp.a.libelle)} » ${sp.a.acceptees} fois d'affilée.${s?.mode !== 'actif' ? ' Tant que l\'agent observe, rien ne bouge : ce sera prêt pour la suite.' : ''}</p>`;
    }
    $('sc-main').innerHTML = `
      <article class="tb-seance-carte is-special" aria-label="${h(titre)}">
        <h2 class="tb-seance-objet">${h(titre)}</h2>
        ${corps}
        <div class="tb-seance-touches">
          ${boutons.map((b, k) => `<button type="button" class="tb-btn${k === 0 ? ' is-primary tb-seance-reco' : ''}" data-sp="${k}"><span class="tb-kbd">${k === 0 ? 'Entrée' : k + 1}</span><span><b>${h(b.libelle)}</b>${b.detail ? `<small>${h(b.detail)}</small>` : ''}</span></button>`).join('')}
          <button type="button" class="tb-btn is-ghost" data-k="Escape"><span class="tb-kbd">Échap</span>Passer</button>
        </div>
      </article>`;
    $('sc-main').querySelectorAll<HTMLButtonElement>('[data-sp]').forEach(b => b.addEventListener('click', () => void faireSpecial(cur, Number(b.dataset.sp))));
    $('sc-main').querySelectorAll<HTMLButtonElement>('[data-k]').forEach(b => b.addEventListener('click', () => touche(b.dataset.k!)));
    ($('sc-main').querySelector('[data-sp="0"]') as HTMLElement | null)?.focus();
  }

  function retirerSpecial(cur: { liste: 'avant' | 'apres'; item: Special }): void {
    if (cur.liste === 'avant') avant = avant.filter(x => x !== cur.item); else apres = apres.filter(x => x !== cur.item);
  }

  async function faireSpecial(cur: { liste: 'avant' | 'apres'; item: Special }, k: number): Promise<void> {
    if (occupe) return;
    const b = boutonsSpecial(cur.item)[k];
    if (!b) return;
    occupe = true;
    root.classList.add('is-busy');
    try {
      const r = await b.faire();
      if (!r) return;
      const sortis = typeof r.sortie === 'number' ? r.sortie : 0;
      if (sortis && s?.dansLaBoite != null) s.dansLaBoite = Math.max(0, s.dansLaBoite - sortis);
      pile.push({ index: i, recommandee: null, faite: `special:${cur.item.kind}`, annuler: r.annuler, sortie: sortis, special: cur });
      traites += cur.item.kind === 'groupe' ? cur.item.g.messageIds.length : 1;
      retirerSpecial(cur);
      toast(r.message, 'success', r.annuler ? () => void defaire() : undefined);
      renderCarte();
    } catch (e) {
      toast(humanError(e), 'error');
    } finally {
      occupe = false;
      root.classList.remove('is-busy');
    }
  }

  function renderCarte(): void {
    majCompteur();
    const main = $('sc-main');
    const sp = specialCourant();
    if (sp) { renderSpecial(sp); return; }
    const c = cartes[i];
    if (!c) { renderFin(); return; }
    const rec = c.recommandee;
    const risque = !!c.risque;
    main.innerHTML = `
      <article class="tb-seance-carte${risque ? ' is-risque' : ''}" aria-label="Mail de ${h(c.from?.name || c.from?.email)}">
        ${risque ? `<div class="tb-seance-risque" role="alert">${icon('alert', 16)}<div><b>${c.risque!.niveau === 'eleve' ? 'Mail dangereux' : 'Mail suspect'} : aucune action en une touche.</b><span>${h(c.risque!.raisons.join(' · ') || 'Signaux de risque relevés par ATLAS.')}</span><span>N'ouvre ni lien ni pièce jointe ; vérifie-le dans Outlook (panneau ATLAS : « Ce mail est sûr » ou quarantaine).</span></div></div>` : ''}
        <div class="tb-seance-de"><b>${h(c.from?.name || c.from?.email || '')}</b><span>${h(c.from?.email || '')}</span><span class="tb-when">${h(ilYA(c.receivedAt, Date.now()))}</span></div>
        <h2 class="tb-seance-objet">${h(c.subject || '(sans objet)')}</h2>
        <p class="tb-seance-resume">${h(c.resume || '')}</p>
        ${c.raisons.length ? `<div class="tb-chips is-gauche">${c.raisons.map(r => `<span class="tb-chip${/Client|Offre|Projet|Prospect|Devis/.test(r) ? ' is-argo' : ''}">${h(r)}</span>`).join('')}</div>` : ''}
        ${c.brouillon && !risque ? `<details class="tb-seance-brouillon"${rec?.type === 'repondre' ? ' open' : ''}><summary>${icon('reply', 14)}Brouillon prêt : ${h(c.brouillon.resumeIntention || 'réponse préparée')}</summary><pre>${h(c.brouillon.texte)}</pre>${(c.brouillon.alertes || []).map(a => `<p class="tb-note${a.gravite === 'bloquant' ? ' is-err' : ''}">${h(a.message)}</p>`).join('')}</details>` : ''}
        ${!risque && c.suggestions?.length ? `<div class="tb-seance-suggestions" aria-label="Réponses courtes">${c.suggestions.map((t, k) => `<button type="button" class="tb-btn is-ghost" data-sugg="${k}"><span class="tb-kbd">${k + 5}</span><span>${h(t)}</span></button>`).join('')}</div>` : !risque && attendReponse(c) ? '<p class="tb-note" id="sc-sugg-attente">Réponses courtes en préparation…</p>' : ''}
        ${!risque && c.miseEnRelation ? `<button type="button" class="tb-btn" id="sc-intro"><span class="tb-kbd">8</span><span><b>Mise en relation : répondre</b><small>L'introducteur passe en copie cachée, avec un merci</small></span></button>` : ''}
        ${rec ? `<button type="button" class="tb-btn is-primary tb-seance-reco" id="sc-reco"><span class="tb-kbd">Entrée</span><span><b>${h(rec.libelle)}</b>${rec.detail ? `<small>${h(rec.detail)}</small>` : ''}</span></button>${c.appris ? `<p class="tb-note">${h(c.appris)}</p>` : ''}` : ''}
        <div class="tb-seance-touches">
          ${risque
            ? `<button type="button" class="tb-btn" data-alt="ouvrir">${icon('external', 14)}Vérifier dans Outlook</button><button type="button" class="tb-btn" data-alt="plus-tard">${icon('clock', 14)}Plus tard (demain)</button><button type="button" class="tb-btn is-ghost" data-k="Escape"><span class="tb-kbd">Échap</span>Passer</button>`
            : `<button type="button" class="tb-btn" data-k="1"><span class="tb-kbd">1</span>Répondre</button>
               <button type="button" class="tb-btn" data-k="2"><span class="tb-kbd">2</span>Confier</button>
               <button type="button" class="tb-btn" data-k="3"><span class="tb-kbd">3</span>Plus tard</button>
               <button type="button" class="tb-btn" data-k="4"><span class="tb-kbd">4</span>Autre</button>
               <button type="button" class="tb-btn is-ghost" data-k="Escape"><span class="tb-kbd">Échap</span>Passer</button>`}
        </div>
      </article>
      <aside class="tb-seance-ctx tb-panel" aria-label="Contexte">${contexteHtml(c)}</aside>`;
    main.querySelector('#sc-reco')?.addEventListener('click', () => { if (rec) void faire(c, rec); });
    main.querySelectorAll<HTMLButtonElement>('[data-sugg]').forEach(b => b.addEventListener('click', () => void repondreSuggestion(c, Number(b.dataset.sugg))));
    main.querySelector('#sc-intro')?.addEventListener('click', () => void miseEnRelation(c));
    carteAfficheeLe = Date.now();
    if (!risque && !c.suggestions?.length && attendReponse(c)) void chargerSuggestions(c);
    main.querySelectorAll<HTMLButtonElement>('[data-k]').forEach(b => b.addEventListener('click', () => touche(b.dataset.k!)));
    main.querySelectorAll<HTMLButtonElement>('[data-alt]').forEach(b => b.addEventListener('click', () => {
      const a = b.dataset.alt === 'ouvrir' ? c.alternatives.find(x => x.type === 'ouvrir') : c.alternatives.find(x => x.type === 'plus-tard');
      if (a) void faire(c, a);
    }));
    (main.querySelector('#sc-reco') as HTMLElement | null)?.focus();
  }

  function attendReponse(c: SeanceCarte): boolean {
    return c.raisons.includes('Attend une réponse') && !suggestionsDemandees.has(c.messageId);
  }

  /** Trois réponses courtes à la demande (cartes au-delà des premières, préparées par le worker). */
  async function chargerSuggestions(c: SeanceCarte): Promise<void> {
    suggestionsDemandees.add(c.messageId);
    try {
      const r = await fetchSuggestions(c.messageId, c.mailbox);
      if (r.textes?.length) c.suggestions = r.textes.slice(0, 3);
    } catch { /* facultatif */ }
    if (cartes[i] === c && !specialCourant()) renderCarte();
    else document.getElementById('sc-sugg-attente')?.remove();
  }

  async function repondreSuggestion(c: SeanceCarte, k: number): Promise<void> {
    const t = c.suggestions?.[k];
    if (!t || occupe) return;
    const a: SeanceAction = { cle: 'repondre', type: 'repondre', libelle: 'Réponse courte', donnees: { pret: true } };
    // Formule de fin de la charte selon la langue du mail (signature ajoutée à l'envoi par le serveur).
    const fin = ({ EN: 'Kind regards,', DE: 'Mit freundlichen Grüßen,', LB: 'Mat beschte Gréiss,' } as Record<string, string>)[String(c.langue || '').toUpperCase()] || 'Bien à vous,';
    await faire(c, a, undefined, `${t}\n\n${fin}`);
  }

  /** Mise en relation : réponse préparée (merci + introducteur en copie cachée), ouverte dans un nouveau message. */
  async function miseEnRelation(c: SeanceCarte): Promise<void> {
    if (occupe) return;
    occupe = true;
    try {
      const r = await preparerMiseEnRelation(c.messageId, c.mailbox);
      const u = `mailto:${r.a.map(encodeURIComponent).join(',')}?bcc=${r.cci.map(encodeURIComponent).join(',')}&subject=${encodeURIComponent(r.sujet)}&body=${encodeURIComponent(r.texte)}`;
      const copie = await copierTexte(r.texte);
      opts.openLink(u);
      toast(`Nouveau message ouvert : ${r.introducteur.name || r.introducteur.email} en copie cachée. Relis puis envoie${copie ? ' (texte aussi copié)' : ''}`, 'success');
    } catch (e) {
      toast(humanError(e), 'error');
    } finally { occupe = false; }
  }

  function renderFin(): void {
    const n = s?.dansLaBoite;
    const obj = s?.objectif ?? 5;
    $('sc-main').innerHTML = `
      <div class="tb-seance-fin tb-panel is-raised">
        <div class="tb-h">${n != null && n <= obj ? 'Objectif atteint' : 'Séance terminée'}</div>
        <p>${traites} mail${traites > 1 ? 's' : ''} trié${traites > 1 ? 's' : ''} en ${minutesLisibles(Date.now() - debut)}.${n != null ? ` Il reste ${n} mail${n > 1 ? 's' : ''} dans la boîte de réception${n > obj ? ` (objectif ${obj})` : ''}.` : ''}</p>
        <p class="tb-note">Les mails mis de côté ou confiés reviennent d'eux-mêmes s'ils n'ont pas bougé.</p>
        <div class="tb-actions"><button type="button" class="tb-btn is-primary" id="sc-relire">Relire la boîte</button><button type="button" class="tb-btn" id="sc-fermer">Fermer</button></div>
      </div>`;
    root.querySelector('#sc-relire')?.addEventListener('click', () => void charger(true));
    root.querySelector('#sc-fermer')?.addEventListener('click', () => void fermer());
    majCompteur();
  }

  // ── Actions ──
  function suivante(sortie: boolean): void {
    if (sortie && s?.dansLaBoite != null) s.dansLaBoite = Math.max(0, s.dansLaBoite - 1);
    i++;
    renderCarte();
  }

  /** Répondre : brouillon déposé (verrou du tableau ouvert), sinon formulaire de réponse d'Outlook, sinon texte copié et mail ouvert. */
  async function repondre(c: SeanceCarte, impose?: string): Promise<boolean> {
    let texte = impose || c.brouillon?.texte || '';
    if (!texte) {
      toast('ARGO rédige la réponse…', 'info');
      const r = await redigerArgo(c.messageId, c.mailbox);
      texte = r.texte || '';
      if (!texte) { toast('Rédaction indisponible : réponds depuis Outlook', 'error'); openAgentMail(c, opts.openLink); return false; }
    }
    if (s?.ecrituresActives) {
      try {
        const d = await deposerBrouillon(c.messageId, c.mailbox, texte);
        if (d.depose && d.webLink) { opts.openLink(d.webLink); toast('Brouillon déposé dans le fil : relis puis envoie depuis Outlook', 'success'); return true; }
      } catch { /* repli ci-dessous */ }
    }
    if (opts.repondreDansOutlook) {
      try { if (await opts.repondreDansOutlook(c.messageId, texteVersHtml(texte))) { toast('Réponse ouverte dans Outlook : relis puis envoie', 'success'); return true; } } catch { /* repli */ }
    }
    const copie = await copierTexte(texte);
    openAgentMail(c, opts.openLink);
    toast(copie ? 'Texte copié : clique « Répondre » dans Outlook, colle, relis puis envoie' : 'Mail ouvert : réponds depuis Outlook (copie impossible ici)', copie ? 'success' : 'info');
    return true;
  }

  async function choisirCollegue(c: SeanceCarte): Promise<{ email: string; nom: string } | null> {
    const liste = s?.collegues || [];
    if (!liste.length) { toast('Aucun collègue suivi par l\'agent', 'error'); return null; }
    const k = await menu(`Confier « ${c.subject || '(sans objet)'} » à…`, liste.map(x => ({ libelle: x.nom, detail: x.email })));
    return k == null ? null : liste[k];
  }

  async function faire(c: SeanceCarte, a: SeanceAction, quand?: string, texteImpose?: string): Promise<void> {
    if (occupe) return;
    if (a.type === 'ouvrir') { if (!openAgentMail(c, opts.openLink)) toast('Lien du mail indisponible : ouvre-le depuis Outlook', 'error'); return; }
    occupe = true;
    root.classList.add('is-busy');
    try {
      let annuler: (() => Promise<void>) | undefined;
      let sortie = false;
      let message = '';
      switch (a.type) {
        case 'repondre': {
          if (!(await repondre(c, texteImpose))) return;
          break;
        }
        case 'classer-projet': {
          const projetId = String(a.donnees?.projetId || '');
          if (!projetId) throw new Error('projet inconnu');
          try {
            await callAtlasWorker('emails/link-projet', {
              email: { id: c.graphId, internetMessageId: c.messageId, subject: c.subject, from: c.from, receivedAt: c.receivedAt },
              projetId, linkedByName: prenomDe(c.mailbox), direction: 'reçu',
            });
          } catch (e) { console.warn('[seance] liaison au projet :', e); }
          const d = await fetchDossierProjet(projetId, c.mailbox, c.messageId);
          const r = d.existant
            ? await rangerDansDossier({ messageId: c.messageId, mailbox: c.mailbox, dossierId: d.existant.id, projetId })
            : await rangerDansDossier({ messageId: c.messageId, mailbox: c.mailbox, creer: { chemin: d.propose || String(a.donnees?.chemin || ''), projetId } });
          if (r.actionId) { const id = r.actionId; annuler = () => annulerAction(id); }
          sortie = true;
          message = `Lié et classé dans ${r.dossier.chemin}`;
          break;
        }
        case 'archiver': {
          const r = await rangerDansDossier({ messageId: c.messageId, mailbox: c.mailbox, dossierId: String(a.donnees?.dossierId || '') });
          if (r.actionId) { const id = r.actionId; annuler = () => annulerAction(id); }
          sortie = true;
          message = 'Archivé';
          break;
        }
        case 'metier': {
          const type = String(a.donnees?.typeMetier || '');
          const r = await executerAction({ messageId: c.messageId, mailbox: c.mailbox, type, donnees: (a.donnees?.proposition as Record<string, unknown>) || undefined });
          if (!r.ok) {
            if (r.erreur === 'choix-requis') { toast(`${r.resume || 'Un choix est nécessaire'} : fais-le depuis le panneau ATLAS du mail`, 'info'); openAgentMail(c, opts.openLink); }
            else toast(r.resume || 'Action impossible', 'error');
            return;
          }
          if (r.actionId) { const id = r.actionId; annuler = () => annulerAction(id); }
          message = r.resume || a.libelle;
          break;
        }
        case 'confier': {
          let qui = a.donnees?.a ? { email: String(a.donnees.a), nom: String(a.donnees.nom || '') } : null;
          if (!qui) qui = await choisirCollegue(c);
          if (!qui) return;
          let delai = quand;
          if (!delai) {
            const k = await menu('Revient si rien ne bouge…', [{ libelle: 'Demain' }, { libelle: 'Lundi' }]);
            if (k == null) return;
            delai = k === 0 ? 'demain' : 'lundi';
          }
          const r = await confierMail({ messageId: c.messageId, mailbox: c.mailbox, a: qui.email, quand: delai });
          annuler = async () => { await annulerConfier(c.messageId, c.mailbox); };
          message = `Confié à ${r.nom || qui.nom}${r.prevenu ? ' (prévenu dans ATLAS)' : ''} : revient ${delai === 'lundi' ? 'lundi' : 'demain'} si rien ne bouge`;
          break;
        }
        case 'plus-tard': {
          const q = quand || String(a.donnees?.quand || 'demain');
          await mettrePlusTard(c.messageId, q);
          annuler = async () => { await retirerPlusTard(c.messageId); };
          message = `De côté jusqu'à ${q === 'lundi' ? 'lundi' : 'demain'}`;
          break;
        }
      }
      noterDecisionSeance(c.recommandee?.cle ?? null, a.cle, false, c.patron, { dureeMs: Date.now() - carteAfficheeLe, client: c.contexte.tiers });
      pile.push({ carte: c, index: i, recommandee: c.recommandee?.cle ?? null, faite: a.cle, annuler, sortie, ...(c.patron ? { patron: c.patron } : {}) });
      traites++;
      if (message) toast(message, 'success', () => void defaire());
      suivante(sortie);
    } catch (e) {
      toast(humanError(e), 'error');
    } finally {
      occupe = false;
      root.classList.remove('is-busy');
    }
  }

  async function defaire(): Promise<void> {
    const d = pile.pop();
    if (!d || occupe) { if (d) pile.push(d); return; }
    occupe = true;
    try {
      if (d.annuler) await d.annuler();
      const n = typeof d.sortie === 'number' ? d.sortie : d.sortie ? 1 : 0;
      if (n && s?.dansLaBoite != null) s.dansLaBoite += n;
      if (d.special) {
        // Carte spéciale : elle revient en tête de sa liste.
        if (d.special.liste === 'avant') avant = [d.special.item, ...avant]; else apres = [d.special.item, ...apres];
        if (d.faite !== 'passer') traites = Math.max(0, traites - (d.special.item.kind === 'groupe' ? d.special.item.g.messageIds.length : 1));
        if (d.special.liste === 'apres') i = Math.min(i, cartes.length);
        toast('Action annulée', 'info');
        return;
      }
      if (!d.carte) return;
      if (d.faite !== 'passer') { noterDecisionSeance(d.recommandee, d.faite, true, d.patron); traites = Math.max(0, traites - 1); }
      // La carte revient là où elle était.
      cartes = cartes.filter(x => x !== d.carte);
      i = Math.min(d.index, cartes.length);
      cartes.splice(i, 0, d.carte);
      toast('Action annulée', 'info');
    } catch (e) {
      pile.push(d);
      toast(humanError(e), 'error');
    } finally {
      occupe = false;
      renderCarte();
    }
  }

  async function autre(c: SeanceCarte): Promise<void> {
    const alts = [...(c.recommandee ? [c.recommandee] : []), ...c.alternatives];
    const k = await menu('Autre action', alts.map(a => ({ libelle: a.libelle, detail: a.detail })));
    if (k != null) await faire(c, alts[k]);
  }

  /** « Fais-le revenir quand… » : conditions possibles d'après le contexte du mail (projet, devis fournisseur). */
  function conditionsRetour(c: SeanceCarte): Array<{ libelle: string; detail?: string; condition: ConditionRetour }> {
    const out: Array<{ libelle: string; detail?: string; condition: ConditionRetour }> = [];
    const p = c.contexte.projet;
    if (p?.id && p.echeance) out.push({ libelle: 'À J-7 de l\'événement', detail: `${p.libelle}${p.dates ? ` · ${p.dates}` : ''}`, condition: { type: 'evenement-j7', projetId: p.id, libelle: p.libelle } });
    if (p?.id) out.push({ libelle: 'Quand le projet passe à l\'étape suivante', detail: `${p.libelle}${p.statut ? ` · aujourd'hui « ${p.statut} »` : ''}`, condition: { type: 'etape-projet', projetId: p.id, statutInitial: p.statut || '', libelle: p.libelle } });
    const d = c.contexte.devis;
    if (d?.id && d.expireLe) out.push({ libelle: 'Avant l\'expiration du devis', detail: `${d.libelle} · expire le ${jourFr(d.expireLe)}`, condition: { type: 'expiration-devis', devisId: d.id, libelle: d.libelle } });
    return out;
  }

  async function revenirQuand(c: SeanceCarte, condition: ConditionRetour, libelle: string): Promise<void> {
    if (occupe) return;
    occupe = true;
    root.classList.add('is-busy');
    try {
      const r = await poserRetour(c.messageId, c.mailbox, condition);
      const annuler = async () => { await retirerRetour(c.messageId, c.mailbox); };
      noterDecisionSeance(c.recommandee?.cle ?? null, 'plus-tard', false, c.patron);
      pile.push({ carte: c, index: i, recommandee: c.recommandee?.cle ?? null, faite: 'plus-tard', annuler, sortie: false, ...(c.patron ? { patron: c.patron } : {}) });
      traites++;
      toast(`De côté : revient ${libelle.charAt(0).toLowerCase()}${libelle.slice(1)}${r.attente ? ` (${r.attente.toLowerCase()})` : ''}`, 'success', () => void defaire());
      suivante(false);
    } catch (e) {
      toast(humanError(e), 'error');
    } finally {
      occupe = false;
      root.classList.remove('is-busy');
    }
  }

  function touche(k: string): void {
    if (occupe) return;
    const sp = specialCourant();
    if (sp) {
      if (k === 'Escape') { pile.push({ index: i, recommandee: null, faite: 'passer', sortie: false, special: sp }); retirerSpecial(sp); renderCarte(); return; }
      if (k === 'Enter') { void faireSpecial(sp, 0); return; }
      const n = Number(k);
      if (n >= 2 && n <= 4) void faireSpecial(sp, n - 1);
      return;
    }
    const c = cartes[i];
    if (!c) return;
    if (k === 'Escape') {
      pile.push({ carte: c, index: i, recommandee: null, faite: 'passer', sortie: false });
      suivante(false);
      return;
    }
    if (c.risque) { toast('Mail à risque : aucune action en une touche. Vérifie-le dans Outlook.', 'error'); return; }
    if (k === 'Enter') { if (c.recommandee) void faire(c, c.recommandee); return; }
    if (k === '1') { const a = [c.recommandee, ...c.alternatives].find(x => x?.type === 'repondre'); if (a) void faire(c, a); return; }
    if (k === '2') {
      const a = [c.recommandee, ...c.alternatives].find(x => x?.type === 'confier' && x.donnees?.a) || c.alternatives.find(x => x.type === 'confier');
      if (a) void faire(c, a);
      return;
    }
    if (k === '3') {
      const conds = conditionsRetour(c);
      void menu('Fais-le revenir…', [{ libelle: 'Demain' }, { libelle: 'Lundi' }, ...conds.map(x => ({ libelle: x.libelle, detail: x.detail }))]).then(n => {
        if (n == null) return;
        if (n >= 2) { void revenirQuand(c, conds[n - 2].condition, conds[n - 2].libelle); return; }
        const a = [c.recommandee, ...c.alternatives].find(x => x?.type === 'plus-tard');
        if (a) void faire(c, a, n === 0 ? 'demain' : 'lundi');
      });
      return;
    }
    if (k === '4') { void autre(c); return; }
    if (k === '5' || k === '6' || k === '7') { const n = Number(k) - 5; if (c.suggestions?.[n]) void repondreSuggestion(c, n); return; }
    if (k === '8' && c.miseEnRelation) void miseEnRelation(c);
  }

  function onKey(e: KeyboardEvent): void {
    if (ferme || document.querySelector('.tb-overlay')) return;
    const cible = e.target as HTMLElement;
    if (cible instanceof HTMLInputElement || cible instanceof HTMLTextAreaElement) return;
    if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'z') { e.preventDefault(); e.stopPropagation(); void defaire(); return; }
    if (e.metaKey || e.ctrlKey || e.altKey) return;
    if (['Enter', 'Escape', '1', '2', '3', '4', '5', '6', '7', '8'].includes(e.key)) {
      // Entrée sur un bouton de la carte : le clic s'en charge.
      if (e.key === 'Enter' && cible instanceof HTMLButtonElement && cible.id !== 'sc-reco' && cible.dataset.sp !== '0') return;
      e.preventDefault();
      e.stopPropagation();
      touche(e.key);
    }
  }
  document.addEventListener('keydown', onKey, true);

  async function fermer(): Promise<void> {
    if (ferme) return;
    ferme = true;
    document.removeEventListener('keydown', onKey, true);
    if (traites || Date.now() - debut > 30_000) {
      try { await finSeance({ dureeMs: Date.now() - debut, traites }); } catch { /* statistiques facultatives */ }
    }
    root.classList.remove('is-on');
    window.setTimeout(() => root.remove(), 160);
    opts.onFerme?.();
  }

  async function charger(frais = false): Promise<void> {
    $('sc-main').innerHTML = '<div class="loading"><div class="spinner"></div><p>ARGO prépare tes cartes…</p></div>';
    try {
      s = await fetchSeance(frais);
      cartes = s.cartes || [];
      avant = [
        ...(s.groupes || []).map(g => ({ kind: 'groupe' as const, g })),
        ...(s.promesses || []).map(p => ({ kind: 'promesse' as const, p })),
        ...(s.offresConsultees || []).map(o => ({ kind: 'offre' as const, o })),
        ...(s.crsASuivre || []).map(r => ({ kind: 'cr' as const, r })),
      ];
      apres = [
        ...(s.desabonnements || []).map(d => ({ kind: 'desabonnement' as const, d })),
        ...(s.sollicitations?.mails.length ? [{ kind: 'sollicitations' as const, s: s.sollicitations }] : []),
        ...(s.autonomie || []).map(a => ({ kind: 'autonomie' as const, a })),
      ];
      i = 0;
      $('sc-pied').innerHTML = renderStats(s.stats || []);
      const conc = s.concentration;
      if (conc?.actif && conc.retenus) $('sc-pied').insertAdjacentHTML('beforeend', `<span>Concentration : ${conc.retenus} mail${conc.retenus > 1 ? 's' : ''} non urgent${conc.retenus > 1 ? 's' : ''} attend${conc.retenus > 1 ? 'ent' : ''} ${h(conc.libelle || 'le prochain créneau')}</span>`);
      if (s.sollicitationsHorsSeance && !s.sollicitations) $('sc-pied').insertAdjacentHTML('beforeend', `<span>${s.sollicitationsHorsSeance} sollicitation${s.sollicitationsHorsSeance > 1 ? 's' : ''} commerciale${s.sollicitationsHorsSeance > 1 ? 's' : ''} hors séance</span>`);
      if (s.mode !== 'actif') $('sc-pied').insertAdjacentHTML('beforeend', '<span class="tb-note">L\'agent observe : rien n\'est fait sans ta touche.</span>');
      renderCarte();
    } catch (e) {
      $('sc-main').innerHTML = `<div class="tb-empty"><strong>Séance indisponible</strong><span>${h(humanError(e))}</span><span class="tb-actions"><button type="button" class="tb-btn is-primary" id="sc-retry">Réessayer</button></span></div>`;
      root.querySelector('#sc-retry')?.addEventListener('click', () => void charger(true));
    }
  }

  $('sc-annuler').addEventListener('click', () => void defaire());
  $('sc-fin').addEventListener('click', () => void fermer());
  void charger(true);
}
