/**
 * agent-panel.ts — Panneau LECTURE « Agent » (phase 1.5 de l'agent d'inbox).
 *
 * Maquette : .claude/BACKLOG-AGENT-INBOX.md §7.
 *   En-tête (pile + catégorie) · résumé en une ligne · faits utiles · action principale +
 *   « Autre action ▾ » · Correspondant (fiche ATLAS) · Brouillon prêt (masqué sans brouillon) ·
 *   « Pourquoi cette proposition ? » (source / règle / catégorie).
 *
 * Règle d'or : AUCUN appel IA à l'ouverture d'un mail. L'état vient de
 * GET /api/plugin/agent/state (calculé une seule fois par l'agent serveur). S'il est « pending »,
 * on affiche « Analyse en cours » (le worker enfile alors le mail en priorité) et on relit à 5 s, 15 s, 30 s, puis bouton « Actualiser ». La ré-analyse
 * reste une action volontaire (onglet IA, bouton « Re-analyser »).
 *
 * Tant que l'état est en mode `blanc` (mode à blanc, phase 1.7), l'action principale et le menu
 * « Autre action » sont désactivés avec la mention « Agent en observation ».
 *
 * Phase 2 : mail d'un expéditeur inconnu (pile `a_filtrer`) → encart « Nouvel expéditeur » avec
 * Accepter / Refuser (POST /api/plugin/agent/expediteur, une fois pour toutes) ; section repliable
 * « Ce que l'agent a fait » (journal + Annuler), lue seulement quand on la déplie.
 *
 * Phase 3 / 4 (04/10/2026, tout à la demande, rien n'est notifié) : « Que faire de ce mail ? »
 * (3 cartes au plus, choix du projet si requis, « Fait » + Annuler, actions déjà faites en grisé),
 * « Répondre plus tard » (mail reçu) ou « Relancer si pas de réponse » (mail envoyé), réunions
 * liées (liens ATLAS), brouillon prêt (intention + alertes du contrôle avant envoi), « Résumer le
 * fil », « Poser une question à ma boîte » : voir agent-outils.ts.
 *
 * Phase 5 (good@ et équipe) : pour un mail de good@ (membres autorisés), bloc « Équipe » (assigné
 * à / pris par, Je prends / Relâcher / Attribuer à…, avertissement de collision, commentaires
 * internes) : voir agent-equipe.ts. L'état est cherché dans toutes les boîtes visibles de la
 * personne (la sienne d'abord, puis good@) ; un mail de good@ garde ses actions métier sur good@.
 *
 * Phase 6 (à la demande) : encart « À savoir » (rebond / départ, absence avec date de retour,
 * candidature : `detections` de l'état), « Traduire » (FR / EN / DE / LB), « Modèles de réponse »
 * (texte à copier ; « Répondre avec ce texte » par displayReplyForm, Mailbox 1.0, hors mobile),
 * fiche mémoire du contact (la personne seule), pièces jointes lourdes → lien court.
 *
 * Mobile : une colonne, boutons larges, aucun survol, aucune fonction de rédaction ; seules des
 * API Mailbox ≤ 1.5 sont appelées sans garde (les autres passent par platform.ts). Les boutons
 * d'écriture (actions, rappels) restent disponibles : ils appellent le worker, pas Office.js.
 */

import { fetchAgentState, fetchEmailSuggestions, decideExpediteur, fetchDetections, definirPreferenceTri, ciblePreference, type EmailActionSuggestion, type AgentPile } from '../api/agent';
import type { InboxMessageState, InboxEngagement } from '../api/inbox-agent.types';
import { renderJournal } from './agent-lists';
import { refreshJourneeBanner } from './journee-banner';
import { showToast } from '../taskpane';
import { getAllContacts, getAllTiers, getProjetsByClient, fetchContactArgoProfile } from '../api/airtable';
import { lireEtatAgent } from '../api/agent';
import { fetchVueMail, fetchCorrespondant, vueMemorisee, memoriserVue, type VueMail, type FicheCorrespondant } from '../api/vue-mail';
import { AtlasError } from '../api/net';
import { isMobile, openExternal, supportsMailbox, ATLAS_BASE } from '../api/platform';
import {
  renderActionsMetier, renderRappels, renderResumeFil, renderQuestion, lienReunion,
  renderTraduction, renderModeles, renderFicheContact, renderDetections, renderPiecesLourdes,
} from './agent-outils';
import { renderEquipe } from './agent-equipe';
import { convertToRestId, getGraphToken } from '../api/graph';
import { renderOffreRecue, renderClasserOutlook } from './agent-dossiers';
import { renderCartePret } from './carte-pret';
import { fetchTraites, type TraiteARanger } from '../api/tableau';
import { renderPlusActions } from './agent-parite';
import { renderSecurite } from './agent-securite';
import { escapeHtml } from '../utils/html';
import { humanError } from '../api/net';
import { icon } from '../ui/icons';

type EtatAvecBrouillon = InboxMessageState;

/** Motif du brouillon préparé, en clair. */
const MOTIF_BROUILLON: Record<string, string> = {
  reponse: 'Réponse',
  'relance-engagement': 'Relance d\'une promesse échue',
  'relance-envoi': 'Relance d\'un envoi sans réponse',
};

/** Relectures espacées de l'état tant que l'agent analyse (puis bouton « Actualiser »). */
const PENDING_RETRY_DELAYS_MS = [5000, 15000, 30000];

const PILE_LABELS: Record<AgentPile, string> = {
  a_traiter: 'À traiter par toi',
  en_attente: 'En attente d\'une réponse',
  pour_info: 'Pour info',
  bruit: 'Bruit',
  a_filtrer: 'À filtrer',
  plus_tard: 'Plus tard',
};

const PILE_CLASSES: Record<AgentPile, string> = {
  a_traiter: 'pile-a-traiter',
  en_attente: 'pile-en-attente',
  pour_info: 'pile-pour-info',
  bruit: 'pile-bruit',
  a_filtrer: 'pile-a-filtrer',
  plus_tard: 'pile-pour-info',
};

const URGENCE_LABELS = ['', 'Urgence faible', 'Urgence moyenne', 'Urgence haute'];

/** Domaines de messagerie grand public : jamais rattachés à un client par le domaine. */
const GENERIC_DOMAINS = new Set([
  'gmail.com', 'outlook.com', 'hotmail.com', 'hotmail.fr', 'live.com', 'yahoo.com', 'yahoo.fr',
  'icloud.com', 'me.com', 'pt.lu', 'gmx.de', 'gmx.net', 'web.de', 'orange.fr', 'free.fr',
]);

/** Statuts de projet considérés comme clos (pour « N projets en cours »). */
const CLOSED_STATUS = /termin|factur|annul|archiv|perdu|refus|clos/i;


function humanize(code: string): string {
  const s = String(code || '').replace(/[_-]+/g, ' ').trim();
  return s ? s.charAt(0).toUpperCase() + s.slice(1) : '';
}

/** Date seule (AAAA-MM-JJ) affichée telle quelle ; horodatage UTC affiché à l'heure du Luxembourg. */
function formatDate(value: string | undefined): string {
  if (!value) return '';
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return value;
  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    return d.toLocaleDateString('fr-FR', { timeZone: 'UTC', day: 'numeric', month: 'short' });
  }
  return d.toLocaleString('fr-FR', {
    timeZone: 'Europe/Luxembourg', day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit',
  });
}

/** 'fait' = promesse faite par la personne de la boîte ; 'recu' = promesse faite par le correspondant. */
function engagementLabel(sens: InboxEngagement['sens']): string {
  return sens === 'fait' ? 'Promis par toi' : 'Promis à toi';
}

interface CurrentItem {
  messageId: string;
  itemId: string;
  conversationId: string;
  subject: string;
  fromEmail: string;
  fromName: string;
  receivedAt: string;
  mailbox: string;
}

function readCurrentItem(): CurrentItem | null {
  try {
    const item = Office.context.mailbox?.item as any;
    if (!item) return null;
    return {
      messageId: String(item.internetMessageId || ''),
      itemId: String(item.itemId || ''),
      conversationId: String(item.conversationId || ''),
      subject: typeof item.subject === 'string' ? item.subject : '',
      fromEmail: String(item.from?.emailAddress || ''),
      fromName: String(item.from?.displayName || ''),
      receivedAt: item.dateTimeCreated ? new Date(item.dateTimeCreated).toISOString() : new Date().toISOString(),
      mailbox: String(Office.context.mailbox?.userProfile?.emailAddress || ''),
    };
  } catch {
    return null;
  }
}

export class AgentPanel {
  private root: HTMLElement;
  private navigate: (tabId: string) => void;
  private mobile = isMobile();
  private destroyed = false;
  private retryTimer: number | null = null;
  private tentatives = 0;
  private item: CurrentItem | null = null;
  private state: EtatAvecBrouillon | null = null;
  private suggestions: EmailActionSuggestion[] | null = null;
  /** Empreinte de l'état affiché (vue mémorisée) : la vue fraîche ne redessine que s'il a changé. */
  private signatureEtat = '';

  constructor(root: HTMLElement, navigate: (tabId: string) => void) {
    this.root = root;
    this.navigate = navigate;
    this.item = readCurrentItem();
    this.renderShell();
    // Lot 1 « vitesse » (09/10/2026) : dernière vue de ce mail affichée tout de suite (avant même la
    // connexion), puis UNE vue fraîche du worker (état, sécurité, projet, correspondant, traité) ;
    // les autres cartes partent en parallèle.
    this.chargerVue();
    this.renderOutils();
  }

  destroy(): void {
    this.destroyed = true;
    if (this.retryTimer !== null) window.clearTimeout(this.retryTimer);
    this.retryTimer = null;
  }

  private $(id: string): HTMLElement | null {
    return this.root.querySelector<HTMLElement>(`#${id}`);
  }

  // ── Squelette ──

  private renderShell(): void {
    if (!this.item) {
      this.root.innerHTML = '<p class="empty-state">Aucun mail sélectionné.</p>';
      return;
    }
    this.root.innerHTML = `
      <div class="agent-panel ${this.mobile ? 'is-mobile' : ''}">
        <section id="agent-securite" class="agent-securite" hidden></section>

        <section id="agent-pret" class="agent-section" hidden></section>

        <div id="agent-state" class="agent-state">
          <div class="state-loading" aria-busy="true" aria-label="Chargement du mail"><div class="skeleton-line skeleton-lg" style="width:58%"></div><div class="skeleton-line" style="width:92%"></div><div class="skeleton-line" style="width:76%"></div></div>
        </div>

        <section id="agent-detections" class="agent-section" hidden></section>

        <section id="agent-equipe" class="agent-section agent-equipe" hidden></section>

        <section id="agent-actions-metier" class="agent-section" hidden></section>

        <section id="agent-rappels" class="agent-section" hidden></section>

        <section id="agent-offre" class="agent-section" hidden></section>

        <section id="agent-classer" class="agent-section" hidden></section>

        <section id="agent-parite" class="agent-section" hidden></section>

        <section class="agent-section">
          <div class="agent-section-title">Correspondant</div>
          <div id="agent-correspondant" class="agent-correspondant">
            <div class="state-loading" aria-busy="true"><div class="skeleton-line" style="width:70%"></div><div class="skeleton-line skeleton-sm" style="width:45%"></div></div>
          </div>
          <div id="agent-fiche-contact" class="agent-fiche-contact" hidden></div>
        </section>

        <section id="agent-pj-lourdes" class="agent-section" hidden></section>

        <section id="agent-reunions" class="agent-section" hidden></section>

        <section id="agent-brouillon" class="agent-section" hidden></section>

        <section id="agent-resume-fil-wrap" class="agent-section" hidden></section>

        <section id="agent-traduction" class="agent-section" hidden></section>

        <section id="agent-modeles" class="agent-section" hidden></section>

        <section id="agent-question" class="agent-section"></section>

        <section class="agent-section">
          <button type="button" class="agent-link" id="agent-journal-toggle" aria-expanded="false">${icon('activity', 14)}Ce que l'agent a fait</button>
          <div id="agent-journal" hidden></div>
        </section>

        <div id="agent-why-wrap" class="agent-why-wrap" hidden>
          <button type="button" class="agent-link" id="agent-why-toggle" aria-expanded="false">${icon('question', 14)}Pourquoi cette proposition ?</button>
          <div id="agent-why" class="agent-why" hidden></div>
        </div>

        <div class="agent-footer">
          <button type="button" class="btn btn-secondary btn-block agent-btn" data-nav="ia"
            title="Thème, urgence, dossier conseillé et correction du classement">${icon('tag', 14)}Classement détaillé</button>
        </div>
      </div>
    `;
    this.root.querySelectorAll<HTMLButtonElement>('button[data-nav]').forEach(btn => {
      btn.addEventListener('click', () => this.navigate(btn.dataset.nav!));
    });
    this.$('agent-journal-toggle')?.addEventListener('click', () => this.toggleJournal());
    this.$('agent-why-toggle')?.addEventListener('click', () => {
      const why = this.$('agent-why');
      const toggle = this.$('agent-why-toggle');
      if (!why || !toggle) return;
      why.hidden = !why.hidden;
      toggle.setAttribute('aria-expanded', String(!why.hidden));
    });
  }

  // ── Vue du mail en un appel (lot 1 « vitesse », 09/10/2026) ──

  private async chargerVue(): Promise<void> {
    const it = this.item;
    if (!it) return;
    const memo = vueMemorisee(it.messageId);
    if (memo) this.appliquerVue(memo, true);
    let v: VueMail;
    try {
      v = await fetchVueMail({ messageId: it.messageId, conversationId: it.conversationId, from: it.fromEmail, nom: it.fromName });
    } catch (e) {
      if (this.destroyed) return;
      // Worker pas encore à jour (route absente) ou vue indisponible : appels séparés, comme avant.
      console.warn('[AgentPanel] vue du mail indisponible :', e instanceof AtlasError ? e.detail || e.message : e);
      this.loadSecurite();
      void this.loadState();
      void this.loadCorrespondant();
      this.chargerClasser();
      return;
    }
    if (this.destroyed || this.item !== it) return;
    memoriserVue(it.messageId, v);
    this.appliquerVue(v, false);
  }

  /** Dessine une vue (mémorisée : état, sécurité et correspondant seulement ; fraîche : tout). */
  private appliquerVue(v: VueMail, memorisee: boolean): void {
    const it = this.item;
    if (!it) return;
    const manque = new Set(v.aCharger);
    // Sécurité : verdict de l'agent ; sinon (mail sans verdict) demandé à part.
    const secu = this.$('agent-securite');
    if (secu) {
      if (v.securite) void renderSecurite(secu, { messageId: it.messageId, onInfo: showToast, reponse: v.securite });
      else if (!memorisee && manque.has('securite')) this.loadSecurite();
    }
    // État identique à la vue mémorisée déjà affichée : rien à redessiner (ni à redemander).
    const signature = JSON.stringify(v.etat);
    if (memorisee || signature !== this.signatureEtat) this.appliquerEtat(lireEtatAgent(v.etat));
    this.signatureEtat = signature;
    if (v.correspondant) this.afficherCorrespondant(v.correspondant);
    else if (!memorisee) {
      if (manque.has('correspondant')) void this.correspondantSeul();
      else this.afficherCorrespondant(null);
    }
    // Lot 2 (10/10/2026) : carte « Prêt » (projet reconnu à l'arrivée, Lier et classer), sans appel de plus.
    const pret = this.$('agent-pret');
    if (pret && !this.mobile) {
      renderCartePret(pret, v.pret || null, {
        messageId: it.messageId, mailbox: it.mailbox, conversationId: it.conversationId, onInfo: showToast,
        onChange: p => this.chargerClasser(p ? { projetId: p.id, projetLibelle: p.libelle } : { sansProjet: true }),
      });
    }
    if (memorisee) return;
    this.chargerClasser({
      ...(v.dossier && v.projet && !v.pret?.auto?.range ? { dossier: v.dossier } : {}),
      ...(v.projet ? { projetId: v.projet.id, projetLibelle: v.projet.libelle, projetPropose: v.projet.source === 'agent' } : manque.has('projet') ? {} : { sansProjet: true }),
      ...(v.traite && v.traite.statut === 'a_ranger' ? { repondu: v.traite } : {}),
      traiteAFaire: manque.has('traite'),
    });
  }

  // ── Sécurité du mail (backlog reczJ0zLhPXqkWF9S) : bandeau rouge en tête, masqué si risque faible ──

  private loadSecurite(): void {
    const host = this.$('agent-securite');
    if (!host || !this.item?.messageId) return;
    void renderSecurite(host, { messageId: this.item.messageId, onInfo: showToast });
  }

  // ── État de l'agent (sans IA) ──

  private async loadState(): Promise<void> {
    if (!this.item) return;
    // Boîte vide : le worker cherche dans toutes les boîtes visibles (la sienne d'abord, puis good@
    // pour ses membres) ; un mail de good@ est ainsi reconnu même ouvert depuis la boîte partagée.
    const r = await fetchAgentState(this.item.messageId, '');
    if (this.destroyed) return;
    this.appliquerEtat(r);
  }

  private appliquerEtat(r: Awaited<ReturnType<typeof fetchAgentState>>): void {
    if (this.destroyed) return;
    const host = this.$('agent-state');
    if (!host) return;

    if (r.kind === 'state') {
      this.state = r.state as EtatAvecBrouillon;
      this.adapterBoitePartagee(this.state);
      this.renderState(host, this.state);
      this.renderBrouillon(this.state);
      this.renderReunions(this.state);
      this.renderDetectionsDuMail(this.state);
      this.renderWhy(this.state);
      if (this.item && !this.item.conversationId && this.state.conversationId) {
        this.item.conversationId = this.state.conversationId;
        const wrap = this.$('agent-resume-fil-wrap');
        if (wrap) renderResumeFil(wrap, { conversationId: this.item.conversationId, mailbox: this.item.mailbox });
      }
      return;
    }

    if (r.kind === 'pending') {
      host.innerHTML = `
        <div class="agent-notice">
          <div class="spinner"></div>
          <div><strong>Analyse en cours</strong><br/><span class="agent-muted">L'agent traite ce mail, résultat dans quelques secondes.</span></div>
        </div>
      `;
      const delai = PENDING_RETRY_DELAYS_MS[this.tentatives];
      if (delai !== undefined) {
        // Relectures espacées (5 s, 15 s, 30 s) ; jamais d'appel IA ici.
        this.tentatives++;
        this.retryTimer = window.setTimeout(() => {
          this.retryTimer = null;
          if (!this.destroyed) this.loadState();
        }, delai);
      } else {
        host.innerHTML = `
          <div class="agent-notice">
            <div><strong>Analyse en cours</strong><br/><span class="agent-muted">L'agent n'a pas encore fini ce mail. Réessaie dans un instant.</span></div>
          </div>
          <button type="button" class="btn btn-secondary btn-block agent-btn" id="agent-reload">Actualiser</button>
        `;
        this.$('agent-reload')?.addEventListener('click', () => {
          this.tentatives = 0;
          host.innerHTML = '<div class="agent-loading"><div class="spinner"></div><span>Lecture…</span></div>';
          this.loadState();
        });
      }
      return;
    }

    if (r.kind === 'non_analyse') {
      host.innerHTML = `
        <div class="agent-notice agent-notice-muted">
          <div>${r.raison === 'ancien'
            ? 'Ce mail est trop ancien pour l\'agent : les autres onglets restent utilisables.'
            : 'L\'agent ne retrouve pas ce mail dans la boîte : les autres onglets restent utilisables.'}</div>
        </div>
      `;
      return;
    }

    const why = r.kind === 'error'
      ? `État de l'agent indisponible (${escapeHtml(r.message)}).`
      : 'L\'agent n\'a pas d\'état pour ce mail (boîte hors périmètre ou agent pas encore actif).';
    host.innerHTML = `
      <div class="agent-notice agent-notice-muted">
        <div>${why}<br/><span class="agent-muted">Les onglets Lier, Projet et IA restent disponibles.</span></div>
      </div>
    `;
  }

  private renderState(host: HTMLElement, s: EtatAvecBrouillon): void {
    const pile = (PILE_LABELS[s.pile as AgentPile] ? s.pile : 'pour_info') as AgentPile;
    const blanc = s.mode !== 'actif';
    const facts: string[] = [];
    if (s.urgence > 0) facts.push(`<li class="fact-urgence-${Math.min(3, s.urgence)}">${URGENCE_LABELS[Math.min(3, s.urgence)]}</li>`);
    if (s.besoinReponse) facts.push('<li>Réponse attendue de ta part</li>');
    for (const e of s.engagements || []) {
      const ech = e.echeance ? ` · échéance ${escapeHtml(formatDate(e.echeance))}` : '';
      facts.push(`<li><strong>${engagementLabel(e.sens)}</strong> : ${escapeHtml(e.texte)}${ech}</li>`);
    }
    if (s.langue && s.langue !== 'FR') facts.push(`<li>Langue : ${s.langue === 'AUTRE' ? 'autre' : escapeHtml(s.langue)}</li>`);

    const action = s.actionProbable;
    host.innerHTML = `
      <div class="agent-head ${PILE_CLASSES[pile]}">
        <span class="agent-pile">${escapeHtml(PILE_LABELS[pile])}</span>
        ${s.categorie ? `<span class="agent-cat">· ${escapeHtml(humanize(s.categorie))}</span>` : ''}
      </div>
      ${pile === 'a_filtrer' ? this.filtreHtml(s) : ''}
      ${this.preferenceHtml(s, pile)}
      ${s.resume ? `<p class="agent-resume">« ${escapeHtml(s.resume)} »</p>` : ''}
      ${facts.length ? `<ul class="agent-facts">${facts.join('')}</ul>` : ''}

      <div class="agent-actions">
        ${action ? `
          <button type="button" class="btn btn-primary btn-block agent-btn" id="agent-main-action" ${blanc ? 'disabled' : ''}>
            ${escapeHtml(action.libelle)}
          </button>` : ''}
        <button type="button" class="btn btn-secondary btn-block agent-btn" id="agent-more-toggle" aria-expanded="false" ${blanc ? 'disabled' : ''}>
          Autre action ▾
        </button>
        <div id="agent-more" class="agent-more" hidden></div>
        ${blanc ? '<div class="agent-observation">Agent en observation : les actions seront proposées après le mode à blanc.</div>' : ''}
      </div>
    `;
    if (pile === 'a_filtrer') this.bindFiltre(s);
    this.bindPreference(s);
    if (blanc) return;
    this.$('agent-main-action')?.addEventListener('click', () => {
      if (action) this.runAction(action.type, action.cible?.id, action.cible?.table);
    });
    this.$('agent-more-toggle')?.addEventListener('click', () => this.toggleMore());
  }

  // ── Filtre des nouveaux expéditeurs (pile « À filtrer ») ──

  private filtreHtml(s: EtatAvecBrouillon): string {
    const email = s.from?.email || this.item?.fromEmail || '';
    return `
      <div class="agent-filtre" id="agent-filtre">
        <div class="agent-filtre-title">Nouvel expéditeur</div>
        <div class="agent-muted">Premier mail de <strong>${escapeHtml(s.from?.name || email)}</strong>${s.from?.name && email ? ` (${escapeHtml(email)})` : ''}. Le laisser arriver dans ta boîte ?</div>
        <div class="agent-filtre-btns">
          <button type="button" class="btn btn-primary agent-btn" data-decision="accepter" ${email ? '' : 'disabled'}>Accepter</button>
          <button type="button" class="btn btn-secondary agent-btn" data-decision="refuser" ${email ? '' : 'disabled'}>Refuser</button>
        </div>
        <div class="agent-muted">Une fois pour toutes : « Refuser » envoie ses prochains mails au bruit (annulable dans « Ce que l'agent a fait »).</div>
      </div>
    `;
  }

  private bindFiltre(s: EtatAvecBrouillon): void {
    const box = this.$('agent-filtre');
    const email = s.from?.email || this.item?.fromEmail || '';
    if (!box || !email) return;
    box.querySelectorAll<HTMLButtonElement>('button[data-decision]').forEach(btn => {
      btn.addEventListener('click', async () => {
        const decision = btn.dataset.decision === 'accepter' ? 'accepter' : 'refuser';
        box.querySelectorAll<HTMLButtonElement>('button').forEach(b => { b.disabled = true; });
        try {
          await decideExpediteur(email, decision);
          if (this.destroyed) return;
          box.innerHTML = decision === 'accepter'
            ? `<div class="agent-filtre-title">Expéditeur accepté</div><div class="agent-muted">Les prochains mails de ${escapeHtml(email)} arriveront normalement.</div>`
            : `<div class="agent-filtre-title">Expéditeur refusé</div><div class="agent-muted">Les prochains mails de ${escapeHtml(email)} iront au bruit.</div>`;
          showToast(decision === 'accepter' ? 'Expéditeur accepté' : 'Expéditeur refusé', 'success');
          refreshJourneeBanner(document.getElementById('journee-host'), { onInfo: showToast }).catch(() => { /* bandeau facultatif */ });
        } catch (e) {
          box.querySelectorAll<HTMLButtonElement>('button').forEach(b => { b.disabled = false; });
          showToast(`${humanError(e)}`, 'error');
        }
      });
    });
  }

  // ── Préférence de tri (09/10/2026) : « Toujours me montrer ce type de mail » / « C'est du bruit » ──

  private preferenceHtml(s: EtatAvecBrouillon, pile: AgentPile): string {
    const email = s.from?.email || this.item?.fromEmail || '';
    const auto = /^(newsletter|notification_systeme|expediteur_refuse)$/.test(String(s.categorie || '')) || /(no-?reply|notifications?|notice|alerts?|mailer|newsletters?|automated|billing|invoice|receipts?)/i.test(email);
    if (!email || (pile !== 'bruit' && !auto && !s.preference)) return '';
    const etat = s.preference ? escapeHtml(s.preference.raison) : pile === 'bruit' ? 'Classé en bruit.' : 'Envoi automatisé.';
    return `
      <div class="agent-filtre" id="agent-preference">
        <div class="agent-muted">${etat}</div>
        <div class="agent-filtre-btns">
          <button type="button" class="btn btn-secondary agent-btn" data-pref="montrer">Toujours me montrer ce type de mail</button>
          <button type="button" class="btn btn-secondary agent-btn" data-pref="bruit">C'est du bruit</button>
        </div>
      </div>
    `;
  }

  private bindPreference(s: EtatAvecBrouillon): void {
    const box = this.$('agent-preference');
    const email = s.from?.email || this.item?.fromEmail || '';
    if (!box || !email) return;
    box.querySelectorAll<HTMLButtonElement>('button[data-pref]').forEach(btn => {
      btn.addEventListener('click', async () => {
        const effet = btn.dataset.pref === 'bruit' ? 'bruit' : 'montrer';
        box.querySelectorAll<HTMLButtonElement>('button').forEach(b => { b.disabled = true; });
        try {
          const r = await definirPreferenceTri(email, effet);
          if (this.destroyed) return;
          const cible = ciblePreference(r);
          box.innerHTML = `<div class="agent-muted">${effet === 'montrer' ? 'Toujours montré' : 'Bruit désormais'} : ${escapeHtml(cible)}${r.reclasses ? ` (${r.reclasses} mail${r.reclasses > 1 ? 's' : ''} récent${r.reclasses > 1 ? 's' : ''} reclassé${r.reclasses > 1 ? 's' : ''})` : ''}.</div>`;
          showToast(effet === 'montrer' ? `Toujours montré : ${cible}` : `Bruit désormais : ${cible}`, 'success');
          refreshJourneeBanner(document.getElementById('journee-host'), { onInfo: showToast }).catch(() => { /* bandeau facultatif */ });
        } catch (e) {
          box.querySelectorAll<HTMLButtonElement>('button').forEach(b => { b.disabled = false; });
          showToast(humanError(e), 'error');
        }
      });
    });
  }

  // ── Journal « Ce que l'agent a fait » (lu au premier dépliage) ──

  private toggleJournal(): void {
    const host = this.$('agent-journal');
    const toggle = this.$('agent-journal-toggle');
    if (!host || !toggle) return;
    host.hidden = !host.hidden;
    toggle.setAttribute('aria-expanded', String(!host.hidden));
    toggle.textContent = `${host.hidden ? '▸' : '▾'} Ce que l'agent a fait`;
    if (!host.hidden && !host.dataset.loaded) {
      host.dataset.loaded = '1';
      renderJournal(host, { limite: 30, onInfo: showToast });
    }
  }

  // ── Actions (mode actif uniquement) ──

  /** Ouvre / ferme le menu « Autre action » ; les suggestions (sans IA) sont chargées au 1er clic. */
  private async toggleMore(): Promise<void> {
    const menu = this.$('agent-more');
    const toggle = this.$('agent-more-toggle');
    if (!menu || !toggle) return;
    if (!menu.hidden) {
      menu.hidden = true;
      toggle.setAttribute('aria-expanded', 'false');
      return;
    }
    menu.hidden = false;
    toggle.setAttribute('aria-expanded', 'true');
    if (this.suggestions === null) {
      menu.innerHTML = '<div class="agent-loading"><div class="spinner"></div><span>Actions…</span></div>';
      try {
        this.suggestions = await fetchEmailSuggestions(this.emailPayload());
      } catch (e) {
        console.warn('[AgentPanel] suggestions indisponibles :', e);
        this.suggestions = [];
      }
      if (this.destroyed) return;
    }
    const local: Array<{ id: string; label: string; icon: string }> = [
      { id: 'nav:link', label: 'Lier à un projet', icon: 'link' },
      { id: 'nav:info', label: 'Voir le projet', icon: 'folder' },
      ...(this.mobile ? [] : [{ id: 'nav:create', label: 'Créer un projet', icon: 'plus' }]),
      { id: 'atlas', label: 'Ouvrir dans ATLAS', icon: 'external' },
    ];
    const sugg = (this.suggestions || []).sort((a, b) => a.priority - b.priority);
    menu.innerHTML = `
      ${sugg.map((s, i) => `<button type="button" class="agent-menu-item" data-sugg="${i}" title="${escapeHtml(s.tooltip || '')}">${escapeHtml(s.label)}</button>`).join('')}
      ${local.map(l => `<button type="button" class="agent-menu-item" data-local="${l.id}">${icon(l.icon, 14)}${escapeHtml(l.label)}</button>`).join('')}
    `;
    menu.querySelectorAll<HTMLButtonElement>('button[data-sugg]').forEach(btn => {
      btn.addEventListener('click', () => {
        const s = sugg[Number(btn.dataset.sugg)];
        if (s) this.runAction(s.type, typeof s.payload?.projetId === 'string' ? s.payload.projetId : undefined);
      });
    });
    menu.querySelectorAll<HTMLButtonElement>('button[data-local]').forEach(btn => {
      btn.addEventListener('click', () => {
        const id = btn.dataset.local!;
        if (id.startsWith('nav:')) this.navigate(id.slice(4));
        else this.openInAtlas();
      });
    });
  }

  /**
   * Exécute une action proposée. Les parcours déjà présents dans le complément (lier, créer, voir
   * un projet, marquer traité) s'ouvrent dans leur onglet ; les autres ouvrent ATLAS sur le mail
   * avec l'action demandée (comme « ATLAS Assistant »). Rien n'est exécuté sans clic.
   */
  private runAction(type: string, cible?: string, table?: string): void {
    const t = (type || '').toLowerCase();
    if (/(view-projet|voir.?projet)/.test(t) || (cible && /projet/i.test(table || '') && !/(link|lier|cr[ée]er|create)/.test(t))) {
      if (cible) openExternal(`${ATLAS_BASE}/projet/${encodeURIComponent(cible)}`);
      else this.navigate('info');
      return;
    }
    if (/(link|lier)/.test(t)) { this.navigate('link'); return; }
    if (!this.mobile && /(create-projet|creer.?projet|cr[ée]er.?projet)/.test(t)) { this.navigate('create'); return; }
    if (/(mark-handled|traite)/.test(t)) { this.navigate('ia'); return; }
    if (/prospection/.test(t)) { openExternal(`${ATLAS_BASE}/prospection`); return; }
    this.openInAtlas(type, cible);
  }

  private openInAtlas(action?: string, cible?: string): void {
    const id = this.item?.messageId || this.item?.itemId || '';
    let url = `${ATLAS_BASE}/?email=${encodeURIComponent(id)}`;
    if (action) url += `&action=${encodeURIComponent(action)}`;
    if (cible) url += `&cible=${encodeURIComponent(cible)}`;
    openExternal(url);
  }

  private emailPayload() {
    const it = this.item!;
    let emailId = this.state?.graphId || '';
    if (!emailId && it.itemId) {
      try {
        emailId = Office.context.mailbox.convertToRestId(it.itemId, Office.MailboxEnums.RestVersion.v2_0);
      } catch { emailId = it.itemId; }
    }
    return {
      emailId,
      conversationId: it.conversationId || undefined,
      subject: it.subject,
      fromEmail: it.fromEmail,
      fromName: it.fromName,
      receivedAt: it.receivedAt,
      bodyPreview: this.state?.resume || '',
    };
  }

  /**
   * Mail d'une boîte partagée (good@, phase 5) : les actions métier visent good@ ; les rappels
   * personnels (plus tard / relancer) ne s'y appliquent pas : masqués.
   */
  private adapterBoitePartagee(s: EtatAvecBrouillon): void {
    const it = this.item;
    if (!it || !s.mailbox || s.mailbox.toLowerCase() === (it.mailbox || '').toLowerCase()) return;
    const actions = this.$('agent-actions-metier');
    if (actions) renderActionsMetier(actions, { messageId: it.messageId, mailbox: s.mailbox, onInfo: showToast, delegue: this.delegue });
    const rappels = this.$('agent-rappels');
    if (rappels) { rappels.hidden = true; rappels.innerHTML = ''; }
  }

  // ── Outils à la demande (actions métier, rappels, résumé du fil, question) ──

  private renderOutils(): void {
    const it = this.item;
    if (!it) return;
    const onInfo = showToast;
    const onChange = () => {
      refreshJourneeBanner(document.getElementById('journee-host'), { onInfo: showToast }).catch(() => { /* bandeau facultatif */ });
    };
    // Fiche mémoire du contact : demandée tout de suite, en parallèle de la vue du mail (lot 1 « vitesse »).
    const ficheContact = this.$('agent-fiche-contact');
    const exp = (it.fromEmail || '').toLowerCase();
    if (ficheContact && exp && !exp.endsWith('@vibes.lu')) { ficheContact.dataset.email = exp; void renderFicheContact(ficheContact, { email: exp }); }
    const equipe = this.$('agent-equipe');
    // Phase 5 : bloc masqué si le mail n'est pas dans good@ ou si la personne n'en est pas membre.
    if (equipe) renderEquipe(equipe, { messageId: it.messageId, onInfo, onChange });
    const actions = this.$('agent-actions-metier');
    if (actions) renderActionsMetier(actions, { messageId: it.messageId, mailbox: it.mailbox, onInfo, delegue: this.delegue });
    // 07/10/2026 (parité Inbox ATLAS) : « Offre fournisseur reçue » et « Classer dans Outlook » (arbre complet, nouveau dossier).
    const offre = this.$('agent-offre');
    if (offre && !this.mobile) renderOffreRecue(offre, { messageId: it.messageId, mailbox: it.mailbox, onInfo, delegue: this.delegue, repondre: (html: string) => this.repondreHtml(html) });
    // « Classer ce mail » : dessinée avec la vue du mail (projet et mail répondu déjà connus), cf. chargerClasser.
    // 07/10/2026 (fin de la parité Inbox ATLAS) : reclasser, pièces → projet, RDV, tiers, prospection.
    const parite = this.$('agent-parite');
    if (parite) {
      const [prenomExp, ...nomExp] = String(it.fromName || '').trim().split(/\s+/);
      renderPlusActions(parite, {
        messageId: it.messageId, mailbox: it.mailbox, onInfo, delegue: this.delegue,
        ...(this.mobile ? {} : { repondre: (html: string) => this.repondreHtml(html) }),
        prefillTiers: { email: '', contactPrenom: prenomExp || '', contactNom: nomExp.join(' '), contactEmail: it.fromEmail || '' },
      });
    }
    const rappels = this.$('agent-rappels');
    // Mail envoyé par la personne (Éléments envoyés, ou lu dans un fil) : relance ; sinon plus tard.
    const envoye = !!it.fromEmail && !!it.mailbox && it.fromEmail.toLowerCase() === it.mailbox.toLowerCase();
    if (rappels) renderRappels(rappels, { messageId: it.messageId, envoye, onInfo, onChange });
    const fil = this.$('agent-resume-fil-wrap');
    if (fil) renderResumeFil(fil, { conversationId: it.conversationId, mailbox: it.mailbox });
    const question = this.$('agent-question');
    if (question) renderQuestion(question, { mailbox: it.mailbox, onInfo });
    // Phase 6 : boîte non précisée → le worker cherche dans les boîtes permises (la sienne, good@).
    const traduction = this.$('agent-traduction');
    if (traduction) renderTraduction(traduction, { messageId: it.messageId, onInfo });
    const modeles = this.$('agent-modeles');
    if (modeles) renderModeles(modeles, { messageId: it.messageId, onInfo, ...(this.mobile ? {} : { repondre: (texte: string) => this.repondreAvec(texte) }) });
    const pj = this.$('agent-pj-lourdes');
    // Dépôt sur le cloud (transit 30 jours) : le worker lit les pièces par Graph (identifiant REST du mail).
    if (pj) {
      void renderPiecesLourdes(pj, {
        pieces: this.piecesJointes(), onInfo,
        depot: {
          source: async () => ({ messageId: this.emailPayload().emailId, ...(it.mailbox ? { mailbox: it.mailbox } : {}) }),
          idRest: (id: string) => convertToRestId(id),
        },
      });
    }
  }

  /**
   * Carte « Classer ce mail » avec ce que la vue du mail sait déjà (projet, mail répondu) ; sinon la
   * carte le demande elle-même. Traités à ranger (08/10/2026) : mail déjà répondu → la carte passe en
   * tête, suggestion mise en avant.
   */
  private chargerClasser(o: { projetId?: string; projetLibelle?: string; projetPropose?: boolean; sansProjet?: boolean; repondu?: TraiteARanger; traiteAFaire?: boolean; dossier?: VueMail['dossier'] } = { traiteAFaire: true }): void {
    const it = this.item;
    const classer = this.$('agent-classer');
    if (!it || !classer) return;
    const base = { messageId: it.messageId, mailbox: it.mailbox, conversationId: it.conversationId, onInfo: showToast, delegue: this.delegue };
    const { traiteAFaire, repondu, ...projet } = o;
    const enTete = () => { const etat = this.$('agent-state'); if (etat?.parentElement) etat.insertAdjacentElement('afterend', classer); };
    if (repondu) {
      enTete();
      renderClasserOutlook(classer, { ...base, repondu, ...(repondu.destination.projetId ? { projetId: repondu.destination.projetId, projetLibelle: repondu.destination.libelle } : projet) });
      return;
    }
    renderClasserOutlook(classer, { ...base, ...projet });
    if (!traiteAFaire) return;
    void fetchTraites(it.messageId).then(r => {
      const t = r.traites?.[0];
      if (!t || t.statut !== 'a_ranger' || this.destroyed || this.item !== it || !classer.isConnected) return;
      enTete();
      renderClasserOutlook(classer, { ...base, repondu: t, ...(t.destination.projetId ? { projetId: t.destination.projetId, projetLibelle: t.destination.libelle } : {}) });
    }).catch(() => { /* suggestion facultative */ });
  }

  /**
   * Repli « jeton de la personne » (comme l'Inbox ATLAS) quand l'agent ne peut pas écrire dans la boîte
   * (double verrou fermé) : SA boîte seulement (jamais une boîte partagée), mail ouvert dans Outlook.
   */
  private delegue = async (): Promise<{ token: string; restId: string } | null> => {
    try {
      const it = this.item;
      if (!it?.itemId) return null;
      const moi = String(Office.context?.mailbox?.userProfile?.emailAddress || '').toLowerCase();
      if (it.mailbox && moi && it.mailbox.toLowerCase() !== moi) return null;
      return { token: await getGraphToken(), restId: convertToRestId(it.itemId) };
    } catch { return null; }
  };

  /** Réponse préremplie (HTML déjà prêt, ex. remerciement ARGO) : rien n'est envoyé. */
  private repondreHtml(html: string): void {
    try { (Office.context.mailbox.item as any).displayReplyForm({ htmlBody: html }); }
    catch (e) { console.warn('[AgentPanel] réponse préremplie impossible :', e); showToast('Réponse préremplie impossible ici : copie le texte.', 'error'); }
  }

  /** Pièces jointes du mail lu (item.attachments, Mailbox 1.0, mobile compris). */
  private piecesJointes(): Array<{ id?: string; nom: string; octets: number; isInline?: boolean }> {
    try {
      const atts = (Office.context.mailbox?.item as any)?.attachments;
      if (!Array.isArray(atts)) return [];
      return atts.map((a: any) => ({ id: a?.id ? String(a.id) : undefined, nom: String(a?.name || ''), octets: Number(a?.size) || 0, isInline: !!a?.isInline }));
    } catch { return []; }
  }

  /**
   * « Répondre avec ce texte » (modèle rempli) : formulaire de réponse d'Outlook prérempli
   * (displayReplyForm, Mailbox 1.0 ; bureau et web, jamais sur mobile). Rien n'est envoyé.
   */
  private repondreAvec(texte: string): void {
    const html = texte.split(/\n{2,}/).map(p => `<p>${escapeHtml(p).replace(/\n/g, '<br/>')}</p>`).join('');
    try {
      (Office.context.mailbox.item as any).displayReplyForm({ htmlBody: html });
    } catch (e) {
      console.warn('[AgentPanel] réponse préremplie impossible :', e);
      showToast('Réponse préremplie impossible ici : copie le texte.', 'error');
    }
  }

  // ── Détections (rebond, absence, candidature) ──

  private renderDetectionsDuMail(s: EtatAvecBrouillon): void {
    const host = this.$('agent-detections');
    if (!host) return;
    if (s.detections) { renderDetections(host, s.detections); return; }
    // Mail lu avant le branchement des détections : relecture côté worker, seulement pour un mail automatique.
    if (s.regle === 'rebond' || s.regle === 'reponse-automatique' || s.categorie === 'rebond' || s.categorie === 'absence') {
      fetchDetections(s.messageId, s.mailbox)
        .then(d => { if (!this.destroyed && host.isConnected) renderDetections(host, d); })
        .catch(() => { /* encart facultatif */ });
    }
  }

  // ── Réunions liées (comptes rendus ATLAS du même tiers / projet, lecture seule) ──

  private renderReunions(s: EtatAvecBrouillon): void {
    const host = this.$('agent-reunions');
    if (!host) return;
    const reunions = (s.reunionsLiees || []).filter(r => r && r.id);
    if (!reunions.length) { host.hidden = true; host.innerHTML = ''; return; }
    host.hidden = false;
    host.innerHTML = `
      <div class="agent-section-title">Réunions liées</div>
      <ul class="agent-reunions">
        ${reunions.slice(0, 5).map((r, i) => `
          <li><button type="button" class="agent-link" data-reunion="${i}">${escapeHtml(r.titre || 'Compte rendu')}</button>
          ${r.date ? `<span class="agent-muted"> · ${escapeHtml(formatDate(r.date))}</span>` : ''}</li>`).join('')}
      </ul>`;
    host.querySelectorAll<HTMLButtonElement>('button[data-reunion]').forEach(btn => {
      btn.addEventListener('click', () => {
        const r = reunions[Number(btn.dataset.reunion)];
        if (r) openExternal(lienReunion(r.id));
      });
    });
  }

  // ── Brouillon prêt (masqué tant qu'il n'y en a pas) ──

  private renderBrouillon(s: EtatAvecBrouillon): void {
    const host = this.$('agent-brouillon');
    if (!host) return;
    const b = s.brouillon;
    if (!b || (!b.graphId && !b.resumeIntention)) { host.hidden = true; host.innerHTML = ''; return; }
    const depose = !!b.cree && !!b.graphId;
    // displayMessageForm (Mailbox 1.0) + convertToEwsId (Mailbox 1.3) : bureau et web seulement ;
    // sur mobile, displayMessageForm n'existe pas : le brouillon est dans le fil / dossier Brouillons.
    const peutOuvrir = depose && !this.mobile && supportsMailbox('1.3');
    const lienWeb = depose && !this.mobile
      ? `https://outlook.office.com/mail/drafts/id/${encodeURIComponent(b.graphId)}`
      : '';
    const alertes = (b.alertes || []).filter(a => a && a.message);
    host.hidden = false;
    host.innerHTML = `
      <div class="agent-section-title">${depose ? 'Brouillon prêt dans le fil' : 'Brouillon (simulé)'}</div>
      ${b.resumeIntention ? `<p class="agent-resume">« ${escapeHtml(b.resumeIntention)} »</p>` : ''}
      ${b.motif && MOTIF_BROUILLON[b.motif] ? `<div class="agent-muted">${escapeHtml(MOTIF_BROUILLON[b.motif])}${b.prepareLe ? ` · préparé ${escapeHtml(formatDate(b.prepareLe))}` : ''}</div>` : ''}
      ${alertes.length ? `
        <ul class="agent-avertissements" aria-label="Contrôle avant envoi">
          ${alertes.map(a => `<li class="${a.gravite === 'bloquant' ? 'is-bloquant' : ''}">${a.gravite === 'bloquant' ? '<strong>À corriger :</strong> ' : ''}${escapeHtml(a.message)}</li>`).join('')}
        </ul>` : ''}
      ${!depose
        ? (b.simuleMotif === 'autonomie'
          ? '<p class="agent-muted">Préparé sans être déposé : tes réglages d\'autonomie (ATLAS › Mon agent) demandent de garder les brouillons ici.</p>'
          : '<p class="agent-muted">Agent en observation : rien n\'a été déposé dans ta boîte.</p>')
        : peutOuvrir
          ? '<button type="button" class="btn btn-secondary btn-block agent-btn" id="agent-open-draft">Ouvrir le brouillon</button>'
          : lienWeb
            ? '<button type="button" class="btn btn-secondary btn-block agent-btn" id="agent-open-draft-web">Ouvrir le brouillon (Outlook sur le web)</button>'
            : '<p class="agent-muted">Le brouillon est dans ton dossier Brouillons, dans le fil de ce mail.</p>'}
      ${depose ? '<div class="agent-muted">Jamais envoyé par l\'agent, sans signature : relis-le et envoie-le toi-même.</div>' : ''}
    `;
    this.$('agent-open-draft')?.addEventListener('click', () => {
      try {
        const ewsId = Office.context.mailbox.convertToEwsId(b.graphId, Office.MailboxEnums.RestVersion.v2_0);
        Office.context.mailbox.displayMessageForm(ewsId);
      } catch (e) {
        console.warn('[AgentPanel] ouverture du brouillon impossible :', e);
        if (lienWeb) openExternal(lienWeb);
        else showToast('Brouillon impossible à ouvrir d\'ici : il est dans ton dossier Brouillons.', 'error');
      }
    });
    this.$('agent-open-draft-web')?.addEventListener('click', () => openExternal(lienWeb));
  }

  // ── Pourquoi cette proposition ? ──

  private renderWhy(s: EtatAvecBrouillon): void {
    const wrap = this.$('agent-why-wrap');
    const why = this.$('agent-why');
    if (!wrap || !why) return;
    const rows: Array<[string, string]> = [
      ['Source', s.source === 'regles' ? 'Règle (sans IA)' : 'Lecture IA unique par l\'agent serveur'],
    ];
    if (s.regle) rows.push(['Règle', s.regle]);
    rows.push(['Pile', PILE_LABELS[s.pile as AgentPile] || s.pile]);
    if (s.categorie) rows.push(['Catégorie', humanize(s.categorie)]);
    rows.push(['Mode', s.mode === 'actif' ? 'Actif (actions proposées)' : 'À blanc (observation, aucune action)']);
    if (s.traiteLe) rows.push(['Traité le', formatDate(s.traiteLe)]);
    if (typeof s.coutEur === 'number') rows.push(['Coût IA', `${s.coutEur.toFixed(4).replace('.', ',')} €`]);
    why.innerHTML = `<dl class="agent-why-list">${rows.map(([k, v]) => `<dt>${escapeHtml(k)}</dt><dd>${escapeHtml(v)}</dd>`).join('')}</dl>`;
    wrap.hidden = false;
  }

  // ── Correspondant (fiche calculée par le worker, sans IA) ──

  /** Fiche du correspondant (vue du mail ou atlas/correspondant) ; null = expéditeur sans adresse. */
  private afficherCorrespondant(f: FicheCorrespondant | null): void {
    const host = this.$('agent-correspondant');
    const it = this.item;
    if (!host || !it || this.destroyed) return;
    if (!f) { host.innerHTML = `<p class="agent-muted">${it.fromEmail ? 'Fiche ATLAS indisponible.' : 'Expéditeur inconnu.'}</p>`; return; }
    if (f.interne) {
      host.innerHTML = `<div class="agent-corr-name">${escapeHtml(it.fromName || f.email)}</div><div class="agent-muted">Collègue GOOD VIBES</div>`;
      return;
    }
    const details: string[] = [];
    const n = f.projetsEnCours.length;
    if (n) details.push(`${n} projet${n > 1 ? 's' : ''} en cours`);
    if (f.ton) details.push(f.ton === 'Amical' ? 'tutoiement' : 'vouvoiement');
    if (f.langue) details.push(`langue ${f.langue}`);
    const fiche = this.$('agent-fiche-contact');
    if (fiche && fiche.dataset.email !== f.email.toLowerCase()) { fiche.dataset.email = f.email.toLowerCase(); void renderFicheContact(fiche, { email: f.email }); }
    host.innerHTML = `
      <div class="agent-corr-name">${escapeHtml(f.nom || it.fromName || f.email)}${f.societe ? ` · ${escapeHtml(f.societe)}` : ''}${f.categorie ? ` <span class="agent-muted">(${escapeHtml(f.categorie.toLowerCase())})</span>` : ''}</div>
      ${f.fonction ? `<div class="agent-muted">${escapeHtml(f.fonction)}</div>` : ''}
      ${details.length ? `<div class="agent-muted">${escapeHtml(details.join(' · '))}</div>` : ''}
      ${f.connu ? '' : '<div class="agent-muted">Pas encore de fiche dans ATLAS.</div>'}
      <button type="button" class="btn btn-secondary btn-block agent-btn" id="agent-corr-open">Ouvrir dans ATLAS</button>
    `;
    this.$('agent-corr-open')?.addEventListener('click', () => {
      if (n === 1) openExternal(`${ATLAS_BASE}/projet/${encodeURIComponent(f.projetsEnCours[0].id)}`);
      else this.openInAtlas();
    });
  }

  /** Fiche seule (la vue ne l'avait pas prête) ; worker plus ancien : calcul d'avant dans le volet. */
  private async correspondantSeul(): Promise<void> {
    const it = this.item;
    if (!it?.fromEmail) { this.afficherCorrespondant(null); return; }
    try {
      const f = await fetchCorrespondant(it.fromEmail, it.fromName);
      if (!this.destroyed && this.item === it) this.afficherCorrespondant(f);
    } catch (e) {
      if (this.destroyed) return;
      if (e instanceof AtlasError && e.status === 404) { void this.loadCorrespondant(); return; }
      this.afficherCorrespondant(null);
    }
  }

  /** Ancien calcul dans le volet (tables entières) : seulement si le worker n'a pas encore la route correspondant. */
  private async loadCorrespondant(): Promise<void> {
    const host = () => this.$('agent-correspondant');
    const it = this.item;
    if (!it || !host()) return;
    const email = (it.fromEmail || '').toLowerCase();
    if (!email) { host()!.innerHTML = '<p class="agent-muted">Expéditeur inconnu.</p>'; return; }
    const domain = email.split('@')[1] || '';
    if (domain === 'vibes.lu') {
      host()!.innerHTML = `<div class="agent-corr-name">${escapeHtml(it.fromName || email)}</div><div class="agent-muted">Collègue GOOD VIBES</div>`;
      return;
    }
    try {
      const [contacts, tiers] = await Promise.all([getAllContacts().catch(() => []), getAllTiers().catch(() => [])]);
      const contact = contacts.find(c => (c.email || '').toLowerCase() === email);
      let tier = contact?.relationSociete
        ? tiers.find(t => (t.relation || '').toLowerCase() === contact.relationSociete.toLowerCase())
        : undefined;
      if (!tier && domain && !GENERIC_DOMAINS.has(domain)) {
        tier = tiers.find(t => (t.email || '').toLowerCase().endsWith(`@${domain}`));
      }
      const company = contact?.relationSociete || tier?.relation || '';
      const [projets, profile] = await Promise.all([
        company ? getProjetsByClient(company).catch(() => []) : Promise.resolve([]),
        fetchContactArgoProfile(email).catch(() => null),
      ]);
      if (this.destroyed || !host()) return;
      const enCours = projets.filter(p => !CLOSED_STATUS.test(p.statut || ''));
      const name = contact?.personneDeContact || it.fromName || email;
      const details: string[] = [];
      if (enCours.length) details.push(`${enCours.length} projet${enCours.length > 1 ? 's' : ''} en cours`);
      if (profile?.tonPrefere) details.push(profile.tonPrefere === 'Amical' ? 'tutoiement' : 'vouvoiement');
      if (profile?.languePreferee) details.push(`langue ${profile.languePreferee}`);
      const known = !!(contact || tier);
      const fiche = this.$('agent-fiche-contact');
      if (fiche) void renderFicheContact(fiche, { email });
      host()!.innerHTML = `
        <div class="agent-corr-name">${escapeHtml(name)}${company ? ` · ${escapeHtml(company)}` : ''}${tier?.categorie ? ` <span class="agent-muted">(${escapeHtml(tier.categorie.toLowerCase())})</span>` : ''}</div>
        ${contact?.fonction ? `<div class="agent-muted">${escapeHtml(contact.fonction)}</div>` : ''}
        ${details.length ? `<div class="agent-muted">${escapeHtml(details.join(' · '))}</div>` : ''}
        ${known ? '' : '<div class="agent-muted">Pas encore de fiche dans ATLAS.</div>'}
        <button type="button" class="btn btn-secondary btn-block agent-btn" id="agent-corr-open">Ouvrir dans ATLAS</button>
      `;
      this.$('agent-corr-open')?.addEventListener('click', () => {
        if (enCours.length === 1) openExternal(`${ATLAS_BASE}/projet/${encodeURIComponent(enCours[0].id)}`);
        else this.openInAtlas();
      });
    } catch (e) {
      if (this.destroyed || !host()) return;
      host()!.innerHTML = `<div class="agent-corr-name">${escapeHtml(it.fromName || email)}</div><div class="agent-muted">Fiche ATLAS indisponible.</div>`;
      const fiche = this.$('agent-fiche-contact');
      if (fiche) void renderFicheContact(fiche, { email });
      console.warn('[AgentPanel] correspondant :', e);
    }
  }
}
