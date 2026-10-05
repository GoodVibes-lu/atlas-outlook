/**
 * Project Info Panel · Onglet « Projet » : le projet ATLAS de ce mail.
 * Détection par #NNN dans le sujet, sinon par la conversation déjà liée.
 * Fiche projet, mails liés, profil ARGO de l'expéditeur, lien vers ATLAS.
 *
 * Refonte 05/10/2026 (retour « Load failed ») : chaque bloc a son propre état. La détection en
 * échec affiche un message lisible + « Réessayer » ; les compléments (nombre de mails, champs
 * détaillés, profil ARGO) sont lus en parallèle avec allSettled : l'un en échec n'empêche pas la
 * fiche de s'afficher. Jamais de texte technique à l'écran (ui/states.ts).
 */

import {
  getAllProjets,
  getLinkedConversationIds,
  countLinkedEmails,
  fetchContactArgoProfile,
  getProjetExtraFields,
} from '../api/airtable';
import type { Projet, ArgoProfile } from '../types';
import { escapeHtml } from '../utils/html';
import { icon } from '../ui/icons';
import { loadingHtml, emptyHtml, renderError } from '../ui/states';

/** Statut → ton sémantique (statut réel seulement, cf. DS : pas de couleur décorative). */
const STATUS_TONE: Record<string, 'neutral' | 'info' | 'success' | 'warning' | 'danger' | 'muted'> = {
  'Demande': 'warning',
  'Devis envoyé': 'info',
  'Confirmé': 'success',
  'En cours': 'info',
  'Terminé': 'muted',
  'Facturé': 'muted',
  'Annulé': 'danger',
  'Archivé': 'muted',
};

function statusBadge(statut: string): string {
  if (!statut) return '';
  const tone = STATUS_TONE[statut] || 'neutral';
  return `<span class="badge badge-${tone}">${escapeHtml(statut)}</span>`;
}

function fmtDate(iso: string | undefined): string {
  if (!iso) return '';
  const d = new Date(iso);
  return Number.isNaN(d.getTime()) ? '' : d.toLocaleDateString('fr-LU', { day: 'numeric', month: 'long', year: 'numeric' });
}

type Extra = { type?: string; budget?: string; descriptif?: string };

export class ProjectInfoPanel {
  private container: HTMLElement;
  private destroyed = false;
  private senderEmail = '';

  constructor(container: HTMLElement) {
    this.container = container;
    this.container.innerHTML = `
      <div class="panel-scroll">
        <div id="project-info-content"></div>
        <div id="argo-contact-card"></div>
      </div>
    `;
    void this.detectProject();
  }

  private get content(): HTMLElement | null { return this.container.querySelector('#project-info-content'); }

  private async detectProject(): Promise<void> {
    const content = this.content;
    if (!content) return;
    content.innerHTML = loadingHtml('Recherche du projet de ce mail…', 3);

    let item: any = null;
    try { item = Office.context.mailbox.item; } catch { item = null; }
    if (!item) {
      content.innerHTML = emptyHtml({ icon: 'mail', title: 'Aucun mail sélectionné', text: 'Ouvre un mail pour voir le projet ATLAS auquel il se rattache.' });
      return;
    }
    const subject: string = item.subject || '';
    this.senderEmail = item.from?.emailAddress || '';

    let projet: Projet | null = null;
    try {
      // 1) #NNN dans le sujet
      const match = subject.match(/#\s*(\d{2,4})/);
      if (match) {
        const projets = await getAllProjets();
        projet = projets.find(p => String(p.noProjet) === match[1]) || null;
      }
      // 2) Conversation déjà liée à un projet
      if (!projet) {
        const conversationId: string = item.conversationId || '';
        if (conversationId) {
          const convMap = await getLinkedConversationIds();
          const linked = convMap.get(conversationId);
          if (linked) {
            const projets = await getAllProjets();
            projet = projets.find(p => p.id === linked.projetId) || null;
          }
        }
      }
    } catch (err) {
      if (this.destroyed) return;
      renderError(content, err, () => { void this.detectProject(); }, { title: 'Projet non vérifié' });
      // Le profil de l'expéditeur ne dépend pas du projet : on le tente quand même.
      void this.loadArgoCard();
      return;
    }
    if (this.destroyed) return;

    if (!projet) {
      content.innerHTML = emptyHtml({
        icon: 'folder',
        title: 'Aucun projet rattaché',
        text: 'Ni numéro de projet (#123) dans le sujet, ni conversation déjà liée. Lie ce mail à un projet depuis l\'onglet « Lier ».',
        action: { id: 'goto-link', label: 'Lier ce mail', icon: 'link' },
      });
      content.querySelector('[data-action="goto-link"]')?.addEventListener('click', () => {
        document.querySelector<HTMLButtonElement>('.nav-tab[data-tab="link"]')?.click();
      });
      void this.loadArgoCard();
      return;
    }

    // Compléments en parallèle : un échec n'empêche pas la fiche.
    const [count, extra, profile] = await Promise.allSettled([
      countLinkedEmails(projet.id),
      getProjetExtraFields(projet.id) as Promise<Extra>,
      this.senderEmail ? fetchContactArgoProfile(this.senderEmail) : Promise.resolve(null),
    ]);
    if (this.destroyed) return;
    this.showProjectInfo(content, projet,
      extra.status === 'fulfilled' ? (extra.value || {}) : {},
      count.status === 'fulfilled' ? count.value : null,
      extra.status === 'rejected' || count.status === 'rejected');
    this.showArgoCard(profile.status === 'fulfilled' ? profile.value : null);
  }

  private showProjectInfo(container: HTMLElement, p: Projet, extra: Extra, emailCount: number | null, partial: boolean): void {
    const ref = p.refProjet ? p.refProjet : `#${p.noProjet}`;
    const descriptif = extra.descriptif
      ? (extra.descriptif.length > 220 ? `${extra.descriptif.slice(0, 220)}…` : extra.descriptif)
      : '';
    const rows: Array<[string, string]> = [];
    if (p.client) rows.push(['Client', p.client]);
    if (extra.type) rows.push(['Type', extra.type]);
    if (extra.budget) rows.push(['Budget', extra.budget]);
    if (p.enCharge) rows.push(['En charge', p.enCharge]);
    if (p.dateDebut) rows.push(['Début', fmtDate(p.dateDebut)]);
    if (p.dateFin) rows.push(['Fin', fmtDate(p.dateFin)]);

    container.innerHTML = `
      <article class="card project-card" aria-labelledby="pi-title">
        <header class="card-head">
          <span class="eyebrow">${icon('folder', 12)}Projet ${escapeHtml(ref)}</span>
          ${statusBadge(p.statut)}
        </header>
        <h2 class="card-title" id="pi-title">${escapeHtml(p.denomination || ref)}</h2>
        ${rows.length ? `<dl class="facts">${rows.map(([k, v]) => `<dt>${escapeHtml(k)}</dt><dd>${escapeHtml(v)}</dd>`).join('')}</dl>` : ''}
        ${descriptif ? `<p class="card-note">${escapeHtml(descriptif)}</p>` : ''}
        <footer class="card-foot">
          <span class="meta">${icon('mail', 14)}${emailCount === null ? 'Mails liés : indisponible' : `${emailCount} mail${emailCount !== 1 ? 's' : ''} lié${emailCount !== 1 ? 's' : ''}`}</span>
          <button id="btn-open-atlas" type="button" class="btn btn-secondary btn-sm">${icon('external', 14)}Ouvrir dans ATLAS</button>
        </footer>
        ${partial ? `<p class="help">Certains détails n'ont pas pu être lus. <button type="button" class="link-btn" id="pi-retry">Réessayer</button></p>` : ''}
      </article>
    `;

    container.querySelector('#btn-open-atlas')?.addEventListener('click', () => {
      const url = `atlas-app://open?entity=projet&id=${encodeURIComponent(p.id)}`;
      window.open(url, '_blank', 'noopener');
    });
    container.querySelector('#pi-retry')?.addEventListener('click', () => { void this.detectProject(); });
  }

  /** Profil ARGO de l'expéditeur (silencieux s'il est absent ou indisponible). */
  private async loadArgoCard(): Promise<void> {
    if (!this.senderEmail) return;
    try {
      const profile = await fetchContactArgoProfile(this.senderEmail);
      if (!this.destroyed) this.showArgoCard(profile);
    } catch { /* bloc facultatif */ }
  }

  private showArgoCard(a: ArgoProfile | null): void {
    const card = this.container.querySelector<HTMLElement>('#argo-contact-card');
    if (!card) return;
    if (!a) { card.innerHTML = ''; return; }
    const fullName = [a.prenom, a.nom].filter(Boolean).join(' ') || this.senderEmail;
    const ton = a.tonPrefere === 'Amical' ? 'Tutoiement' : a.tonPrefere === 'Professionnel' ? 'Vouvoiement' : '';
    card.innerHTML = `
      <section class="section" aria-labelledby="pi-argo">
        <h2 class="section-heading" id="pi-argo">Expéditeur, selon ARGO</h2>
        <div class="card person-card">
          <div class="person-avatar" aria-hidden="true">${icon('user', 16)}</div>
          <div class="person-body">
            <p class="person-name">${escapeHtml(fullName)}</p>
            <p class="meta">${escapeHtml(this.senderEmail)}</p>
            ${ton || a.languePreferee ? `<div class="chips">${ton ? `<span class="chip">${ton}</span>` : ''}${a.languePreferee ? `<span class="chip">${escapeHtml(a.languePreferee)}</span>` : ''}</div>` : ''}
            ${a.tutoiementAvec.length > 0 ? `<p class="help">Tutoie : ${a.tutoiementAvec.map(t => escapeHtml(t)).join(', ')}</p>` : ''}
          </div>
        </div>
      </section>
    `;
  }

  destroy(): void {
    this.destroyed = true;
    this.container.innerHTML = '';
  }
}
