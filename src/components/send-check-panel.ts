/**
 * send-check-panel.ts — Onglet « 🛡️ Vérifier » du panneau RÉDACTION (bureau et web uniquement) :
 * contrôles avant l'envoi de l'agent d'inbox (phase 2, `.claude/BACKLOG-AGENT-INBOX.md`).
 *
 * Lit le message en cours (objet, destinataires, corps HTML avec le fil cité, pièces jointes) et
 * applique les contrôles PURS de `utils/send-check.ts` : pièce jointe annoncée mais absente,
 * homonyme / client d'un autre projet que celui du fil, « répondre à tous » inutile, tutoiement
 * incohérent avec le fil. Les données ATLAS (contacts, clients, projet lié au fil, profil de
 * conversation) viennent des routes existantes `/api/plugin/atlas/*` (cache 5 min) ; si elles
 * sont indisponibles, seuls les contrôles locaux sont faits. Aucun appel IA, rien n'est envoyé ni
 * modifié.
 *
 * Contrôle manuel (« Vérifier avant d'envoyer ») : l'événement d'envoi `OnMessageSend` (Smart
 * Alerts, Mailbox 1.12) est prêt dans commands.ts mais pas activé dans le manifeste (il imposerait
 * Mailbox 1.12 à tout le complément, mobile compris) : voir le commentaire de manifest.xml.
 *
 * API Office : getAsync des destinataires / objet (Mailbox 1.1), body.getAsync (1.3),
 * getAttachmentsAsync (1.8, sous garde : sinon le contrôle de pièce jointe est sauté).
 *
 * « Relire avant envoi » (Assistant inbox, 07/10/2026) : en plus, tutoiement d'un client, langue
 * habituelle du correspondant (profil de conversation), date incohérente avec le projet du fil,
 * fournisseur visible d'un client, « Cordialement ». ENVOI AU BON MOMENT : destinataire hors de ses
 * heures (fuseau d'après le domaine) → conseil ; « Différer l'envoi » par `item.delayDeliveryTime`
 * (Mailbox 1.13 : nouvel Outlook pour Mac 1.1 à 1.14, web, Windows), sinon rappel de la commande
 * d'Outlook « Programmer l'envoi ». L'annulation d'envoi est un réglage d'Outlook lui-même.
 * Vérification Microsoft (doc du 16/09/2026) : `OnMessageSend` (Smart Alerts, 1.12) est pris en
 * charge par le nouvel Outlook pour Mac avec le manifeste XML, mais exige un déploiement par
 * l'administrateur et Mailbox 1.12 dans les Requirements du VersionOverrides (coupe le mobile) :
 * le bouton manuel reste la voie par défaut (décision à prendre par Charles).
 */

import { checkAvantEnvoi, hasBlocking, momentEnvoi, GENERIC_DOMAINS, type SendCheckProblem, type SendCheckInput, type SendCheckRecipient } from '../utils/send-check';
import { getAllContacts, getAllTiers, getAllProjets, getLinkedConversationIds, fetchContactArgoProfile } from '../api/airtable';
import { supportsMailbox } from '../api/platform';
import { escapeHtml } from './agent-lists';
import { humanError } from '../api/net';
import { icon } from '../ui/icons';

function getAsyncValue<T>(getter: ((cb: (r: Office.AsyncResult<T>) => void) => void) | undefined, fallback: T): Promise<T> {
  return new Promise((resolve) => {
    if (!getter) { resolve(fallback); return; }
    try {
      getter((r) => resolve(r.status === Office.AsyncResultStatus.Succeeded ? (r.value ?? fallback) : fallback));
    } catch { resolve(fallback); }
  });
}

const toRecipients = (list: Office.EmailAddressDetails[] | undefined): SendCheckRecipient[] =>
  (list || []).map(r => ({ email: String(r?.emailAddress || ''), name: String(r?.displayName || '') })).filter(r => r.email);

/** Message en cours de rédaction → entrée des contrôles (sans les données ATLAS). */
export async function readComposeItem(): Promise<SendCheckInput & { conversationId: string }> {
  const item = Office.context.mailbox?.item as any;
  if (!item) throw new Error('Aucun message en cours de rédaction.');
  const [subject, to, cc, bcc, html] = await Promise.all([
    getAsyncValue<string>(item.subject?.getAsync?.bind(item.subject), ''),
    getAsyncValue<Office.EmailAddressDetails[]>(item.to?.getAsync?.bind(item.to), []),
    getAsyncValue<Office.EmailAddressDetails[]>(item.cc?.getAsync?.bind(item.cc), []),
    getAsyncValue<Office.EmailAddressDetails[]>(item.bcc?.getAsync?.bind(item.bcc), []),
    getAsyncValue<string>(item.body?.getAsync ? (cb: any) => item.body.getAsync(Office.CoercionType.Html, cb) : undefined, ''),
  ]);
  // Pièces jointes : Mailbox 1.8 en rédaction ; sinon inconnu (contrôle sauté plutôt que faux positif).
  let piecesJointes: number | null = null;
  if (supportsMailbox('1.8') && item.getAttachmentsAsync) {
    const list = await getAsyncValue<Office.AttachmentDetailsCompose[] | null>(item.getAttachmentsAsync.bind(item), null);
    if (list) piecesJointes = list.filter(a => !a.isInline).length;
  }
  return {
    subject,
    html,
    to: toRecipients(to),
    cc: toRecipients(cc),
    bcc: toRecipients(bcc),
    moi: String(Office.context.mailbox?.userProfile?.emailAddress || ''),
    piecesJointes,
    conversationId: String(item.conversationId || ''),
  };
}

/** Complète l'entrée avec les données ATLAS (contacts, clients, projet du fil, tutoiement connu). */
async function enrichWithAtlas(input: SendCheckInput & { conversationId: string }): Promise<{ input: SendCheckInput; atlas: boolean }> {
  try {
    const [contacts, tiers, projets, linked] = await Promise.all([
      getAllContacts(), getAllTiers(), getAllProjets(), getLinkedConversationIds().catch(() => new Map()),
    ]);
    // Projet du fil : conversation liée dans ATLAS, sinon « #NNN » dans l'objet.
    let projet = input.conversationId ? projets.find(p => p.id === linked.get(input.conversationId)?.projetId) : undefined;
    if (!projet) {
      const no = (input.subject || '').match(/#\s*(\d{2,4})\b/)?.[1];
      if (no) projet = projets.find(p => String(p.noProjet).replace(/\D/g, '') === no);
    }
    const clients = Array.from(new Set(projets.map(p => p.client).filter(Boolean)));
    const domainesClients: Record<string, string> = {};
    const domainesFournisseurs: Record<string, string> = {};
    const estFournisseur = (cat: string) => /fournisseur|prestataire|supplier|traiteur|r[ée]gie/i.test(cat || '');
    for (const t of tiers) {
      const dom = (t.email || '').toLowerCase().split('@')[1];
      if (dom && !GENERIC_DOMAINS.has(dom) && !domainesClients[dom]) domainesClients[dom] = t.relation;
      if (dom && !GENERIC_DOMAINS.has(dom) && estFournisseur(t.categorie) && !domainesFournisseurs[dom]) domainesFournisseurs[dom] = t.relation;
    }
    const fournisseurs = tiers.filter(t => estFournisseur(t.categorie)).map(t => t.relation).filter(Boolean)
      // Un tiers à la fois client et fournisseur n'est pas « caché » au client.
      .filter(f => !clients.some(c => c.toLowerCase() === f.toLowerCase()));
    // Tutoiement connu : profil de conversation du destinataire principal (un seul).
    let tutoiementConnu = false;
    let langueHabituelle: string | null = null;
    const principal = [...(input.to || [])].find(r => !/@vibes\.lu$/i.test(r.email));
    if (principal) {
      const profile = await fetchContactArgoProfile(principal.email).catch(() => null);
      if ((input.to || []).length === 1) tutoiementConnu = profile?.tonPrefere === 'Amical';
      langueHabituelle = profile?.languePreferee || null;
    }
    return {
      atlas: true,
      input: {
        ...input,
        contacts: contacts.map(c => ({ email: c.email, nom: c.personneDeContact, societe: c.relationSociete })),
        projetDuFil: projet ? { nom: projet.noProjet ? `#${projet.noProjet} ${projet.denomination}` : projet.denomination, client: projet.client } : null,
        clients,
        domainesClients,
        tutoiementConnu,
        langueHabituelle,
        fournisseurs,
        domainesFournisseurs,
        datesProjet: projet?.dateDebut ? { debut: projet.dateDebut, fin: projet.dateFin || projet.dateDebut } : null,
      },
    };
  } catch (e) {
    console.warn('[SendCheck] données ATLAS indisponibles, contrôles locaux seulement :', e);
    return { input, atlas: false };
  }
}

export class SendCheckPanel {
  private container: HTMLElement;
  private busy = false;
  private destroyed = false;

  constructor(container: HTMLElement) {
    this.container = container;
    this.render();
    this.run();
  }

  destroy(): void {
    this.destroyed = true;
    this.container.innerHTML = '';
  }

  private render(): void {
    this.container.innerHTML = `
      <div class="panel-scroll">
        <div class="section-heading">Relire avant envoi</div>
        <p class="agent-muted" style="margin-bottom:8px;">Pièce jointe annoncée, mauvais destinataire, fournisseur visible du client, tutoiement d'un client, langue, dates du projet, « Cordialement », heure d'envoi. Sans IA, rien n'est envoyé.</p>
        <div id="send-check-result" class="send-check-result"></div>
        <div id="send-check-moment"></div>
        <button type="button" class="btn btn-primary btn-block agent-btn" id="send-check-run">Relire avant envoi</button>
      </div>
    `;
    this.container.querySelector('#send-check-run')?.addEventListener('click', () => this.run());
  }

  private async run(): Promise<void> {
    if (this.busy) return;
    const host = this.container.querySelector<HTMLElement>('#send-check-result');
    const btn = this.container.querySelector<HTMLButtonElement>('#send-check-run');
    if (!host) return;
    this.busy = true;
    if (btn) btn.disabled = true;
    host.innerHTML = '<div class="agent-loading"><div class="spinner"></div><span>Vérification…</span></div>';
    try {
      const local = await readComposeItem();
      const { input, atlas } = await enrichWithAtlas(local);
      if (this.destroyed) return;
      const problems = checkAvantEnvoi(input);
      this.renderResult(host, problems, { atlas, pjConnue: input.piecesJointes !== null && input.piecesJointes !== undefined });
      this.renderMoment(input);
    } catch (e) {
      if (this.destroyed) return;
      host.innerHTML = `<p class="agent-muted">Vérification impossible (${escapeHtml(humanError(e))}).</p>`;
    } finally {
      this.busy = false;
      if (btn) { btn.disabled = false; btn.textContent = 'Revérifier'; }
    }
  }

  /** Envoi au bon moment : heure du destinataire principal, « Différer l'envoi » (Mailbox 1.13) ou rappel. */
  private renderMoment(input: SendCheckInput): void {
    const host = this.container.querySelector<HTMLElement>('#send-check-moment');
    if (!host) return;
    const dest = [...(input.to || [])].find(r => !/@vibes\.lu$/i.test(r.email));
    const m = dest ? momentEnvoi(Date.now(), dest.email) : null;
    if (!m?.horsHeures || !m.conseil) { host.innerHTML = ''; return; }
    const item = Office.context.mailbox?.item as any;
    const differable = supportsMailbox('1.13') && typeof item?.delayDeliveryTime?.setAsync === 'function';
    host.innerHTML = `
      <div class="send-check-item">${icon('clock', 14)}<span>Il est ${escapeHtml(m.heureLocale)} chez ${escapeHtml(dest!.name || dest!.email)}${m.fuseau !== 'Europe/Luxembourg' ? ` (${escapeHtml(m.fuseau)})` : ''}, hors de ses heures de bureau. Mieux reçu le ${escapeHtml(m.libelle || '')}.</span></div>
      ${differable
        ? `<button type="button" class="btn btn-secondary btn-block" id="send-check-differer">${icon('clock', 14)}Différer l'envoi au ${escapeHtml(m.libelle || '')}</button>`
        : '<p class="agent-muted">Pour différer : flèche à côté d\'« Envoyer », puis « Programmer l\'envoi ».</p>'}
      <p class="agent-muted">Le délai pour annuler un envoi se règle dans les paramètres d'Outlook (rédaction), pas dans ATLAS.</p>`;
    host.querySelector('#send-check-differer')?.addEventListener('click', () => {
      try {
        item.delayDeliveryTime.setAsync(new Date(m.conseil!), (r: Office.AsyncResult<void>) => {
          const ok = r.status === Office.AsyncResultStatus.Succeeded;
          host.insertAdjacentHTML('beforeend', `<p class="agent-muted">${ok ? 'Envoi différé : clique « Envoyer », Outlook le garde jusqu\'à l\'heure prévue.' : `Report impossible (${escapeHtml(r.error?.message || 'Outlook')}) : utilise « Programmer l'envoi ».`}</p>`);
        });
      } catch (e) { host.insertAdjacentHTML('beforeend', `<p class="agent-muted">Report impossible (${escapeHtml(humanError(e))}).</p>`); }
    });
  }

  private renderResult(host: HTMLElement, problems: SendCheckProblem[], ctx: { atlas: boolean; pjConnue: boolean }): void {
    const notes: string[] = [];
    if (!ctx.atlas) notes.push('Fiches ATLAS indisponibles : homonymes et clients d\'autres projets non vérifiés.');
    if (!ctx.pjConnue) notes.push('Pièces jointes illisibles dans cette version d\'Outlook : contrôle sauté.');
    const notesHtml = notes.map(n => `<p class="agent-muted">${escapeHtml(n)}</p>`).join('');
    if (!problems.length) {
      host.innerHTML = `<div class="send-check-ok">${icon('check-circle', 16)}Rien à signaler : tu peux envoyer.</div>${notesHtml}`;
      return;
    }
    const bloquant = hasBlocking(problems);
    host.innerHTML = `
      <div class="send-check-head ${bloquant ? 'is-blocking' : ''}">${bloquant ? 'À corriger avant d\'envoyer' : 'À vérifier'} (${problems.length})</div>
      <ul class="send-check-list">
        ${problems.map(p => `<li class="send-check-item ${p.gravite === 'bloquant' ? 'is-blocking' : ''}">${icon(p.gravite === 'bloquant' ? 'x' : 'alert', 14)}<span>${escapeHtml(p.message)}</span></li>`).join('')}
      </ul>
      ${notesHtml}
    `;
  }
}
