/**
 * commands.ts — Actions ATLAS exécutées directement depuis la ribbon Outlook,
 * SANS ouvrir le task-pane.
 *
 * Boutons exposés via le manifest (ExecuteFunction) :
 *   • atlasDoneCommand       : marque le mail comme Traité ✓
 *   • atlasSnoozeCommand     : reporte le mail à demain 8h ⏰
 *   • atlasArchiveCommand    : archive le mail 📦 (rangé dans « Archives » par le worker)
 *
 * Chaque commande passe par le worker (api/actions-mail.ts, lot 1 « vitesse » du 09/10/2026) :
 * fiche de tri ATLAS + action de l'agent dans la boîte de la personne (marquer lu, « plus tard »,
 * ranger dans Archives), puis une notification Outlook (NotificationMessage InfoBar). Plus aucun accès
 * direct à la boîte (catégories, déplacement) : Microsoft ne donne plus ce jeton au complément.
 *
 * AVANTAGE : actions accessibles en permanence depuis le bandeau Outlook,
 * sans dépendre du pin de task-pane (non supporté sur Outlook Mac sideload).
 *
 *   • atlasTableauCommand    : ancien bouton « Tableau de bord » (fenêtre de dialogue, retirée le
 *                              10/10/2026) : indique le volet ; le manifeste à jour ouvre le volet.
 *
 * Phase 2 de l'agent d'inbox : `atlasOnMessageSend`, gestionnaire de l'événement d'envoi
 * (Smart Alerts, `OnMessageSend`, mode « soft block », Mailbox 1.12), ACTIVÉ dans manifest.xml le
 * 08/10/2026 (décision de la direction). Ce fichier est chargé par commands.html (runtime
 * navigateur : nouvel Outlook pour Mac, web, nouvel Outlook Windows) : mêmes contrôles que l'onglet
 * « Relire avant envoi », données ATLAS comprises, dans un budget de 5 s ; réseau en panne ou budget
 * dépassé = contrôles locaux, puis envoi autorisé. Outlook classique Windows charge launch-event.ts.
 */

import { upsertEmailTag } from './api/airtable';
import { analyzeEmailWithClaude } from './api/claude';
import { convertToRestId, ATLAS_IA_CATEGORIES } from './api/graph';
import { archiverMail, ficheDeTri, marquerTraite, reporterDemain, type MailCourant } from './api/actions-mail';
import { initRoamingStorage } from './api/roaming-storage';
import { supportsMailbox } from './api/platform';
import { gererOnMessageSend } from './utils/send-check-office';
import { enrichWithAtlas } from './components/send-check-panel';
import { humanError } from './api/net';

Office.onReady(async () => {
  // Hydrate les réglages depuis roamingSettings et efface les anciens secrets (clé Anthropic,
  // jeton Airtable) : les commandes passent désormais par le worker.
  await initRoamingStorage();
  // Les actions sont attachées via Office.actions.associate ci-dessous,
  // mais Outlook Mac fallback : aussi exposer en globals.
  (window as any).atlasDoneCommand = atlasDoneCommand;
  (window as any).atlasSnoozeCommand = atlasSnoozeCommand;
  (window as any).atlasArchiveCommand = atlasArchiveCommand;
  (window as any).atlasReanalyzeCommand = atlasReanalyzeCommand;
  (window as any).atlasOnMessageSend = atlasOnMessageSend;
  (window as any).atlasTableauCommand = atlasTableauCommand;
  try {
    Office.actions.associate('atlasDoneCommand', atlasDoneCommand);
    Office.actions.associate('atlasSnoozeCommand', atlasSnoozeCommand);
    Office.actions.associate('atlasArchiveCommand', atlasArchiveCommand);
    Office.actions.associate('atlasReanalyzeCommand', atlasReanalyzeCommand);
    Office.actions.associate('atlasOnMessageSend', atlasOnMessageSend);
    Office.actions.associate('atlasTableauCommand', atlasTableauCommand);
  } catch (e) {
    console.warn('[ATLAS commands] Office.actions.associate not available:', e);
  }
});

// ── Helpers ────────────────────────────────────────────────────────────────

function showInfoBar(message: string, isError = false): void {
  try {
    const item = Office.context.mailbox?.item as any;
    if (!item || !item.notificationMessages) return;
    const key = 'atlas-cmd-' + Date.now();
    item.notificationMessages.addAsync(key, {
      type: isError
        ? Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
        : Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: message.slice(0, 150),
      icon: 'icon16',
      persistent: false,
    });
    // Auto-clear après 5s
    setTimeout(() => {
      try { item.notificationMessages.removeAsync(key); } catch { /* noop */ }
    }, 5000);
  } catch (e) {
    console.warn('[ATLAS commands] showInfoBar failed:', e);
  }
}

function getCurrentMailContext(): (MailCourant & { senderEmail: string }) | null {
  const item = Office.context.mailbox?.item as any;
  if (!item) return null;
  const ewsId: string = item.itemId || '';
  if (!ewsId) return null;
  return {
    restId: convertToRestId(ewsId),
    messageId: String(item.internetMessageId || ''),
    conversationId: item.conversationId || '',
    senderEmail: item.from?.emailAddress || '',
  };
}

/**
 * Boutons du ruban (lot 1 « vitesse », 09/10/2026) : tout passe par le worker (api/actions-mail.ts),
 * plus par le jeton de boîte du complément que Microsoft ne donne plus (catégories et déplacement
 * échouaient). Le rangement se fait dans la boîte de la personne, sur son clic (rangementClicAutorise).
 */
async function commandeMail(event: Office.AddinCommands.Event, faire: (m: MailCourant) => Promise<{ message: string }>): Promise<void> {
  try {
    const ctx = getCurrentMailContext();
    if (!ctx) { showInfoBar('Aucun mail sélectionné', true); return; }
    const r = await faire(ctx);
    showInfoBar(r.message);
  } catch (e) {
    showInfoBar(humanError(e), true);
  } finally {
    event.completed();
  }
}

/** Traité : fiche de tri « Traité » et mail marqué lu (worker). */
export async function atlasDoneCommand(event: Office.AddinCommands.Event): Promise<void> {
  await commandeMail(event, marquerTraite);
}

/** Reporter : « répondre plus tard » de l'agent, demain 8 h (jours ouvrés), et fiche « Reporté ». */
export async function atlasSnoozeCommand(event: Office.AddinCommands.Event): Promise<void> {
  await commandeMail(event, reporterDemain);
}

/**
 * 🔄 Re-analyser — Force l'analyse Claude sur le mail courant.
 * Crée ou met à jour le tag IA. Applique ensuite la catégorie correspondante.
 */
export async function atlasReanalyzeCommand(event: Office.AddinCommands.Event): Promise<void> {
  try {
    const item = Office.context.mailbox?.item as any;
    if (!item) { showInfoBar('Aucun mail sélectionné', true); event.completed(); return; }
    const ctx = getCurrentMailContext();
    if (!ctx) { showInfoBar('Mail non identifiable', true); event.completed(); return; }

    const userEmail = Office.context.mailbox?.userProfile?.emailAddress || '';
    const subject: string = item.subject || '';
    const from = {
      name: item.from?.displayName || '',
      email: item.from?.emailAddress || '',
    };
    const toRecipients = (item.to || []).map((r: any) => ({ name: r.displayName || '', email: r.emailAddress || '' }));
    const ccRecipients = (item.cc || []).map((r: any) => ({ name: r.displayName || '', email: r.emailAddress || '' }));
    const receivedAt = item.dateTimeCreated ? new Date(item.dateTimeCreated).toISOString() : new Date().toISOString();

    // Body via Office.js
    const body: string = await new Promise((resolve) => {
      try {
        item.body.getAsync(Office.CoercionType.Text, (res: any) => {
          resolve(res?.status === Office.AsyncResultStatus.Succeeded ? (res.value || '') : '');
        });
      } catch { resolve(''); }
    });

    showInfoBar('ARGO classe ce mail…');

    const analysis = await analyzeEmailWithClaude({
      subject, from, toRecipients, ccRecipients, body, receivedAt, userEmail,
    });

    const existing = await ficheDeTri(ctx);
    const upserted = await upsertEmailTag({
      oldTagId: existing?.id,
      emailId: ctx.restId,
      conversationId: ctx.conversationId,
      subject,
      fromEmail: from.email,
      fromName: from.name,
      receivedAt,
      category: analysis.category,
      urgencyScore: analysis.urgencyScore,
      summary: analysis.summary,
      detectedLanguage: analysis.detectedLanguage,
      userEmail,
    });

    // Applique les catégories (type IA + urgence si haute)
    void upserted; // catégorie Outlook posée par l'agent du worker (plus d'accès direct à la boîte ici)
    const iaCat = ATLAS_IA_CATEGORIES[analysis.category]?.name || analysis.category;
    showInfoBar(`Classé : ${iaCat} (urgence ${analysis.urgencyScore}/5)`);
  } catch (e) {
    showInfoBar(`${humanError(e)}`, true);
  } finally {
    event.completed();
  }
}

/** Archiver : rangé dans « Archives » par le worker (annulable dans le panneau) et fiche « Archivé ». */
export async function atlasArchiveCommand(event: Office.AddinCommands.Event): Promise<void> {
  await commandeMail(event, archiverMail);
}

/**
 * Tableau de bord (ancien manifeste) : depuis le 10/10/2026 il s'affiche dans le volet ATLAS, plus
 * dans une fenêtre de dialogue (elle restait au-dessus de toutes les applications sur Outlook Mac).
 * Le manifeste à jour ouvre directement le volet (ShowTaskpane, `?tab=tableau`) ; tant qu'il n'est
 * pas redéployé, ce bouton indique le chemin.
 */
export function atlasTableauCommand(event: Office.AddinCommands.Event): void {
  showInfoBar('Ouvre le volet ATLAS puis « Tableau de bord » en haut du volet.', false);
  try { event.completed(); } catch { /* déjà terminé */ }
}

// ── Smart Alerts : contrôles avant l'envoi (OnMessageSend, Mailbox 1.12) ──

/**
 * Gestionnaire de l'événement d'envoi (Smart Alerts, `SendMode="SoftBlock"`), runtime navigateur.
 * Mêmes contrôles que le bouton manuel (utils/send-check-office.ts), enrichis des fiches ATLAS quand
 * le worker répond à temps ; s'il y a quelque chose à signaler, l'envoi est retenu avec le message et
 * la personne peut corriger ou envoyer quand même. Erreur, API absente, réseau en panne ou budget
 * dépassé : l'envoi passe (jamais bloqué par ATLAS).
 */
async function atlasOnMessageSend(event: any): Promise<void> {
  await gererOnMessageSend(event, { enrichir: async (input) => (await enrichWithAtlas(input)).input });
}
