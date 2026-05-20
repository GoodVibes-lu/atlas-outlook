/**
 * commands.ts — Actions ATLAS exécutées directement depuis la ribbon Outlook,
 * SANS ouvrir le task-pane.
 *
 * Boutons exposés via le manifest (ExecuteFunction) :
 *   • atlasDoneCommand       : marque le mail comme Traité ✓
 *   • atlasSnoozeCommand     : reporte le mail à demain 8h ⏰
 *   • atlasArchiveCommand    : archive le mail 📦 (+ move dans dossier habituel)
 *
 * Chaque commande :
 *   1. Trouve/crée le tag IA pour le mail courant (via Airtable EmailTags)
 *   2. Met à jour le statut
 *   3. Applique la catégorie Outlook colorée (visuel inbox)
 *   4. Pour Archive : déplace le mail dans le dossier appris (index sender)
 *   5. Affiche une notification Outlook (NotificationMessage InfoBar)
 *
 * AVANTAGE : actions accessibles en permanence depuis le bandeau Outlook,
 * sans dépendre du pin de task-pane (non supporté sur Outlook Mac sideload).
 */

import {
  getEmailTagByEmailId,
  getEmailTagByConversationId,
  markTagDone,
  snoozeTag,
  archiveTag,
} from './api/airtable';
import {
  convertToRestId,
  setMessageCategories,
  moveMessageToFolder,
  ATLAS_CATEGORIES,
} from './api/graph';
import { lookupSenderFolder, recordSenderFolder } from './api/sender-folder-index';

// Office.js doit être prêt avant que les commandes soient invoquées.
// On register les handlers globalement (window) — manifest les référence par nom.
Office.onReady(() => {
  // Les actions sont attachées via Office.actions.associate ci-dessous,
  // mais Outlook Mac fallback : aussi exposer en globals.
  (window as any).atlasDoneCommand = atlasDoneCommand;
  (window as any).atlasSnoozeCommand = atlasSnoozeCommand;
  (window as any).atlasArchiveCommand = atlasArchiveCommand;
  try {
    Office.actions.associate('atlasDoneCommand', atlasDoneCommand);
    Office.actions.associate('atlasSnoozeCommand', atlasSnoozeCommand);
    Office.actions.associate('atlasArchiveCommand', atlasArchiveCommand);
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

async function findTagForCurrentMail(restId: string, conversationId: string) {
  let tag = await getEmailTagByEmailId(restId);
  if (!tag && conversationId) tag = await getEmailTagByConversationId(conversationId);
  return tag;
}

function getCurrentMailContext() {
  const item = Office.context.mailbox?.item as any;
  if (!item) return null;
  const ewsId: string = item.itemId || '';
  if (!ewsId) return null;
  const restId = convertToRestId(ewsId);
  const conversationId: string = item.conversationId || '';
  const senderEmail: string = item.from?.emailAddress || '';
  return { restId, conversationId, senderEmail };
}

// ── Commandes exposées ─────────────────────────────────────────────────────

/**
 * ✓ Traité — Marque le mail comme Traité (Airtable) + catégorie verte.
 * Pas de déplacement de mail (juste le statut + le tag visuel).
 */
export async function atlasDoneCommand(event: Office.AddinCommands.Event): Promise<void> {
  try {
    const ctx = getCurrentMailContext();
    if (!ctx) { showInfoBar('Aucun mail sélectionné', true); event.completed(); return; }

    const tag = await findTagForCurrentMail(ctx.restId, ctx.conversationId);
    if (!tag) {
      showInfoBar('Mail pas encore taggé — ouvre ATLAS pour analyser', true);
      event.completed();
      return;
    }

    const ok = await markTagDone(tag.id);
    if (!ok) { showInfoBar('Échec marquage Traité', true); event.completed(); return; }

    // Catégorie ✓ verte
    try {
      await setMessageCategories(ctx.restId, [ATLAS_CATEGORIES.DONE.name]);
    } catch (e) {
      console.warn('[ATLAS commands] setCategories done failed:', e);
    }

    showInfoBar('✓ Traité — tag vert appliqué');
  } catch (e) {
    showInfoBar(`Erreur : ${(e as Error).message?.slice(0, 80)}`, true);
  } finally {
    event.completed();
  }
}

/**
 * ⏰ Reporter — Snooze à demain 8h (Airtable) + catégorie bleue visible
 * dans la liste inbox.
 */
export async function atlasSnoozeCommand(event: Office.AddinCommands.Event): Promise<void> {
  try {
    const ctx = getCurrentMailContext();
    if (!ctx) { showInfoBar('Aucun mail sélectionné', true); event.completed(); return; }

    const tag = await findTagForCurrentMail(ctx.restId, ctx.conversationId);
    if (!tag) {
      showInfoBar('Mail pas encore taggé — ouvre ATLAS pour analyser', true);
      event.completed();
      return;
    }

    const ok = await snoozeTag(tag.id);
    if (!ok) { showInfoBar('Échec snooze', true); event.completed(); return; }

    try {
      await setMessageCategories(ctx.restId, [ATLAS_CATEGORIES.SNOOZED.name]);
    } catch (e) {
      console.warn('[ATLAS commands] setCategories snooze failed:', e);
    }

    showInfoBar('⏰ Reporté à demain 8h — tag bleu visible dans l\'inbox');
  } catch (e) {
    showInfoBar(`Erreur : ${(e as Error).message?.slice(0, 80)}`, true);
  } finally {
    event.completed();
  }
}

/**
 * 📦 Archiver — Archive le tag (Airtable) + déplace le mail dans le
 * dossier habituel via l'index sender → dossier + catégorie violette.
 */
export async function atlasArchiveCommand(event: Office.AddinCommands.Event): Promise<void> {
  try {
    const ctx = getCurrentMailContext();
    if (!ctx) { showInfoBar('Aucun mail sélectionné', true); event.completed(); return; }

    const tag = await findTagForCurrentMail(ctx.restId, ctx.conversationId);
    if (!tag) {
      showInfoBar('Mail pas encore taggé — ouvre ATLAS pour analyser', true);
      event.completed();
      return;
    }

    const ok = await archiveTag(tag.id);
    if (!ok) { showInfoBar('Échec archive', true); event.completed(); return; }

    try {
      await setMessageCategories(ctx.restId, [ATLAS_CATEGORIES.ARCHIVED.name]);
    } catch (e) {
      console.warn('[ATLAS commands] setCategories archive failed:', e);
    }

    // Déplace le mail dans le dossier habituel (sender → dossier index)
    let movedTo = '';
    if (ctx.senderEmail) {
      const hit = lookupSenderFolder(ctx.senderEmail);
      if (hit) {
        try {
          const { getApiContext } = await import('./api/graph');
          const apiCtx = await getApiContext();
          await moveMessageToFolder(apiCtx.token, ctx.restId, hit.folderId);
          recordSenderFolder(ctx.senderEmail, hit.folderId, hit.folderPath, 1);
          movedTo = hit.folderPath;
        } catch (e) {
          console.warn('[ATLAS commands] move failed:', e);
        }
      }
    }

    showInfoBar(movedTo ? `📦 Archivé + rangé dans ${movedTo}` : '📦 Archivé (pas de dossier habituel)');
  } catch (e) {
    showInfoBar(`Erreur : ${(e as Error).message?.slice(0, 80)}`, true);
  } finally {
    event.completed();
  }
}
