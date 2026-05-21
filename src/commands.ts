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
  upsertEmailTag,
} from './api/airtable';
import { analyzeEmailWithClaude, hasAnthropicToken } from './api/claude';
import {
  convertToRestId,
  setMessageCategories,
  moveMessageToFolder,
  ATLAS_CATEGORIES,
  ATLAS_IA_CATEGORIES,
} from './api/graph';
import { lookupSenderFolder, recordSenderFolder } from './api/sender-folder-index';
import { initRoamingStorage } from './api/roaming-storage';

/**
 * Construit la liste des catégories à appliquer pour un état donné.
 * Inclut : type IA + urgence (si >= 4) + état (snoozed/done/archived).
 */
function buildCategoriesFor(tag: any, state: 'done' | 'snoozed' | 'archived'): string[] {
  const cats: string[] = [];
  if (tag?.category) {
    const ia = ATLAS_IA_CATEGORIES[tag.category];
    if (ia) cats.push(ia.name);
  }
  const u = tag?.urgencyScore || 0;
  if (u >= 5) cats.push(ATLAS_CATEGORIES.URGENCE_5.name);
  else if (u === 4) cats.push(ATLAS_CATEGORIES.URGENCE_4.name);
  if (state === 'snoozed') cats.push(ATLAS_CATEGORIES.SNOOZED.name);
  else if (state === 'done') cats.push(ATLAS_CATEGORIES.DONE.name);
  else if (state === 'archived') cats.push(ATLAS_CATEGORIES.ARCHIVED.name);
  return cats;
}

// Office.js doit être prêt avant que les commandes soient invoquées.
// On register les handlers globalement (window) — manifest les référence par nom.
Office.onReady(async () => {
  // Hydrate les clés depuis roamingSettings (les commandes ribbon ont
  // besoin du PAT Airtable + Anthropic même après cache clear)
  await initRoamingStorage();
  // Les actions sont attachées via Office.actions.associate ci-dessous,
  // mais Outlook Mac fallback : aussi exposer en globals.
  (window as any).atlasDoneCommand = atlasDoneCommand;
  (window as any).atlasSnoozeCommand = atlasSnoozeCommand;
  (window as any).atlasArchiveCommand = atlasArchiveCommand;
  (window as any).atlasReanalyzeCommand = atlasReanalyzeCommand;
  try {
    Office.actions.associate('atlasDoneCommand', atlasDoneCommand);
    Office.actions.associate('atlasSnoozeCommand', atlasSnoozeCommand);
    Office.actions.associate('atlasArchiveCommand', atlasArchiveCommand);
    Office.actions.associate('atlasReanalyzeCommand', atlasReanalyzeCommand);
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
      await setMessageCategories(ctx.restId, buildCategoriesFor(tag, 'done'));
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
      await setMessageCategories(ctx.restId, buildCategoriesFor(tag, 'snoozed'));
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
 * 🔄 Re-analyser — Force l'analyse Claude sur le mail courant.
 * Crée ou met à jour le tag IA. Applique ensuite la catégorie correspondante.
 */
export async function atlasReanalyzeCommand(event: Office.AddinCommands.Event): Promise<void> {
  try {
    if (!hasAnthropicToken()) {
      showInfoBar('Clé Anthropic non configurée (Settings ATLAS)', true);
      event.completed();
      return;
    }
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

    showInfoBar('Analyse Claude en cours…');

    const analysis = await analyzeEmailWithClaude({
      subject, from, toRecipients, ccRecipients, body, receivedAt, userEmail,
    });

    const existing = await findTagForCurrentMail(ctx.restId, ctx.conversationId);
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
    const newTag = { id: upserted.id, category: analysis.category, urgencyScore: analysis.urgencyScore };
    try {
      await setMessageCategories(ctx.restId, buildCategoriesFor(newTag, 'done' as any).filter((c) =>
        // Pour le re-analyze : on applique IA + urgence, PAS d'état
        !c.includes('Traité') && !c.includes('Reporté') && !c.includes('Archivé')
      ));
    } catch (e) {
      console.warn('[ATLAS commands] setCategories reanalyze failed:', e);
    }

    const iaCat = ATLAS_IA_CATEGORIES[analysis.category]?.name || analysis.category;
    showInfoBar(`✓ Analysé : ${iaCat} (urgence ${analysis.urgencyScore}/5)`);
  } catch (e) {
    showInfoBar(`Erreur re-analyse : ${(e as Error).message?.slice(0, 80)}`, true);
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
      await setMessageCategories(ctx.restId, buildCategoriesFor(tag, 'archived'));
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
