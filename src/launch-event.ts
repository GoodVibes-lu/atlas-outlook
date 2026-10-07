/**
 * launch-event.ts — Gestionnaire Smart Alerts (`OnMessageSend`, SendMode « SoftBlock ») pour le
 * RUNTIME JAVASCRIPT SEUL d'Outlook CLASSIQUE Windows (manifest.xml : `<Override type="javascript">`).
 *
 * Ce runtime n'a ni DOM, ni fenêtre, ni MSAL : seuls Office.js et le code pur sont disponibles. Les
 * contrôles sont donc LOCAUX (pièce jointe annoncée, destinataires, registre, « Cordialement »…),
 * les mêmes que le bouton manuel sans les données ATLAS. Le nouvel Outlook (Mac, Windows) et le
 * web chargent commands.html (runtime navigateur) : là, commands.ts enrichit avec les fiches ATLAS.
 *
 * Construit en un seul fichier IIFE `launch-event.js` par vite.config.ts (esbuild, sans import),
 * exigence d'Outlook classique. Jamais bloquant : erreur ou budget dépassé = envoi autorisé.
 */
import { gererOnMessageSend } from './utils/send-check-office';

function atlasOnMessageSend(event: any): Promise<void> {
  return gererOnMessageSend(event);
}

// Enregistrement au chargement (pas d'Office.onReady : le runtime JS seul appelle le gestionnaire
// par son nom dès que le fichier est évalué).
(globalThis as any).atlasOnMessageSend = atlasOnMessageSend;
try {
  Office.actions.associate('atlasOnMessageSend', atlasOnMessageSend);
} catch (e) {
  console.warn('[ATLAS launch-event] Office.actions.associate indisponible :', e);
}
