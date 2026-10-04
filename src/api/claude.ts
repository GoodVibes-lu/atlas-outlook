/**
 * claude.ts — Classification IA d'un mail depuis le complément Outlook.
 *
 * Sert au bouton « Re-analyser ce mail » : purge un tag erroné et relance l'analyse SANS attendre
 * le scan périodique. Résultat immédiat.
 *
 * Depuis le 03/10/2026, l'analyse est faite par le worker (`/api/plugin/atlas/ai/classify`) via
 * l'Agency Brain (journal des coûts, plafond, cache) : le complément ne détient plus de clé
 * Anthropic et n'appelle plus l'API en direct. Même prompt et mêmes garde-fous qu'avant.
 */

import { callAtlasWorker } from './worker';

export type ClaudeCategory =
  | 'demande_devis' | 'validation_client' | 'refus_client' | 'question_staff'
  | 'facture_fournisseur' | 'prospection_entrante' | 'prospection_sortante' | 'rdv_planning'
  | 'newsletter' | 'notification_systeme' | 'spam' | 'autre'
  | 'federation_association' | 'demande_interne_staff' | 'fournisseur';

export interface ClaudeAnalysis {
  category: ClaudeCategory;
  urgencyScore: number;
  summary: string;
  detectedLanguage: 'FR' | 'EN' | 'DE' | 'LU' | 'AUTRE';
}

/** L'IA est servie par le worker : toujours disponible côté complément (aucune clé à configurer). */
export function isAiAvailable(): boolean {
  return true;
}

export interface AnalyzeInput {
  subject: string;
  from: { name: string; email: string };
  toRecipients: Array<{ name?: string; email: string }>;
  ccRecipients: Array<{ name?: string; email: string }>;
  body: string;
  receivedAt: string;
  /** Conservé pour compat : le worker utilise l'adresse du jeton Microsoft. */
  userEmail: string;
  /** Analyse lancée automatiquement (pas par un clic) : tâche de fond, soumise au plafond IA du jour. */
  auto?: boolean;
}

/** Lance l'analyse IA d'un mail (catégorie, urgence, résumé, langue). */
export async function analyzeEmailWithClaude(input: AnalyzeInput): Promise<ClaudeAnalysis> {
  const r = await callAtlasWorker<{ analysis: ClaudeAnalysis }>('ai/classify', {
    subject: input.subject,
    from: input.from,
    toRecipients: input.toRecipients,
    ccRecipients: input.ccRecipients,
    // Même plafond qu'avant (4000 caractères) : inutile d'envoyer davantage.
    body: (input.body || '').slice(0, 4000),
    receivedAt: input.receivedAt,
    auto: !!input.auto,
  });
  if (!r?.analysis) throw new Error('Analyse IA indisponible.');
  return r.analysis;
}
