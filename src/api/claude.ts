/**
 * claude.ts — Appel Claude direct depuis l'addin Outlook pour la classification IA.
 *
 * Sert au bouton "Re-analyser ce mail" : permet de purger un tag erroné +
 * relancer l'analyse Claude SANS attendre que l'app desktop ATLAS scan
 * périodiquement. Résultat immédiat.
 *
 * La clé Anthropic est stockée en localStorage côté addin (clé séparée du
 * Airtable token). À configurer via Settings.
 */

const ANTHROPIC_API = 'https://api.anthropic.com/v1/messages';
const MODEL = 'claude-haiku-4-5';

const CATEGORIES = [
  'demande_devis', 'validation_client', 'refus_client', 'question_staff',
  'facture_fournisseur', 'prospection_entrante', 'rdv_planning',
  'newsletter', 'notification_systeme', 'spam', 'autre',
  'federation_association', 'demande_interne_staff', 'fournisseur',
] as const;
export type ClaudeCategory = typeof CATEGORIES[number];

export interface ClaudeAnalysis {
  category: ClaudeCategory;
  urgencyScore: number;
  summary: string;
  detectedLanguage: 'FR' | 'EN' | 'DE' | 'LU' | 'AUTRE';
}

export function getAnthropicToken(): string {
  return localStorage.getItem('atlas_addin_anthropic_key') || '';
}

export function hasAnthropicToken(): boolean {
  return !!getAnthropicToken();
}

export interface AnalyzeInput {
  subject: string;
  from: { name: string; email: string };
  toRecipients: Array<{ name?: string; email: string }>;
  ccRecipients: Array<{ name?: string; email: string }>;
  body: string;
  receivedAt: string;
  userEmail: string;
}

/**
 * Lance une analyse Claude pour un mail. Reflète la logique de
 * `src/services/inbound-scanner.service.ts` (prompt + garde-fous) côté addin
 * pour permettre la re-analyse en 1 clic depuis le task-pane.
 */
export async function analyzeEmailWithClaude(input: AnalyzeInput): Promise<ClaudeAnalysis> {
  const token = getAnthropicToken();
  if (!token) {
    throw new Error('Clé Anthropic non configurée dans Settings de l\'addin.');
  }

  // Pré-calcul du flag "thread interne" — utilisé pour bloquer demande_interne_staff
  // sur les mails avec destinataire externe (cas client + collègue en CC).
  const senderDom = (input.from.email.split('@')[1] || '').toLowerCase();
  const toExternal = input.toRecipients
    .map((r) => (r.email.split('@')[1] || '').toLowerCase())
    .filter((d) => d && d !== 'vibes.lu');
  const isInternalThread = senderDom === 'vibes.lu' && toExternal.length === 0;

  // Tronque le body cleaned. On garde l'historique cité car parfois utile,
  // mais on plafonne à 4000 chars pour éviter de cramer des tokens.
  const cleanedBody = input.body.slice(0, 4000);

  const toList = input.toRecipients.map((r) => `${r.name || ''} <${r.email}>`).join(', ') || '(aucun)';
  const ccList = input.ccRecipients.map((r) => `${r.name || ''} <${r.email}>`).join(', ') || '(aucun)';

  const userIsInTo = input.toRecipients.some((r) => r.email.toLowerCase() === input.userEmail.toLowerCase());
  const userIsInCc = input.ccRecipients.some((r) => r.email.toLowerCase() === input.userEmail.toLowerCase());
  const role = userIsInTo ? 'to (destinataire principal — action attendue)'
    : userIsInCc ? 'cc (pour info uniquement, action généralement PAS attendue)'
    : 'unknown (BCC ou alias)';

  const prompt = `Tu analyses un email reçu par Good Vibes, agence événementielle au Luxembourg.

CONTEXTE GOOD VIBES :
- Charles Maes = managing director (charles@vibes.lu).
- Équipe interne : @vibes.lu (good@, charles@, audric@, baptiste@, ellora@).
- Clients récurrents : Paperjam, Neimënster, Arendt & Medernach, Luxreal, LIST, CNAPA, etc.

RÈGLE CRITIQUE — ANALYSE DU MAIL COURANT :
- L'expéditeur de CE mail est : ${input.from.name} <${input.from.email}>
- Concentre-toi sur le contenu du mail COURANT, pas sur les messages cités dans la conversation (replies).
- NE PAS classer "demande_interne_staff" si l'expéditeur OU un destinataire principal est externe — c'est un échange client/fournisseur même si un collègue interne est cité dans l'historique.
- Thread interne (tous @vibes.lu) : ${isInternalThread ? 'OUI' : 'NON'}
- Si Thread interne = NON → catégorie NE PEUT PAS être "demande_interne_staff".

Email :
De : ${input.from.name} <${input.from.email}>
À : ${toList}
Cc : ${ccList}
Sujet : ${input.subject}
Date : ${input.receivedAt}
Rôle de Charles : ${role}

Corps du mail (nettoyé) :
"""
${cleanedBody}
"""

INSTRUCTIONS :
1. Tagger selon le contenu du MAIL COURANT (pas l'historique cité).
2. Si "Thread interne = NON" → catégorie ≠ demande_interne_staff.
3. Si Charles est en Cc, plafonne l'urgence à 2/5 sauf urgence explicite dans le corps.
4. Sois STRICT sur urgence 5 (vraies urgences avec deadline < 24h).
5. Résume en 1-2 phrases factuelles, en français. Le résumé doit refléter ce que dit l'EXPÉDITEUR du mail courant, pas l'historique.`;

  const tool = {
    name: 'tagger_email_entrant',
    description: 'Catégorise un email entrant pour Good Vibes.',
    input_schema: {
      type: 'object',
      properties: {
        category: {
          type: 'string',
          enum: CATEGORIES as unknown as string[],
        },
        urgencyScore: { type: 'integer', minimum: 1, maximum: 5 },
        summary: { type: 'string' },
        detectedLanguage: { type: 'string', enum: ['FR', 'EN', 'DE', 'LU', 'AUTRE'] },
      },
      required: ['category', 'urgencyScore', 'summary', 'detectedLanguage'],
    },
  };

  const res = await fetch(ANTHROPIC_API, {
    method: 'POST',
    headers: {
      'x-api-key': token,
      'anthropic-version': '2023-06-01',
      'anthropic-dangerous-direct-browser-access': 'true',
      'content-type': 'application/json',
    },
    body: JSON.stringify({
      model: MODEL,
      max_tokens: 512,
      tools: [tool],
      tool_choice: { type: 'tool', name: 'tagger_email_entrant' },
      messages: [{ role: 'user', content: prompt }],
    }),
  });

  if (!res.ok) {
    const err = await res.text().catch(() => '');
    throw new Error(`Anthropic ${res.status}: ${err.slice(0, 200)}`);
  }

  const data: any = await res.json();
  const toolUse = (data.content || []).find((b: any) => b.type === 'tool_use');
  if (!toolUse) throw new Error('Réponse Anthropic invalide (pas de tool_use).');

  const out = toolUse.input as Record<string, string | number | undefined>;
  let category = (out.category as ClaudeCategory) || 'autre';
  let urgencyScore = Math.max(1, Math.min(5, Number(out.urgencyScore) || 2));

  // Garde-fou anti "demande_interne_staff" mal classé (même règle que le service backend)
  if (category === 'demande_interne_staff' && !isInternalThread) {
    console.info('[claude addin] override demande_interne_staff → autre (destinataire externe)');
    category = 'autre';
  }

  // Plafond Cc urgence 2
  if (userIsInCc && !userIsInTo && urgencyScore > 2 && !/urgent|asap|deadline|aujourd['’]hui|today/i.test(cleanedBody)) {
    urgencyScore = 2;
  }

  return {
    category,
    urgencyScore,
    summary: String(out.summary || '').slice(0, 500),
    detectedLanguage: (out.detectedLanguage as ClaudeAnalysis['detectedLanguage']) || 'AUTRE',
  };
}
