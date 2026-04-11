/**
 * ARGO AI client for ATLAS Outlook Add-in
 * Email analysis + tone adaptation via Claude API
 */

import type { ArgoProfile } from '../types';

function getAnthropicKey(): string {
  return localStorage.getItem('atlas_addin_anthropic_key') || '';
}

async function callClaude(prompt: string, maxTokens = 1024): Promise<string> {
  const key = getAnthropicKey();
  if (!key) throw new Error('No Anthropic API key configured');

  const res = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': key,
      'anthropic-version': '2023-06-01',
      'anthropic-dangerous-direct-browser-access': 'true',
    },
    body: JSON.stringify({
      model: 'claude-haiku-4-5-20250929',
      max_tokens: maxTokens,
      messages: [{ role: 'user', content: prompt }],
    }),
  });

  if (!res.ok) throw new Error(`Claude API ${res.status}`);
  const data = await res.json();
  return (data.content?.[0]?.text || '').trim();
}

// ── Email Analysis ──

export interface EmailAnalysis {
  denomination: string;
  client?: string;
  typeEvenement?: string;
  debut?: string;
  lieu?: string;
  descriptif?: string;
}

export async function analyzeEmailForProjet(
  subject: string, fromName: string, fromEmail: string, bodyPreview: string
): Promise<EmailAnalysis> {
  const prompt = `Analyse cet email pour en extraire les infos projet/événement. Retourne un JSON avec les champs: denomination, client, typeEvenement, debut (YYYY-MM-DD), lieu, descriptif. Sois concis.

De: ${fromName} <${fromEmail}>
Sujet: ${subject}
Aperçu: ${bodyPreview?.slice(0, 500)}

Retourne UNIQUEMENT le JSON, sans markdown.`;

  try {
    const raw = await callClaude(prompt, 512);
    return JSON.parse(raw.replace(/```json?\s*\n?/i, '').replace(/\n?```\s*$/i, ''));
  } catch {
    return { denomination: subject, client: fromName };
  }
}

// ── Analyze received email for response tone ──

export async function analyzeReceivedEmail(
  subject: string, body: string, senderName: string
): Promise<{ sentiment: string; urgence: string; tonUtilise: string; suggestions: string[] }> {
  const key = getAnthropicKey();
  if (!key) return { sentiment: 'neutre', urgence: 'normal', tonUtilise: 'professionnel', suggestions: [] };

  const prompt = `Analyse cet email reçu pour déterminer le ton de la réponse appropriée.

De: ${senderName}
Sujet: ${subject}
Corps: ${body.slice(0, 800)}

Retourne un JSON avec:
- sentiment: "positif" | "neutre" | "négatif" | "urgent"
- urgence: "faible" | "normal" | "élevé" | "critique"
- tonUtilise: "tutoyé" | "vouvoyé" | "mixte" | "formel" | "informel"
- suggestions: [3 phrases-clés à inclure dans la réponse]

UNIQUEMENT le JSON, sans markdown.`;

  try {
    const raw = await callClaude(prompt, 512);
    return JSON.parse(raw.replace(/```json?\s*\n?/i, '').replace(/\n?```\s*$/i, ''));
  } catch {
    return { sentiment: 'neutre', urgence: 'normal', tonUtilise: 'professionnel', suggestions: [] };
  }
}

// ── Tone Adaptation ──

export function getSalutation(profile: ArgoProfile | null, senderName?: string): string {
  if (!profile) return 'Bonjour,';

  const lang = profile.languePreferee || 'FR';
  const isTu = profile.tonPrefere === 'Amical' ||
    (senderName && profile.tutoiementAvec.some(n => n.toLowerCase().includes(senderName.toLowerCase())));
  const prenom = profile.prenom || '';

  switch (lang) {
    case 'EN': return prenom ? `Dear ${prenom},` : 'Dear Sir/Madam,';
    case 'DE': return isTu ? `Hallo ${prenom},` : `Sehr geehrte Damen und Herren,`;
    case 'LU': return prenom ? `Gudde Moien ${prenom},` : 'Gudde Moien,';
    default:
      if (isTu && prenom) return `Salut ${prenom},`;
      if (prenom) return `Bonjour ${prenom},`;
      return 'Bonjour,';
  }
}

export function getClosing(profile: ArgoProfile | null, senderName?: string): string {
  if (!profile) return 'Cordialement,';

  const lang = profile.languePreferee || 'FR';
  const isTu = profile.tonPrefere === 'Amical' ||
    (senderName && profile.tutoiementAvec.some(n => n.toLowerCase().includes(senderName.toLowerCase())));

  switch (lang) {
    case 'EN': return isTu ? 'Best regards,' : 'Kind regards,';
    case 'DE': return isTu ? 'Viele Grüße,' : 'Mit freundlichen Grüßen,';
    case 'LU': return 'Mat beschte Gréiss,';
    default: return isTu ? 'A bientôt,' : 'Cordialement,';
  }
}

export async function adaptEmailBody(
  html: string, profile: ArgoProfile | null, senderName?: string
): Promise<string> {
  if (!profile || !html) return html;

  const key = getAnthropicKey();
  if (!key) return simpleAdapt(html, profile, senderName);

  const lang = profile.languePreferee || 'FR';
  const isTu = profile.tonPrefere === 'Amical' ||
    (senderName && profile.tutoiementAvec.some(n => n.toLowerCase().includes((senderName || '').toLowerCase())));

  const langLabel = lang === 'LU' ? 'Luxembourgeois (salutation et closing) + Français tutoyé (corps)'
    : lang === 'DE' ? 'Allemand' : lang === 'EN' ? 'Anglais'
    : isTu ? 'Français tutoyé' : 'Français vouvoyé';

  const prompt = `Réécris cet email HTML de manière naturelle et fluide pour le destinataire. Retourne UNIQUEMENT le HTML brut, sans bloc markdown.

Destinataire : ${profile.prenom} ${profile.nom}
Style : ${langLabel}
Tonalité : ${isTu ? 'Informel, amical, tutoyé' : 'Professionnel, formel, vouvoyé'}

Instructions :
- Réécris les phrases naturellement (pas de substitution mécanique)
- PRÉSERVE exactement le HTML (balises, liens, styles)
- PRÉSERVE les liens, mots de passe, numéros de référence mot pour mot

HTML :
${html}`;

  try {
    const res = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': key,
        'anthropic-version': '2023-06-01',
        'anthropic-dangerous-direct-browser-access': 'true',
      },
      body: JSON.stringify({
        model: 'claude-sonnet-4-5-20250929',
        max_tokens: 1024,
        messages: [{ role: 'user', content: prompt }],
      }),
    });

    if (!res.ok) return simpleAdapt(html, profile, senderName);
    const data = await res.json();
    let adapted = (data.content?.[0]?.text || '').trim()
      .replace(/^```html?\s*\n?/i, '').replace(/\n?```\s*$/i, '').trim();

    if (adapted && (adapted.includes('<p>') || adapted.includes('<'))) return adapted;
    return simpleAdapt(html, profile, senderName);
  } catch {
    return simpleAdapt(html, profile, senderName);
  }
}

function simpleAdapt(html: string, profile: ArgoProfile, senderName?: string): string {
  const salutation = getSalutation(profile, senderName);
  const closing = getClosing(profile, senderName);
  const isTu = profile.tonPrefere === 'Amical' ||
    (senderName && profile.tutoiementAvec.some(n => n.toLowerCase().includes((senderName || '').toLowerCase())));

  let adapted = html;
  adapted = adapted.replace(/<p>Bonjour,<\/p>/i, `<p>${salutation}</p>`);
  adapted = adapted.replace(/^Bonjour,/im, salutation);
  adapted = adapted.replace(/<p>Cordialement,<\/p>/i, `<p>${closing}</p>`);
  adapted = adapted.replace(/Cordialement,/i, closing);

  if (isTu) {
    adapted = adapted.replace(/Veuillez trouver/gi, 'Tu trouveras');
    adapted = adapted.replace(/Veuillez cliquer/gi, 'Clique');
    adapted = adapted.replace(/Veuillez/gi, "N'hésite pas à");
    adapted = adapted.replace(/votre espace/gi, 'ton espace');
    adapted = adapted.replace(/votre demande/gi, 'ta demande');
    adapted = adapted.replace(/votre email/gi, 'ton email');
    adapted = adapted.replace(/Vous trouverez/gi, 'Tu trouveras');
    adapted = adapted.replace(/vous trouverez/gi, 'tu trouveras');
    adapted = adapted.replace(/N'hésitez pas/gi, "N'hésite pas");
    adapted = adapted.replace(/n'hésitez pas/gi, "n'hésite pas");
    adapted = adapted.replace(/je vous répondrai/gi, 'je te répondrai');
  }

  return adapted;
}
