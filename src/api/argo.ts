/**
 * ARGO AI client for ATLAS Outlook Add-in
 * Email analysis + tone adaptation via Claude API
 */

import type { ArgoProfile } from '../types';

function getAnthropicKey(): string {
  return localStorage.getItem('atlas_addin_anthropic_key') || '';
}

/**
 * Détection rapide de langue par scoring sur mots-clés fréquents.
 * Retourne 'FR' | 'EN' | 'DE' | 'LU' | null (si aucun signal).
 * Pas de dépendance externe. Suffit pour distinguer FR/EN/DE/LU dans nos
 * mails business.
 */
export function detectLanguage(text: string): 'FR' | 'EN' | 'DE' | 'LU' | null {
  const t = ' ' + text.toLowerCase().replace(/[^a-zàâäéèêëïîôöùûüÿœæç\s]/gi, ' ') + ' ';
  // Mots discriminants courts (entourés d'espaces pour éviter substring fortuit)
  const markers = {
    FR: [' le ', ' la ', ' les ', ' de ', ' du ', ' des ', ' un ', ' une ', ' est ', ' pour ', ' bonjour ', ' merci ', ' cordialement ', ' avec ', ' vous ', ' nous ', ' votre ', ' notre ', ' bien ', ' à '],
    EN: [' the ', ' is ', ' are ', ' for ', ' with ', ' you ', ' your ', ' hello ', ' hi ', ' thanks ', ' regards ', ' best ', ' please ', ' would ', ' could '],
    DE: [' der ', ' die ', ' das ', ' und ', ' ist ', ' mit ', ' für ', ' wir ', ' sie ', ' ihre ', ' bitte ', ' danke ', ' guten ', ' viele ', ' grüße ', ' freundlichen ', ' hallo ', ' sehr ', ' geehrte ', ' moien '],
    LU: [' moien ', ' merci ', ' wann ', ' ech ', ' mir ', ' dir ', ' net ', ' fir ', ' awer ', ' hatt ', ' antwerten ', ' farv ', ' gréiss ', ' grouss '],
  } as const;

  const scores = { FR: 0, EN: 0, DE: 0, LU: 0 } as Record<string, number>;
  for (const [lang, words] of Object.entries(markers)) {
    for (const w of words) {
      // count occurrences
      let i = 0; let n = 0;
      while ((i = t.indexOf(w, i)) !== -1) { n++; i += w.length; }
      scores[lang] += n;
    }
  }
  // Le LU partage beaucoup avec DE — on boost le LU si "moien" présent (très spécifique)
  if (t.includes(' moien ') || t.includes(' gréiss ')) scores.LU += 5;

  const best = Object.entries(scores).sort((a, b) => b[1] - a[1])[0];
  if (!best || best[1] < 2) return null; // pas assez de signal
  return best[0] as 'FR' | 'EN' | 'DE' | 'LU';
}

/**
 * Détecte si le mail / draft est en tutoiement (informel) ou vouvoiement.
 * Heuristique par scoring de markers :
 *   FR : "tu/ton/ta/tes/te" / "salut" → tu | "vous/votre/cordialement" → vous
 *   DE : "du/dein/dir/dich" / "hallo/hi/lieber" → tu | "Sie/Ihre/sehr geehrte" → vous
 *   EN : "hi <prenom>" / "hey" → tu | "Dear Mr/Mrs" → vous
 *   LU : "du/äis/säi" → tu | (rarement vouvoiement)
 * Retourne true si tutoiement, false si vouvoiement, null si pas de signal.
 */
export function detectTutoiement(text: string, lang: 'FR' | 'EN' | 'DE' | 'LU' | null): boolean | null {
  const t = ' ' + text.toLowerCase().replace(/[^a-zàâäéèêëïîôöùûüÿœæçß\s]/gi, ' ') + ' ';
  let scoreTu = 0;
  let scoreVous = 0;

  // FR
  if (!lang || lang === 'FR') {
    for (const m of [' tu ', ' ton ', ' ta ', ' tes ', ' te ', ' t\'', ' salut ', ' coucou ', ' bisous ', ' bises ']) {
      if (t.includes(m)) scoreTu += m === ' tu ' ? 3 : 1;
    }
    for (const m of [' vous ', ' votre ', ' vos ', ' cordialement ', ' madame ', ' monsieur ', ' bien à vous ']) {
      if (t.includes(m)) scoreVous += m === ' vous ' ? 3 : 1;
    }
  }
  // DE
  if (!lang || lang === 'DE') {
    for (const m of [' du ', ' dein ', ' deine ', ' deinen ', ' dir ', ' dich ', ' hallo ', ' hi ', ' lieber ', ' liebe ', ' viele grüße ']) {
      if (t.includes(m)) scoreTu += m === ' du ' ? 3 : 1;
    }
    // 'Sie' en majuscule = vouvoiement DE, mais on travaille en lowercase →
    // marker 'sehr geehrte' / 'ihnen' / 'ihre' (avec Maj on lowercase ça reste)
    for (const m of [' ihnen ', ' ihre ', ' ihren ', ' sehr geehrte ', ' freundlichen grüßen ', ' mit freundlichen ']) {
      if (t.includes(m)) scoreVous += 2;
    }
  }
  // EN
  if (!lang || lang === 'EN') {
    for (const m of [' hi ', ' hey ', ' cheers ', ' thanks ', ' xoxo ']) {
      if (t.includes(m)) scoreTu += 1;
    }
    for (const m of [' dear mr ', ' dear mrs ', ' dear ms ', ' sincerely ', ' kindly ', ' to whom it may concern ']) {
      if (t.includes(m)) scoreVous += 2;
    }
  }
  // LU — généralement tutoiement par défaut
  if (lang === 'LU') {
    for (const m of [' du ', ' däi ', ' deng ', ' moien ', ' gréiss ']) {
      if (t.includes(m)) scoreTu += 1;
    }
  }

  if (scoreTu === 0 && scoreVous === 0) return null;
  return scoreTu > scoreVous;
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
      model: 'claude-haiku-4-5',
      max_tokens: maxTokens,
      messages: [{ role: 'user', content: prompt }],
    }),
  });

  if (!res.ok) {
    const body = await res.text().catch(() => '');
    throw new Error(`Claude API ${res.status}: ${body.slice(0, 200)}`);
  }
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

// ── Email Summary ──

export async function summarizeEmail(subject: string, body: string, senderName: string): Promise<string> {
  const key = getAnthropicKey();
  if (!key) return '';

  const prompt = `Resume cet email en 2-3 lignes concises et utiles pour un gestionnaire de projet. Mentionne l'action demandee s'il y en a une. Pas de guillemets, pas de prefixe "Resume:", juste le texte.

De: ${senderName}
Sujet: ${subject}
Corps: ${body.slice(0, 1500)}`;

  try {
    return await callClaude(prompt, 200);
  } catch { return ''; }
}

// ── Quick Reply Suggestions ──

export interface QuickReplySuggestion {
  label: string;
  tone: string;
  body: string;
}

export async function generateQuickReplies(
  subject: string, body: string, senderName: string, profile: ArgoProfile | null, userName: string
): Promise<QuickReplySuggestion[]> {
  const key = getAnthropicKey();
  if (!key) return [];

  const detectedLang: 'FR' | 'EN' | 'DE' | 'LU' = (detectLanguage(body.slice(0, 800)) || (profile?.languePreferee as any) || 'FR');
  const lang = detectedLang;
  const detectedTu = detectTutoiement(body.slice(0, 800), detectedLang);
  const profileTu = profile?.tonPrefere === 'Amical' ||
    (profile?.tutoiementAvec?.some(n => n.toLowerCase().includes(userName.toLowerCase())) ?? false);
  const isTu = detectedTu !== null ? detectedTu : profileTu;
  const prenom = profile?.prenom || senderName.split(' ')[0] || '';

  const prompt = `Tu es un assistant email pour une agence de communication luxembourgeoise (GOOD VIBES).
Genere exactement 3 reponses courtes et naturelles a cet email.

De: ${senderName}
Sujet: ${subject}
Corps: ${body.slice(0, 800)}

Contexte:
- Repondre dans la MEME LANGUE que le mail reçu : ${lang === 'EN' ? 'anglais' : lang === 'DE' ? 'allemand' : lang === 'LU' ? 'luxembourgeois' : 'francais'}
- ${isTu ? 'Tutoyer' : 'Vouvoyer'} le destinataire (prenom: ${prenom})
- Ton professionnel mais chaleureux (agence de com)
- INTERDIT : ne JAMAIS ajouter de signature, nom ${userName}, "GOOD VIBES" ou closing ("Viele Grüße/Cordialement/Best regards"). Exclaimer gère la signature corporate automatiquement, sinon doublon dans le mail envoyé.
- La réponse DOIT se terminer sur la dernière phrase du corps, sans formule de fin.

Retourne un JSON array de 3 objets: [{"label": "2-3 mots max (emoji + description)", "tone": "positif|neutre|formel", "body": "le HTML de la reponse complete avec <p> tags, salutation et closing inclus"}]

UNIQUEMENT le JSON, sans markdown.`;

  try {
    const raw = await callClaude(prompt, 1500);
    const parsed = JSON.parse(raw.replace(/```json?\s*\n?/i, '').replace(/\n?```\s*$/i, ''));
    return Array.isArray(parsed) ? parsed.slice(0, 3) : [];
  } catch { return []; }
}

// ── Free-form AI Reply ──

export async function generateFreeReply(
  subject: string, body: string, senderName: string,
  instruction: string, profile: ArgoProfile | null, userName: string
): Promise<string> {
  const key = getAnthropicKey();
  if (!key) throw new Error('Cle API Anthropic requise');

  const prenom = profile?.prenom || senderName.split(' ')[0] || '';

  // Détection langue : on prend la langue du DRAFT en priorité (intent
  // explicite de Charles), sinon celle du MAIL REÇU (auto-match), sinon
  // le profile, sinon FR.
  const combinedText = `${instruction}\n\n${body.slice(0, 800)}`;
  const detectedLang: 'FR' | 'EN' | 'DE' | 'LU' = (detectLanguage(combinedText) || (profile?.languePreferee as any) || 'FR');
  const langLabel = detectedLang === 'EN' ? 'anglais' : detectedLang === 'DE' ? 'allemand' : detectedLang === 'LU' ? 'luxembourgeois' : 'francais';

  // Détection tutoiement : depuis le draft + mail reçu (du/tu/hi vs Sie/vous/Dear).
  // Si détecté → prime sur le profile. Sinon fallback au profile.
  const detectedTu = detectTutoiement(combinedText, detectedLang);
  const profileTu = profile?.tonPrefere === 'Amical' ||
    (profile?.tutoiementAvec?.some(n => n.toLowerCase().includes(userName.toLowerCase())) ?? false);
  const isTu = detectedTu !== null ? detectedTu : profileTu;

  const prompt = `Tu es l'éditeur des emails de ${userName} (GOOD VIBES events & communications, agence de com au Luxembourg).

Ton job : prendre le brouillon brut de ${userName} ci-dessous et le POLIR pour qu'il soit prêt à envoyer. Garde l'INTENTION et le TON de l'auteur, n'invente pas de contenu, ne reformule pas tout — corrige juste :
  • Fautes d'orthographe / grammaire / ponctuation
  • Tournures maladroites ou trop télégraphiques
  • Ajoute une salutation d'ouverture (Hallo Luca, / Salut X, etc) si absente
  • Améliore le flow si nécessaire (sans changer le sens)
  • Garde la même longueur ± 30%, pas de blabla supplémentaire

Email reçu (contexte uniquement, pour comprendre la conversation) :
De: ${senderName}
Sujet: ${subject}
${body.slice(0, 800)}

BROUILLON DE ${userName} À POLIR :
"""
${instruction}
"""

Règles CRITIQUES :
- LANGUE : ${langLabel}. Détecté : ${detectedLang}. La réponse doit être DANS CETTE LANGUE.
- FORME D'ADRESSE : ${isTu ? `TUTOIEMENT (du/tu/hi informel) — ${prenom} et ${userName} se tutoient.` : `VOUVOIEMENT (vous/Sie/Dear).`}
- SALUTATION D'OUVERTURE :
${isTu && detectedLang === 'DE' ? `  → "Hallo ${prenom},"` : ''}
${isTu && detectedLang === 'FR' ? `  → "Salut ${prenom},"` : ''}
${isTu && detectedLang === 'EN' ? `  → "Hi ${prenom},"` : ''}
${!isTu && detectedLang === 'DE' ? `  → "Sehr geehrte/r ${prenom},"` : ''}
${!isTu && detectedLang === 'FR' ? `  → "Bonjour ${prenom},"` : ''}
${!isTu && detectedLang === 'EN' ? `  → "Dear ${prenom},"` : ''}

⛔ INTERDIT — NE JAMAIS AJOUTER :
- Closing/signature ("Viele Grüße / Bien à toi / Best regards / Cordialement")
- Nom de ${userName}
- "GOOD VIBES events & communications"
- Logo, contact, footer

Ces éléments sont gérés par Exclaimer (signature corporate auto). Si tu les ajoutes, ils seront DUPLIQUÉS dans le mail envoyé. La réponse DOIT se terminer sur la dernière phrase du corps, sans formule de fin.

Format de sortie : HTML avec <p> tags. UNIQUEMENT le HTML, sans markdown, sans commentaires.`;

  const raw = await callClaude(prompt, 1500);
  return raw.replace(/^```html?\s*\n?/i, '').replace(/\n?```\s*$/i, '').trim();
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
        model: 'claude-sonnet-4-5',
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
