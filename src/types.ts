// ── ATLAS Outlook Add-in — Shared Types ──

export interface Projet {
  id: string;
  noProjet: string;
  denomination: string;
  client: string;
  statut: string;
  enCharge: string;
  refProjet: string;
  dateDebut: string;
  dateFin: string;
}

export interface Tier {
  id: string;
  relation: string;      // Company name
  categorie: string;
  email: string;
  telephone: string;
}

export interface Contact {
  id: string;
  personneDeContact: string;
  email: string;
  relationSociete: string;  // Parent company
  fonction: string;
}

export interface MailMessageFull {
  id: string;
  subject: string;
  from: { name: string; email: string };
  receivedAt: string;
  isRead: boolean;
  hasAttachments: boolean;
  bodyPreview: string;
  webLink: string;
  conversationId?: string;
  internetMessageId?: string;
  toRecipients: Array<{ name: string; email: string }>;
  ccRecipients: Array<{ name: string; email: string }>;
  bodyHtml: string | null;
  bodyText: string;
  attachments: MailAttachment[];
}

export interface MailAttachment {
  id: string;
  name: string;
  size: number;
  contentType: string;
  isInline: boolean;
}

export interface ArgoProfile {
  prenom: string;
  nom: string;
  tonPrefere: string;       // "Amical" (tu) | "Professionnel" (vous) | ""
  languePreferee: string;   // "FR" | "EN" | "DE" | "LU" | ""
  tutoiementAvec: string[];
}

export interface EmailTemplate {
  id: string;
  nom: string;
  type: string;
  ton: string;
  marque: string;
  categorie: string;
  sujetFR: string;
  corpsFR: string;
  sujetEN: string;
  corpsEN: string;
  sujetDE: string;
  corpsDE: string;
  sujetLU: string;
  corpsLU: string;
  variables: string;
  statut: string;
}

export interface SearchResult {
  type: 'projet' | 'tiers' | 'contact';
  id: string;
  label: string;
  detail: string;
}

export type AddinMode = 'read' | 'compose' | 'link' | 'create' | 'view';

export interface AddinState {
  mode: AddinMode;
  isConfigured: boolean;
  airtableToken: string;
  graphToken: string | null;
  anthropicKey: string;
  userName: string;
  userEmail: string;
}
