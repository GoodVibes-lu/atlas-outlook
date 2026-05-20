/**
 * Airtable API client for ATLAS Outlook Add-in
 * Lightweight version of the main app's Airtable services
 */

import type { Projet, Tier, Contact, MailMessageFull, EmailTemplate } from '../types';

const API_URL = 'https://api.airtable.com/v0';

// ── Base & Table IDs ──
const PROJETS_BASE = 'appKiJY0qjI4UTrWU';
const RELATIONS_BASE = 'app4TQws4kxKZTPts';
const ATLAS_BASE = 'appjtMG7hCTZqsG02';

const TABLES = {
  PROJETS: 'tblKBSumqrxAQFt2u',
  CLIENTS: 'tbl5zL8euh9HRH7bj',
  CONTACTS_CLIENTS: 'tblGJfIYAHaPE3thh',
  EMAILS_PROJET: 'tblsQnNwCG9QJn9fh',
  CONTACTS_RELATIONS: 'tblgrzBY9UTIONhtj',
  COMMUNICATIONS: 'tbl5e6la54kMiFnkV',
} as const;

// ── Email Projet Field IDs ──
const EF = {
  SUJET: 'fldOTTeY1RjPIZxRg',
  PROJET: 'fldFxTWpWmtS8YrfR',
  TIERS: 'fldBYpC5RECe5d9d1',
  CONTACT: 'fldotoFwwtRxMW0QG',
  CONVERSATION_ID: 'fldYxGo2kwoXtQOFS',
  INTERNET_MSG_ID: 'fld8Qj0zM8ACTx1Ne',
  GRAPH_MSG_ID: 'fldzbWxcZNcFWnfZi',
  DE_NOM: 'fldWerqRALCPfOD8Z',
  DE_EMAIL: 'fldQSslG9Z8mc98ge',
  DESTINATAIRES: 'fldjgFs94JdxQJYCO',
  CC: 'fldKghgunVcpPJYWf',
  DATE: 'fldOdN5XlsaY3pCIx',
  CORPS_HTML: 'fldsAixRYdLScpXOM',
  CORPS_TEXT: 'fld0BhszZiwmsMEIb',
  DIRECTION: 'fldfFRSwaThWVE6s4',
  A_PIECES_JOINTES: 'fldsWZPGJXj9G1aNE',
  PIECES_JOINTES_JSON: 'fldknSspuyihMSOGg',
  LIE_PAR: 'fldRpqJg9Tbs9FTyD',
  LIE_LE: 'fldmYBQkYUaiJlPe4',
  MAILBOX_SOURCE: 'fldoFmqL37IkBtZjL',
  PRIVE: 'fldq9zZOMpTR1P4Qx',
  PRIVE_PAR: 'fldT7udFfEiEDZK09',
} as const;

// ── Projet Field IDs ──
const PF = {
  NO_PROJET: 'fldGjQVMntdHceWLa',       // No Projet (autoNumber)
  DENOMINATION: 'fldaVDut8RijfsorS',     // Dénomination du projet (singleLineText)
  CLIENT: 'fld7Aa90eAmcYvY71',           // Client (multipleRecordLinks)
  STATUT: 'fld0JYd0AHLcVfhaT',           // Statut (singleSelect)
  EN_CHARGE: 'fldwyFpDMHjNUT2Cr',        // En charge (multipleRecordLinks)
  NOM_EN_CHARGE: 'fldW3ycEJFPMsKGyB',    // Nom_prénom_en_charge (formula — nom lisible)
  REF_PROJET: 'fldORNsU9KLxxGxAN',       // Ref Projet (formula — ex: "GV-724")
  NOM_CLIENT: 'fld0giuubfxZd45sq',        // NomClient (singleLineText)
  DATE_DEBUT: 'fld1UJIUV8yNmWhxL',       // Début (date)
  DATE_FIN: 'fldtTsomztPLywxf0',         // Fin (date)
} as const;

// ── Contact ARGO Profile Field IDs ──
const CF = {
  PRENOM: 'fldlIYQP2usQ1TnzE',
  NOM: 'fld0I4v3efCf8pnHg',
  TON_PREFERE: 'fldcpQVHivq26y0qM',
  LANGUE_PREFEREE: 'fldBQLHQ2VSO3MdLv',
  TUTOIEMENT_AVEC: 'fldiGoJ5p41KvSUQW',
} as const;

// ── Communications Field IDs (verified 2026-04-11) ──
const COMMS = {
  NOM: 'fldgAZ4fpwPzNws3n',         // Nom (singleLineText)
  TYPE: 'fld4zgRXS0vHC65x9',         // Type (singleSelect)
  TON: 'fldJlQlFiBMqspqif',          // Ton (singleSelect) ✓
  MARQUE: 'fldfEDDbRqzc9qnEt',       // Marque (singleSelect) ✓
  STATUT: 'flda4SlFpeHuQfY0M',       // Statut (singleSelect)
  CATEGORIE: 'fldQxDsh9coxZN9Sk',    // Categorie (singleSelect)
  OBJET_FR: 'fld5rMGYjK6rnd3xC',     // ✓
  CORPS_FR: 'fldH0tQ7SoVtDKsDe',     // ✓
  OBJET_EN: 'fldEaEn5Ehma9PE00',     // ✓
  CORPS_EN: 'fldbVEfyksg8k9UxP',     // ✓
  OBJET_DE: 'fldyJs1RdRSU9Cn10',     // Objet_DE
  CORPS_DE: 'fldm7fYSCeG1ZZtPi',     // Corps_DE
  OBJET_LU: 'fldKRsMj1CfjlBiFZ',     // Objet_LU
  CORPS_LU: 'fldk1a0LCvuw4IE7p',     // Corps_LU
  VARIABLES: 'fldA9eojxIBUElgxk',    // ✓
  DESTINATAIRES: 'fldSDwAhN3dUzoCIl', // Destinataires (multipleSelects)
  TAGS: 'fld3vQ65xIfem0Wcx',         // Tags (multipleSelects)
} as const;

// ── Client Field IDs (Projets base Clients table) ──
const CLF = {
  RELATION: 'fldYyVtj5Rh5TDgOb',     // Relation (company name)
  NOM: 'fldq4OWNWYiS7XUIg',          // Nom
  EMAIL: 'fldCNqiCExXOpLQuI',         // Email
  STATUT: 'fld3JlvlZP7t0aa80',       // Statut (singleSelect)
  SECTEUR: 'fldQZFM9TuuIEvcgD',      // Secteur (multipleSelects)
} as const;

// ── Contact Client Field IDs (Projets base Contacts_Clients table) ──
const CCF = {
  PERSONNE: 'fldDFZzW5cris8sva',     // Personne de contact
  EMAIL: 'fldzZSIPtUkkKJtVi',         // Email
  SOCIETE: 'fldlSSa8wDIQvoSey',      // Société
  FONCTION: 'fldHiZeUeyv4913jv',     // Fonction
  PRENOM: 'fldGnI3ERmuztxoX2',       // Prénom
  NOM: 'fldTTEiu6sS3Bf8rU',          // Nom
  STATUT: 'fld8hB9syCzrX5o0k',       // Statut
} as const;

// ── Projet creation additional fields ──
const PCF = {
  TYPE: 'fldzcy1D0UhFDnusK',              // Type (multipleSelects)
  FOURCHETTE_BUDGET: 'fldB7LExyJIQySlG6', // Fourchette budgetaire
  MOIS_REALISATION: 'flduVRo1sbIYv8uOY',  // Mois_Realisation_prévue
  ANNEE_REALISATION: 'fldyouUoKmDFIUL7O', // Année_Realisation_prévue
  CONTACT_CLIENTS: 'fldJiH0bNfUmiHKyh',   // Contact clients (multipleRecordLinks)
  DESCRIPTIF: 'fldo2A4Ja7UaiQ3QO',        // Descriptif (multilineText)
} as const;

// ── Employés table (Projets base) for "En charge" dropdown ──
const EMPLOYES_TABLE = 'tblajI3x4CuLW0tBO';
const EMF = {
  NAME: 'fld8MMLPqgx14wgpz',         // Name (primary)
  NOM_TXT: 'fldyX4SPWldcSxGvW',       // Nom_TXT (formula)
  EMAIL: 'fldB28x4sPa5QVp6L',         // Email
} as const;

// ── Helpers ──

function getToken(): string {
  return localStorage.getItem('atlas_addin_airtable_token') || '';
}

function headers(): Record<string, string> {
  return {
    'Authorization': `Bearer ${getToken()}`,
    'Content-Type': 'application/json',
  };
}

async function airtableFetch<T>(url: string, opts?: RequestInit): Promise<T> {
  const res = await fetch(url, { ...opts, headers: { ...headers(), ...opts?.headers } });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(`Airtable ${res.status}: ${JSON.stringify(err)}`);
  }
  return res.json();
}

function selectName(val: unknown): string {
  if (!val) return '';
  if (typeof val === 'string') return val;
  if (typeof val === 'object' && val !== null && 'name' in val) return (val as { name: string }).name;
  if (Array.isArray(val)) return val[0]?.name || val[0] || '';
  return '';
}

// ── Cache ──

const cache = new Map<string, { data: unknown; ts: number }>();
const CACHE_TTL = 5 * 60 * 1000; // 5 min

function getCached<T>(key: string): T | null {
  const c = cache.get(key);
  if (c && Date.now() - c.ts < CACHE_TTL) return c.data as T;
  return null;
}

function setCache<T>(key: string, data: T): void {
  cache.set(key, { data, ts: Date.now() });
}

// ── Projets ──

export async function getAllProjets(): Promise<Projet[]> {
  const cached = getCached<Projet[]>('projets');
  if (cached) return cached;

  const fields = Object.values(PF).map(f => `fields%5B%5D=${f}`).join('&');
  let all: Projet[] = [];
  let offset = '';

  do {
    const url = `${API_URL}/${PROJETS_BASE}/${TABLES.PROJETS}?returnFieldsByFieldId=true&${fields}${offset ? `&offset=${offset}` : ''}`;
    const data = await airtableFetch<{ records: any[]; offset?: string }>(url);

    all = all.concat(data.records.map((r: any) => {
      const f = r.fields || {};
      return {
        id: r.id,
        noProjet: f[PF.NO_PROJET] || '',
        denomination: f[PF.DENOMINATION] || '',
        client: f[PF.NOM_CLIENT] || selectName(f[PF.CLIENT]),
        statut: selectName(f[PF.STATUT]),
        enCharge: f[PF.NOM_EN_CHARGE] || '',
        refProjet: f[PF.REF_PROJET] || '',
        dateDebut: f[PF.DATE_DEBUT] || '',
        dateFin: f[PF.DATE_FIN] || '',
      };
    }));
    offset = data.offset || '';
  } while (offset);

  setCache('projets', all);
  return all;
}

// ── Tiers (from Projets base Clients) ──

export async function getAllTiers(): Promise<Tier[]> {
  const cached = getCached<Tier[]>('tiers');
  if (cached) return cached;

  let all: Tier[] = [];
  let offset = '';

  do {
    const url = `${API_URL}/${PROJETS_BASE}/${TABLES.CLIENTS}?returnFieldsByFieldId=true${offset ? `&offset=${offset}` : ''}`;
    const data = await airtableFetch<{ records: any[]; offset?: string }>(url);

    all = all.concat(data.records.map((r: any) => {
      const f = r.fields || {};
      return {
        id: r.id,
        relation: f[CLF.RELATION] || f[CLF.NOM] || '',
        categorie: selectName(f[CLF.SECTEUR] || ''),
        email: f[CLF.EMAIL] || '',
        telephone: '',
      };
    }));
    offset = data.offset || '';
  } while (offset);

  setCache('tiers', all);
  return all;
}

// ── Contacts (from Projets base Contacts_Clients) ──

export async function getAllContacts(): Promise<Contact[]> {
  const cached = getCached<Contact[]>('contacts');
  if (cached) return cached;

  let all: Contact[] = [];
  let offset = '';

  do {
    const url = `${API_URL}/${PROJETS_BASE}/${TABLES.CONTACTS_CLIENTS}?returnFieldsByFieldId=true${offset ? `&offset=${offset}` : ''}`;
    const data = await airtableFetch<{ records: any[]; offset?: string }>(url);

    all = all.concat(data.records.map((r: any) => {
      const f = r.fields || {};
      return {
        id: r.id,
        personneDeContact: f[CCF.PERSONNE] || '',
        email: f[CCF.EMAIL] || '',
        relationSociete: f[CCF.SOCIETE] || '',
        fonction: f[CCF.FONCTION] || '',
      };
    }));
    offset = data.offset || '';
  } while (offset);

  setCache('contacts', all);
  return all;
}

// ── Linked Conversation IDs ──

export async function getLinkedConversationIds(): Promise<Map<string, { projetId: string; projetName: string }>> {
  const cached = getCached<Map<string, { projetId: string; projetName: string }>>('convIds');
  if (cached) return cached;

  const map = new Map<string, { projetId: string; projetName: string }>();
  const fields = [EF.CONVERSATION_ID, EF.PROJET, EF.SUJET].map(f => `fields%5B%5D=${f}`).join('&');
  let offset = '';

  do {
    const url = `${API_URL}/${PROJETS_BASE}/${TABLES.EMAILS_PROJET}?returnFieldsByFieldId=true&${fields}${offset ? `&offset=${offset}` : ''}`;
    const data = await airtableFetch<{ records: any[]; offset?: string }>(url);

    for (const r of data.records) {
      const f = r.fields || {};
      const convId = f[EF.CONVERSATION_ID];
      const projetIds = f[EF.PROJET] || [];
      if (convId && projetIds.length > 0) {
        map.set(convId, { projetId: projetIds[0], projetName: f[EF.SUJET] || '' });
      }
    }
    offset = data.offset || '';
  } while (offset);

  setCache('convIds', map);
  return map;
}

// ── Check if email already linked ──

export async function getAllLinkedEmailIds(): Promise<{ graphIds: Set<string>; internetIds: Set<string> }> {
  const cached = getCached<{ graphIds: Set<string>; internetIds: Set<string> }>('linkedIds');
  if (cached) return cached;

  const graphIds = new Set<string>();
  const internetIds = new Set<string>();
  const fields = [EF.GRAPH_MSG_ID, EF.INTERNET_MSG_ID].map(f => `fields%5B%5D=${f}`).join('&');
  let offset = '';

  do {
    const url = `${API_URL}/${PROJETS_BASE}/${TABLES.EMAILS_PROJET}?returnFieldsByFieldId=true&${fields}${offset ? `&offset=${offset}` : ''}`;
    const data = await airtableFetch<{ records: any[]; offset?: string }>(url);

    for (const r of data.records) {
      const f = r.fields || {};
      if (f[EF.GRAPH_MSG_ID]) graphIds.add(f[EF.GRAPH_MSG_ID]);
      if (f[EF.INTERNET_MSG_ID]) internetIds.add(f[EF.INTERNET_MSG_ID]);
    }
    offset = data.offset || '';
  } while (offset);

  const result = { graphIds, internetIds };
  setCache('linkedIds', result);
  return result;
}

// ── Resolve Tiers name → Projets base Client record ID ──

export async function resolveClientIdInProjetsBase(tiersName: string): Promise<string | null> {
  const formula = encodeURIComponent(`{Relation} = "${tiersName.replace(/"/g, '\\"')}"`);
  const url = `${API_URL}/${PROJETS_BASE}/${TABLES.CLIENTS}?filterByFormula=${formula}&maxRecords=1`;
  try {
    const data = await airtableFetch<{ records: Array<{ id: string }> }>(url);
    return data.records?.[0]?.id || null;
  } catch { return null; }
}

// ── Resolve Contact name → Projets base Contacts_Clients record ID ──

export async function resolveContactIdInProjetsBase(contactName: string): Promise<string | null> {
  const formula = encodeURIComponent(`{Personne de contact} = "${contactName.replace(/"/g, '\\"')}"`);
  const url = `${API_URL}/${PROJETS_BASE}/${TABLES.CONTACTS_CLIENTS}?filterByFormula=${formula}&maxRecords=1`;
  try {
    const data = await airtableFetch<{ records: Array<{ id: string }> }>(url);
    return data.records?.[0]?.id || null;
  } catch { return null; }
}

// ── Link email to project ──

const AIRTABLE_TEXT_LIMIT = 95_000;

function sanitize(text: string | null | undefined, limit = AIRTABLE_TEXT_LIMIT): string {
  if (!text) return '';
  let clean = text.replace(/data:[^;]+;base64,[A-Za-z0-9+/=]+/g, '[image inline supprimée]');
  if (clean.length > limit) clean = clean.slice(0, limit) + '\n\n[… contenu tronqué]';
  return clean;
}

export async function linkEmailToProject(
  email: MailMessageFull,
  projetRecordId: string,
  linkedByName: string,
  direction: 'reçu' | 'envoyé' = 'reçu',
  tiersRecordId?: string,
  options?: { prive?: boolean; privePar?: string },
): Promise<string> {
  const fields: Record<string, unknown> = {
    [EF.SUJET]: email.subject,
    [EF.PROJET]: [projetRecordId],
    [EF.CONVERSATION_ID]: email.conversationId || '',
    [EF.INTERNET_MSG_ID]: email.internetMessageId || '',
    [EF.GRAPH_MSG_ID]: email.id,
    [EF.DE_NOM]: email.from.name,
    [EF.DE_EMAIL]: email.from.email,
    [EF.DESTINATAIRES]: JSON.stringify(email.toRecipients || []),
    [EF.CC]: JSON.stringify(email.ccRecipients || []),
    [EF.DATE]: email.receivedAt,
    [EF.CORPS_HTML]: sanitize(email.bodyHtml),
    [EF.CORPS_TEXT]: sanitize(email.bodyText),
    [EF.DIRECTION]: direction,
    [EF.A_PIECES_JOINTES]: email.hasAttachments,
    [EF.PIECES_JOINTES_JSON]: JSON.stringify(
      email.attachments.filter(a => !a.isInline).map(a => ({ name: a.name, size: a.size, contentType: a.contentType }))
    ),
    [EF.LIE_PAR]: linkedByName,
    [EF.LIE_LE]: new Date().toISOString(),
    [EF.MAILBOX_SOURCE]: 'outlook-addin',
  };

  if (tiersRecordId) fields[EF.TIERS] = [tiersRecordId];
  if (options?.prive) {
    fields[EF.PRIVE] = true;
    fields[EF.PRIVE_PAR] = options.privePar || linkedByName;
  }

  const url = `${API_URL}/${PROJETS_BASE}/${TABLES.EMAILS_PROJET}`;
  const result = await airtableFetch<{ id: string }>(url, {
    method: 'POST',
    body: JSON.stringify({ fields, returnFieldsByFieldId: true }),
  });

  // Invalidate cache
  cache.delete('linkedIds');
  cache.delete('convIds');

  return result.id;
}

// ── Link email to contact ──

export async function linkEmailToContact(
  email: MailMessageFull,
  contactName: string,
  linkedByName: string,
  direction: 'reçu' | 'envoyé' = 'reçu',
  tiersName?: string,
  options?: { prive?: boolean; privePar?: string },
): Promise<string> {
  const contactRecordId = await resolveContactIdInProjetsBase(contactName);
  if (!contactRecordId) throw new Error(`Contact "${contactName}" introuvable`);

  const fields: Record<string, unknown> = {
    [EF.SUJET]: email.subject,
    [EF.CONTACT]: [contactRecordId],
    [EF.CONVERSATION_ID]: email.conversationId || '',
    [EF.INTERNET_MSG_ID]: email.internetMessageId || '',
    [EF.GRAPH_MSG_ID]: email.id,
    [EF.DE_NOM]: email.from.name,
    [EF.DE_EMAIL]: email.from.email,
    [EF.DESTINATAIRES]: JSON.stringify(email.toRecipients || []),
    [EF.CC]: JSON.stringify(email.ccRecipients || []),
    [EF.DATE]: email.receivedAt,
    [EF.CORPS_HTML]: sanitize(email.bodyHtml),
    [EF.CORPS_TEXT]: sanitize(email.bodyText),
    [EF.DIRECTION]: direction,
    [EF.A_PIECES_JOINTES]: email.hasAttachments,
    [EF.PIECES_JOINTES_JSON]: JSON.stringify(
      email.attachments.filter(a => !a.isInline).map(a => ({ name: a.name, size: a.size, contentType: a.contentType }))
    ),
    [EF.LIE_PAR]: linkedByName,
    [EF.LIE_LE]: new Date().toISOString(),
    [EF.MAILBOX_SOURCE]: 'outlook-addin',
  };

  if (tiersName) {
    const tiersId = await resolveClientIdInProjetsBase(tiersName);
    if (tiersId) fields[EF.TIERS] = [tiersId];
  }
  if (options?.prive) {
    fields[EF.PRIVE] = true;
    fields[EF.PRIVE_PAR] = options.privePar || linkedByName;
  }

  const url = `${API_URL}/${PROJETS_BASE}/${TABLES.EMAILS_PROJET}`;
  const result = await airtableFetch<{ id: string }>(url, {
    method: 'POST',
    body: JSON.stringify({ fields, returnFieldsByFieldId: true }),
  });

  cache.delete('linkedIds');
  cache.delete('convIds');
  return result.id;
}

// ── Fetch ARGO profile for contact ──

export async function fetchContactArgoProfile(contactEmail: string): Promise<import('../types').ArgoProfile | null> {
  if (!contactEmail) return null;
  try {
    const formula = encodeURIComponent(`{Email} = "${contactEmail.replace(/"/g, '\\"')}"`);
    const fields = Object.values(CF).map(f => `fields%5B%5D=${f}`).join('&');
    const url = `${API_URL}/${RELATIONS_BASE}/${TABLES.CONTACTS_RELATIONS}?filterByFormula=${formula}&${fields}&returnFieldsByFieldId=true&maxRecords=1`;

    const data = await airtableFetch<{ records: any[] }>(url);
    const record = data.records?.[0];
    if (!record) return null;

    const f = record.fields || {};
    return {
      prenom: f[CF.PRENOM] || '',
      nom: f[CF.NOM] || '',
      tonPrefere: selectName(f[CF.TON_PREFERE]),
      languePreferee: selectName(f[CF.LANGUE_PREFEREE]),
      tutoiementAvec: Array.isArray(f[CF.TUTOIEMENT_AVEC])
        ? f[CF.TUTOIEMENT_AVEC].map((v: unknown) => typeof v === 'string' ? v : selectName(v))
        : [],
    };
  } catch { return null; }
}

// ── Fetch Communication templates ──

export async function getEmailTemplates(): Promise<EmailTemplate[]> {
  const cached = getCached<EmailTemplate[]>('templates');
  if (cached) return cached;

  const fields = Object.values(COMMS).map(f => `fields%5B%5D=${f}`).join('&');
  const url = `${API_URL}/${ATLAS_BASE}/${TABLES.COMMUNICATIONS}?returnFieldsByFieldId=true&${fields}`;
  const data = await airtableFetch<{ records: any[] }>(url);

  const templates = data.records
    .map((r: any) => {
      const f = r.fields || {};
      return {
        id: r.id,
        nom: f[COMMS.NOM] || '',
        type: selectName(f[COMMS.TYPE]),
        ton: selectName(f[COMMS.TON]),
        marque: selectName(f[COMMS.MARQUE]),
        categorie: selectName(f[COMMS.CATEGORIE]),
        statut: selectName(f[COMMS.STATUT]),
        sujetFR: f[COMMS.OBJET_FR] || '',
        corpsFR: f[COMMS.CORPS_FR] || '',
        sujetEN: f[COMMS.OBJET_EN] || '',
        corpsEN: f[COMMS.CORPS_EN] || '',
        sujetDE: f[COMMS.OBJET_DE] || '',
        corpsDE: f[COMMS.CORPS_DE] || '',
        sujetLU: f[COMMS.OBJET_LU] || '',
        corpsLU: f[COMMS.CORPS_LU] || '',
        variables: f[COMMS.VARIABLES] || '',
      };
    })
    .filter(t => t.statut === 'Actif' && (t.type === 'Email' || t.type === 'Réponse type'));

  setCache('templates', templates);
  return templates;
}

// ── Get projects for a specific client ──

export async function getProjetsByClient(tiersName: string): Promise<Projet[]> {
  const all = await getAllProjets();
  return all.filter(p => p.client.toLowerCase().includes(tiersName.toLowerCase()));
}

// ── Folder Mappings (learned Outlook folder paths per user) ──

const FM_TABLE = 'tblaK3BjfSmPduFXH';
const FMF = {
  CLE: 'fldJsVLkNWi0KhyEJ',
  USER_EMAIL: 'fldFWBn7ovTxwLIMj',
  CLIENT: 'fld8yDAo7vCmoVss8',
  PROJET: 'fldLSOiFzusnmTCQb',
  FOLDER_PATH: 'fld6alxyhIiSPoc6S',
  FOLDER_ID: 'fldIXbWRi092LnVLA',
  SCOPE: 'fldOjR7K8efkElKjQ',
} as const;

export interface FolderMapping {
  id: string;
  cle: string;
  userEmail: string;
  folderPath: string;
  folderId: string;
  scope: 'client' | 'projet';
}

/** Get folder mapping for a user + entity (client or projet) */
export async function getFolderMapping(
  userEmail: string, scope: 'client' | 'projet', entityId: string
): Promise<FolderMapping | null> {
  const cle = `${userEmail}|${entityId}`;
  const formula = encodeURIComponent(`{${FMF.CLE}} = "${cle}"`);
  const url = `${API_URL}/${PROJETS_BASE}/${FM_TABLE}?returnFieldsByFieldId=true&filterByFormula=${formula}&pageSize=1`;
  try {
    const data = await airtableFetch<{ records: any[] }>(url);
    if (data.records.length === 0) return null;
    const f = data.records[0].fields || {};
    return {
      id: data.records[0].id,
      cle: f[FMF.CLE] || '',
      userEmail: f[FMF.USER_EMAIL] || '',
      folderPath: f[FMF.FOLDER_PATH] || '',
      folderId: f[FMF.FOLDER_ID] || '',
      scope: f[FMF.SCOPE] || scope,
    };
  } catch { return null; }
}

/** Save or update a folder mapping */
export async function saveFolderMapping(
  userEmail: string, scope: 'client' | 'projet', entityId: string,
  folderPath: string, folderId: string
): Promise<void> {
  const cle = `${userEmail}|${entityId}`;
  const existing = await getFolderMapping(userEmail, scope, entityId);

  const fields: Record<string, unknown> = {
    [FMF.CLE]: cle,
    [FMF.USER_EMAIL]: userEmail,
    [FMF.FOLDER_PATH]: folderPath,
    [FMF.FOLDER_ID]: folderId,
    [FMF.SCOPE]: scope,
  };

  if (scope === 'client') fields[FMF.CLIENT] = [entityId];
  else fields[FMF.PROJET] = [entityId];

  if (existing) {
    await airtableFetch(`${API_URL}/${PROJETS_BASE}/${FM_TABLE}/${existing.id}`, {
      method: 'PATCH',
      body: JSON.stringify({ fields: { [FMF.FOLDER_PATH]: folderPath, [FMF.FOLDER_ID]: folderId }, returnFieldsByFieldId: true }),
    });
  } else {
    await airtableFetch(`${API_URL}/${PROJETS_BASE}/${FM_TABLE}`, {
      method: 'POST',
      body: JSON.stringify({ fields, returnFieldsByFieldId: true }),
    });
  }
}

// ── Create new Tiers (Client) ──

export async function createTiers(
  relation: string,
  email?: string,
): Promise<{ id: string; relation: string }> {
  const fields: Record<string, unknown> = {
    [CLF.RELATION]: relation,
    [CLF.STATUT]: 'Actif',
  };
  if (email) fields[CLF.EMAIL] = email;

  const result = await airtableFetch<{ id: string; fields: Record<string, unknown> }>(
    `${API_URL}/${PROJETS_BASE}/${TABLES.CLIENTS}`,
    {
      method: 'POST',
      body: JSON.stringify({ fields, typecast: true, returnFieldsByFieldId: true }),
    },
  );

  // Invalidate tiers cache
  cache.delete('tiers');
  return { id: result.id, relation };
}

// ── Create new Projet ──

export interface CreateProjetInput {
  denomination: string;
  clientRecordId: string;
  types?: string[];
  budget?: string;
  mois?: string;
  annee?: string;
  enChargeRecordId?: string;
  contactClientRecordId?: string;
  dateDebut?: string;
  descriptif?: string;
}

export async function createProjet(input: CreateProjetInput): Promise<{ id: string; noProjet: number }> {
  const fields: Record<string, unknown> = {
    [PF.DENOMINATION]: input.denomination,
    [PF.CLIENT]: [input.clientRecordId],
    [PF.STATUT]: 'Demande',
  };

  if (input.types?.length) fields[PCF.TYPE] = input.types;
  if (input.budget) fields[PCF.FOURCHETTE_BUDGET] = input.budget;
  if (input.mois) fields[PCF.MOIS_REALISATION] = input.mois;
  if (input.annee) fields[PCF.ANNEE_REALISATION] = input.annee;
  if (input.enChargeRecordId) fields[PF.EN_CHARGE] = [input.enChargeRecordId];
  if (input.contactClientRecordId) fields[PCF.CONTACT_CLIENTS] = [input.contactClientRecordId];
  if (input.dateDebut) fields[PF.DATE_DEBUT] = input.dateDebut;
  if (input.descriptif) fields[PCF.DESCRIPTIF] = input.descriptif;

  const result = await airtableFetch<{ id: string; fields: Record<string, unknown> }>(
    `${API_URL}/${PROJETS_BASE}/${TABLES.PROJETS}`,
    {
      method: 'POST',
      body: JSON.stringify({ fields, typecast: true, returnFieldsByFieldId: true }),
    },
  );

  // Invalidate projets cache
  cache.delete('projets');
  const noProjet = result.fields?.[PF.NO_PROJET] as number || 0;
  return { id: result.id, noProjet };
}

// ── Get Employés for "En charge" dropdown ──

export interface Employe {
  id: string;
  name: string;
  email: string;
}

export async function getAllEmployes(): Promise<Employe[]> {
  const cached = getCached<Employe[]>('employes');
  if (cached) return cached;

  const fields = [EMF.NAME, EMF.NOM_TXT, EMF.EMAIL].map(f => `fields%5B%5D=${f}`).join('&');
  const url = `${API_URL}/${PROJETS_BASE}/${EMPLOYES_TABLE}?returnFieldsByFieldId=true&${fields}`;
  const data = await airtableFetch<{ records: any[] }>(url);

  const employes = data.records.map((r: any) => {
    const f = r.fields || {};
    return {
      id: r.id,
      name: f[EMF.NOM_TXT] || f[EMF.NAME] || '',
      email: f[EMF.EMAIL] || '',
    };
  }).filter(e => e.name);

  setCache('employes', employes);
  return employes;
}

// ── Count linked emails for a project ──

export async function countLinkedEmails(projetRecordId: string): Promise<number> {
  try {
    // Use SEARCH on linked record field — Airtable stores linked records as comma-separated IDs
    const formula = encodeURIComponent(`SEARCH("${projetRecordId}", ARRAYJOIN(${EF.PROJET}, ","))`);
    const url = `${API_URL}/${PROJETS_BASE}/${TABLES.EMAILS_PROJET}?filterByFormula=${formula}&fields%5B%5D=${EF.SUJET}&returnFieldsByFieldId=true&pageSize=100`;
    const data = await airtableFetch<{ records: any[] }>(url);
    return data.records.length;
  } catch { return 0; }
}

// ── Email Tags (Inbound Scanner IA) ──────────────────────────────────────────

const ATLAS_EMAIL_TAGS_TABLE = 'tblZsl8roAPbydyYH';
const ETF = {
  EMAIL_ID:         'fldmsOwKA0G13CYtQ',
  CONVERSATION_ID:  'fld4iM2eL4ZYJzvIN',
  SUBJECT:          'fldmq2A4YvXaXJmyY',
  FROM_EMAIL:       'fld1dgBaT6m5RzwDI',
  CATEGORY:         'fldbX9OLwThezZHvT',
  URGENCY_SCORE:    'fldWmGpEMlwDUnwsE',
  SUMMARY:          'fldcXdX0YEbJqRDjQ',
  INBOX_STATUS:     'fldfYly6qkcNVyCht',
  SNOOZED_UNTIL:    'fldtJPsIRxbcKLx9f',
  ACTIONED_AT:      'fldb3sfQUmy2BngIw',
  ARCHIVED:         'fldHcDBghYdZl8FUE',
  USER_EMAIL:       'fldlMxGx7ovezklHH',
  LINKED_PROJET_ID: 'flddfFthABUkASkao',
} as const;

export interface EmailTag {
  id: string;
  emailId: string;
  category: string;        // EmailTagCategory
  urgencyScore: number;    // 1-5
  summary: string;
  inboxStatus: 'inbox' | 'done' | 'snoozed' | 'archived';
  linkedProjetId?: string; // Projet auquel le mail est rattaché (pour folder mapping)
}

function parseTagRecord(r: any): EmailTag {
  const f = r.fields || {};
  return {
    id: r.id,
    emailId: f[ETF.EMAIL_ID] || '',
    category: selectName(f[ETF.CATEGORY]) || 'autre',
    urgencyScore: Number(f[ETF.URGENCY_SCORE]) || 2,
    summary: f[ETF.SUMMARY] || '',
    inboxStatus: (f[ETF.INBOX_STATUS] || 'inbox') as EmailTag['inboxStatus'],
    linkedProjetId: f[ETF.LINKED_PROJET_ID] || undefined,
  };
}

/** Récupère le tag IA pour un email (par EmailId Outlook). Null si pas encore taggé. */
export async function getEmailTagByEmailId(emailId: string): Promise<EmailTag | null> {
  if (!emailId) return null;
  const formula = encodeURIComponent(`{${ETF.EMAIL_ID}} = "${emailId.replace(/"/g, '\\"')}"`);
  const url = `${API_URL}/${ATLAS_BASE}/${ATLAS_EMAIL_TAGS_TABLE}?returnFieldsByFieldId=true&filterByFormula=${formula}&pageSize=1`;
  try {
    const data = await airtableFetch<{ records: any[] }>(url);
    if (!data.records?.length) return null;
    return parseTagRecord(data.records[0]);
  } catch { return null; }
}

/**
 * Récupère le tag IA pour la conversation entière (fallback quand l'EmailId
 * ne match pas — typiquement à cause des différences EWS / Graph REST ID
 * entre l'addin Outlook et le scanner backend).
 *
 * Le conversationId est stable cross-client (Office.js, Graph, EWS), donc on
 * peut s'appuyer dessus pour retrouver le tag du dernier message de la conv.
 */
export async function getEmailTagByConversationId(conversationId: string): Promise<EmailTag | null> {
  if (!conversationId) return null;
  const formula = encodeURIComponent(`{${ETF.CONVERSATION_ID}} = "${conversationId.replace(/"/g, '\\"')}"`);
  const url = `${API_URL}/${ATLAS_BASE}/${ATLAS_EMAIL_TAGS_TABLE}?returnFieldsByFieldId=true&filterByFormula=${formula}&pageSize=10`;
  try {
    const data = await airtableFetch<{ records: any[] }>(url);
    if (!data.records?.length) return null;
    return parseTagRecord(data.records[0]);
  } catch { return null; }
}

/** Marque un tag comme "Traité" (Done) + actioned_at = now. */
export async function markTagDone(tagRecordId: string): Promise<boolean> {
  if (!tagRecordId) return false;
  const url = `${API_URL}/${ATLAS_BASE}/${ATLAS_EMAIL_TAGS_TABLE}/${tagRecordId}`;
  try {
    await airtableFetch(url, {
      method: 'PATCH',
      body: JSON.stringify({
        fields: { [ETF.INBOX_STATUS]: 'done', [ETF.ACTIONED_AT]: new Date().toISOString() },
        returnFieldsByFieldId: true,
      }),
    });
    return true;
  } catch (e) { console.warn('[addin] markTagDone failed:', e); return false; }
}

/** Snooze un tag jusqu'à `until` (par défaut : demain matin 8h). */
export async function snoozeTag(tagRecordId: string, until?: Date): Promise<boolean> {
  if (!tagRecordId) return false;
  const target = until || (() => { const d = new Date(); d.setDate(d.getDate() + 1); d.setHours(8, 0, 0, 0); return d; })();
  const url = `${API_URL}/${ATLAS_BASE}/${ATLAS_EMAIL_TAGS_TABLE}/${tagRecordId}`;
  try {
    await airtableFetch(url, {
      method: 'PATCH',
      body: JSON.stringify({
        fields: {
          [ETF.INBOX_STATUS]: 'snoozed',
          [ETF.SNOOZED_UNTIL]: target.toISOString(),
          [ETF.ACTIONED_AT]: new Date().toISOString(),
        },
        returnFieldsByFieldId: true,
      }),
    });
    return true;
  } catch (e) { console.warn('[addin] snoozeTag failed:', e); return false; }
}

/** Archive un tag + ARCHIVED=true (legacy). */
export async function archiveTag(tagRecordId: string): Promise<boolean> {
  if (!tagRecordId) return false;
  const url = `${API_URL}/${ATLAS_BASE}/${ATLAS_EMAIL_TAGS_TABLE}/${tagRecordId}`;
  try {
    await airtableFetch(url, {
      method: 'PATCH',
      body: JSON.stringify({
        fields: {
          [ETF.INBOX_STATUS]: 'archived',
          [ETF.ARCHIVED]: true,
          [ETF.ACTIONED_AT]: new Date().toISOString(),
        },
        returnFieldsByFieldId: true,
      }),
    });
    return true;
  } catch (e) { console.warn('[addin] archiveTag failed:', e); return false; }
}

/**
 * Crée ou remplace un tag email dans Airtable.
 * Utilisé par le bouton "Re-analyser" : on supprime l'ancien tag puis on
 * écrit le nouveau résultat Claude. Si pas d'ancien tag → simple création.
 */
export async function upsertEmailTag(input: {
  oldTagId?: string;
  emailId: string;
  conversationId: string;
  subject: string;
  fromEmail: string;
  fromName: string;
  receivedAt: string;
  category: string;
  urgencyScore: number;
  summary: string;
  detectedLanguage: string;
  userEmail: string;
}): Promise<{ id: string }> {
  // 1. DELETE ancien (best-effort)
  if (input.oldTagId) {
    try {
      await airtableFetch(`${API_URL}/${ATLAS_BASE}/${ATLAS_EMAIL_TAGS_TABLE}/${input.oldTagId}`, { method: 'DELETE' });
    } catch (e) { console.warn('[upsertEmailTag] DELETE old failed:', e); }
  }

  // 2. CREATE nouveau
  const fields = {
    [ETF.EMAIL_ID]: input.emailId,
    [ETF.CONVERSATION_ID]: input.conversationId,
    [ETF.SUBJECT]: input.subject,
    [ETF.FROM_EMAIL]: input.fromEmail,
    [ETF.CATEGORY]: input.category,
    [ETF.URGENCY_SCORE]: input.urgencyScore,
    [ETF.SUMMARY]: input.summary,
    [ETF.USER_EMAIL]: input.userEmail,
    [ETF.INBOX_STATUS]: 'inbox',
    [ETF.ACTIONED_AT]: new Date().toISOString(),
  };
  const res = await airtableFetch<{ id: string }>(`${API_URL}/${ATLAS_BASE}/${ATLAS_EMAIL_TAGS_TABLE}`, {
    method: 'POST',
    body: JSON.stringify({ fields, returnFieldsByFieldId: true, typecast: true }),
  });
  return { id: res.id };
}

/** Override la catégorie d'un tag (apprentissage). */
export async function correctTagCategory(tagRecordId: string, newCategory: string): Promise<boolean> {
  if (!tagRecordId || !newCategory) return false;
  const url = `${API_URL}/${ATLAS_BASE}/${ATLAS_EMAIL_TAGS_TABLE}/${tagRecordId}`;
  try {
    await airtableFetch(url, {
      method: 'PATCH',
      body: JSON.stringify({
        fields: { [ETF.CATEGORY]: newCategory },
        returnFieldsByFieldId: true,
        typecast: true,
      }),
    });
    return true;
  } catch (e) { console.warn('[addin] correctTagCategory failed:', e); return false; }
}
