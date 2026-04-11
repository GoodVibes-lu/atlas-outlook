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

// ── Communications Field IDs ──
const COMMS = {
  NOM: 'fldZHHGiNclP9x3vR',
  TYPE: 'fldGcXqLaP1w3ccQ2',
  TON: 'fldJlQlFiBMqspqif',
  MARQUE: 'fldfEDDbRqzc9qnEt',
  STATUT: 'flddJ48SsOXj3qJ8R',
  OBJET_FR: 'fld5rMGYjK6rnd3xC',
  CORPS_FR: 'fldH0tQ7SoVtDKsDe',
  OBJET_EN: 'fldEaEn5Ehma9PE00',
  CORPS_EN: 'fldbVEfyksg8k9UxP',
  VARIABLES: 'fldA9eojxIBUElgxk',
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
        commercial: selectName(f[PF.EN_CHARGE]),
        chefDeProjet: '',
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
        relation: f['fldYyVtj5Rh5TDgOb'] || f['fldq4OWNWYiS7XUIg'] || '', // Relation or Nom
        categorie: selectName(f['fldQZFM9TuuIEvcgD'] || ''),               // Secteur
        email: f['fldCNqiCExXOpLQuI'] || '',                                // Email
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
        personneDeContact: f['fldDFZzW5cris8sva'] || '',  // Personne de contact
        email: f['fldzZSIPtUkkKJtVi'] || '',               // Email
        relationSociete: f['fldlSSa8wDIQvoSey'] || '',     // Société
        fonction: f['fldHiZeUeyv4913jv'] || '',             // Fonction
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
        statut: selectName(f[COMMS.STATUT]),
        sujetFR: f[COMMS.OBJET_FR] || '',
        corpsFR: f[COMMS.CORPS_FR] || '',
        sujetEN: f[COMMS.OBJET_EN] || '',
        corpsEN: f[COMMS.CORPS_EN] || '',
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
