/**
 * agent-equipe.ts — Bloc « Équipe » du panneau Agent pour un mail de good@ (phase 5 de l'agent
 * d'inbox, backlog `.claude/BACKLOG-AGENT-INBOX.md` §4) :
 *
 *  • assigné à / pris par (« Pris par X à HH:MM », même principe que « Je prends » de TACTIK) ;
 *  • boutons « Je prends », « Relâcher » (la personne qui a pris), « Attribuer à… » (membres de good@) ;
 *  • avertissement de COLLISION (quelqu'un d'autre répond déjà : jamais bloquant) ;
 *  • délai de réponse (échéance) et escalades déjà parties ;
 *  • COMMENTAIRES INTERNES (stockés par le worker, jamais dans le mail, invisibles du client),
 *    visibles des seuls membres autorisés de la boîte partagée.
 *
 * Données : GET /api/plugin/agent/equipe (404 / 403 → bloc masqué : mail hors good@ ou personne non
 * membre), POST | DELETE …/equipe/prendre, POST …/equipe/attribuer, POST …/equipe/commentaires.
 * Office.js : AUCUNE API ici (mobile compris) ; tout passe par le worker. Aucun appel IA.
 */

import { fetchEquipe, prendreMail, relacherMail, attribuerMail, ajouterCommentaire } from '../api/agent';
import type { InboxEquipe, InboxEquipeVue } from '../api/inbox-agent.types';
import { escapeHtml } from './agent-lists';

type InfoFn = (message: string, type?: 'success' | 'error' | 'info') => void;

export interface EquipeOptions {
  messageId: string;
  /** Boîte partagée si elle est connue (état du mail) ; sinon le worker la cherche. */
  mailbox?: string;
  onInfo?: InfoFn;
  /** Après une prise / attribution (compteurs du bandeau à relire). */
  onChange?: () => void;
}

const heure = (iso?: string) => {
  if (!iso) return '';
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return '';
  return d.toLocaleString('fr-FR', { timeZone: 'Europe/Luxembourg', day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' });
};
const prenom = (email?: string) => {
  const l = String(email || '').split('@')[0].split(/[._+]/)[0] || '';
  return l ? l[0].toUpperCase() + l.slice(1) : '';
};

function etatHtml(v: InboxEquipeVue): string {
  const e: InboxEquipe = v.equipe;
  const moi = v.moi;
  const lignes: string[] = [];
  if (e.prisPar) {
    lignes.push(`<div class="equipe-ligne equipe-pris"><strong>${e.prisPar === moi ? 'Pris par toi' : `Pris par ${escapeHtml(e.prisParNom || prenom(e.prisPar))}`}</strong>${e.prisLe ? ` <span class="agent-muted">· ${escapeHtml(heure(e.prisLe))}</span>` : ''}</div>`);
  }
  if (e.assigneA) {
    lignes.push(`<div class="equipe-ligne">${e.assigneA === moi ? 'Pour toi' : `Pour ${escapeHtml(e.assigneNom || prenom(e.assigneA))}`}${e.motif ? ` <span class="agent-muted">· ${escapeHtml(e.motif)}</span>` : ''}</div>`);
  } else if (!e.prisPar) {
    lignes.push(`<div class="equipe-ligne equipe-a-attribuer"><strong>À attribuer</strong>${e.motif && e.motif !== 'À attribuer' ? ` <span class="agent-muted">· ${escapeHtml(e.motif.replace(/^À attribuer : /, ''))}</span>` : ''}</div>`);
  }
  if (e.escalade?.echeance && !e.prisPar) {
    const depasse = Date.parse(e.escalade.echeance) < Date.now();
    lignes.push(`<div class="agent-muted">Réponse attendue ${depasse ? 'depuis' : 'avant'} ${escapeHtml(heure(e.escalade.echeance))} (${e.escalade.delaiHeures} h ouvrées)</div>`);
  }
  for (const n of e.escalade?.niveaux || []) {
    lignes.push(`<div class="agent-muted">Escalade ${n.niveau === 1 ? 'au responsable' : 'à la direction'} ${escapeHtml(heure(n.at))}${n.simule ? ' (simulée : agent en observation)' : n.notifie ? '' : ' (en attente de ses horaires)'}</div>`);
  }
  return lignes.join('');
}

function boutonsHtml(v: InboxEquipeVue): string {
  const e = v.equipe;
  const autres = v.membres.filter(m => m.email !== e.assigneA || !!e.prisPar);
  return `
    <div class="equipe-btns">
      ${!e.prisPar ? '<button type="button" class="btn btn-primary agent-btn" data-equipe="prendre">Je prends</button>' : ''}
      ${e.prisPar === v.moi ? '<button type="button" class="btn btn-secondary agent-btn" data-equipe="relacher">Relâcher</button>' : ''}
      <button type="button" class="btn btn-secondary agent-btn" data-equipe="attribuer-toggle" aria-expanded="false">Attribuer à…</button>
    </div>
    <div class="equipe-attribuer" hidden>
      <select class="equipe-select" aria-label="Membre de la boîte partagée">
        ${autres.map(m => `<option value="${escapeHtml(m.email)}">${escapeHtml(m.nom || m.email)}${m.email === v.moi ? ' (moi)' : ''}</option>`).join('')}
      </select>
      <button type="button" class="btn btn-primary agent-btn" data-equipe="attribuer">Attribuer</button>
    </div>`;
}

function commentairesHtml(v: InboxEquipeVue): string {
  const cs = v.equipe.commentaires || [];
  return `
    <div class="equipe-commentaires">
      <div class="agent-section-subtitle">Commentaires internes <span class="agent-muted">· invisibles du client</span></div>
      ${cs.length ? `<ul class="equipe-comment-list">${cs.slice(-20).map(c => `
        <li><span class="equipe-comment-auteur">${escapeHtml(c.auteurNom || prenom(c.auteur))}</span> <span class="agent-muted">${escapeHtml(heure(c.at))}</span><div class="equipe-comment-texte">${escapeHtml(c.texte)}</div></li>`).join('')}</ul>`
        : '<p class="agent-muted">Aucun commentaire.</p>'}
      <textarea class="equipe-comment-saisie" rows="2" maxlength="1000" placeholder="Note pour l'équipe (jamais envoyée au client)"></textarea>
      <button type="button" class="btn btn-secondary btn-block agent-btn" data-equipe="commenter">Ajouter le commentaire</button>
    </div>`;
}

/** Affiche le bloc « Équipe » dans `host` (masqué si le mail n'est pas dans good@ ou si la personne n'est pas membre). */
export async function renderEquipe(host: HTMLElement, opts: EquipeOptions): Promise<void> {
  let v: InboxEquipeVue | null;
  try {
    v = await fetchEquipe(opts.messageId, opts.mailbox);
  } catch (e) {
    host.hidden = false;
    host.innerHTML = `<div class="agent-section-title">Équipe</div><p class="agent-error" role="alert">Équipe indisponible : ${escapeHtml((e as Error).message)}</p>`;
    return;
  }
  if (!host.isConnected) return;
  if (!v) { host.hidden = true; host.innerHTML = ''; return; }
  afficher(host, v, opts);
}

function afficher(host: HTMLElement, v: InboxEquipeVue, opts: EquipeOptions): void {
  const col = v.collision;
  host.hidden = false;
  host.innerHTML = `
    <div class="agent-section-title">Équipe · ${escapeHtml(v.mailbox)}</div>
    ${col ? `<div class="equipe-collision ${col.niveau === 'avertissement' ? 'is-warning' : 'is-info'}" role="${col.niveau === 'avertissement' ? 'alert' : 'status'}">${col.niveau === 'avertissement' ? '⚠️ ' : ''}${escapeHtml(col.message)}</div>` : ''}
    ${etatHtml(v)}
    ${boutonsHtml(v)}
    ${commentairesHtml(v)}
    ${v.mode !== 'actif' ? '<div class="agent-muted">Agent en observation : aucune catégorie « ATLAS · Pour … » n\'est posée dans la boîte ; la prise et les commentaires sont partagés.</div>' : ''}
  `;
  const btn = (k: string) => host.querySelector<HTMLButtonElement>(`button[data-equipe="${k}"]`);
  const occupe = (on: boolean) => host.querySelectorAll<HTMLButtonElement>('button').forEach(b => { b.disabled = on; });
  const recharger = async (message?: string) => {
    if (message) opts.onInfo?.(message, 'success');
    opts.onChange?.();
    const nv = await fetchEquipe(v.messageId, v.mailbox).catch(() => null);
    if (nv && host.isConnected) afficher(host, nv, opts);
  };
  const agir = async (f: () => Promise<{ ok: boolean; equipe?: InboxEquipe; error?: string }>, succes: string) => {
    occupe(true);
    try {
      const r = await f();
      if (!r.ok) opts.onInfo?.(r.error || 'Refusé', 'info');
      await recharger(r.ok ? succes : undefined);
    } catch (e) {
      occupe(false);
      opts.onInfo?.(`Erreur : ${(e as Error).message}`, 'error');
    }
  };
  btn('prendre')?.addEventListener('click', () => agir(() => prendreMail(v.messageId, v.mailbox), 'Tu as pris ce mail : l\'équipe le voit'));
  btn('relacher')?.addEventListener('click', () => agir(() => relacherMail(v.messageId, v.mailbox), 'Mail relâché'));
  btn('attribuer-toggle')?.addEventListener('click', () => {
    const zone = host.querySelector<HTMLElement>('.equipe-attribuer');
    if (!zone) return;
    zone.hidden = !zone.hidden;
    btn('attribuer-toggle')?.setAttribute('aria-expanded', String(!zone.hidden));
  });
  btn('attribuer')?.addEventListener('click', () => {
    const a = host.querySelector<HTMLSelectElement>('.equipe-select')?.value || '';
    if (!a) return;
    agir(() => attribuerMail(v.messageId, v.mailbox, a), `Mail attribué à ${prenom(a)}`);
  });
  btn('commenter')?.addEventListener('click', async () => {
    const zone = host.querySelector<HTMLTextAreaElement>('.equipe-comment-saisie');
    const texte = (zone?.value || '').trim();
    if (!texte) { opts.onInfo?.('Écris d\'abord le commentaire.', 'info'); return; }
    occupe(true);
    try {
      await ajouterCommentaire(v.messageId, v.mailbox, texte);
      await recharger('Commentaire interne ajouté');
    } catch (e) {
      occupe(false);
      opts.onInfo?.(`Erreur : ${(e as Error).message}`, 'error');
    }
  });
}
