/**
 * agent-parite.ts · Fin de la PARITÉ avec l'Inbox ATLAS dans Outlook (demande de Charles du 07/10/2026),
 * volet du mail ET tableau de bord (mêmes cartes) :
 *
 *  • « Ce n'est pas X qui a répondu : rattacher Y à la demande » après le dépôt d'une offre (annulable) ;
 *  • « Importer les mails de ce dossier dans le projet » (par lots, progression, jamais de doublon) ;
 *  • « Reclasser vers un autre projet » : lié à B, délié de A, mail déplacé s'il était dans le dossier
 *    de A (dossier de B, ou créé au nommage ATLAS), « Annuler » remet tout ;
 *  • « Pièces jointes → projet » : champ fichier du projet choisi comme dans l'Inbox (Documents, Devis
 *    signé…), en plus du dépôt sur le NAS (action « Pièces jointes » de l'agent) ;
 *  • « Proposer un RDV » sur n'importe quel mail : créneaux du mail et de l'agenda, lieu, durée, texte
 *    repris dans la réponse (rien n'est envoyé) ;
 *  • « Créer un tiers » : formulaire complet d'ATLAS (doublons, code BOB, contact principal) ;
 *  • « Déplacer vers Prospection » : rattaché au contact de l'expéditeur, suivi de prospection, puis
 *    rangé dans « Prospection » (créé s'il manque), annulable.
 * Écritures ATLAS par le worker (services d'ATLAS) ; boîte : même rangement que « Classer dans Outlook »
 * (double verrou, repli jeton de la personne). Erreurs toujours affichées.
 */
import {
  fetchEcartContactDevis, rattacherContactDevis, annulerContactDevis, importerLotDossier, fetchProjetDuMail, reclasserMail,
  fetchPiecesProjet, joindreAuProjet, fetchCreneauxRdv, propositionRdv, preparerTiers, creerTiers, lierProspection, donneesErreur,
  LIBELLES_LIEUX, CHEMIN_PROSPECTION, type LieuRdv, type SaisieTiersOutlook, type PreparationTiers,
} from '../api/parite';
import { annulerAction, estVerrouFerme, type DossierOutlook } from '../api/agent';
import { getAllProjets } from '../api/airtable';
import { rangerMail, renderDossierProjet, type CtxDossiers } from './agent-dossiers';
import { escapeHtml } from '../utils/html';
import { humanError } from '../api/net';
import { icon } from '../ui/icons';

const erreur = (quoi: string, e: unknown) => `<p class="agent-error" role="alert">${escapeHtml(quoi)} : ${escapeHtml(humanError(e))}</p>`;
const chargement = (t: string) => `<div class="agent-loading"><div class="spinner"></div><span>${escapeHtml(t)}</span></div>`;
const verrouHtml = '<p class="agent-error" role="alert">L\'agent n\'écrit pas encore dans cette boîte (double verrou fermé) : rien n\'a été déplacé.</p>';

/** Projets actifs (liste ATLAS) en options de <select>. */
async function optionsProjets(selectionne?: string | null, exclu?: string): Promise<string> {
  const tous = (await getAllProjets()).filter(p => !/clôtur|clotur|annul/i.test(p.statut || '') && p.id !== exclu);
  return `<option value="">Choisir un projet…</option>${tous.map(p => `<option value="${escapeHtml(p.id)}" ${p.id === selectionne ? 'selected' : ''}>${escapeHtml(`#${p.noProjet || p.refProjet} ${p.denomination}${p.client ? ` · ${p.client}` : ''}`)}</option>`).join('')}`;
}

// ── 1. Contact de la demande de devis ───────────────────────────────────────────

/** Après le dépôt d'une offre : le répondant n'est pas le contact de la demande ? (rien si identique) */
export async function renderContactDevis(host: HTMLElement, ctx: CtxDossiers & { devisId: string }): Promise<void> {
  host.hidden = false;
  host.innerHTML = chargement('Contact de la demande…');
  try {
    const e = await fetchEcartContactDevis(ctx.messageId, ctx.mailbox || undefined, ctx.devisId);
    if (!host.isConnected) return;
    if (!e) { host.hidden = true; host.innerHTML = ''; return; }
    if (!e.newContact) {
      host.innerHTML = `<div class="agent-carte"><div class="agent-carte-apercu">${icon('info', 14)} La réponse vient de ${escapeHtml(e.senderName || e.senderEmail)}, pas du contact de la demande (${escapeHtml(e.originalEmail)}). Aucune fiche contact trouvée pour ${escapeHtml(e.senderEmail)} : crée-la dans ATLAS pour la rattacher.</div></div>`;
      return;
    }
    const c = e.newContact;
    host.innerHTML = `<div class="agent-carte"><div class="agent-carte-titre">${icon('user', 14)} Ce n'est pas ${escapeHtml(e.originalEmail)} qui a répondu</div>
      <div class="agent-carte-apercu">Rattacher ${escapeHtml(c.name)}${c.societe ? ` (${escapeHtml(c.societe)})` : ''} à la demande de devis ? Le remerciement partira au bon contact.</div>
      <div class="btn-row"><button type="button" class="btn btn-primary agent-btn" data-oui>Rattacher ${escapeHtml(c.name)}</button><button type="button" class="btn btn-secondary agent-btn" data-non>Garder le contact actuel</button></div>
      <div data-res hidden></div></div>`;
    const res = host.querySelector<HTMLElement>('[data-res]')!;
    host.querySelector<HTMLButtonElement>('[data-non]')!.addEventListener('click', () => { host.hidden = true; host.innerHTML = ''; });
    host.querySelector<HTMLButtonElement>('[data-oui]')!.addEventListener('click', async ev => {
      const b = ev.currentTarget as HTMLButtonElement;
      b.disabled = true;
      try {
        await rattacherContactDevis(ctx.messageId, ctx.mailbox || undefined, ctx.devisId, c.id);
        host.innerHTML = `<div class="agent-fait" role="status"><span>${icon('check-circle', 14)} Contact de la demande : ${escapeHtml(c.name)}</span><span class="agent-fait-btns"><button type="button" class="btn btn-secondary agent-btn" data-annuler>Annuler</button></span></div>`;
        ctx.onInfo?.(`Contact mis à jour : ${c.name}`, 'success');
        host.querySelector<HTMLButtonElement>('[data-annuler]')!.addEventListener('click', async ev2 => {
          const a = ev2.currentTarget as HTMLButtonElement;
          a.disabled = true;
          try { await annulerContactDevis(ctx.messageId, ctx.mailbox || undefined, ctx.devisId); a.replaceWith(Object.assign(document.createElement('span'), { className: 'agent-muted', textContent: 'Annulé : contact d\'origine remis' })); }
          catch (err) { a.disabled = false; ctx.onInfo?.(`Annulation impossible : ${humanError(err)}`, 'error'); }
        });
      } catch (err) { b.disabled = false; res.hidden = false; res.innerHTML = erreur('Rattachement impossible', err); }
    });
  } catch (e) { if (host.isConnected) host.innerHTML = erreur('Contact de la demande indisponible', e); }
}

// ── 2. Import des mails d'un dossier dans le projet ─────────────────────────────

export function renderImportDossier(host: HTMLElement, ctx: CtxDossiers & { projetId: string; dossier: DossierOutlook }): void {
  host.hidden = false;
  host.innerHTML = `<button type="button" class="btn btn-secondary btn-block agent-btn" data-importer>${icon('inbox', 14)}Importer les mails de ce dossier dans le projet</button>
    <div data-progres class="agent-muted" role="status" aria-live="polite" hidden></div>`;
  const b = host.querySelector<HTMLButtonElement>('[data-importer]')!;
  const prog = host.querySelector<HTMLElement>('[data-progres]')!;
  b.addEventListener('click', async () => {
    b.disabled = true;
    prog.hidden = false;
    let curseur: number | null = 0, importes = 0, deja = 0;
    try {
      while (curseur !== null) {
        prog.textContent = `Import en cours… ${importes} importé(s), ${deja} déjà lié(s)`;
        const r = await importerLotDossier({ mailbox: ctx.mailbox || undefined, projetId: ctx.projetId, dossierId: ctx.dossier.id, curseur });
        importes += r.importes; deja += r.deja;
        prog.textContent = `${r.traites}${r.total ? ` / ${r.total}` : ''} mails lus · ${importes} importé(s), ${deja} déjà lié(s)`;
        curseur = r.suivant;
      }
      prog.textContent = `Terminé : ${importes} mail${importes > 1 ? 's' : ''} importé${importes > 1 ? 's' : ''} dans le projet${deja ? `, ${deja} déjà lié${deja > 1 ? 's' : ''}` : ''}.`;
      ctx.onInfo?.(`${importes} mail(s) importé(s) dans le projet`, 'success');
      b.textContent = 'Réimporter (sans doublon)';
    } catch (e) {
      prog.innerHTML = `${escapeHtml(`${importes} importé(s) avant l'arrêt.`)} ${erreur('Import interrompu', e)}`;
      ctx.onInfo?.('Import interrompu : relance-le, rien ne sera importé deux fois', 'error');
    } finally { b.disabled = false; }
  });
}

// ── 3. Reclasser vers un autre projet ───────────────────────────────────────────

export async function renderReclasser(host: HTMLElement, ctx: CtxDossiers): Promise<void> {
  host.hidden = false;
  host.innerHTML = chargement('Projet du mail…');
  let actuel: { id: string; libelle: string } | null = null;
  try { actuel = await fetchProjetDuMail(ctx.messageId, ctx.mailbox || undefined); }
  catch (e) { if (host.isConnected) host.innerHTML = erreur('Projet du mail indisponible', e); return; }
  if (!host.isConnected) return;
  if (!actuel) { host.innerHTML = '<p class="agent-muted">Ce mail n\'est rattaché à aucun projet : utilise « Lier » (ou « Rattacher à un projet ») pour le rattacher.</p>'; return; }
  const de = actuel;
  host.innerHTML = `<div class="agent-carte"><div class="agent-carte-apercu">Rattaché à <strong>${escapeHtml(de.libelle)}</strong></div>
    <label class="agent-muted" for="rc-projet">Reclasser vers</label>
    <select class="agent-select" id="rc-projet"><option value="">Chargement des projets…</option></select>
    <button type="button" class="btn btn-primary btn-block agent-btn" data-go disabled>Reclasser</button>
    <div data-res hidden></div></div>`;
  const sel = host.querySelector<HTMLSelectElement>('#rc-projet')!;
  const go = host.querySelector<HTMLButtonElement>('[data-go]')!;
  const res = host.querySelector<HTMLElement>('[data-res]')!;
  try { sel.innerHTML = await optionsProjets(null, de.id); } catch (e) { res.hidden = false; res.innerHTML = erreur('Projets indisponibles', e); }
  sel.addEventListener('change', () => { go.disabled = !sel.value; });
  go.addEventListener('click', async () => {
    const vers = sel.value;
    const versLib = sel.selectedOptions[0]?.textContent || vers;
    if (!vers) return;
    go.disabled = true;
    res.hidden = true;
    try {
      const r = await reclasserMail({ messageId: ctx.messageId, mailbox: ctx.mailbox || undefined, deProjetId: de.id, versProjetId: vers });
      let rangement: { chemin: string; actionId?: string } | null = null;
      let note = '';
      if (r.deplacer === 'ranger' && r.dossierVers) {
        try { rangement = await rangerMail(ctx, { dossier: r.dossierVers, projetId: vers }); }
        catch (e) { note = estVerrouFerme(e) ? 'Mail laissé dans le dossier de l\'ancien projet (double verrou fermé).' : `Mail non déplacé : ${humanError(e)}`; }
      }
      host.innerHTML = `<div class="agent-fait" role="status"><span>${icon('check-circle', 14)} Reclassé vers ${escapeHtml(versLib)}${r.delies ? ` (délié de ${escapeHtml(de.libelle)})` : ''}${rangement ? ` · rangé dans « ${escapeHtml(rangement.chemin)} »` : ''}</span>
        ${note ? `<span class="agent-muted">${escapeHtml(note)}</span>` : ''}
        <span class="agent-fait-btns"><button type="button" class="btn btn-secondary agent-btn" data-annuler>Annuler</button></span></div><div data-dossier hidden></div>`;
      ctx.onInfo?.(`Reclassé vers ${versLib}`, 'success');
      // Mail dans le dossier de A et B sans dossier : « Créer le dossier et ranger » au nommage ATLAS.
      if (r.deplacer === 'creer') void renderDossierProjet(host.querySelector<HTMLElement>('[data-dossier]')!, { ...ctx, projetId: vers, propose: r.dossierACreer || '' });
      host.querySelector<HTMLButtonElement>('[data-annuler]')!.addEventListener('click', async ev => {
        const a = ev.currentTarget as HTMLButtonElement;
        a.disabled = true;
        try {
          if (rangement?.actionId) await annulerAction(rangement.actionId);
          await reclasserMail({ messageId: ctx.messageId, mailbox: ctx.mailbox || undefined, deProjetId: vers, versProjetId: de.id, deplacer: false });
          a.replaceWith(Object.assign(document.createElement('span'), { className: 'agent-muted', textContent: `Annulé : de nouveau rattaché à ${de.libelle}` }));
          ctx.onInfo?.('Reclassement annulé', 'success');
        } catch (e) { a.disabled = false; ctx.onInfo?.(`Annulation impossible : ${humanError(e)}`, 'error'); }
      });
    } catch (e) { go.disabled = false; res.hidden = false; res.innerHTML = erreur('Reclassement impossible', e); }
  });
}

// ── 4. Pièces jointes → champ du projet ─────────────────────────────────────────

export async function renderPiecesProjet(host: HTMLElement, ctx: CtxDossiers): Promise<void> {
  host.hidden = false;
  host.innerHTML = chargement('Pièces jointes…');
  try {
    const c = await fetchPiecesProjet(ctx.messageId, ctx.mailbox || undefined);
    if (!host.isConnected) return;
    if (!c.pieces.length) { host.innerHTML = '<p class="agent-muted">Ce mail n\'a pas de pièce jointe.</p>'; return; }
    host.innerHTML = `<div class="agent-carte">
      <label class="agent-muted" for="pp-projet">Projet</label><select class="agent-select" id="pp-projet"><option value="">Chargement…</option></select>
      <label class="agent-muted" for="pp-champ">Champ du projet</label>
      <select class="agent-select" id="pp-champ">${c.destinations.map(d => `<option value="${escapeHtml(d.cle)}">${escapeHtml(d.libelle)}</option>`).join('')}</select>
      <div class="agent-muted">Pièces jointes</div>
      ${c.pieces.map((p, i) => `<label class="agent-check"><input type="checkbox" data-pj="${i}" checked/> ${escapeHtml(p.nom)}</label>`).join('')}
      <button type="button" class="btn btn-primary btn-block agent-btn" data-go>Attacher au projet</button>
      <div data-res hidden></div></div>`;
    const sel = host.querySelector<HTMLSelectElement>('#pp-projet')!;
    const res = host.querySelector<HTMLElement>('[data-res]')!;
    try { sel.innerHTML = await optionsProjets(c.projetId); } catch (e) { res.hidden = false; res.innerHTML = erreur('Projets indisponibles', e); }
    host.querySelector<HTMLButtonElement>('[data-go]')!.addEventListener('click', async ev => {
      const b = ev.currentTarget as HTMLButtonElement;
      const ids = [...host.querySelectorAll<HTMLInputElement>('[data-pj]')].filter(x => x.checked).map(x => c.pieces[Number(x.dataset.pj)]?.id).filter(Boolean) as string[];
      if (!sel.value || !ids.length) { res.hidden = false; res.innerHTML = '<p class="agent-error" role="alert">Choisis le projet et au moins une pièce jointe.</p>'; return; }
      b.disabled = true; b.textContent = 'Envoi…';
      try {
        const r = await joindreAuProjet({ messageId: ctx.messageId, mailbox: ctx.mailbox || undefined, projetId: sel.value, champ: host.querySelector<HTMLSelectElement>('#pp-champ')!.value, pieceJointeIds: ids });
        host.innerHTML = `<div class="agent-fait" role="status"><span>${icon('check-circle', 14)} ${r.joints} fichier(s) → ${escapeHtml(r.champ)}</span>${r.refus.length ? `<span class="agent-muted">${r.refus.length} pièce(s) non ajoutée(s) : ${escapeHtml(r.refus.map(x => x.motif).join(' ; '))}</span>` : ''}</div>`;
        ctx.onInfo?.(`${r.joints} fichier(s) → ${r.champ}`, 'success');
      } catch (e) { b.disabled = false; b.textContent = 'Attacher au projet'; res.hidden = false; res.innerHTML = erreur('Envoi impossible', e); }
    });
  } catch (e) { if (host.isConnected) host.innerHTML = erreur('Pièces jointes indisponibles', e); }
}

// ── 5. Proposer un RDV ──────────────────────────────────────────────────────────

/**
 * `repondre(html)` : réponse préremplie (volet du mail) ; `deposer(texte)` : brouillon de réponse
 * (tableau de bord). Rien n'est envoyé.
 */
export function renderRdv(host: HTMLElement, ctx: CtxDossiers & { repondre?: (html: string) => void; deposer?: (texte: string) => Promise<void> }): void {
  host.hidden = false;
  let lieu: LieuRdv = 'teams';
  let duree = 60;
  host.innerHTML = `<div class="agent-carte">
    <label class="agent-muted" for="rdv-lieu">Lieu</label>
    <select class="agent-select" id="rdv-lieu">${(Object.keys(LIBELLES_LIEUX) as LieuRdv[]).map(k => `<option value="${k}" ${k === lieu ? 'selected' : ''}>${escapeHtml(LIBELLES_LIEUX[k])}</option>`).join('')}</select>
    <label class="agent-muted" for="rdv-duree">Durée</label>
    <select class="agent-select" id="rdv-duree">${[30, 45, 60, 90].map(d => `<option value="${d}" ${d === duree ? 'selected' : ''}>${d} min</option>`).join('')}</select>
    <div data-creneaux>${chargement('Créneaux libres de l\'agenda…')}</div>
    <button type="button" class="btn btn-primary btn-block agent-btn" data-go disabled>${icon('calendar-plus', 14)}${ctx.repondre ? 'Proposer dans la réponse' : 'Déposer la proposition en brouillon'}</button>
    <div data-res hidden></div></div>`;
  const zone = host.querySelector<HTMLElement>('[data-creneaux]')!;
  const go = host.querySelector<HTMLButtonElement>('[data-go]')!;
  const res = host.querySelector<HTMLElement>('[data-res]')!;
  let creneaux: Array<{ debut: number; libelle: string; propose: boolean; conflit: boolean }> = [];
  const charger = async () => {
    zone.innerHTML = chargement('Créneaux libres de l\'agenda…');
    go.disabled = true;
    try {
      const r = await fetchCreneauxRdv(ctx.messageId, ctx.mailbox || undefined, duree);
      creneaux = r.creneaux;
      const libresDuMail = creneaux.filter(c => c.propose && !c.conflit);
      zone.innerHTML = creneaux.length
        ? `${r.agendaLu ? '' : '<p class="agent-muted">Agenda illisible : créneaux proposés sans vérification.</p>'}${creneaux.map((c, i) => `<label class="agent-check"><input type="checkbox" data-c="${i}" ${(libresDuMail.length ? libresDuMail.includes(c) : !c.conflit) ? 'checked' : ''}/> ${escapeHtml(c.libelle)}${c.propose ? ' <span class="agent-muted">(proposé dans le mail)</span>' : ''}${c.conflit ? ' <span class="agent-muted">(en conflit avec l\'agenda)</span>' : ''}</label>`).join('')}`
        : '<p class="agent-muted">Aucun créneau trouvé : la proposition demandera les disponibilités.</p>';
      go.disabled = false;
    } catch (e) { zone.innerHTML = erreur('Créneaux indisponibles', e); go.disabled = false; }
  };
  host.querySelector<HTMLSelectElement>('#rdv-lieu')!.addEventListener('change', ev => { lieu = (ev.target as HTMLSelectElement).value as LieuRdv; });
  host.querySelector<HTMLSelectElement>('#rdv-duree')!.addEventListener('change', ev => { duree = Number((ev.target as HTMLSelectElement).value) || 60; void charger(); });
  go.addEventListener('click', async () => {
    const choisis = [...zone.querySelectorAll<HTMLInputElement>('[data-c]')].filter(x => x.checked).map(x => creneaux[Number(x.dataset.c)]?.debut).filter((x): x is number => typeof x === 'number');
    go.disabled = true;
    res.hidden = true;
    try {
      const p = await propositionRdv({ messageId: ctx.messageId, mailbox: ctx.mailbox || undefined, lieu, duree, creneaux: choisis });
      if (ctx.repondre) { ctx.repondre(p.html); res.hidden = false; res.innerHTML = '<p class="agent-muted">Réponse ouverte avec la proposition : relis puis envoie.</p>'; }
      else if (ctx.deposer) { await ctx.deposer(p.texte); }
      else { res.hidden = false; res.innerHTML = `<textarea class="agent-input" rows="6" readonly>${escapeHtml(p.texte)}</textarea>`; }
    } catch (e) { res.hidden = false; res.innerHTML = erreur('Proposition impossible', e); }
    finally { go.disabled = false; }
  });
  void charger();
}

// ── 6. Créer un tiers ───────────────────────────────────────────────────────────

const champ = (id: string, libelle: string, valeur = '', type = 'text') =>
  `<label class="agent-muted" for="${id}">${escapeHtml(libelle)}</label><input type="${type}" class="agent-input" id="${id}" value="${escapeHtml(valeur)}" maxlength="200">`;

export function renderCreerTiers(host: HTMLElement, ctx: CtxDossiers & { prefill?: { nom?: string; email?: string; contactPrenom?: string; contactNom?: string; contactEmail?: string } }): void {
  host.hidden = false;
  host.innerHTML = chargement('Formulaire « Nouveau tiers »…');
  const vide: SaisieTiersOutlook = { nom: '', categories: [], pays: 'Luxembourg' };
  void preparerTiers(vide).then(p0 => {
    if (!host.isConnected) return;
    const o = p0.options;
    const pf = ctx.prefill || {};
    host.innerHTML = `<div class="agent-carte">
      ${champ('nt-nom', 'Nom (société ou nom de famille)', pf.nom || '')}
      <label class="agent-check"><input type="checkbox" id="nt-part"/> Particulier ?</label>
      <div data-prenom hidden>${champ('nt-prenom', 'Prénom')}</div>
      <div class="agent-muted">Catégories</div>
      ${o.categories.map(c => `<label class="agent-check"><input type="checkbox" data-cat value="${escapeHtml(c)}"/> ${escapeHtml(c)}</label>`).join('')}
      ${champ('nt-pays', 'Pays', 'Luxembourg')}
      ${champ('nt-email', 'E-mail', pf.email || '', 'email')}${champ('nt-tel', 'Téléphone')}${champ('nt-tva', 'N° de TVA')}
      ${champ('nt-adresse', 'Adresse')}${champ('nt-cp', 'Code postal')}${champ('nt-ville', 'Localité')}${champ('nt-web', 'Site web')}${champ('nt-matricule', 'Matricule')}
      <label class="agent-muted" for="nt-nb">Nombre d'employés</label><select class="agent-select" id="nt-nb"><option value="">Non renseigné</option>${o.nbEmployes.map(n => `<option>${escapeHtml(n)}</option>`).join('')}</select>
      <div data-contact>
        <div class="agent-section-subtitle">Contact principal (facultatif)</div>
        ${champ('nt-ct-prenom', 'Prénom', pf.contactPrenom || '')}${champ('nt-ct-nom', 'Nom', pf.contactNom || '')}
        <label class="agent-muted" for="nt-ct-genre">Genre</label><select class="agent-select" id="nt-ct-genre"><option value="">Non renseigné</option>${o.genres.map(g => `<option>${escapeHtml(g)}</option>`).join('')}</select>
        ${champ('nt-ct-fonction', 'Fonction')}
        <label class="agent-muted" for="nt-ct-langue">Langue</label><select class="agent-select" id="nt-ct-langue">${o.langues.map(l => `<option ${l === 'Français' ? 'selected' : ''}>${escapeHtml(l)}</option>`).join('')}</select>
        ${champ('nt-ct-email', 'E-mail', pf.contactEmail || '', 'email')}${champ('nt-ct-tel', 'Téléphone')}${champ('nt-ct-gsm', 'GSM')}
      </div>
      <button type="button" class="btn btn-secondary btn-block agent-btn" data-verifier>Vérifier (doublons, code BOB)</button>
      <div data-verif hidden></div>
      <div data-res hidden></div></div>`;
    const v = (id: string) => (host.querySelector<HTMLInputElement | HTMLSelectElement>(`#${id}`)?.value || '').trim();
    const part = host.querySelector<HTMLInputElement>('#nt-part')!;
    part.addEventListener('change', () => { host.querySelector<HTMLElement>('[data-prenom]')!.hidden = !part.checked; host.querySelector<HTMLElement>('[data-contact]')!.hidden = part.checked; });
    const saisie = (): SaisieTiersOutlook => ({
      nom: v('nt-nom'), prenom: v('nt-prenom'), particulier: part.checked,
      categories: [...host.querySelectorAll<HTMLInputElement>('[data-cat]')].filter(x => x.checked).map(x => x.value),
      pays: v('nt-pays'), email: v('nt-email'), tel: v('nt-tel'), tva: v('nt-tva'), adresse: v('nt-adresse'), cp: v('nt-cp'), ville: v('nt-ville'),
      web: v('nt-web'), matricule: v('nt-matricule'), nbEmployes: v('nt-nb'),
      contact: part.checked ? undefined : { prenom: v('nt-ct-prenom'), nom: v('nt-ct-nom'), genre: v('nt-ct-genre'), fonction: v('nt-ct-fonction'), langue: v('nt-ct-langue'), email: v('nt-ct-email'), tel: v('nt-ct-tel'), gsm: v('nt-ct-gsm') },
    });
    const verif = host.querySelector<HTMLElement>('[data-verif]')!;
    const res = host.querySelector<HTMLElement>('[data-res]')!;
    host.querySelector<HTMLButtonElement>('[data-verifier]')!.addEventListener('click', async ev => {
      const b = ev.currentTarget as HTMLButtonElement;
      b.disabled = true;
      verif.hidden = false; verif.innerHTML = chargement('Vérification…');
      try { afficherVerif(await preparerTiers(saisie())); }
      catch (e) { verif.innerHTML = erreur('Vérification impossible', e); }
      finally { b.disabled = false; }
    });
    const afficherVerif = (p: PreparationTiers) => {
      if (p.erreurs.length) { verif.innerHTML = `<ul class="agent-avertissements">${p.erreurs.map(x => `<li>${escapeHtml(x)}</li>`).join('')}</ul>`; return; }
      verif.innerHTML = `
        ${p.doublons.length ? `<div class="agent-error" role="alert">Tiers peut-être déjà dans ATLAS : ${p.doublons.map(d => `${escapeHtml(d.nom)} (${Math.round(d.score * 100)} %)`).join(', ')}</div>
          <label class="agent-check"><input type="checkbox" id="nt-doublons"/> Ce n'est aucun d'eux : créer quand même</label>` : ''}
        <div class="agent-muted">Code BOB</div>
        ${p.candidatsBob.map((c, i) => `<label class="agent-check"><input type="radio" name="nt-bob" value="${escapeHtml(c.refBOB)}" ${i === 0 ? 'checked' : ''}/> ${escapeHtml(c.refBOB)} · ${escapeHtml(c.nom)} <span class="agent-muted">(${c.reason === 'nom' ? 'nom proche, à confirmer' : c.reason})</span></label>`).join('')}
        <label class="agent-check"><input type="radio" name="nt-bob" value="" ${p.candidatsBob.length ? '' : 'checked'}/> ${p.codePropose ? `Nouveau code ${escapeHtml(p.codePropose)}` : 'Sans code BOB (liste BOB illisible)'}</label>
        <button type="button" class="btn btn-primary btn-block agent-btn" data-creer>Créer le tiers</button>`;
      verif.querySelector<HTMLButtonElement>('[data-creer]')!.addEventListener('click', async ev => {
        const b = ev.currentTarget as HTMLButtonElement;
        const doublonsVus = !p.doublons.length || !!verif.querySelector<HTMLInputElement>('#nt-doublons')?.checked;
        if (!doublonsVus) { res.hidden = false; res.innerHTML = '<p class="agent-error" role="alert">Coche « créer quand même » après avoir vérifié les doublons.</p>'; return; }
        const refBob = verif.querySelector<HTMLInputElement>('input[name="nt-bob"]:checked')?.value || '';
        b.disabled = true; b.textContent = 'Création…';
        try {
          const r = await creerTiers(saisie(), { ...(refBob ? { refBob } : {}), doublonsVus: true });
          host.innerHTML = `<div class="agent-fait" role="status"><span>${icon('check-circle', 14)} Tiers créé : ${escapeHtml(r.nom)}${r.codeBob ? ` (BOB ${escapeHtml(r.codeBob)})` : ''}${r.contactCree ? ' · contact principal créé' : ''}</span>
            ${r.aCompleter?.length ? `<span class="agent-muted">À compléter dans ATLAS (option absente de la liste) : ${escapeHtml(r.aCompleter.join(', '))}</span>` : ''}
            ${r.lien ? `<span class="agent-fait-btns"><button type="button" class="btn btn-secondary agent-btn" data-ouvrir>Ouvrir dans ATLAS</button></span>` : ''}</div>`;
          host.querySelector<HTMLButtonElement>('[data-ouvrir]')?.addEventListener('click', () => { try { Office.context.ui.openBrowserWindow(r.lien); } catch { window.open(r.lien, '_blank', 'noopener'); } });
          ctx.onInfo?.(`Tiers créé : ${r.nom}`, 'success');
          try { document.dispatchEvent(new CustomEvent('atlas:tiers-cree')); } catch { /* rafraîchissement facultatif */ }
        } catch (e) {
          b.disabled = false; b.textContent = 'Créer le tiers';
          const d = donneesErreur(e);
          res.hidden = false;
          res.innerHTML = d?.doublons ? '<p class="agent-error" role="alert">Doublons probables : coche « créer quand même » après vérification.</p>' : erreur('Création impossible', e);
        }
      });
    };
  }).catch(e => { if (host.isConnected) host.innerHTML = erreur('Formulaire indisponible', e); });
}

// ── 7. Déplacer vers Prospection ────────────────────────────────────────────────

export function renderProspection(host: HTMLElement, ctx: CtxDossiers): void {
  host.hidden = false;
  host.innerHTML = `<button type="button" class="btn btn-secondary btn-block agent-btn" data-go>${icon('folder-move', 14)}Déplacer vers Prospection</button><div data-res hidden></div>`;
  const b = host.querySelector<HTMLButtonElement>('[data-go]')!;
  const res = host.querySelector<HTMLElement>('[data-res]')!;
  b.addEventListener('click', async () => {
    b.disabled = true; b.textContent = 'En cours…';
    res.hidden = true;
    try {
      // 1) comme l'Inbox : mail rattaché au contact de l'expéditeur, suivi « Email », pipeline « Contacté ».
      const l = await lierProspection(ctx.messageId, ctx.mailbox || undefined);
      // 2) rangé dans « Prospection » sous la boîte de réception (créé s'il manque), annulable.
      const r = await rangerMail(ctx, { creer: { chemin: l.chemin || CHEMIN_PROSPECTION } });
      host.innerHTML = `<div class="agent-fait" role="status"><span>${icon('check-circle', 14)} ${l.contact ? `Lié à ${escapeHtml(l.contact.nom)}${l.contact.tiers ? ` (${escapeHtml(l.contact.tiers)})` : ''} et d` : 'D'}éplacé dans « ${escapeHtml(r.chemin)} »</span>
        ${l.contact ? '' : '<span class="agent-muted">Expéditeur inconnu d\'ATLAS : aucun contact rattaché (crée-le pour suivre la prospection).</span>'}
        ${r.actionId ? '<span class="agent-fait-btns"><button type="button" class="btn btn-secondary agent-btn" data-annuler>Annuler le déplacement</button></span>' : ''}</div>`;
      ctx.onInfo?.('Mail déplacé vers Prospection', 'success');
      host.querySelector<HTMLButtonElement>('[data-annuler]')?.addEventListener('click', async ev => {
        const a = ev.currentTarget as HTMLButtonElement;
        a.disabled = true;
        try { await annulerAction(r.actionId!); a.replaceWith(Object.assign(document.createElement('span'), { className: 'agent-muted', textContent: 'Annulé : mail remis à sa place' })); }
        catch (e) { a.disabled = false; ctx.onInfo?.(`Annulation impossible : ${humanError(e)}`, 'error'); }
      });
    } catch (e) {
      b.disabled = false; b.textContent = 'Déplacer vers Prospection';
      res.hidden = false;
      res.innerHTML = estVerrouFerme(e) ? verrouHtml : erreur('Déplacement impossible', e);
    }
  });
}

// ── Bloc « Plus d'actions » (volet du mail et tableau de bord) ──────────────────

export interface CtxPlus extends CtxDossiers {
  repondre?: (html: string) => void;
  deposer?: (texte: string) => Promise<void>;
  prefillTiers?: { nom?: string; email?: string; contactPrenom?: string; contactNom?: string; contactEmail?: string };
}

/** Boutons qui ouvrent chaque carte à la demande (rien n'est chargé tant qu'on n'ouvre pas). */
export function renderPlusActions(host: HTMLElement, ctx: CtxPlus): void {
  host.hidden = false;
  const items: Array<{ cle: string; libelle: string; ic: string; rendre: (z: HTMLElement) => void }> = [
    { cle: 'reclasser', libelle: 'Reclasser vers un autre projet', ic: 'folder-move', rendre: z => void renderReclasser(z, ctx) },
    { cle: 'pieces', libelle: 'Pièces jointes → projet', ic: 'paperclip', rendre: z => void renderPiecesProjet(z, ctx) },
    { cle: 'rdv', libelle: 'Proposer un RDV', ic: 'calendar-plus', rendre: z => renderRdv(z, ctx) },
    { cle: 'tiers', libelle: 'Créer un tiers', ic: 'building', rendre: z => renderCreerTiers(z, { ...ctx, prefill: ctx.prefillTiers }) },
    { cle: 'prospection', libelle: 'Déplacer vers Prospection', ic: 'inbox', rendre: z => renderProspection(z, ctx) },
  ];
  host.innerHTML = `<div class="agent-section-title">Comme dans l'Inbox ATLAS</div>${items.map(i => `
    <button type="button" class="btn btn-secondary btn-block agent-btn" data-ouvrir="${i.cle}" aria-expanded="false">${icon(i.ic, 14)}${escapeHtml(i.libelle)}</button>
    <div data-zone="${i.cle}" hidden></div>`).join('')}`;
  for (const i of items) {
    const b = host.querySelector<HTMLButtonElement>(`[data-ouvrir="${i.cle}"]`)!;
    const z = host.querySelector<HTMLElement>(`[data-zone="${i.cle}"]`)!;
    let rendu = false;
    b.addEventListener('click', () => {
      const ouvrir = z.hidden;
      z.hidden = !ouvrir;
      b.setAttribute('aria-expanded', String(ouvrir));
      if (ouvrir && !rendu) { rendu = true; i.rendre(z); }
    });
  }
}
