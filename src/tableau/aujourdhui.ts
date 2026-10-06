/**
 * aujourdhui.ts · colonne de droite sans mail choisi : « Aujourd'hui ».
 *  - Prochains rendez-vous (externes, aujourd'hui et prochain jour ouvré) avec la préparation
 *    DÉJÀ calculée par meeting-prep.service (même source que le brief 5 min avant et le mail de la veille) ;
 *  - Engagements : ce que j'ai promis / ce qu'on me doit (table Email Promises), « en retard »
 *    seulement si la vérification ne les a pas trouvés tenus ; « Relancer » / « Rédiger » (ARGO),
 *    « Tenu » ;
 *  - Réactivité aux clients : mails sans réponse au-delà de 24 h ouvrées, moyennes par personne et
 *    par client (calcul worker ; la direction voit tout) ;
 *  - Nouveaux expéditeurs à filtrer et journal de l'agent avec Annuler (agent-lists.ts, inchangés) ;
 *  - Outils ARGO existants : question à sa boîte, rattrapage (agent-outils.ts).
 */
import type { Ctx } from './app';
import type { Engagement, TableauMail } from '../api/tableau';
import { fetchEngagements, fetchRdv, fetchReactivite, marquerTenu, redigerEngagement } from '../api/tableau';
import { renderQuestion, renderRattrapage } from '../components/agent-outils';
import { renderMailList, renderJournal } from '../components/agent-lists';
import { humanError } from '../api/net';
import { icon } from '../ui/icons';
import { dureeHeures, prenomDe, quandLisible } from './logique';
import { h, toast } from './ui';
import { renderDetail, composer } from './detail';

const heure = (iso: string) => new Date(iso).toLocaleTimeString('fr-FR', { timeZone: 'Europe/Luxembourg', hour: '2-digit', minute: '2-digit' });
const jourLisible = (ymd: string) => new Date(`${ymd}T12:00:00Z`).toLocaleDateString('fr-FR', { timeZone: 'Europe/Luxembourg', weekday: 'long', day: 'numeric', month: 'long' });

export function renderAujourdhui(host: HTMLElement, ctx: Ctx): void {
  host.innerHTML = `
    <div class="tb-panel is-raised" id="tb-rdv"><div class="tb-h">Prochains rendez-vous</div><div class="tb-skel"></div></div>
    <div class="tb-panel" id="tb-eng"><div class="tb-h">Engagements</div><div class="tb-skel"></div></div>
    <div class="tb-panel" id="tb-react"><div class="tb-h">Réactivité aux clients</div><div class="tb-skel"></div></div>
    <div class="tb-panel" id="tb-filtrer"><div class="tb-h">Nouveaux expéditeurs à filtrer</div><div id="tb-filtrer-liste"></div></div>
    <div class="tb-panel"><div class="tb-h">Ce que l'agent a fait</div><div id="tb-journal"></div></div>
    <div class="tb-panel"><div class="tb-tabs" role="tablist"><button type="button" role="tab" data-outil="question" aria-selected="true">Question à ma boîte</button><button type="button" role="tab" data-outil="rattrapage" aria-selected="false">Rattrapage</button></div><div id="tb-outil"></div></div>`;
  void rdv(host.querySelector('#tb-rdv')!);
  void engagements(host.querySelector('#tb-eng')!, ctx);
  void reactivite(host.querySelector('#tb-react')!, ctx);
  // Repris du tableau de bord précédent : filtre des nouveaux expéditeurs (Accepter / Refuser) et journal avec Annuler.
  void renderMailList(host.querySelector('#tb-filtrer-liste')!, 'a_filtrer', { limite: 8, openLink: ctx.openLink, onInfo: (msg, type) => toast(msg, type), onChange: () => void ctx.rafraichir(true) });
  void renderJournal(host.querySelector('#tb-journal')!, { limite: 8, onInfo: (msg, type) => toast(msg, type), onChange: () => void ctx.rafraichir(true) });
  const outil = host.querySelector<HTMLElement>('#tb-outil')!;
  const montrer = (quoi: string) => {
    host.querySelectorAll<HTMLElement>('[data-outil]').forEach(b => b.setAttribute('aria-selected', String(b.dataset.outil === quoi)));
    outil.innerHTML = '';
    if (quoi === 'question') renderQuestion(outil, { openLink: ctx.openLink, onInfo: (msg, type) => toast(msg, type) });
    else renderRattrapage(outil, { openLink: ctx.openLink, onInfo: (msg, type) => toast(msg, type) });
    outil.querySelector<HTMLElement>('input, textarea')?.focus();
  };
  host.querySelectorAll<HTMLElement>('[data-outil]').forEach(b => b.addEventListener('click', () => montrer(b.dataset.outil!)));
  renderQuestion(outil, { openLink: ctx.openLink, onInfo: (msg, type) => toast(msg, type) });
}

async function rdv(host: HTMLElement): Promise<void> {
  try {
    const r = await fetchRdv();
    if (!r.rdv.length) { host.innerHTML = '<div class="tb-h">Prochains rendez-vous</div><p class="tb-note">Aucun rendez-vous externe aujourd\'hui ni au prochain jour ouvré.</p>'; return; }
    const parJour = [r.jours.aujourdhui, r.jours.lendemain].map(j => ({ j, l: r.rdv.filter(x => x.jour === j) })).filter(x => x.l.length);
    host.innerHTML = `<div class="tb-h">Prochains rendez-vous<span class="tb-h-n">${r.rdv.length}</span></div>${parJour.map(({ j, l }) => `
      <p class="tb-note">${h(j === r.jours.aujourdhui ? 'Aujourd\'hui' : jourLisible(j))}</p>
      ${l.map(x => `<div class="tb-rdv"><time>${h(heure(x.debut))}</time><details${x.utile && j === r.jours.aujourdhui ? ' open' : ''}>
        <summary>${h(x.sujet)}<br><span class="tb-note">${h(x.avec.join(', ') || 'externes')}${x.lieu ? ` · ${h(x.lieu)}` : ''}</span></summary>
        <div class="tb-prep">${x.prep && (x.prep.sections.length || x.prep.notesRdv.length) ? [
          x.prep.viaIntitule ? `<p class="tb-note">${h(x.prep.viaIntitule)}</p>` : '',
          ...x.prep.notesRdv.map(n => `<h5>${h(n.titre)}</h5><p>${h(n.contenu)}</p>`),
          ...x.prep.sections.map(s => `<h5>${h(s.titre)}</h5><ul>${s.lignes.map(li => `<li>${h(li.texte)}${li.meta ? ` <span class="tb-note">${h(li.meta)}</span>` : ''}</li>`).join('')}</ul>`),
        ].join('') : '<p class="tb-note">Rien de particulier à préparer.</p>'}</div>
      </details></div>`).join('')}`).join('')}`;
  } catch (e) {
    host.innerHTML = `<div class="tb-h">Prochains rendez-vous</div><p class="tb-note">${h(humanError(e))}</p>`;
  }
}

async function engagements(host: HTMLElement, ctx: Ctx): Promise<void> {
  try {
    const r = await fetchEngagements();
    const ligne = (e: Engagement) => `<div class="tb-eng">
      <p>${h(e.promiseText)}<br><small>${e.direction === 'made' ? 'à' : 'de'} ${h(e.counterpartName || e.counterpart)} · ${e.enRetard ? `<span class="tb-chip is-hot">en retard de ${Math.max(1, e.jours)} j</span>` : `échéance ${h(quandLisible(e.dueDate, Date.now()) || e.dueDate.slice(0, 10))}`}</small></p>
      <span class="tb-actions"><button type="button" class="tb-btn${e.enRetard ? ' is-argo' : ''}" data-r="${h(e.recordId)}">${e.direction === 'received' ? 'Relancer' : 'Rédiger'}</button><button type="button" class="tb-btn is-ghost" data-k="${h(e.recordId)}" title="Marquer tenu">${icon('check', 14)}</button></span>
    </div>`;
    const total = r.promis.length + r.attendus.length;
    host.innerHTML = `<div class="tb-h">Engagements<span class="tb-h-n">${total}</span></div>
      ${!total ? '<p class="tb-note">Aucun engagement en suspens : tout ce qui a été promis est tenu.</p>' : ''}
      ${r.promis.length ? `<p class="tb-note">Ce que j'ai promis</p>${r.promis.slice(0, 6).map(ligne).join('')}` : ''}
      ${r.attendus.length ? `<p class="tb-note">Ce qu'on me doit</p>${r.attendus.slice(0, 6).map(ligne).join('')}` : ''}
      ${r.verificationEnCours ? '<p class="tb-note">Vérification des engagements échus en cours : ceux qui sont tenus disparaîtront.</p>' : ''}`;
    host.querySelectorAll<HTMLButtonElement>('[data-k]').forEach(b => b.addEventListener('click', async () => {
      b.disabled = true;
      try { await marquerTenu(b.dataset.k!); toast('Engagement tenu', 'success'); void engagements(host, ctx); } catch (e) { toast(humanError(e), 'error'); b.disabled = false; }
    }));
    host.querySelectorAll<HTMLButtonElement>('[data-r]').forEach(b => b.addEventListener('click', async () => {
      const e = [...r.promis, ...r.attendus].find(x => x.recordId === b.dataset.r)!;
      b.disabled = true;
      try {
        const rep = await redigerEngagement(e.recordId);
        const cle = `${(rep.mailbox || '').toLowerCase()}|${rep.messageId}`;
        const t = ctx.etat.t;
        const connu = t ? [...t.personnes, ...t.deCote, ...t.notifications, ...t.factures].find(m => `${(m.mailbox || '').toLowerCase()}|${m.messageId}` === cle) : undefined;
        const m: TableauMail = connu || {
          messageId: rep.messageId || '', graphId: '', mailbox: rep.mailbox || t?.moi || '', conversationId: e.sourceConversationId,
          from: { email: e.counterpart, name: e.counterpartName }, subject: e.sourceSubject, receivedAt: e.createdAt, pile: 'a_traiter',
          resume: e.promiseText, urgence: 0, categorie: '', famille: 'personnes', priorite: 0, raisons: [e.direction === 'received' ? 'Promesse reçue' : 'Ma promesse'],
        };
        const droite = document.getElementById('tb-right')!;
        if (connu) ctx.choisir(cle); else renderDetail(droite, m, ctx);
        void composer(droite, m, ctx, 'argo', rep);
      } catch (err) { toast(humanError(err), 'error'); }
      b.disabled = false;
    }));
  } catch (e) {
    host.innerHTML = `<div class="tb-h">Engagements</div><p class="tb-note">${h(humanError(e))}</p>`;
  }
}

async function reactivite(host: HTMLElement, ctx: Ctx): Promise<void> {
  try {
    const r = await fetchReactivite();
    const retards = r.personnes.reduce((n, p) => n + p.enRetard.length, 0);
    if (!r.personnes.length) { host.innerHTML = '<div class="tb-h">Réactivité aux clients</div><p class="tb-note">Aucun mail client à mesurer sur 30 jours.</p>'; return; }
    host.innerHTML = `<div class="tb-h">Réactivité aux clients<span class="tb-h-n">${retards ? `${retards} > ${r.seuilHeures} h ouvrées` : 'à jour'}</span></div>
      <table class="tb-react"><thead><tr><th>Personne</th><th>Délai moyen</th><th>Sans réponse</th></tr></thead><tbody>
      ${r.personnes.map(p => `<tr class="${p.enRetard.length ? 'is-late' : ''}"><td>${h(prenomDe(p.personne))}</td><td>${h(dureeHeures(p.moyenneHeures))}</td><td>${p.enRetard.length}</td></tr>`).join('')}
      </tbody></table>
      ${retards ? `<div>${r.personnes.flatMap(p => p.enRetard.slice(0, 4).map(x => `<div class="tb-eng"><p><b>${h(x.client)}</b> · ${h(x.subject)}<br><small>${h(prenomDe(x.mailbox))} · ${h(dureeHeures(x.heures))} ouvrées · responsable ${h(prenomDe(x.responsable))}</small></p>${x.webLink ? `<button type="button" class="tb-btn is-ghost" data-l="${h(x.webLink)}">${icon('external', 14)}</button>` : ''}</div>`)).join('')}</div>` : ''}
      ${r.clients.length ? `<table class="tb-react"><thead><tr><th>Client</th><th>Délai moyen</th><th>Mails</th></tr></thead><tbody>${r.clients.slice(0, 6).map(c => `<tr class="${c.enRetard ? 'is-late' : ''}"><td>${h(c.client)}</td><td>${h(dureeHeures(c.moyenneHeures))}</td><td>${c.mails}</td></tr>`).join('')}</tbody></table>` : ''}
      <p class="tb-note">Heures ouvrées de chacun (horaires, congés, fériés LU) ; le responsable du client est prévenu une fois par mail.</p>`;
    host.querySelectorAll<HTMLButtonElement>('[data-l]').forEach(b => b.addEventListener('click', () => ctx.openLink(b.dataset.l!)));
  } catch (e) {
    host.innerHTML = `<div class="tb-h">Réactivité aux clients</div><p class="tb-note">${h(humanError(e))}</p>`;
  }
}
