/**
 * boite-zero.ts · « Boîte zéro du soir » (lot 6 du complément, 10/10/2026), dans le volet et le
 * tableau de bord.
 *
 * Le soir d'un jour ouvré (après 17 h) : « Ranger 3 mails répondus, archiver 12 mails sans intérêt,
 * rappeler 2 mails à traiter demain matin » en un clic, puis « Annuler » (tout est remis en place).
 * Si la personne a réglé « Classer dans le dossier du projet : seul », ATLAS l'a fait seul et la carte
 * en donne le résumé, avec « Annuler ». Masquée en journée et quand la boîte est déjà propre.
 */
import { annulerBoiteZero, appliquerBoiteZero, fetchBoiteZero, type BoiteZero } from '../api/tableau';
import { humanError } from '../api/net';
import { escapeHtml } from '../utils/html';
import { icon } from '../ui/icons';

type InfoFn = (msg: string, kind?: 'success' | 'error' | 'info') => void;

/** Classes du volet (agent-*) ou du tableau de bord (tb-*). */
const CLASSES = {
  volet: { titre: 'agent-section-title', carte: 'agent-carte', principal: 'btn btn-primary btn-block agent-btn', lien: 'agent-link', note: 'agent-muted', fait: 'status-linked' },
  tableau: { titre: 'tb-h', carte: '', principal: 'tb-btn is-primary', lien: 'tb-btn is-ghost', note: 'tb-note', fait: 'tb-note' },
};

const n = (k: number, un: string, plusieurs: string) => `${k} ${k > 1 ? plusieurs : un}`;

function lignesPlan(b: BoiteZero): string[] {
  return [
    b.plan.ranger.length ? `Ranger ${n(b.plan.ranger.length, 'mail répondu', 'mails répondus')} dans leur dossier` : '',
    b.plan.archiver.length ? `Archiver ${n(b.plan.archiver.length, 'mail sans intérêt', 'mails sans intérêt')}` : '',
    b.plan.rappeler.length ? `Rappeler demain matin ${n(b.plan.rappeler.length, 'mail à traiter', 'mails à traiter')} depuis plus de 48 h` : '',
  ].filter(Boolean);
}

/** Soir d'un jour de semaine à Luxembourg (le worker tranche aussi les jours fériés) : rien n'est demandé en journée. */
export function soirDeSemaine(now = new Date()): boolean {
  const f = new Intl.DateTimeFormat('en-GB', { timeZone: 'Europe/Luxembourg', weekday: 'short', hour: '2-digit', hour12: false }).formatToParts(now);
  const jour = f.find(x => x.type === 'weekday')?.value || '';
  const heure = Number(f.find(x => x.type === 'hour')?.value);
  return !['Sat', 'Sun'].includes(jour) && heure >= 17 && heure < 23;
}

export async function renderBoiteZero(host: HTMLElement, o: { style: 'volet' | 'tableau'; onInfo: InfoFn; onChange?: () => void }): Promise<void> {
  const c = CLASSES[o.style];
  let b: BoiteZero;
  try { b = await fetchBoiteZero(); } catch { host.hidden = true; return; }
  if (!host.isConnected) return;
  const lignes = lignesPlan(b);
  const dernier = b.dernier;
  if (!dernier?.resume && (!b.fenetre || !lignes.length)) { host.hidden = true; host.innerHTML = ''; return; }
  host.hidden = false;
  const annuler = dernier?.annulable ? `<button type="button" class="${c.lien}" data-annuler>Annuler</button>` : '';
  host.innerHTML = `
    <div class="${c.titre}">${icon('broom', 14)} Boîte zéro</div>
    <div class="${c.carte}">
      ${dernier?.resume ? `<p class="${c.fait}">${escapeHtml(dernier.seul ? `Fait seul ce soir : ${dernier.resume}` : dernier.resume)}</p>${annuler}` : ''}
      ${b.fenetre && lignes.length ? `
        <ul class="${c.note}">${lignes.map(l => `<li>${escapeHtml(l)}</li>`).join('')}</ul>
        <button type="button" class="${c.principal}" data-nettoyer>Tout faire en un clic</button>
        <div class="${c.note}">Annulable. Aucun mail n'est supprimé ni envoyé.</div>` : ''}
    </div>`;
  host.querySelector<HTMLButtonElement>('[data-nettoyer]')?.addEventListener('click', async ev => {
    const btn = ev.currentTarget as HTMLButtonElement;
    btn.disabled = true; btn.textContent = 'Nettoyage…';
    try {
      const r = await appliquerBoiteZero();
      o.onInfo(r.erreurs ? `${r.resume} (${r.erreurs} non fait${r.erreurs > 1 ? 's' : ''})` : r.resume, r.erreurs ? 'info' : 'success');
      o.onChange?.();
      void renderBoiteZero(host, o);
    } catch (e) { btn.disabled = false; btn.textContent = 'Tout faire en un clic'; o.onInfo(humanError(e), 'error'); }
  });
  host.querySelector<HTMLButtonElement>('[data-annuler]')?.addEventListener('click', async ev => {
    const btn = ev.currentTarget as HTMLButtonElement;
    btn.disabled = true;
    try {
      const r = await annulerBoiteZero();
      o.onInfo(r.erreurs ? `Annulé en partie (${r.erreurs} à remettre à la main)` : 'Annulé : tout est remis en place', r.erreurs ? 'info' : 'success');
      o.onChange?.();
      void renderBoiteZero(host, o);
    } catch (e) { btn.disabled = false; o.onInfo(humanError(e), 'error'); }
  });
}
