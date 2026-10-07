/**
 * assistant-inbox-reglages.ts · réglages de l'« Assistant inbox » dans les Réglages du complément
 * (backlog recbcLIUXh5Y5FMoc, 07/10/2026) : mode concentration (créneaux de distribution), règles
 * écrites en français (interprétation à confirmer, pause, suppression) et tâches confiées à ATLAS
 * (« Je le fais seul désormais »). Tout est gardé par le worker, par personne (aucun jeton de boîte).
 */
import { showToast } from '../taskpane';
import {
  creerRegle, enregistrerConcentration, etatRegle, fetchReglagesAssistant, retirerDelegation, supprimerRegle,
  type ReglagesAssistant, type RegleAssistant,
} from '../api/seance';
import { humanError } from '../api/net';
import { escapeHtml } from '../utils/html';
import { icon } from '../ui/icons';
import { inlineLoadingHtml } from '../ui/states';

const ETATS: Record<RegleAssistant['etat'], string> = { 'a-confirmer': 'à confirmer', active: 'active', pause: 'en pause' };

export function monterReglagesAssistant(host: HTMLElement): void {
  let r: ReglagesAssistant | null = null;

  const rendre = (): void => {
    if (!r) return;
    const c = r.concentration;
    const observation = !r.ecrituresAgent;
    host.innerHTML = `
      <div class="tool-row">
        <div class="tool-row-icon">${icon('clock', 18)}</div>
        <div class="tool-row-body">
          <p class="tool-row-title">Mode concentration</p>
          <p class="help">Les mails non urgents arrivent à heures fixes. Clients actifs, direction, alertes de sécurité et urgences passent toujours.${observation ? ' Tant que l\'agent observe, rien ne bouge dans Outlook : la séance de tri et les vues ATLAS les gardent pour le créneau.' : ' Entre deux créneaux, ils attendent dans « ATLAS · Plus tard ».'}</p>
          <label class="form-label checkbox-row"><input type="checkbox" id="as-conc-actif" ${c.actif ? 'checked' : ''}/> Activer</label>
          <label class="form-label" for="as-conc-creneaux">Créneaux (heure de Luxembourg)</label>
          <input type="text" class="form-input" id="as-conc-creneaux" value="${escapeHtml(c.creneaux.join(', '))}" placeholder="9:00, 13:00, 17:00" />
          <button class="btn btn-secondary btn-sm" id="as-conc-save">${icon('check', 14)}Enregistrer</button>
        </div>
      </div>

      <div class="tool-row">
        <div class="tool-row-icon">${icon('tag', 18)}</div>
        <div class="tool-row-body">
          <p class="tool-row-title">Mes règles en français</p>
          <p class="help">Écris la règle comme tu la dirais, par exemple « Les BAT de Paperjam vont sur la créa, puis classés ». ATLAS te montre ce qu'il a compris : rien ne s'applique avant ta confirmation.${observation ? ' Pendant l\'observation, la règle est proposée dans la séance de tri.' : ''}</p>
          <textarea class="form-input" id="as-regle-texte" rows="2" maxlength="400" placeholder="Les notifications Metricool sont archivées"></textarea>
          <button class="btn btn-secondary btn-sm" id="as-regle-creer">${icon('bolt', 14)}Comprendre la règle</button>
          <div id="as-regles" class="tool-row-result">${r.regles.length ? r.regles.map(x => `
            <div class="as-regle" data-id="${escapeHtml(x.id)}">
              <p><b>${escapeHtml(x.texte)}</b><br/><span class="help">${escapeHtml(x.interpretation)}</span><br/><small>${ETATS[x.etat]}${x.appliquee ? ` · appliquée ${x.appliquee} fois` : ''}</small></p>
              <div class="btn-row">
                ${x.etat !== 'active' ? `<button class="btn btn-primary btn-sm" data-a="active">${x.etat === 'a-confirmer' ? 'C\'est bien ça' : 'Réactiver'}</button>` : '<button class="btn btn-ghost btn-sm" data-a="pause">Pause</button>'}
                <button class="btn btn-ghost btn-sm" data-a="suppr">Supprimer</button>
              </div>
            </div>`).join('') : '<p class="help">Aucune règle pour l\'instant.</p>'}</div>
        </div>
      </div>

      ${r.delegations.length ? `<div class="tool-row">
        <div class="tool-row-icon">${icon('check-circle', 18)}</div>
        <div class="tool-row-body">
          <p class="tool-row-title">Ce qu'ATLAS fait seul</p>
          ${r.delegations.map(d => `<div class="as-regle" data-p="${escapeHtml(d.patron)}"><p>${escapeHtml(d.libelle)}<br/><small>${d.etat === 'appliquer' ? `en service${d.appliquees ? ` · ${d.appliquees} fois` : ''}` : 'prêt, dès que l\'agent sort du mode observation'}</small></p><div class="btn-row"><button class="btn btn-ghost btn-sm" data-a="retirer">Reprendre la main</button></div></div>`).join('')}
        </div>
      </div>` : ''}`;

    host.querySelector('#as-conc-save')?.addEventListener('click', async () => {
      const actif = (host.querySelector('#as-conc-actif') as HTMLInputElement).checked;
      const creneaux = (host.querySelector('#as-conc-creneaux') as HTMLInputElement).value.split(/[,;\s]+/).map(x => x.trim()).filter(Boolean);
      try {
        const res = await enregistrerConcentration({ actif, creneaux });
        r!.concentration = res.concentration;
        showToast(actif ? `Concentration : ${res.concentration.creneaux.join(', ')}` : 'Concentration coupée', 'success');
        rendre();
      } catch (e) { showToast(humanError(e), 'error'); }
    });

    host.querySelector('#as-regle-creer')?.addEventListener('click', async () => {
      const ta = host.querySelector('#as-regle-texte') as HTMLTextAreaElement;
      const btn = host.querySelector('#as-regle-creer') as HTMLButtonElement;
      const texte = ta.value.trim();
      if (texte.length < 8) { showToast('Écris la règle en une phrase.', 'error'); return; }
      btn.disabled = true;
      btn.innerHTML = inlineLoadingHtml('ARGO lit la règle…');
      try {
        const res = await creerRegle(texte);
        r!.regles = [...r!.regles, res.regle];
        showToast('Vérifie ce qu\'ATLAS a compris, puis confirme', 'info');
        rendre();
      } catch (e) {
        showToast(humanError(e), 'error');
        btn.disabled = false;
        btn.innerHTML = `${icon('bolt', 14)}Comprendre la règle`;
      }
    });

    host.querySelectorAll<HTMLElement>('.as-regle[data-id]').forEach(el => el.querySelectorAll<HTMLButtonElement>('[data-a]').forEach(b => b.addEventListener('click', async () => {
      const id = el.dataset.id!;
      try {
        if (b.dataset.a === 'suppr') {
          if (!confirm('Supprimer cette règle ?')) return;
          await supprimerRegle(id);
          r!.regles = r!.regles.filter(x => x.id !== id);
        } else {
          const etat = b.dataset.a as 'active' | 'pause';
          await etatRegle(id, etat);
          r!.regles = r!.regles.map(x => (x.id === id ? { ...x, etat } : x));
        }
        rendre();
      } catch (e) { showToast(humanError(e), 'error'); }
    })));

    host.querySelectorAll<HTMLElement>('.as-regle[data-p]').forEach(el => el.querySelector('[data-a="retirer"]')?.addEventListener('click', async () => {
      try {
        await retirerDelegation(el.dataset.p!);
        r!.delegations = r!.delegations.filter(d => d.patron !== el.dataset.p);
        showToast('ATLAS te le proposera de nouveau', 'info');
        rendre();
      } catch (e) { showToast(humanError(e), 'error'); }
    }));
  };

  host.innerHTML = inlineLoadingHtml('Chargement des réglages…');
  fetchReglagesAssistant()
    .then(x => {
      r = x;
      if (!x.boiteSuivie) { host.innerHTML = '<p class="help">L\'assistant inbox s\'active quand l\'agent ATLAS suit ta boîte. Demande à Charles.</p>'; return; }
      rendre();
    })
    .catch(e => { host.innerHTML = `<p class="help">${escapeHtml(humanError(e))}</p>`; });
}
