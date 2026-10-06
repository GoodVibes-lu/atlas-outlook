/**
 * Settings Panel · Réglages : profil, outils de la boîte (catégories, index des dossiers),
 * Diagnostic (origine, serveur ATLAS, jeton, dernières erreurs) et à propos.
 *
 * Plus de clé Anthropic ni de jeton Airtable (03/10/2026) : données ATLAS et ARGO passent par le
 * worker avec le compte Microsoft de l'utilisateur (cf. api/worker.ts).
 * Refonte 05/10/2026 : DS ATLAS, icônes SVG, Diagnostic discret (cf. api/net.ts).
 */

import { showToast } from '../taskpane';
import { scanMailboxBuildIndex } from '../api/scan-mailbox';
import { indexStats, clearIndex } from '../api/sender-folder-index';
import { createAtlasCategoriesVerbose } from '../api/graph';
import { persistKey, readGraphToken, saveGraphToken } from '../api/roaming-storage';
import { escapeHtml } from '../utils/html';
import { getDiag, humanError, noteDiag } from '../api/net';
import { callAtlasWorker, resetWorkerToken } from '../api/worker';
import { icon } from '../ui/icons';
import { inlineLoadingHtml } from '../ui/states';

/** Version affichée (manifeste 1.4.0 : menu « Tableau de bord », 07/10/2026). */
const ADDIN_VERSION = '1.4.0 · tableau de bord en fenêtre, 07/10/2026';

export class SettingsPanel {
  private container: HTMLElement;
  private onSave: () => void;

  constructor(container: HTMLElement, onSave: () => void) {
    this.container = container;
    this.onSave = onSave;
    this.render();
  }

  private render(): void {
    const graphToken = readGraphToken();
    // Pré-rempli avec le profil Outlook (plus d'écran de configuration initiale).
    const profile = (() => { try { return Office.context?.mailbox?.userProfile; } catch { return undefined; } })();
    const userName = localStorage.getItem('atlas_addin_user_name') || profile?.displayName || '';
    const userEmail = localStorage.getItem('atlas_addin_user_email') || profile?.emailAddress || '';

    this.container.innerHTML = `
      <div class="panel-scroll settings">
        <section class="section" aria-labelledby="set-profil">
          <h2 class="section-heading" id="set-profil">Ton profil</h2>
          <div class="form-group">
            <label class="form-label" for="setting-name">Nom affiché</label>
            <input type="text" class="form-input" id="setting-name" value="${escapeHtml(userName)}" placeholder="Prénom Nom" autocomplete="name" />
          </div>
          <div class="form-group">
            <label class="form-label" for="setting-email">Adresse e-mail</label>
            <input type="email" class="form-input" id="setting-email" value="${escapeHtml(userEmail)}" placeholder="prenom@vibes.lu" autocomplete="email" />
          </div>
          <p class="help">Projets, liaisons et ARGO passent par le serveur ATLAS avec ton compte Microsoft : aucune clé à saisir ici.</p>

          <details class="disclosure">
            <summary>${icon('chevron-right', 14)}Accès avancé à la boîte (Outlook Mac)</summary>
            <div class="disclosure-body">
              <label class="form-label" for="setting-graph">Jeton Microsoft Graph</label>
              <textarea class="form-input mono" id="setting-graph" rows="3" placeholder="Coller le jeton ici">${escapeHtml(graphToken)}</textarea>
              <p class="help">Outlook Mac ne laisse pas toujours ATLAS lire la liste de tes dossiers : sans ce jeton, l'apprentissage des dossiers ne marche pas sur ce poste.
              Dans ATLAS desktop, ouvre la console (Cmd+Opt+I), tape <code>copy(localStorage.getItem('atlas_ms_access_token'))</code> puis colle ici.
              Le jeton dure environ une heure et n'est gardé que pour cette session.</p>
            </div>
          </details>

          <button class="btn btn-primary btn-block" id="save-settings-btn">${icon('check', 14)}Enregistrer</button>
        </section>

        <section class="section" aria-labelledby="set-outils">
          <h2 class="section-heading" id="set-outils">Outils de la boîte</h2>

          <div class="tool-row">
            <div class="tool-row-icon">${icon('tag', 18)}</div>
            <div class="tool-row-body">
              <p class="tool-row-title">Catégories ATLAS dans Outlook</p>
              <p class="help">Crée les 20 catégories (14 thèmes ARGO et 6 états) pour voir les étiquettes dans ta liste de mails. À faire une fois.</p>
              <button class="btn btn-secondary btn-sm" id="create-cats-btn">${icon('tag', 14)}Créer les catégories</button>
              <div id="cats-result" class="tool-row-result" aria-live="polite"></div>
            </div>
          </div>

          <div class="tool-row">
            <div class="tool-row-icon">${icon('folder', 18)}</div>
            <div class="tool-row-body">
              <p class="tool-row-title">Où tu ranges chaque expéditeur</p>
              <div id="index-stats" class="help">${this.renderIndexStats()}</div>
              <p class="help">Parcourt ta boîte pour apprendre tes dossiers habituels (30 secondes à 2 minutes). L'index s'enrichit ensuite à chaque rangement.</p>
              <div class="btn-row">
                <button class="btn btn-secondary btn-sm" id="scan-mailbox-btn">${icon('search', 14)}Analyser ma boîte</button>
                <button class="btn btn-ghost btn-sm" id="clear-index-btn">Vider l'index</button>
              </div>
              <div id="scan-progress" class="tool-row-result" aria-live="polite" hidden></div>
            </div>
          </div>
        </section>

        <section class="section">
          <details class="disclosure" id="diag-details">
            <summary>${icon('activity', 14)}Diagnostic</summary>
            <div class="disclosure-body">
              <p class="help">À transmettre à Charles en cas de souci. Aucun contenu de mail n'y figure.</p>
              <dl class="diag-list" id="diag-list"></dl>
              <div class="btn-row">
                <button class="btn btn-secondary btn-sm" id="diag-test">${icon('refresh', 14)}Tester la connexion</button>
                <button class="btn btn-ghost btn-sm" id="diag-copy">${icon('copy', 14)}Copier</button>
              </div>
              <div id="diag-test-res" class="tool-row-result" aria-live="polite"></div>
            </div>
          </details>
        </section>

        <footer class="settings-about">
          ATLAS pour Outlook · ${ADDIN_VERSION}<br/>
          GOOD VIBES events &amp; communications · <a href="https://vibes.lu" target="_blank" rel="noopener">vibes.lu</a>
        </footer>
      </div>
    `;

    document.getElementById('save-settings-btn')?.addEventListener('click', () => this.saveSettings());

    document.getElementById('create-cats-btn')?.addEventListener('click', async () => {
      const btn = document.getElementById('create-cats-btn') as HTMLButtonElement;
      const out = document.getElementById('cats-result');
      btn.disabled = true;
      btn.innerHTML = inlineLoadingHtml('Création en cours…');
      if (out) out.textContent = '';
      try {
        const r = await createAtlasCategoriesVerbose();
        const parts = [`${r.created} créée${r.created > 1 ? 's' : ''}`, `${r.existed} déjà présente${r.existed > 1 ? 's' : ''}`];
        if (r.failed > 0) parts.push(`${r.failed} en échec`);
        if (out) out.textContent = parts.join(' · ');
        if (r.created > 0) showToast(`${r.created} catégories créées dans Outlook`, 'success');
        else if (r.existed > 0 && r.failed === 0) showToast('Toutes les catégories existent déjà', 'info');
        else showToast('Outlook n\'a pas accepté la création des catégories : voir Diagnostic.', 'error');
        if (r.errors.length) noteDiag({ at: Date.now(), service: 'outlook', label: 'catégories', status: 0, ms: 0, kind: 'requete', detail: r.errors.slice(0, 3).join(' | ') });
      } catch (e) {
        if (out) out.textContent = humanError(e);
        showToast(humanError(e), 'error');
      } finally {
        btn.disabled = false;
        btn.innerHTML = `${icon('refresh', 14)}Recréer les catégories`;
        this.renderDiag();
      }
    });

    document.getElementById('scan-mailbox-btn')?.addEventListener('click', async () => {
      const btn = document.getElementById('scan-mailbox-btn') as HTMLButtonElement;
      const progDiv = document.getElementById('scan-progress');
      btn.disabled = true;
      btn.innerHTML = inlineLoadingHtml('Analyse en cours…');
      if (progDiv) { progDiv.hidden = false; progDiv.textContent = 'Préparation…'; }
      try {
        const r = await scanMailboxBuildIndex((p) => {
          if (progDiv) {
            progDiv.textContent = `${p.foldersScanned} / ${p.foldersTotal} dossiers · ${p.mailsIndexed} mails · ${p.currentFolder.slice(0, 30)}`;
          }
        });
        showToast(`Analyse terminée : ${r.foldersScanned} dossiers, ${r.mailsIndexed} mails`, 'success');
        const stats = document.getElementById('index-stats');
        if (stats) stats.innerHTML = this.renderIndexStats();
        if (progDiv) progDiv.hidden = true;
      } catch (e) {
        if (progDiv) progDiv.textContent = humanError(e);
        showToast(humanError(e), 'error');
      } finally {
        btn.disabled = false;
        btn.innerHTML = `${icon('refresh', 14)}Relancer l'analyse`;
      }
    });

    document.getElementById('clear-index-btn')?.addEventListener('click', () => {
      if (confirm('Vider tout l\'index des expéditeurs et de leurs dossiers ?')) {
        clearIndex();
        showToast('Index vidé', 'info');
        const stats = document.getElementById('index-stats');
        if (stats) stats.innerHTML = this.renderIndexStats();
      }
    });

    // Diagnostic : lu à l'ouverture du bloc (et après un test).
    document.getElementById('diag-details')?.addEventListener('toggle', () => this.renderDiag());
    document.getElementById('diag-copy')?.addEventListener('click', () => {
      const text = JSON.stringify(getDiag(), null, 2);
      const done = () => showToast('Diagnostic copié', 'success');
      try {
        navigator.clipboard.writeText(text).then(done, () => { this.copyFallback(text); done(); });
      } catch { this.copyFallback(text); done(); }
    });
    document.getElementById('diag-test')?.addEventListener('click', async () => {
      const btn = document.getElementById('diag-test') as HTMLButtonElement;
      const out = document.getElementById('diag-test-res');
      btn.disabled = true;
      if (out) out.innerHTML = inlineLoadingHtml('Test en cours…');
      try {
        resetWorkerToken();
        await callAtlasWorker('templates/list', {});
        if (out) out.innerHTML = `<span class="ok-text">${icon('check-circle', 14)}ATLAS répond, ton compte est reconnu.</span>`;
      } catch (e) {
        if (out) out.textContent = humanError(e);
      } finally {
        btn.disabled = false;
        this.renderDiag();
      }
    });
  }

  /** Bloc Diagnostic : origine, URL du worker, jeton, dernière erreur, derniers appels. */
  private renderDiag(): void {
    const dl = document.getElementById('diag-list');
    if (!dl) return;
    const d = getDiag();
    const fmt = (t: number) => (t ? new Date(t).toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit', second: '2-digit' }) : '');
    const tokenState = d.token.until > Date.now()
      ? `valide jusqu'à ${fmt(d.token.until)} (${d.token.source === 'naa' ? 'connexion automatique' : 'SSO Office'})`
      : d.token.lastError ? `échec à ${fmt(d.token.lastErrorAt)}` : 'pas encore demandé';
    const rows: Array<[string, string]> = [
      ['Origine', d.origin || '?'],
      ['Serveur ATLAS', d.workerUrl],
      ['Réseau', d.online ? 'en ligne' : 'hors ligne'],
      ['Outlook', [d.host, d.platform, d.version].filter(Boolean).join(' · ') || '?'],
      ['Connexion automatique', d.token.naaSupported === null ? '?' : d.token.naaSupported ? 'proposée' : 'non proposée'],
      ['Jeton', tokenState],
    ];
    if (d.token.lastError) rows.push(['Détail jeton', d.token.lastError.slice(0, 220)]);
    if (d.lastError) {
      rows.push(['Dernière erreur', `${fmt(d.lastError.at)} · ${d.lastError.label} · ${d.lastError.status ? `HTTP ${d.lastError.status}` : 'sans réponse'} · ${d.lastError.kind}`]);
      if (d.lastError.detail) rows.push(['Détail', d.lastError.detail.slice(0, 220)]);
    } else rows.push(['Dernière erreur', 'aucune']);
    const recents = d.recent.slice(0, 6).map(r => `${fmt(r.at)} ${r.label} ${r.status || 'sans réponse'} ${r.ms} ms`).join('\n');
    if (recents) rows.push(['Derniers appels', recents]);
    dl.innerHTML = rows.map(([k, v]) => `<dt>${escapeHtml(k)}</dt><dd>${escapeHtml(v)}</dd>`).join('');
  }

  private copyFallback(text: string): void {
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed';
    ta.style.opacity = '0';
    document.body.appendChild(ta);
    ta.select();
    try { document.execCommand('copy'); } catch { /* sans presse-papiers */ }
    ta.remove();
  }

  private saveSettings(): void {
    const name = (document.getElementById('setting-name') as HTMLInputElement).value.trim();
    const email = (document.getElementById('setting-email') as HTMLInputElement).value.trim();
    const graph = (document.getElementById('setting-graph') as HTMLTextAreaElement).value.trim();

    if (!name || !email) {
      showToast('Indique ton nom et ton adresse e-mail.', 'error');
      return;
    }

    // Persiste à la fois en localStorage (rapide) ET roamingSettings
    // (survit aux cache clears + sync multi-device). Best effort en parallèle.
    Promise.all([
      persistKey('atlas_addin_user_name', name),
      persistKey('atlas_addin_user_email', email),
    ]).catch((e) => console.warn('[Settings] persistKey failed:', e));
    saveGraphToken(graph); // session seulement

    showToast('Réglages enregistrés', 'success');
    this.onSave();
  }

  private renderIndexStats(): string {
    const s = indexStats();
    if (s.senders === 0 && s.domains === 0) {
      return 'Index vide : lance une analyse pour commencer.';
    }
    const lastUpdate = s.lastUpdate ? new Date(s.lastUpdate).toLocaleString('fr-FR', { dateStyle: 'short', timeStyle: 'short' }) : 'jamais';
    return `<strong>${s.senders}</strong> expéditeurs · <strong>${s.domains}</strong> domaines · <strong>${s.totalMails}</strong> mails · mis à jour le ${escapeHtml(lastUpdate)}`;
  }

  destroy(): void {
    this.container.innerHTML = '';
  }
}
