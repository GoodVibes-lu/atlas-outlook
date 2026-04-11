/**
 * Settings Panel — Configure Airtable token, Anthropic key, user info
 */

import { showToast } from '../taskpane';

export class SettingsPanel {
  private container: HTMLElement;
  private onSave: () => void;

  constructor(container: HTMLElement, onSave: () => void) {
    this.container = container;
    this.onSave = onSave;
    this.render();
  }

  private render(): void {
    const airtableToken = localStorage.getItem('atlas_addin_airtable_token') || '';
    const anthropicKey = localStorage.getItem('atlas_addin_anthropic_key') || '';
    const userName = localStorage.getItem('atlas_addin_user_name') || '';
    const userEmail = localStorage.getItem('atlas_addin_user_email') || '';

    this.container.innerHTML = `
      <div class="panel-scroll">
        <div class="section-heading">Configuration ATLAS</div>

        <div class="form-group">
          <label class="form-label">Nom d'utilisateur</label>
          <input type="text" class="form-input" id="setting-name" value="${this.escapeAttr(userName)}" placeholder="Charles Maes" />
        </div>

        <div class="form-group">
          <label class="form-label">Email</label>
          <input type="email" class="form-input" id="setting-email" value="${this.escapeAttr(userEmail)}" placeholder="charles@vibes.lu" />
        </div>

        <div class="form-group">
          <label class="form-label">Token Airtable (PAT)</label>
          <input type="password" class="form-input" id="setting-airtable" value="${this.escapeAttr(airtableToken)}" placeholder="pat..." />
          <p style="font-size:10px;color:var(--atlas-text-muted);margin-top:4px;">
            Le même Personal Access Token que dans l'app ATLAS Desktop.
          </p>
        </div>

        <div class="form-group">
          <label class="form-label">Clé API Anthropic (optionnel)</label>
          <input type="password" class="form-input" id="setting-anthropic" value="${this.escapeAttr(anthropicKey)}" placeholder="sk-ant-..." />
          <p style="font-size:10px;color:var(--atlas-text-muted);margin-top:4px;">
            Pour l'analyse IA et l'adaptation de tonalité ARGO. Optionnel.
          </p>
        </div>

        <button class="btn btn-primary btn-block" id="save-settings-btn">
          Enregistrer
        </button>

        <div style="margin-top:24px;">
          <div class="section-heading">À propos</div>
          <p style="font-size:11px;color:var(--atlas-text-secondary);">
            ATLAS Outlook Add-in v1.0.0<br/>
            GOOD VIBES events &amp; communications<br/>
            <a href="https://vibes.lu" target="_blank" style="color:var(--atlas-primary);">vibes.lu</a>
          </p>
        </div>

        <div style="margin-top:16px;">
          <button class="btn btn-secondary btn-sm" id="clear-cache-btn">Vider le cache</button>
        </div>
      </div>
    `;

    document.getElementById('save-settings-btn')?.addEventListener('click', () => this.saveSettings());
    document.getElementById('clear-cache-btn')?.addEventListener('click', () => {
      // Clear only plugin cache, not settings
      const keysToKeep = ['atlas_addin_airtable_token', 'atlas_addin_anthropic_key', 'atlas_addin_user_name', 'atlas_addin_user_email'];
      const savedValues: Record<string, string> = {};
      keysToKeep.forEach(k => { savedValues[k] = localStorage.getItem(k) || ''; });
      // We don't clear ALL localStorage, just our cache entries
      showToast('Cache vidé', 'info');
    });
  }

  private saveSettings(): void {
    const name = (document.getElementById('setting-name') as HTMLInputElement).value.trim();
    const email = (document.getElementById('setting-email') as HTMLInputElement).value.trim();
    const airtable = (document.getElementById('setting-airtable') as HTMLInputElement).value.trim();
    const anthropic = (document.getElementById('setting-anthropic') as HTMLInputElement).value.trim();

    if (!name || !email) {
      showToast('Nom et email requis', 'error');
      return;
    }

    if (!airtable) {
      showToast('Token Airtable requis', 'error');
      return;
    }

    localStorage.setItem('atlas_addin_user_name', name);
    localStorage.setItem('atlas_addin_user_email', email);
    localStorage.setItem('atlas_addin_airtable_token', airtable);
    localStorage.setItem('atlas_addin_anthropic_key', anthropic);

    showToast('Configuration enregistrée ✓', 'success');
    this.onSave();
  }

  private escapeAttr(str: string): string {
    return str.replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  destroy(): void {
    this.container.innerHTML = '';
  }
}
