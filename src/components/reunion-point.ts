/**
 * reunion-point.ts · onglet « Proposer un point » d'une invitation de réunion ATLAS (04/10/2026).
 *
 * Affiché quand l'élément ouvert est une invitation / un rendez-vous (src/api/reunion.ts). Si ATLAS
 * ne suit pas cette réunion, l'onglet l'explique et ne propose rien. Sinon : formulaire (titre,
 * durée, projet lié, pièce à projeter) jusqu'à la veille 17 h, liste de MES propositions avec leur
 * sort, et « confier mon point » (collègue, note écrite ou audio de 3 minutes au plus). Aucun
 * innerHTML avec une valeur saisie ou reçue : tout passe par textContent.
 */
import { lireElementReunion, resoudreReunion, proposerPoint, retirerPoint, deleguerPoint, deposerAudio, chargerReunion, type ReunionVue, type ReunionPoint } from '../api/reunion';
import { getAllEmployes } from '../api/airtable';
import { humanError } from '../api/net';

const STATUT: Record<string, string> = { proposee: 'En attente', acceptee: 'Accepté', refusee: 'Refusé', reportee: 'Reporté à la prochaine réunion' };
const fmt = (iso: string) => new Date(iso).toLocaleString('fr-FR', { timeZone: 'Europe/Luxembourg', weekday: 'long', day: 'numeric', month: 'long', hour: '2-digit', minute: '2-digit' });

function el<K extends keyof HTMLElementTagNameMap>(tag: K, cls?: string, text?: string): HTMLElementTagNameMap[K] {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text != null) e.textContent = text;
  return e;
}
function champ(label: string, input: HTMLElement): HTMLElement {
  const g = el('div', 'form-group');
  g.appendChild(el('label', undefined, label));
  g.appendChild(input);
  return g;
}
function saisie(placeholder: string, max: number, type = 'text'): HTMLInputElement {
  const i = el('input', 'form-input');
  i.type = type; i.placeholder = placeholder; i.maxLength = max;
  return i;
}

export class ReunionPointPanel {
  private root: HTMLElement;
  private vue: ReunionVue | null = null;
  private destroyed = false;

  constructor(host: HTMLElement) {
    this.root = el('div', 'agent-panel');
    host.innerHTML = '';
    host.appendChild(this.root);
    void this.charger();
  }

  destroy(): void { this.destroyed = true; this.root.remove(); }

  private message(t: string): void {
    this.root.innerHTML = '';
    this.root.appendChild(el('p', 'agent-muted', t));
  }

  private async charger(): Promise<void> {
    this.message('Recherche de la réunion dans ATLAS…');
    const element = lireElementReunion();
    if (!element) { this.message('Ouvre une invitation ou un rendez-vous de réunion pour proposer un point.'); return; }
    try {
      const v = await resoudreReunion(element);
      if (this.destroyed) return;
      if (!v) { this.message('Cette réunion n\'est pas suivie par ATLAS : l\'ordre du jour collaboratif n\'est pas ouvert ici.'); return; }
      this.vue = v;
      this.dessiner();
    } catch (e) { this.message(`ATLAS est injoignable pour le moment : ${humanError(e)}.`); }
  }

  private async recharger(): Promise<void> {
    if (!this.vue) return;
    try { this.vue = await chargerReunion(this.vue.key); this.dessiner(); } catch { /* garde l'affichage */ }
  }

  private dessiner(): void {
    const v = this.vue!;
    this.root.innerHTML = '';
    const tete = el('div', 'agent-state');
    tete.appendChild(el('strong', undefined, v.titre));
    tete.appendChild(el('span', 'agent-muted', fmt(v.debutISO)));
    tete.appendChild(el('span', 'agent-muted', v.ouverte ? `Propositions ouvertes jusqu'à ${fmt(v.limiteISO)}.` : 'Les propositions sont closes : demande directement à l\'organisateur.'));
    if (v.estOrganisateur) tete.appendChild(el('span', 'agent-muted', 'Tu es l\'organisateur : l\'arbitrage (accepter, refuser, reporter) se fait dans ATLAS.'));
    this.root.appendChild(tete);
    if (v.ouverte) this.root.appendChild(this.formulaire(v));
    if (v.mesPoints.length) {
      this.root.appendChild(el('strong', undefined, 'Mes propositions'));
      for (const p of v.mesPoints) this.root.appendChild(this.carte(p, v));
    }
  }

  private formulaire(v: ReunionVue): HTMLElement {
    const f = el('div', 'agent-state');
    f.appendChild(el('strong', undefined, 'Proposer un point'));
    const titre = saisie('Titre du point', 120);
    const duree = saisie('10', 3, 'number'); duree.value = '10'; duree.min = '1'; duree.max = '180';
    const projet = saisie('N° de projet (facultatif)', 6);
    const pieceNom = saisie('Nom de la pièce (facultatif)', 120);
    const pieceUrl = saisie('Lien https de la pièce (cloud ATLAS)', 500);
    const err = el('div', 'agent-muted');
    const go = el('button', 'btn btn-primary', 'Proposer');
    go.type = 'button';
    go.addEventListener('click', async () => {
      go.disabled = true; err.textContent = '';
      try {
        this.vue = await proposerPoint(v.key, { titre: titre.value, dureeMin: Number(duree.value), projetNo: projet.value.replace(/\D/g, ''), pieceNom: pieceNom.value, pieceUrl: pieceUrl.value.trim() });
        this.dessiner();
      } catch (e) { err.textContent = humanError(e); go.disabled = false; }
    });
    f.append(champ('Titre', titre), champ('Durée (minutes)', duree), champ('Projet lié', projet), champ('Pièce à projeter', pieceNom), champ('Lien de la pièce', pieceUrl), go, err);
    return f;
  }

  private carte(p: ReunionPoint, v: ReunionVue): HTMLElement {
    const c = el('div', 'agent-state');
    const t = el('div'); t.appendChild(el('strong', undefined, p.titre)); t.appendChild(el('span', 'agent-muted', ` · ${p.dureeMin} min${p.projetNo ? ` · projet #${p.projetNo}` : ''}`));
    c.appendChild(t);
    c.appendChild(el('span', 'agent-muted', STATUT[p.statut] || p.statut));
    if (p.motif) c.appendChild(el('span', 'agent-muted', `Organisateur : ${p.motif}`));
    if (p.delegation) {
      const d = p.delegation;
      c.appendChild(el('span', 'agent-muted', d.mode === 'collegue' ? `Confié à ${d.nom || d.email}` : d.mode === 'note' ? 'Note laissée pour la lecture en séance' : 'Audio laissé pour la réunion'));
    }
    const actions = el('div'); actions.style.cssText = 'display:flex;flex-wrap:wrap;gap:6px';
    const zone = el('div', 'agent-state');
    const bouton = (label: string, fn: () => void) => { const b = el('button', 'btn btn-secondary btn-sm', label); b.type = 'button'; b.addEventListener('click', fn); actions.appendChild(b); };
    if (p.statut !== 'refusee') {
      bouton('Confier à un collègue', () => this.choixCollegue(p, zone));
      bouton('Laisser une note', () => this.choixNote(p, zone));
      bouton('Laisser un audio', () => this.choixAudio(p, zone));
    }
    if (p.statut === 'proposee' && v.ouverte) bouton('Retirer', async () => { try { await retirerPoint(p.id); await this.recharger(); } catch (e) { zone.textContent = humanError(e); } });
    c.append(actions, zone);
    return c;
  }

  private choixCollegue(p: ReunionPoint, zone: HTMLElement): void {
    zone.innerHTML = '';
    const sel = el('select', 'form-input');
    sel.appendChild(Object.assign(el('option', undefined, 'Choisir un collègue…'), { value: '' }));
    const msg = el('div', 'agent-muted');
    zone.append(sel, msg);
    getAllEmployes().then(list => {
      for (const e of list.filter(x => /@vibes\.lu$/i.test(x.email || ''))) sel.appendChild(Object.assign(el('option', undefined, e.name), { value: e.email }));
    }).catch(() => { msg.textContent = 'Liste des collègues indisponible.'; });
    sel.addEventListener('change', async () => {
      if (!sel.value) return;
      try { await deleguerPoint(p.id, { mode: 'collegue', email: sel.value, nom: sel.selectedOptions[0]?.textContent || '' }); await this.recharger(); } catch (e) { msg.textContent = humanError(e); }
    });
  }

  private choixNote(p: ReunionPoint, zone: HTMLElement): void {
    zone.innerHTML = '';
    const ta = el('textarea', 'form-input'); ta.maxLength = 2000; ta.rows = 4; ta.placeholder = 'Ta note : elle sera lue en séance par le présentateur.';
    const go = el('button', 'btn btn-primary btn-sm', 'Enregistrer la note'); go.type = 'button';
    const msg = el('div', 'agent-muted');
    go.addEventListener('click', async () => {
      go.disabled = true;
      try { await deleguerPoint(p.id, { mode: 'note', texte: ta.value }); await this.recharger(); } catch (e) { msg.textContent = humanError(e); go.disabled = false; }
    });
    zone.append(ta, go, msg);
  }

  private choixAudio(p: ReunionPoint, zone: HTMLElement): void {
    zone.innerHTML = '';
    const msg = el('div', 'agent-muted', 'Trois minutes au plus. Le fichier est rangé sur le cloud, jamais sur le NAS.');
    const go = el('button', 'btn btn-primary btn-sm', 'Enregistrer'); go.type = 'button';
    zone.append(go, msg);
    let rec: MediaRecorder | null = null; let t0 = 0; let timer: number | undefined;
    const fin = () => { try { rec?.stop(); } catch { /* */ } if (timer) window.clearInterval(timer); };
    go.addEventListener('click', async () => {
      if (rec && rec.state === 'recording') { fin(); return; }
      try {
        const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
        const morceaux: Blob[] = [];
        rec = new MediaRecorder(stream, { audioBitsPerSecond: 32000 });
        rec.ondataavailable = e => { if (e.data.size) morceaux.push(e.data); };
        rec.onstop = async () => {
          stream.getTracks().forEach(t => t.stop());
          const blob = new Blob(morceaux, { type: rec?.mimeType || 'audio/webm' });
          if (blob.size > 2_400_000) { msg.textContent = 'Audio trop long : 3 minutes au plus.'; go.textContent = 'Enregistrer'; return; }
          const b64: string = await new Promise(res => { const fr = new FileReader(); fr.onloadend = () => { const r = String(fr.result || ''); res(r.slice(r.indexOf(',') + 1)); }; fr.readAsDataURL(blob); });
          try { msg.textContent = 'Envoi…'; await deposerAudio(p.id, b64, Math.max(1, Math.round((Date.now() - t0) / 1000))); await this.recharger(); } catch (e) { msg.textContent = humanError(e); go.textContent = 'Enregistrer'; }
        };
        t0 = Date.now(); rec.start(); go.textContent = 'Arrêter et envoyer (0 s)';
        timer = window.setInterval(() => { const s = Math.round((Date.now() - t0) / 1000); go.textContent = `Arrêter et envoyer (${s} s)`; if (s >= 180) fin(); }, 500);
      } catch { msg.textContent = 'Micro indisponible dans ce volet Outlook : laisse plutôt une note écrite.'; }
    });
  }
}
