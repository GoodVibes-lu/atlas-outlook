/**
 * roaming-storage.ts — Persistance des réglages du complément survivant au cache clear.
 *
 * Problème : localStorage Outlook Mac est wipé à chaque "rm -rf WebKitWebsiteData*"
 * (nécessaire pour appliquer les mises à jour du bundle JS). Charles devait
 * re-coller ses réglages à CHAQUE clear.
 *
 * SÉCURITÉ (03/10/2026) : la clé Anthropic et le jeton Airtable ne sont PLUS stockés (ni ici ni
 * en localStorage) : données et IA passent par le worker. `initRoamingStorage` efface ceux
 * laissés par les anciennes versions (purgeLegacySecrets).
 *
 * Solution : `Office.context.roamingSettings` — storage Microsoft qui :
 *   • Persiste dans la mailbox (côté serveur Exchange)
 *   • Survit aux cache clears locaux
 *   • Sync automatiquement entre Mac, Windows, Web (multi-device)
 *   • Limite 32KB total → largement assez
 *
 * Stratégie : lire roamingSettings au démarrage, hydrater localStorage,
 * et chaque setItem de clé écrit aussi dans roamingSettings (best effort).
 * Compatibilité ascendante : si la roamingSettings est vide, lit le
 * localStorage existant (premier run après migration).
 */

import { purgeLegacySecrets } from './worker';

const GRAPH_TOKEN_KEY = 'atlas_addin_graph_token';

/**
 * Jeton Graph collé à la main (Mac) : durée de vie ~1 h, donc gardé en sessionStorage seulement.
 * Plus de localStorage ni de roamingSettings (origine partagée goodvibes-lu.github.io, audit M26).
 * Il ne sert qu'aux appels directs à la boîte (graph.ts), jamais au worker (audit M8).
 */
export function readGraphToken(): string {
  try { return sessionStorage.getItem(GRAPH_TOKEN_KEY) || ''; } catch { return ''; }
}
export function saveGraphToken(value: string): void {
  try {
    if (value) sessionStorage.setItem(GRAPH_TOKEN_KEY, value);
    else sessionStorage.removeItem(GRAPH_TOKEN_KEY);
  } catch { /* stockage indisponible */ }
}

const PERSISTED_KEYS = [
  'atlas_addin_user_name',
  'atlas_addin_user_email',
] as const;

let roamingReady = false;

/**
 * Initialisation au démarrage. Hydrate localStorage depuis roamingSettings.
 * Si roamingSettings vide mais localStorage rempli (migration depuis ancien
 * code), copie le sens inverse.
 * Idempotent. À appeler dans Office.onReady().
 */
export async function initRoamingStorage(): Promise<void> {
  // Toujours en premier : les anciens secrets ne doivent pas rester sur le poste ni dans la boîte.
  await purgeLegacySecrets();
  await purgeStoredGraphToken();
  if (roamingReady) return;
  try {
    const rs = Office.context?.roamingSettings;
    if (!rs) {
      console.warn('[roaming-storage] Office.context.roamingSettings indispo');
      return;
    }

    let needsSave = false;
    for (const key of PERSISTED_KEYS) {
      const roamingVal = rs.get(key);
      const localVal = localStorage.getItem(key);

      if (roamingVal && !localVal) {
        // Restaure depuis roaming (cas typique : cache local wipé)
        localStorage.setItem(key, roamingVal);
      } else if (!roamingVal && localVal) {
        // Migration : on a une valeur en local mais pas en roaming → push
        rs.set(key, localVal);
        needsSave = true;
      }
      // Sinon : déjà sync (both set OR both empty)
    }

    if (needsSave) {
      await new Promise<void>((resolve) => {
        rs.saveAsync((res) => {
          if (res.status !== Office.AsyncResultStatus.Succeeded) {
            console.warn('[roaming-storage] saveAsync failed:', res.error);
          }
          resolve();
        });
      });
    }

    roamingReady = true;
    console.info('[roaming-storage] hydraté depuis roamingSettings');
  } catch (e) {
    console.warn('[roaming-storage] init failed:', e);
  }
}

/** Efface les copies persistantes du jeton Graph laissées par les versions précédentes. */
async function purgeStoredGraphToken(): Promise<void> {
  try { localStorage.removeItem(GRAPH_TOKEN_KEY); } catch { /* rien */ }
  try {
    const rs = Office.context?.roamingSettings;
    if (!rs || rs.get(GRAPH_TOKEN_KEY) == null) return;
    rs.remove(GRAPH_TOKEN_KEY);
    await new Promise<void>((resolve) => rs.saveAsync(() => resolve()));
  } catch { /* rien */ }
}

/**
 * Écrit une clé dans localStorage ET roamingSettings (best effort).
 * À utiliser dans Settings après save.
 */
export async function persistKey(key: string, value: string): Promise<void> {
  try {
    if (value) localStorage.setItem(key, value);
    else localStorage.removeItem(key);

    const rs = Office.context?.roamingSettings;
    if (!rs) return;
    if (value) rs.set(key, value);
    else rs.remove(key);

    await new Promise<void>((resolve) => {
      rs.saveAsync((res) => {
        if (res.status !== Office.AsyncResultStatus.Succeeded) {
          console.warn('[roaming-storage] persistKey saveAsync failed:', res.error);
        }
        resolve();
      });
    });
  } catch (e) {
    console.warn('[roaming-storage] persistKey failed:', e);
  }
}
