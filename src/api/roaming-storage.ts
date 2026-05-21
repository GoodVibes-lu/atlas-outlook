/**
 * roaming-storage.ts — Persistance des clés API survivant au cache clear.
 *
 * Problème : localStorage Outlook Mac est wipé à chaque "rm -rf WebKitWebsiteData*"
 * (nécessaire pour appliquer les mises à jour du bundle JS). Charles devait
 * re-coller PAT Airtable, clé Anthropic, token Graph à CHAQUE clear.
 *
 * Solution : `Office.context.roamingSettings` — storage Microsoft qui :
 *   • Persiste dans la mailbox (côté serveur Exchange)
 *   • Survit aux cache clears locaux
 *   • Sync automatiquement entre Mac, Windows, Web (multi-device)
 *   • Limite 32KB total → largement assez pour 3 clés
 *
 * Stratégie : lire roamingSettings au démarrage, hydrater localStorage,
 * et chaque setItem de clé écrit aussi dans roamingSettings (best effort).
 * Compatibilité ascendante : si la roamingSettings est vide, lit le
 * localStorage existant (premier run après migration).
 */

const PERSISTED_KEYS = [
  'atlas_addin_airtable_token',
  'atlas_addin_anthropic_key',
  'atlas_addin_graph_token',
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
