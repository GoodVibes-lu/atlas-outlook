# Application ATLAS dans la barre de gauche d'Outlook (tableau de bord de l'agent, phase 2)

Ce dossier contient l'application ATLAS épinglée dans la barre de gauche d'Outlook (comme Calendrier ou To Do), décidée le 03/10/2026 pour le tableau de bord de l'agent d'inbox (`.claude/BACKLOG-AGENT-INBOX.md` §7). Squelette d'essai en phase 1.5, **tableau de bord réel en phase 2** (page `../tableau-de-bord.html`).

Ce n'est **pas** le complément : le panneau ATLAS (lecture et rédaction) reste déclaré par `../manifest.xml` (manifeste XML classique, seul pris en charge sur mobile). L'application de la barre de gauche demande un **second manifeste**, au format « application Microsoft 365 » (`manifest.json` de ce dossier : un onglet personnel, `staticTabs`, portée `personal`).

## Contenu

| Fichier | Rôle |
|---|---|
| `manifest.json` | Manifeste d'application Microsoft 365 (schéma 1.17), un onglet personnel « ATLAS » → `tableau-de-bord.html` |
| `color.png` | Icône couleur 192 × 192 (copie de `../assets/icon-color.png`) |
| `outline.png` | Icône contour 32 × 32, blanche sur fond transparent (copie de `../assets/icon-outline.png`) |
| `../tableau-de-bord.html` + `../src/tableau-de-bord.ts` | La page affichée (publiée avec le complément, même build Vite) |

## Ce que montre le tableau de bord (0.3.0, 06/10/2026)

Page pleine largeur (`../src/tableau/`), langage visuel ATLAS, clavier d'abord. Détail du cadrage : `.claude/CADRAGE-OUTLOOK-DASHBOARD.md` (B, C, D).

- **Boîte triée, toutes les boîtes autorisées** (la sienne ; good@ pour ses membres) : priorités ARGO (score expliqué), personnes et clients, mandats / associations (section masquée si vide, aucune action commerciale), en attente, mis de côté, factures, newsletters en digest (désinscription), notifications. `GET /api/plugin/agent/tableau`, empreinte relue toutes les 20 s (`tableau/version`) : mise à jour sans rechargement, nouveautés signalées.
- **Un mail** : résumé et raisons, ouverture dans Outlook (Entrée), actions proposées par l'agent (`agent/actions`), **réponse ARGO** dans le style de la personne ou **modèle Communications** (`tableau/reponse`), **déposée dans le brouillon Outlook** (`tableau/brouillon`), **envoi programmé** par remise différée d'Exchange avec relance facultative (`tableau/envois`), « mettre de côté » (`agent/plus-tard`), « relancer si pas de réponse » (`agent/relancer`), rattacher à un projet (`atlas/emails/link-projet`).
- **Équipe** (good@) : « Je prends », « Attribuer à… » (cloche à la personne), commentaires internes, qui traite quoi ; la prise est propagée aux cloches (`data.priseInbox`).
- **Aujourd'hui** : prochains rendez-vous externes avec la préparation de meeting-prep (`tableau/rdv`), engagements promis / attendus vérifiés (`tableau/engagements`), réactivité aux clients en heures ouvrées (`tableau/reactivite`), question à sa boîte et rattrapage.
- **Clavier** : j/k, Entrée, r (ARGO), t (modèle), e / s (de côté), p (je prends), l (projet), c (commentaire), 1 à 9 (sections), ⌘K ou / (palette), ? (aide). **Glisser-déposer** d'un mail vers « Je prends », une date ou un projet ; aperçu au survol ; thème d'Outlook suivi ; mouvement réduit respecté.
- **Écritures dans une boîte** (brouillon, envoi programmé) : seulement sur clic, sous le verrou `INBOX_AGENT_TABLEAU_ECRITURES=actif` du worker ; fermé, la page propose « Copier ».

## Connexion

Même jeton que le panneau : `getWorkerToken` (`../src/api/worker.ts`), voie **nested app authentication** (MSAL `createNestablePublicClientApplication`, application Entra « ATLAS Outlook (complément) », URI de redirection SPA `brk-multihub://goodvibes-lu.github.io` déjà déclarée). La page initialise d'abord **TeamsJS** (chargé depuis le CDN Microsoft, version à vérifier à la publication) puis active cette voie hors Office.js (`enableHostedNaa()`).

**Limite** : la connexion automatique dans une application de la barre de gauche dépend de l'hôte (nouvel Outlook Windows et Outlook sur le web récents : prévu ; Mac : à confirmer). Si l'hôte ne la fournit pas, MSAL tente une fenêtre de connexion ; si elle est bloquée, la page affiche « Connexion impossible » avec « Réessayer » et « Ouvrir ATLAS » (le bandeau « Ma journée » du panneau reste disponible). Pour que la fenêtre de repli fonctionne, ajouter aussi `https://goodvibes-lu.github.io/atlas-outlook/tableau-de-bord.html` comme URI de redirection **SPA** de l'application Entra. Le SSO Teams classique (`getAuthToken`, qui exigerait de pré-autoriser les clients Teams / Outlook sur l'application Entra) n'est pas utilisé.

## Avant l'essai

1. **Publier le complément** (build Vite) : `tableau-de-bord.html` est une entrée du build (`vite.config.ts`) et doit répondre sur `https://goodvibes-lu.github.io/atlas-outlook/tableau-de-bord.html` (HTTPS, affichable dans un cadre : pas d'en-tête `X-Frame-Options: DENY`).
2. **`webApplicationInfo`** : renseigné avec l'ID d'application (client) de « ATLAS Outlook (complément) » (`fc36080c-…`, public) et son URI d'ID d'application (`docs/agent-inbox-entra-id.md`, étape 5) : même application que le complément, le worker accepte donc le même jeton. Pour un essai d'**affichage** seul, on peut retirer le bloc.
3. **`version`** : à augmenter à chaque nouvel envoi du paquet (le Centre d'administration refuse un paquet de même version). Actuelle : 0.3.0 (tableau de bord pleine largeur, 06/10/2026).
4. Le champ `id` (GUID) est propre à cette application : ne pas le changer entre deux versions, ne pas réutiliser celui du complément.

## Construire le paquet

Le paquet est un **zip** contenant les trois fichiers à la racine (pas de dossier) :

```bash
cd outlook-addin/teams-app
zip -X ../ATLAS-barre-gauche-0.3.0.zip manifest.json color.png outline.png
```

Validation : schéma Microsoft 1.17 (`$schema` du manifeste), vérifié avec ajv avant chaque envoi.

(Le zip n'est pas versionné ; aucune étape de build ni de déploiement n'est lancée par le dépôt.)

## Installer pour l'essai (une seule personne)

- **Charger l'application personnalisée** : dans Teams, **Applications** › **Gérer vos applications** › **Charger une application** › **Charger une application personnalisée** › choisir le zip. Il faut que le « chargement d'applications personnalisées » soit autorisé pour ce compte (Centre d'administration Teams › **Applications Teams** › **Stratégies de configuration**, ou stratégie d'autorisation d'applications).
- **Ou** par le **Centre d'administration Microsoft 365** › **Paramètres** › **Applications intégrées** › **Charger des applications personnalisées** › type « application Teams » › le zip › attribuer à **un seul utilisateur** (Charles) pour l'essai.
- L'application apparaît ensuite dans la barre de gauche d'Outlook (bouton **Autres applications** « … », puis **épingler**), après quelques minutes à quelques heures.
- Mise à jour : nouvelle `version` dans `manifest.json`, nouveau zip, « Mettre à jour » au même endroit. Une simple modification de la page (`tableau-de-bord.html`) ne demande **pas** de nouveau paquet : il suffit de republier le complément.

## Où l'essayer

| Application | Attendu (documentation Microsoft, à vérifier à l'essai) |
|---|---|
| Nouvel Outlook Windows | oui |
| Outlook sur le web | oui |
| Outlook Mac (nouvel Outlook) | **non** (Microsoft : pas d'applications Teams dans la barre de gauche sur Mac) → grande fenêtre du complément, ci-dessous |
| Outlook classique Windows | a priori non → bandeau « Ma journée » du panneau |
| Outlook iPhone / Android | non → bandeau « Ma journée » du panneau |

La liste de contrôle complète est dans `docs/agent-inbox-essai-complement.md` (sections 3 et 5).

## Outlook Mac : le tableau de bord en grande fenêtre (07/10/2026)

Le nouvel Outlook pour Mac n'affiche pas les applications Teams / Microsoft 365 dans la barre de gauche. Sur Mac (et partout ailleurs aussi), le **complément** ouvre la même page dans une grande fenêtre Office (`displayDialogAsync`, 95 % × 90 % de l'écran) :

- **Ruban** : menu **ATLAS** d'un mail › **Tableau de bord** (commande `atlasTableauCommand`, manifeste du complément 1.4.0) ;
- **Panneau ATLAS** : bouton **Tableau de bord** en tête.

Même page (`tableau-de-bord.html?hote=office`), même design, mêmes raccourcis, même mise à jour en temps réel. Différences propres à la fenêtre (`../src/api/dialogue-tableau.ts`) :

- **Connexion** : le jeton est fourni par la page qui a ouvert la fenêtre (panneau ou commande du ruban, sa connexion habituelle), par `messageParent` / `messageChild` (DialogApi 1.2). Si elle ne répond pas, la fenêtre se connecte elle-même par **redirection** Microsoft : il faut alors que `https://goodvibes-lu.github.io/atlas-outlook/tableau-de-bord.html` soit déclarée comme URI de redirection **SPA** de l'application Entra « ATLAS Outlook (complément) ».
- **Ouvrir un mail** : la fenêtre demande à la page parente de l'ouvrir dans Outlook (`displayMessageForm`) ; sinon le lien Outlook sur le web.
- **Durée de vie** : la fenêtre vit tant que sa page parente vit. Ouverte depuis le panneau, elle se ferme si le panneau se ferme ou se recharge (changer de mail sans panneau épinglé) ; ouverte depuis le ruban, elle reste ouverte (la commande se termine à sa fermeture).
