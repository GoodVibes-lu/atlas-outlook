/**
 * tableau-volet.ts : le tableau de bord ATLAS DANS LE VOLET du complément (10/10/2026).
 *
 * Pourquoi : la fenêtre de dialogue Office (displayDialogAsync) restait au-dessus de toutes les
 * applications sur le nouvel Outlook pour Mac, et l'ouverture dans le navigateur a été refusée
 * (l'équipe travaille dans l'application Outlook, jamais sur le web). Le volet épinglé fait partie
 * de la fenêtre d'Outlook : il ne flotte jamais au-dessus des autres applications.
 *
 * Fonctionnement : le volet a deux modes. « Ce mail » (#app, les onglets habituels) et « Tableau de
 * bord » (#tdb-volet, ce module). Le tableau est démarré une seule fois puis gardé en mémoire :
 * changer de mail (volet épinglé, `SupportsPinning`, Mailbox 1.5, pris en charge sur Mac) ne le
 * recharge pas, revenir sur « Ce mail » le met en sommeil (ni clavier ni rafraîchissement).
 * Les mails s'ouvrent dans l'application Outlook (`displayMessageForm`, Mailbox 1.1), la réponse
 * préparée dans le formulaire d'Outlook (`displayReplyForm`, Mailbox 1.1) ; « Envoyer » passe par le
 * worker comme avant. Mise en page : celle du tableau en largeur étroite (une colonne, détail en
 * écran suivant avec « Retour »), src/tableau.css.
 */

import './tableau.css';
import { demarrerTableau } from './tableau/app';
import { idDuLienOutlook, ouvrirDepuisTableau, repondreSurElementCourant } from './api/dialogue-tableau';
import { icon } from './ui/icons';

let hote: HTMLElement | null = null;
let demarre = false;

/** Ouvre une adresse demandée par le tableau : un mail dans l'application Outlook, jamais Outlook sur le web. */
function ouvrirLien(url: string, onInfo: (m: string, t?: 'success' | 'error' | 'info') => void): void {
  const outlookWeb = (() => { try { return /(^|\.)outlook\.(office|office365|live)\.com$/i.test(new URL(url).hostname); } catch { return false; } })();
  if (outlookWeb && !idDuLienOutlook(url)) {
    onInfo('Ce mail ne peut pas être ouvert d\'ici : retrouve-le dans ta boîte Outlook.', 'info');
    return;
  }
  ouvrirDepuisTableau(url);
}

export interface OptionsVolet {
  /** Retour au mode « Ce mail ». */
  onRetour: () => void;
  /** Message bref dans le volet. */
  onInfo: (message: string, type?: 'success' | 'error' | 'info') => void;
  /** Ouvre directement la séance de tri. */
  seance?: boolean;
}

/** Affiche le tableau de bord dans le volet (démarré au premier appel, ensuite simplement réaffiché). */
export function afficherTableauVolet(opts: OptionsVolet): void {
  if (!hote) {
    hote = document.createElement('div');
    hote.id = 'tdb-volet';
    hote.className = 'tdb-volet';
    document.body.appendChild(hote);
  }
  hote.hidden = false;
  document.body.classList.add('tb', 'en-tableau');

  if (!demarre) {
    demarre = true;
    if (opts.seance) { try { sessionStorage.setItem('atlas_tdb_seance', '1'); } catch { /* stockage indisponible */ } }
    demarrerTableau(hote, {
      openLink: (url) => ouvrirLien(url, opts.onInfo),
      repondreDansOutlook: async (messageId, html) => repondreSurElementCourant(messageId, html),
    });
    // Retour « Ce mail » en tête de la barre du tableau.
    const barre = hote.querySelector('.tb-top');
    if (barre) {
      const retour = document.createElement('button');
      retour.type = 'button';
      retour.className = 'tb-btn tdb-volet-retour';
      retour.id = 'tdb-volet-retour';
      retour.title = 'Revenir au mail affiché';
      retour.innerHTML = `${icon('chevron-right', 14)}<span>Ce mail</span>`;
      retour.addEventListener('click', () => opts.onRetour());
      barre.insertBefore(retour, barre.firstChild);
    }
  } else if (opts.seance) {
    hote.querySelector<HTMLButtonElement>('#tb-seance-btn')?.click();
  }
}

/** Met le tableau en sommeil (caché, gardé en mémoire) et rend le volet au mode « Ce mail ». */
export function masquerTableauVolet(): void {
  if (hote) hote.hidden = true;
  document.body.classList.remove('tb', 'en-tableau');
}
