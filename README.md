# Relance IA Outlook Add-in (Office.js)

Add-in Outlook orienté **relance manuelle 1 clic** + **auto-relance assistée**.

## Fonctionnalités incluses

- Import Excel `.xlsx` (première feuille).
- Dashboard clients avec statuts : `À relancer`, `Relancé`, `En attente`, `Clôturé`.
- **Bouton Relancer** : ouvre un email Outlook natif pré-rempli (`To`, `Subject`, `Body`).
- **Relance automatique (assistée)** : calcule la file des clients à relancer puis ouvre des brouillons prêts à valider.
- Templates intelligents :
  - Douce (J+3)
  - Standard (J+7)
  - Ferme (J+14)
- Persistance locale via `localStorage`.
- Commentaire d'architecture pour mode avancé Microsoft Graph API (backend requis).

## Structure

- `manifest.xml`
- `taskpane.html`
- `taskpane.js`
- `commands.html`
- `styles.css`

## Pré-requis

- Outlook Desktop (Windows/Mac) ou Outlook on the web.
- Un serveur HTTPS local exposant les fichiers (ex: `https://localhost:3000`).
- Certificat local approuvé (pour sideload Outlook desktop).

## Démarrage local (exemple)

Servez les fichiers depuis `https://localhost:3000` avec n'importe quel serveur statique HTTPS.

> Les URLs du manifest pointent vers `https://localhost:3000/*`.

## Sideload - Outlook Web

1. Ouvrez Outlook Web.
2. Paramètres ⚙️ → **Gérer les intégrations** / **Mes compléments**.
3. **Ajouter un complément personnalisé** → **Ajouter à partir d’un fichier**.
4. Sélectionnez `manifest.xml`.
5. Ouvrez un email : le bouton **Ouvrir Relance IA** apparaît dans le ruban.

## Sideload - Outlook Desktop

1. Assurez-vous que `https://localhost:3000` est accessible et que le certificat est approuvé.
2. Outlook → **Obtenir des compléments** → **Mes compléments**.
3. **Ajouter un complément personnalisé** → **Ajouter à partir d’un fichier**.
4. Sélectionnez `manifest.xml`.

## Contrainte Outlook respectée

- Office.js **ne peut pas** envoyer un mail automatiquement sans interaction utilisateur.
- L’add-in implémente donc :
  - mode natif pré-rempli,
  - mode auto-relance assisté (brouillons en chaîne),
  - et documente le mode backend + Graph API pour l’envoi automatique réel.
