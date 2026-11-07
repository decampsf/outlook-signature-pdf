# Signature de documents dans Outlook

Complément personnel pour signer les PDF, Word (.docx) et Excel (.xlsx) joints à un message, puis préparer une réponse avec les fichiers signés.

- PDF : placement du tampon sur les pages.
- Word : insertion dans un paragraphe, y compris dans un tableau.
- Excel : insertion dans une feuille à partir d’une cellule, avec conservation des formules et graphiques.
- Choix par fichier Office : format d’origine ou PDF converti par Microsoft 365.
- La réponse reste à vérifier et à envoyer par l’utilisateur.

[Installation et utilisation](INSTALLATION.md) · [Configuration Microsoft 365](CONFIGURATION_MICROSOFT_365.md)

## Statut de la conversion

L’inscription Microsoft Entra est à configurer dans `config.js`. La connexion et la conversion réelles doivent ensuite être testées dans le tenant professionnel. La présence de OneDrive dans le compte ne configure pas automatiquement le complément.

## Développement

```sh
npm ci
npm run build
npx playwright install chromium
python -m pip install python-docx openpyxl pillow lxml
npm test
```

Les scripts générés dans `vendor/` sont versionnés pour l’hébergement GitHub Pages sans compilation côté serveur. Les dépendances JSZip et MSAL sont épinglées dans `package-lock.json`. `vendor/THIRD_PARTY_LICENSES.txt` conserve leurs mentions de licence.

Les tests génèrent des documents fictifs, signent des DOCX/XLSX avec les moteurs réels, les relisent avec python-docx/openpyxl et vérifient la conservation des autres parties des fichiers. Les parcours Outlook/Graph sont simulés : préparation de la réponse sans envoi, conversion, nettoyage en cas d’erreur, absence du jeton Graph sur le téléchargement externe. Aucun document utilisateur ni identifiant Microsoft réel n’est nécessaire aux tests.
