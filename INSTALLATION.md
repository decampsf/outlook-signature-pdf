# Installation du complément « Signer les PDF »

## Ce qu’il faut préparer

- Un PC sur lequel Outlook est déjà configuré avec le compte professionnel.
- Une image PNG de la signature de l’utilisatrice.
- Le fichier `manifest.xml` téléchargé depuis ce dépôt.

La signature, le nom et la fonction restent dans le stockage local du navigateur intégré à Outlook. Chaque utilisatrice dispose donc de ses propres données.

## 1. Télécharger le fichier d’installation

1. Ouvrir la page <https://github.com/decampsf/outlook-signature-pdf/blob/main/manifest.xml>.
2. Cliquer sur **Télécharger le fichier brut** ou **Download raw file**.
3. Vérifier que le fichier téléchargé s’appelle bien `manifest.xml` et non `manifest.xml.txt`.

## 2. Ajouter le complément à Outlook

1. Ouvrir <https://aka.ms/olksideload> dans le navigateur.
2. Se connecter avec le même compte professionnel que celui utilisé dans Outlook, si cela est demandé.
3. Dans la fenêtre **Compléments pour Outlook**, choisir **Mes compléments**.
4. Descendre jusqu’à la section **Compléments personnalisés**.
5. Cliquer sur **Ajouter un complément personnalisé**, puis **Ajouter à partir d’un fichier**.
6. Sélectionner le fichier `manifest.xml` téléchargé précédemment.
7. Accepter l’avertissement d’installation du complément personnalisé.

Le complément est associé au compte Outlook. Dans Outlook classique, son apparition peut prendre du temps à cause du cache ; fermer puis rouvrir Outlook peut accélérer son affichage.

## 3. Vérifier que le complément apparaît

1. Ouvrir dans Outlook un message reçu contenant au moins une pièce jointe PDF réelle. Un lien vers un fichier en ligne ne suffit pas.
2. Rechercher **Signer les PDF** dans la barre d’actions du message ou dans le menu **Applications** / **…**.
3. Ouvrir le complément. Le volet **Signer les PDF** doit apparaître à droite du message.

Si l’icône est proposée dans le menu **Applications**, elle peut être épinglée pour être plus facile à retrouver.

## 4. Effectuer la configuration personnelle

1. Dans **Ajouter une signature PNG**, sélectionner l’image de signature personnelle.
2. Dans **Nom et fonction affichés sous la signature**, remplacer la valeur existante par les informations de l’utilisatrice, par exemple `M.DUPONT-Cheffe de service`.
3. Cliquer ailleurs dans le volet pour enregistrer cette information.

Ces informations sont enregistrées uniquement sur ce PC. Elles ne sont pas ajoutées au dépôt GitHub et ne remplacent pas celles des autres utilisatrices.

## 5. Faire un essai complet

1. Ouvrir un message de test avec un PDF joint.
2. Lancer **Signer les PDF**.
3. Sélectionner le PDF et la signature.
4. Régler la taille, puis déplacer la signature à l’endroit souhaité.
5. Vérifier que le nom, la fonction, la date et l’heure apparaissent sous la signature.
6. Cliquer sur **Ajouter ici**. Répéter l’opération sur d’autres pages si nécessaire.
7. Cliquer sur **Préparer la réponse avec les PDF signés**.
8. Dans la réponse créée par Outlook, ouvrir le fichier dont le nom se termine par `- signé.pdf` et contrôler le résultat.
9. Vérifier les destinataires avant l’envoi.

## Dépannage

### Le complément n’apparaît pas

- Fermer complètement Outlook, puis le rouvrir.
- Vérifier que le compte utilisé sur <https://aka.ms/olksideload> est bien celui configuré dans Outlook.
- Dans Outlook classique, patienter : Microsoft indique que le cache peut retarder l’apparition d’un complément chargé manuellement.
- Vérifier dans **Mes compléments** que **Signer les PDF** figure toujours parmi les compléments personnalisés.

### Le volet indique qu’il faut l’ouvrir depuis Outlook

La page Web a été ouverte directement dans un navigateur. Il faut ouvrir le complément depuis un message reçu dans Outlook.

### Aucun PDF n’est détecté

- Vérifier que le message contient un fichier `.pdf` joint et non un simple lien OneDrive ou SharePoint.
- Ouvrir un message reçu, et non une fenêtre indépendante sans message sélectionné.

### La signature ou le nom ne sont plus mémorisés

Le stockage appartient au navigateur intégré à Outlook. Une suppression des données de navigation, une réinitialisation d’Outlook ou un changement de profil peut l’effacer. Il suffit alors de sélectionner de nouveau le PNG et de ressaisir le nom et la fonction.

### Mise à jour de l’application

Les évolutions de l’application sont publiées sur GitHub Pages. Il n’est normalement pas nécessaire de réinstaller le manifeste. Fermer puis rouvrir le volet, ou Outlook, permet de charger la version la plus récente.
