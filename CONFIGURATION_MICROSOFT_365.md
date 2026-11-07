# Activer la conversion Word / Excel vers PDF

Le complément signe les PDF, DOCX et XLSX sur le poste. Seuls l’aperçu PDF d’un document Office et son renvoi en PDF nécessitent Microsoft 365. L’accès de l’utilisateur à OneDrive ne suffit pas : le complément doit disposer de sa propre inscription Microsoft Entra.

## Configuration par le responsable Microsoft 365

1. Ouvrir [Microsoft Entra](https://entra.microsoft.com/) avec le compte professionnel.
2. Aller dans **Entra ID → Inscriptions d’applications → Nouvelle inscription**. Si cette fonction est interdite, demander l’inscription à la DSIN.
3. Nom : **Signature de documents Outlook**.
4. Comptes pris en charge : **Comptes dans cet annuaire d’organisation uniquement** (monolocataire).
5. URI de redirection : choisir la plateforme **Application monopage (SPA)** et saisir exactement :

   `https://decampsf.github.io/outlook-signature-pdf/auth.html`

6. Dans **Autorisations API**, ajouter **Microsoft Graph → Autorisations déléguées → Files.ReadWrite**. Ne pas ajouter d’autorisation d’application. Si `User.Read` a été ajouté par défaut, il n’est pas utilisé par ce complément et peut être retiré.
7. Si la politique de l’association exige un consentement administrateur, le faire accorder par un administrateur habilité.
8. Relever dans **Vue d’ensemble** les deux identifiants publics : **ID d’application (client)** et **ID de l’annuaire (locataire)**.
9. Les reporter dans `config.js` :

   ```js
   window.M365_CONFIG = Object.freeze({
     clientId: 'ID-APPLICATION-CLIENT',
     tenantId: 'ID-ANNUAIRE-LOCATAIRE'
   });
   ```

Aucun secret client, mot de passe, certificat ni clé privée n’est nécessaire. Ne jamais en ajouter dans GitHub. Le flux MSAL utilise le code d’autorisation avec PKCE dans une fenêtre de dialogue Office ; ne pas activer l’ancien flux implicite.

## Données et autorisations

- `Files.ReadWrite` autorise techniquement l’accès en lecture/écriture aux fichiers du compte connecté. Le code l’utilise pour ses propres copies temporaires ; il ne parcourt pas les documents existants.
- Chaque conversion dépose une copie **déjà signée**, nommée `Signature-Outlook-temp-<identifiant>.docx` ou `.xlsx`, à la racine du OneDrive du compte affiché dans le complément.
- Le fichier original joint au message n’est jamais modifié ni remplacé.
- Après la conversion (ou après une erreur), le complément demande la suppression de sa copie. OneDrive la déplace dans la corbeille ; il ne s’agit pas d’un effacement définitif. Les politiques de rétention de l’organisation restent applicables.
- En cas d’échec du nettoyage, un avertissement est affiché. Une reconnexion au même compte réessaie la suppression. Seuls les noms générés par le complément sont éligibles à ce nettoyage.
- Fermer Outlook pendant une conversion peut interrompre le nettoyage. Les noms des copies en attente, sans contenu ni jeton, sont mémorisés localement. Vérifier OneDrive si les données locales ont été effacées.
- Le jeton d’accès reste en mémoire dans le volet ; MSAL utilise le stockage de session de la fenêtre d’authentification pour son protocole. Le bouton de déconnexion vide le jeton du volet, sans déconnecter Office ni le navigateur de Microsoft.
- Les documents et signatures ne sont pas envoyés au dépôt GitHub. La sortie DOCX/XLSX n’a pas besoin de OneDrive.

## Contrôle avant mise en service

L’intégration ne peut pas être déclarée opérationnelle tant que ces essais n’ont pas été effectués dans le tenant réel :

1. Charger la version configurée sur l’URL HTTPS enregistrée ; fermer puis rouvrir le volet Outlook.
2. Ouvrir un mail de test avec un DOCX et un XLSX non sensibles de moins de 2 Mo.
3. Ajouter une signature puis cliquer sur **Connecter mon OneDrive professionnel**. Contrôler le compte affiché.
4. Demander l’aperçu PDF signé ; vérifier pagination, polices, tableaux et position de la signature. Pour Excel, contrôler toutes les feuilles imprimées et les zones d’impression.
5. Préparer une réponse dans chacun des deux formats. Ouvrir les pièces jointes avant envoi ; vérifier le DOCX dans Word et le XLSX dans Excel (formules et graphiques compris).
6. Vérifier que les copies temporaires ont quitté la racine OneDrive et se trouvent dans sa corbeille.
7. Vérifier une annulation de connexion, un refus d’autorisation et une interruption réseau. La signature des PDF et le renvoi DOCX/XLSX doivent rester utilisables.

Le renommage du bouton en **Signer les documents** dépend du chargement du nouveau manifeste `1.1.0.0`. Un ancien manifeste garde son ancien intitulé mais pointe toujours vers la même application.

## Limites

- La documentation Microsoft indique une limite de **2 Mo** pour la conversion. Cette limite est vérifiée sur le document signé, qui peut être plus volumineux que l’original. Le renvoi au format d’origine reste disponible jusqu’à la limite appliquée par le complément (25 Mo par fichier).
- Seuls les formats `.docx` et `.xlsx` standard sont pris en charge. `.doc`, `.xls`, fichiers à macros, fichiers chiffrés, documents numériquement signés, feuilles protégées et certains formats Strict sont exclus.
- Word : l’image est insérée dans le paragraphe choisi, sous son texte. Elle peut décaler les paragraphes suivants. Les zones de texte et contrôles de contenu ne sont pas proposés comme points d’insertion ; le suivi des modifications doit être désactivé dans une copie.
- Excel : l’image est ancrée à une cellule ; elle peut recouvrir les cellules voisines. Les valeurs, formules, graphiques et réglages d’impression sont conservés. Choisir une zone vide et imprimable.
- L’aperçu PDF est le résultat de Microsoft 365. Il doit être vérifié, notamment pour la pagination Word et les zones d’impression Excel. Il ne constitue pas une garantie de rendu identique dans toutes les versions Office.
- La signature ajoutée est une image avec nom, fonction et horodatage, comme dans le fonctionnement PDF existant ; aucun certificat de signature numérique n’est ajouté.

## Références techniques

- [Authentification des compléments par la boîte de dialogue Office](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/auth-with-office-dialog-api)
- [Inscription et URI de redirection SPA](https://learn.microsoft.com/en-us/entra/identity-platform/how-to-add-redirect-uri)
- [Conversion Microsoft Graph](https://learn.microsoft.com/en-us/graph/api/driveitem-get-content-format?view=graph-rest-1.0)
- [Redirections et téléchargement JavaScript](https://learn.microsoft.com/en-us/graph/api/driveitem-get-content?view=graph-rest-1.0)
- [Réponses groupées Microsoft Graph](https://learn.microsoft.com/en-us/graph/json-batching)

L’API de conversion est appelée dans un batch JSON pour récupérer son en-tête de redirection. Le téléchargement de l’URL préauthentifiée s’effectue ensuite sans jeton Graph. Ce parcours est couvert par les tests simulés, mais doit être validé dans le tenant réel. Microsoft annonce des évolutions des URL préauthentifiées à partir de 2027 ; revalider ce parcours lors d’un changement de politique Microsoft 365.
