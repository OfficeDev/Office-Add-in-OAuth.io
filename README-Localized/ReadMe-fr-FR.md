# <a name="archived-office-add-in-that-uses-the-oauthio-service-to-get-authorization-to-external-services"></a>[ARCHIVÉ] Complément Office utilisant le service OAuth.io en vue d’obtenir une autorisation pour les services externes

> **Remarque :** Cette repo est archivée et n’est plus activement conservée. Les failles de sécurité peuvent exister dans le projet ou ses dépendances. Si vous envisagez de réutiliser ou exécutez n’importe quel code à partir de cette emprunteuses, veillez à effectuer les vérifications de sécurité appropriés sur le code ou des dépendances tout d’abord. N’utilisez pas ce projet comme point de départ d’un complément Office de production. Commencez toujours votre code de production à l’aide de la charge de travail de développement Office/SharePoint dans Visual Studio, ou le [générateur Yeoman de compléments Office](https://github.com/OfficeDev/generator-office), puis suivez les meilleures pratiques de sécurité lorsque vous développez le complément.

Le service OAuth.io simplifie le processus d’obtention de jetons d’accès OAuth 2.0 à partir de services en ligne populaires, tels que Facebook et Google. Cet exemple présente son utilisation dans un complément Office. 

## <a name="table-of-contents"></a>Sommaire
* [Historique des modifications](#change-history)
* [Conditions préalables](#prerequisites)
* [Configuration du projet](#configure-the-project)
* [Téléchargement de la bibliothèque JavaScript OAuth.io](#download-the-oauth-io-javascript-library)
* [Création d’un projet Google](#create-a-google-project)
* [Création d’une application Facebook](#create-a-facebook-app)
* [Création d’une application OAuth.io, et configuration pour Google et Facebook](#create-an-OAuth-io-app-and-configure-it-to-use-google-and-facebook)
* [Ajout de la clé publique à l’exemple de code](#add-the-public-key-to-the-sample-code)
* [Déploiement du complément](#deploy-the-add-in)
* [Exécution du projet](#run-the-project)
* [Démarrage du complément](#start-the-add-in)
* [Test du complément](#test-the-add-in)
* [Questions et commentaires](#questions-and-comments)
* [Ressources supplémentaires](#additional-resources)

## <a name="change-history"></a>Historique des modifications

5 janvier 2017 :

* Version d’origine.

## <a name="prerequisites"></a>Conditions préalables

* Un compte avec [OAuth.io](https://oauth.io/home)
* Des comptes de développeur avec [Google](https://developers.google.com/) et [Facebook](https://developers.facebook.com/)
* Word 2016 pour Windows, version 16.0.6727.1000 ou ultérieure.
* [Node et npm](https://nodejs.org/en/) Le projet est configuré pour utiliser npm à la fois comme gestionnaire de package et exécuteur de tâches. Il est également configuré pour utiliser Lite Server en tant que serveur web hébergeant le complément lors du développement, afin que le complément soit rapidement opérationnel. N’hésitez pas à utiliser un autre exécuteur de tâches ou serveur web.
* [GIT Bash](https://git-scm.com/downloads) (ou un autre client Git)

## <a name="configure-the-project"></a>Configurer le projet

Dans le dossier où vous souhaitez placer le projet, exécutez les commandes suivantes dans l’interpréteur de commande Git Bash :

1. ```git clone {URL of this repo}``` pour cloner ce référentiel sur votre ordinateur local.
2. ```npm install``` pour installer toutes les dépendances détaillées dans le fichier package.json.
3. ```bash gen-cert.sh``` pour créer le certificat nécessaire à l’exécution de cet exemple. 

Définissez le certificat comme appartenant à une autorité racine approuvée. Sur un ordinateur Windows, procédez comme suit :

1. Dans le dossier de référentiel de votre ordinateur local, double-cliquez sur ca.crt et sélectionnez **Installer le certificat**. 
2. Sélectionnez **Ordinateur local** et choisissez **Suivant** pour continuer. 
3. Sélectionnez **Placer tous les certificats dans le magasin suivant**, puis **Parcourir**.
4. Sélectionnez **Autorités de certification racines de confiance** et **OK**. 
5. Sélectionnez **Suivant**, puis **Terminer**. 

## <a name="download-the-oauthio-javascript-library"></a>Téléchargement de la bibliothèque JavaScript OAuth.io

2. Dans le projet, créez un sous-dossier nommé OAuth.io sous le dossier Scripts.
3. Téléchargez la bibliothèque JavaScript OAuth.io à partir de [oauth-js](https://github.com/oauth-io/oauth-js).
2. À partir du dossier \dist de la bibliothèque, copiez le fichier oauth.js ou le fichier oauth.min.js dans le dossier \Scripts\OAuth.io du projet. 
3. Si vous avez choisi le fichier oauth.min.js à l’étape précédente, ouvrez le fichier \Scripts\popupRedirect.js et remplacez la ligne `url: 'Scripts/OAuth.io/oauth.js',` par `url: 'Scripts/OAuth.io/oauth.min.js',`.

## <a name="create-a-google-project"></a>Création d’un projet Google

1. Dans votre [console de développeur Google](https://console.developers.google.com/apis/dashboard), créez un projet nommé **Complément Office**.
2. Activez l’API Google Plus pour le projet. Pour plus de détails, affichez l’aide de Google.
3. Créez des informations d’identification d’application pour le projet, puis choisissez **ID client OAuth** comme type d’informations d’identification et **Application web** comme type d’application. Utilisez « Office-Add-in-OAuth.io » en tant que **nom** du client (et en tant que nom de produit si vous devez en indiquer un).
4. Définissez l’option **URI de redirection autorisées** sur ```https://oauth.io/auth```.
5. Créez une copie de la clé secrète client et de l’ID client affecté.

## <a name="create-a-facebook-app"></a>Création d’une application Facebook

1. Dans votre [tableau de bord de développement Facebook](https://developers.facebook.com/), créez une application nommée **Complément Office**. Pour plus de détails, consultez l’aide de Facebook.
3. Créez une copie de la clé secrète client et de l’ID client affecté. 
4. Indiquez ```oauth.io``` dans la zone **Domaines de l’application** et saisissez ```https://oauth.io/``` en tant qu’**URL du site**. 

## <a name="create-an-oauthio-app-and-configure-it-to-use-google-and-facebook"></a>Création d’une application OAuth.io, et configuration pour Google et Facebook

1. Dans votre tableau de bord OAuth.io, créez une application et notez la **clé publique** qu’OAuth.io lui affecte.
3. Si elle n’est pas déjà dans la liste, ajoutez `localhost` aux URL de la liste approuvée pour l’application.
2. Ajoutez **Google Plus** en tant que fournisseur (également appelé « API intégrée ») à l’application. Renseignez l’ID client et la clé secrète obtenus lors de la création du projet Google, puis indiquez `profile` en tant qu’**étendue**. Si la configuration du fournisseur (API intégrée) ne s’enregistre pas dans IE, essayez de rouvrir le tableau de bord OAuth.io dans un autre navigateur.
3. Répétez l’étape précédente pour **Facebook**, mais indiquez `user_about_me` en tant qu’**étendue**.

## <a name="add-the-public-key-to-the-sample-code"></a>Ajout de la clé publique à l’exemple de code

L’API `OAuth.initialize` est appelée dans le fichier popup.js. Recherchez la ligne et transmettez la clé publique obtenue dans la section précédente sous forme de chaîne à cette fonction.

## <a name="deploy-the-add-in"></a>Déploiement du complément

Vous devez maintenant indiquer à Microsoft Word où trouver le complément.

1. Créez un partage réseau, ou [partagez un dossier sur le réseau](https://technet.microsoft.com/fr-fr/library/cc770880.aspx).
2. Placez une copie du fichier manifeste Office-Add-in-OAuth.io.xml, depuis la racine du projet, dans le dossier partagé.
3. Lancez Word et ouvrez un document.
4. Choisissez l’onglet **Fichier**, puis choisissez **Options**.
5. Choisissez l’onglet **Fichier**, puis choisissez **Options**.
6. Choisissez **Catalogues de compléments approuvés**.
7. Dans le champ **URL du catalogue**, saisissez le chemin réseau permettant d’accéder au partage de dossier qui contient le fichier Office-Add-in-OAuth.io.xml, puis sélectionnez **Ajouter un catalogue**.
8. Activez la case à cocher **Afficher dans le menu**, puis cliquez sur **OK**.
9. Un message vous informe que vos paramètres seront appliqués lors du prochain démarrage de Microsoft Office. Fermez Word.

## <a name="run-the-project"></a>Exécution du projet

1. Ouvrez une fenêtre de commande de nœud dans le dossier du projet et exécutez ```npm start``` pour démarrer le service web. Laissez la fenêtre de commande ouverte.
2. Ouvrez Internet Explorer ou Edge, et saisissez ```https://localhost:3000``` dans la zone d’adresse. Si vous ne recevez aucun avertissement concernant le certificat, fermez le navigateur et passez à la section suivante intitulée **Démarrer le complément**. Si vous recevez un message d’avertissement indiquant que le certificat n’est pas approuvé, passez aux étapes suivantes :
3. Le navigateur vous fournit un lien vous permettant d’ouvrir la page malgré l’avertissement. Ouvrez-la.
4. Une fois la page ouverte, une erreur de certificat affichée en rouge apparaît dans la barre d’adresses. Double-cliquez sur l’erreur.
5. Sélectionnez **Afficher le certificat**.
5. Sélectionnez **Installer le certificat**.
4. Sélectionnez **Ordinateur local** et choisissez **Suivant** pour continuer. 
3. Sélectionnez **Placer tous les certificats dans le magasin suivant**, puis **Parcourir**.
4. Sélectionnez **Autorités de certification racines de confiance** et **OK**. 
5. Sélectionnez **Suivant**, puis **Terminer**.
6. Fermez le navigateur.

## <a name="start-the-add-in"></a>Démarrer le complément

1. Redémarrez Word et ouvrez un document Word.
2. Dans l’onglet **Insertion** de Word 2016, choisissez **Mes compléments**.
3. Sélectionnez l’onglet **DOSSIER PARTAGÉ**.
4. Sélectionnez **Authentification avec OAuthIO**, puis **OK**.
5. Si les commandes de complément sont prises en charge par votre version de Word, l’interface utilisateur vous informe que le complément a été chargé.
6. Dans le ruban Accueil, un nouveau groupe appelé **OAuth.io** apparaît avec un bouton intitulé **Afficher** et une icône. Cliquez sur ce bouton pour ouvrir le complément.

 > Remarque : le complément se charge dans un volet Office si les commandes de complément ne sont pas prises en charge par votre version de Word.

## <a name="test-the-add-in"></a>Test du complément

1. Cliquez sur le bouton **Obtenir un nom Facebook**.
2. Une fenêtre contextuelle s’ouvre et vous invite à vous connecter avec Facebook (sauf si vous l’êtes déjà).
3. Une fois que vous êtes connecté, votre nom d’utilisateur Facebook est inséré dans le document Word.
4. Répétez les étapes ci-dessus avec le bouton **Obtenir un nom Google**.

## <a name="questions-and-comments"></a>Questions et commentaires

Nous serions ravis de connaître votre opinion sur cet exemple. Vous pouvez nous envoyer vos commentaires via la section *Problèmes* de ce référentiel.

Les questions générales sur le développement de Microsoft Office 365 doivent être publiées sur [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API). Si votre question concerne les API Office JavaScript, assurez-vous qu’elle comporte les balises [office-js] et [API].

## <a name="additional-resources"></a>Ressources supplémentaires

* 
  [Documentation de complément Office](https://msdn.microsoft.com/fr-fr/library/office/jj220060.aspx)
* [Centre de développement Office](http://dev.office.com/)
* Plus d’exemples de complément Office sur [OfficeDev sur Github](https://github.com/officedev)

## <a name="copyright"></a>Copyright
Copyright (c) 2017 Microsoft Corporation. Tous droits réservés.

