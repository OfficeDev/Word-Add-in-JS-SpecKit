# Complément Word JavaScript SpecKit

Découvrez comment vous pouvez créer un complément qui capture et insère du texte réutilisable, ainsi que la façon dont vous pouvez implémenter un processus de validation de document simple.

## Table des matières
* [Historique des modifications](#change-history)
* [Conditions requises](#prerequisites)
* [Configurer le projet](#configure-the-project)
* [Exécuter le projet](#run-the-project)
* [Comprendre le code](#understand-the-code)
* [Questions et commentaires](#questions-and-comments)
* [Ressources supplémentaires](#additional-resources)

## Historique des modifications

31 mars 2016 :
* Version de l’exemple initial.

## Conditions requises

* Word 2016 pour Windows, build 16.0.6727.1000 ou ultérieur.
* [NÅ“ud et npm](https://nodejs.org/en/)
* [GIT Bash](https://git-scm.com/downloads) - Vous devez utiliser une version ultérieure, car les builds antérieurs peuvent afficher une erreur lors de la génération de certificats.

## Configurer le projet

À partir de votre environnement Bash, exécutez les commandes suivantes à la racine de ce projet :

1. Dupliquez ce référentiel sur votre ordinateur local.
2. ```npm install``` pour installer toutes les dépendances dans package.json.
3. ```bash gen-cert.sh``` pour créer les certificats nécessaires à l’exécution de cet exemple. Ensuite, dans le référentiel sur votre ordinateur local, cliquez deux fois sur ca.crt et sélectionnez **Installer le certificat**. Sélectionnez **Ordinateur local** et choisissez **Suivant** pour continuer. Sélectionnez **Placer tous les certificats dans le magasin suivant**, puis **Parcourir**.  Sélectionnez **Autorités de certification racines de confiance** et **OK**. Sélectionnez **Suivant** et **Terminer**. Maintenant, votre certificat d’autorité de certification apparaît dans votre magasin de certificats.
4. ```npm start``` pour démarrer le service.

À ce stade, vous avez déployé cet exemple de complément. Vous devez maintenant indiquer à Microsoft Word où trouver le complément.

1. Créez un partage réseau ou [partagez un dossier sur le réseau](https://technet.microsoft.com/fr-fr/library/cc770880.aspx), puis placez-y le fichier manifeste [word-add-in-javascript-speckit-manifest.xml](word-add-in-javascript-speckit-manifest.xml).
3. Lancez Word et ouvrez un document.
4. Choisissez l’onglet **Fichier**, puis choisissez **Options**.
5. Choisissez **Centre de gestion de la confidentialité**, puis cliquez sur le bouton **Paramètres du Centre de gestion de la confidentialité**.
6. Choisissez **Catalogues de compléments approuvés**.
7. Dans le champ **URL du catalogue**, saisissez le chemin réseau pour accéder au partage de dossier qui contient le fichier word-add-in-javascript-speckit-manifest.xml, puis choisissez **Ajouter un catalogue**.
8. Activez la case à cocher **Afficher dans le menu**, puis cliquez sur **OK**.
9. Un message vous informe que vos paramètres seront appliqués lors du prochain démarrage de Microsoft Office. Fermez et redémarrez Word.

## Exécuter le projet

1. Ouvrez un document Word.
2. Dans l’onglet **Insertion** de Word 2016, choisissez **Mes compléments**.
3. Sélectionnez l’onglet **DOSSIER PARTAGÉ**.
4. Choisissez **Complément Word SpecKit**, puis sélectionnez **OK**.
5. Si les commandes de complément sont prises en charge par votre version de Word, l’interface utilisateur vous informe que le complément a été chargé.

### Interface utilisateur du ruban
Dans le ruban, vous pouvez :
* Sélectionner l’onglet **Complément SpecKit** pour lancer le complément dans l’interface utilisateur.
* Sélectionner **Insérer un modèle spec** pour lancer le volet Office et insérer un modèle spec dans le document.
* Utiliser les boutons de validation du ruban ou du menu contextuel pour valider le document par rapport à une liste noire de mots.

 > Remarque : Le complément se charge dans un volet Office si les commandes de complément ne sont pas prises en charge par votre version de Word.

### Interface utilisateur de volet Office
Dans le volet Office, vous pouvez :
* Enregistrer une phrase en plaçant le curseur dans une phrase, lui donner un nom dans le champ ci-dessus (**Ajouter une phrase à réutiliser* dans le volet Office), puis sélectionner **Ajouter une phrase à réutiliser**. Vous pouvez faire de même pour les paragraphes.
* L’enregistrement des phrases et paragraphes va également rendre les éléments réutilisables disponibles dans la liste déroulante **Insérer des éléments réutilisables**.
* Placer le curseur dans le document. Sélectionnez un texte réutilisable dans la liste déroulante afin qu’il soit inséré dans le document.
* Modifier la propriété *Auteur* du document en modifiant le nom d’auteur et en sélectionnant le bouton **Mettre à jour le nom d’auteur**. Cela mettra à jour la propriété du document, ainsi que le contenu d’un contrôle de contenu lié.

## Comprendre le code

Cet exemple utilise l’[ensemble de conditions requises](http://dev.office.com/reference/add-ins/office-add-in-requirement-sets?product=word) 1.2 lors de la période de visualisation, mais il exige l’ensemble de conditions requises 1.3 dès que ce dernier est disponible.

### Volet de tâches

La fonctionnalité de volet Office est configurée dans sample.js. Ce fichier contient les fonctionnalités suivantes :

* Configuration des gestionnaires d’interface utilisateur et d’événement.
* Extraction du modèle spec à partir d’un service et insertion dans le document.
* Chargement d’une liste noire qui contient des mots qui sont utilisés pour valider le document. Ces termes sont considérés comme des mots incorrects dans le cadre de cet exemple.
* Chargement d’un élément réutilisable par défaut à partir d’un service pour le mettre en cache dans le stockage local.
* Code squelette pour tester le code de fichier de fonction. Vous pouvez développer votre code de commande de complément dans le volet Office avant de le déplacer vers un fichier de fonction, car vous ne pouvez pas joindre un débogueur au fichier de fonction.
* Chargement du nom d’auteur par défaut à partir des propriétés du document dans le volet Office. Cela vous indique comment accéder à une partie XML personnalisée dans un document et comment la modifier.
* Publication d’éléments réutilisables dans le service.

### Validation du document et API de boîte de dialogue

Le fichier validation.js contient le code permettant de valider le document par rapport à une liste noire de mots. La méthode validateContentAgainstBlacklist() utilise la nouvelle méthode splitTextRanges pour fractionner le document en plages en fonction de séparateurs. Les séparateurs de cette fonction identifient les mots dans le document. Nous comparons les mots du document à ceux de la liste noire, puis nous transmettons ces résultats au stockage local. Ensuite, nous utilisons la méthode displayDialogAsync pour ouvrir une boîte de dialogue (dialog.html). La boîte de dialogue obtient les résultats de la validation auprès du stockage local et affiche les résultats.

### Fonctionnalité de texte réutilisable

Le fichier boilerplate.js contient des exemples de façons dont vous pouvez enregistrer du texte réutilisable dans le stockage local, mettre à jour une liste déroulante Structure avec des entrées qui correspondent aux éléments réutilisables enregistrés, et insérer des éléments réutilisables sélectionnés dans une liste déroulante. Ce fichier contient des exemples des éléments suivants :
* splitTextRanges (nouveauté dans l’ensemble de conditions requises WordApi 1.3) - Cette API sera remplacée par split() dans une version ultérieure.
* compareLocationWith (nouveauté dans l’ensemble de conditions requises WordApi 1.3)
* Mise à jour de la liste déroulante Structure avec des nouvelles entrées.
* Insertion de texte réutilisable dans le document.

### Liaison XML personnalisée aux principales propriétés du document

Le fichier authorCustomXml.js contient des méthodes pour obtenir et définir les propriétés du document par défaut.

* Chargez la propriété d’auteur dans le volet Office lors du chargement du volet Office. Notez que le document contient également la valeur de la propriété d’auteur. Cela vient du fait que le modèle contient un contrôle de contenu lié à cette propriété de document. Cela vous permet de définir des valeurs par défaut dans le document en fonction du contenu d’une partie XML personnalisée.
* Mettez à jour la propriété d’auteur du volet Office. Cela mettra à jour la propriété du document et le contrôle de contenu lié dans le document.

### Commandes de compléments

Les déclarations de commande de complément se trouvent dans word-add-in-javascript-speckit-manifest.xml. Cet exemple montre comment créer des commandes de complément dans le ruban et dans un menu contextuel.

## Questions et commentaires

Nous serions ravis de connaître votre opinion sur l’exemple Word SpecKit. Vous pouvez nous faire part de vos suggestions dans la rubrique *Problèmes* de ce référentiel.

Si vous avez des questions générales sur le développement de Microsoft Office 365, envoyez-les sur [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API). Posez vos questions avec les balises [API] et [office-js].

## Ressources supplémentaires

* [Documentation de complément Office](https://msdn.microsoft.com/fr-fr/library/office/jj220060.aspx)
* [Centre de développement Office](http://dev.office.com/)
* [Projets de démarrage et exemples de code des API Office 365](http://msdn.microsoft.com/en-us/office/office365/howto/starter-projects-and-code-samples)

## Copyright
Copyright (c) 2016 Microsoft Corporation. Tous droits réservés.


