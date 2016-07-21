# Exemple de complément de volet Office - Suivi du budget universitaire pour Excel 2016

_S’applique à : Excel 2016_

Ce complément de volet Office montre comment créer un outil de suivi du budget universitaire à l’aide des API JavaScript dans Excel 2016. Il a deux versions : éditeur de code et Visual Studio.

![Exemple d’outil de suivi de budget universitaire](../images/CollegeBudgetTracker_tracker.PNG)

## Essayez !
### Version d’éditeur de code

Pour déployer et tester votre complément, le plus simple consiste à copier le fichier manifeste sur un partage réseau.

1.  Créez un dossier sur un partage réseau (par exemple, \\MyShare\CollegeBudgetTracker).  
2.  Copiez le fichier manifeste (CollegeBudgetTrackerManifest.xml) dans un partage réseau (par exemple, \\\MyShare\MyManifests).
3.  Ajoutez l’emplacement de partage qui contient le fichier manifeste sous forme de catalogue d’applications approuvées dans Excel.

    a. Lancez Excel et ouvrez une feuille de calcul vide.  
    
    b. Choisissez l’onglet **Fichier**, puis choisissez **Options**.
    
    c. Choisissez **Centre de gestion de la confidentialité**, puis cliquez sur le bouton **Paramètres du Centre de gestion de la confidentialité**.
    
    d. Choisissez **Catalogues de compléments approuvés**.
    
    e. Dans la zone **URL du catalogue**, entrez le chemin d’accès du partage réseau que vous avez créé à l’étape 3, puis choisissez **Ajouter un catalogue**.
    
   Activez la case à cocher **Afficher dans le menu**, puis cliquez sur **OK**. Un message s’affiche pour vous informer que vos paramètres seront appliqués la prochaine fois que vous démarrerez Office. 
        
4.  Testez et exécutez le complément. 

    a. Dans l’onglet **Insertion** d’Excel 2016, choisissez **Mes compléments**. 
    
    b. Dans la boîte de dialogue **Compléments Office**, choisissez **Dossier partagé**.
    
    c. Cliquez sur la commande **Suivi du budget universitaire** dans l’onglet Accueil. Le complément s’ouvre dans un volet Office et crée le suivi du budget universitaire dans la feuille active, comme indiqué sur le diagramme. 
      
   ![Exemple d’outil de suivi de budget universitaire](../images/CollegeBudgetTracker_tracker.PNG) 

    d. Ajoutez des dépenses et des revenus à l’aide des onglets **Ajouter des dépenses** et **Ajouter des revenus**, puis observez la façon dont les données et les graphiques changent de manière dynamique.
    
      ![Exemple de suivi de budget universitaire](../images/CollegeBudgetTracker_taskpane1.PNG) 

Pour utiliser le fichier manifeste dans votre propre complément, modifiez l’élément <SourceLocation> du fichier manifeste afin qu’il pointe vers l’emplacement de partage de votre fichier Home.html.
    
### Version de Visual Studio
1.  Copiez le projet dans un dossier local et ouvrez le fichier Excel-Add-in-JS-CollegeBudgetTracker.sln dans Visual Studio.
2.  Appuyez sur F5 pour créer et déployer l’exemple de complément. Excel démarre et le complément s’ouvre dans un volet Office à droite de la feuille de calcul active, comme indiqué dans l’illustration suivante. 
        
  ![Exemple d’outil de suivi de budget universitaire](../images/CollegeBudgetTracker_tracker.PNG) 

3.  Ajoutez des dépenses et des revenus à l’aide des onglets **Ajouter des dépenses** et **Ajouter des revenus**, puis observez la façon dont les données et les graphiques changent de manière dynamique.

  ![Exemple d’outil de suivi de budget universitaire](../images/CollegeBudgetTracker_taskpane1.PNG) 


### En savoir plus

Les API JavaScript pour Excel peuvent vous offrir beaucoup pour l’élaboration de vos compléments. Voici quelques-unes des ressources disponibles : 

1.  [Présentation de la programmation pour les compléments Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-programming-overview.md)
2.  [Explorateur d’extraits de code pour Excel](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
3.  [Exemples de code pour les compléments Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-code-samples.md) 
4.  [Référence de l’API JavaScript pour les compléments Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-javascript-reference.md)
5.  [Créer son premier complément Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/build-your-first-excel-add-in.md)
