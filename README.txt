-TarificatorV2- 

fonctionnalitées : 
-Création d'un dossier annexe pour éviter les fausses manipulations
-Détection du FAB-DIS
-A partir du FAB-DIS, on récupère tout l'onglet "01-COMMERCE" qu'on ajuste pour obtenir 
les colonnes qu'on souhaite pour le fournisseur que l'on souhaite.
-On récupère aussi toute la DEEE dans l'onglet "04-RÉGLEMENTAIRE" qu'on ajuste et tri (RTYP, CONTRIB, DEEE)
-Insertion des colonnes "PHOTO", "FICHE", D3E, D3EC, D3EU, D3EV et "SKUSOCODA" qui vont accueillir le résultat 
de la recherche verticale
-Copie du SKUSOCODA et de l'onglet MEDIA grace au fichier tarif précédent
-Ecriture de la DEEE, "ECL" pour l'éclairage et "ECO" pour le reste
-Prise en charge dans la DEEE du nombre
-Mise en place de la recherche V pour générer les colonnes "PHOTO", "FICHE", "SKUSOCODA", "D3EC", "D3EV", "D3EU"
-Calcul du temps d'exécution de tarificator pour chaque tarif + temps global