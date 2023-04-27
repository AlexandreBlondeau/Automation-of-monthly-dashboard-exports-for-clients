# README - French Version - [Link to the English version](README_en.md)

## **Projet :** Automatisation des exportations mensuelles de tableaux de bord pour des clients, avec une interface utilisateur graphique responsive

### **Auteur :** Alexandre Blondeau - [alex991@live.fr](mailto:alex991@live.fr) - [GitHub](https://github.com/AlexandreBlondeau)  

*Ce [script](Script.py) python effectue les actions suivantes :*

#### **I.**  

1. Création d'une fenêtre principale pour l'application avec des dimensions spécifiques et un titre approprié. Configuration de la fenêtre pour qu'elle soit redimensionnable, permettant ainsi une interface utilisateur responsive.
2. Ajout de widgets à la fenêtre principale, tels que :
    * Un label pour décrire l'objectif du champ de saisie (par exemple, "Entrez la cellule de départ de la colonne M").
    * Un champ de saisie pour permettre à l'utilisateur de spécifier la cellule M à partir de laquelle commencer les itérations.
    * Un bouton pour démarrer le script avec un gestionnaire d'événements associé qui déclenche les actions du script lorsque l'utilisateur clique sur le bouton.
3. Configuration de la disposition des widgets à l'aide d'une méthode appropriée (par exemple, pack, grid ou place), afin de rendre l'interface facile à comprendre et à utiliser. Utilisation des poids de colonnes et de lignes pour garantir une répartition équilibrée de l'espace lors du redimensionnement de la fenêtre.
4. Gestion des erreurs lors de la saisie de l'utilisateur, comme s'assurer que la cellule spécifiée est valide et afficher un message d'erreur si ce n'est pas le cas.
5. Mise en place d'une boucle d'événements pour maintenir l'interface utilisateur active et réagir aux actions de l'utilisateur, telles que la boucle principale mainloop.
6. Lorsque l'utilisateur clique sur le bouton pour démarrer le script, création et démarrage d'un second thread pour exécuter la deuxième partie du script. Ceci évite que le GUI ne gèle pendant le traitement des actions du script et assure une expérience utilisateur fluide.

#### **II.**  

La fonction « second_part » prend en entrée un paramètre start_iteration, qui détermine à partir de quelle itération le processus va commencer. La valeur start_iteration est définie par l’utilisateur dans le GUI. Voici un résumé détaillé du fonctionnement de cette fonction :

1. Initialise Python COM library avec pythoncom.CoInitialize().
2. Définit le chemin du fichier Excel à manipuler : chemin_fichier.
3. Vérifie si le fichier Excel existe déjà, et le supprime s'il existe. Sinon, affiche un message indiquant que le fichier n'existe pas.
4. Ferme l'application Excel en effectuant deux tentatives maximum.
5. Définit le chemin du fichier Excel source file_path et le chemin du dossier de destination destination_folder.
6. Entame une boucle while pour effectuer les opérations sur les cellules de la colonne M, en commençant par l'itération indiquée par start_iteration.
7. Vérifie si le dossier "TBM" existe. S'il n'existe pas, le crée. Sinon, affiche un message indiquant que le dossier existe déjà.
8. Copie le fichier Excel source dans le dossier de destination pour effectuer les actions suivantes en local sur le disque "C:\".
9. Lance l'application Excel en mode invisible, désactive les alertes Excel et ouvre le fichier copié en mode lecture/écriture.
10. Sélectionne la feuille "Paramètres et Instructions" du fichier Excel.
11. Récupère les valeurs à exclure lors de la première itération (qui est défine par l'utilisateur dans le GUI.
12. Copie la valeur de la cellule Mi dans la cellule fusionnée "B1:C1".
13. Vérifie si la valeur de la cellule en cours doit être exclue et passe à la cellule suivante si c'est le cas.
14. Affiche un message indiquant le début de l'itération et la cellule Mi utilisée.
15. Lit la valeur de la cellule suivante dans la colonne M et la valeur de la cellule après la cellule suivante.
16. Marque l'itération comme dernière itération si la cellule suivante est vide ou None ou dans les valeurs à exclure et si la cellule d’encore après est vide ou None.
17. Affiche un message indiquant l'itération en cours pour le client concerné.
18. Actualise toutes les données du fichier Excel et attend que l'actualisation soit terminée.
19. Exécute la macro "Publication" et attend que la publication du fichier PDF soit terminée.
20. Ferme l'application Excel en effectuant deux tentatives maximum.
21. Supprime le fichier Excel créé dans le dossier "C:\TBM".
22. Vérifie si c'est la dernière itération. Si c'est le cas, affiche un message indiquant la fin des itérations et quitte la boucle while.
23. Ferme le classeur Excel sans enregistrer les modifications et quitte l'application Excel si c'était la dernière itération.

**Dépendances**  

Ce programme utilise les bibliothèques suivantes :
- os
- time
- win32com.client
- psutil
- shutil
- pywintypes
- sys
- pythoncom
- tkinter

Pour installer les dépendances manquantes, vous pouvez utiliser la commande pip. (Pas besoin si vous utilisez le fichier .exe qui est fourni).  

python -m pip install XXXX


**Utilisation**  

Si vous disposez du fichier exécutable, exécutez seulement le « .exe » et passer à l’étape 2.

1. Assurez-vous que toutes les dépendances sont installées.
Exécutez le fichier Export_Client_.py pour démarrer le programme.

2. L'interface utilisateur apparaîtra, indiquant la cellule où la première itération commence. Par défaut, il s'agit de la cellule M2.
Entrez la cellule de départ souhaitée et cliquez sur "Démarrer le script".
Le programme lira les valeurs de la colonne M et les traitera itérativement, en générant et en exportant des fichiers PDF.

**Problèmes connus**  

- Ne pas changer le nom du fichier Excel original situé dans le disque « X:\ », car le script ne sera alors pas capable de le retrouver.
- Le code effectue beaucoup de requêtes à un serveur distant via «workbook.RefreshAll() » , il arrive donc que le Firewall bloque les requêtes au bout d’un certain nombre d’itérations et que les erreurs suivantes apparaissent dans le GUI :
  * Erreur de l’actualisation des données (-2147023170, ‘Echec de l’appel de la procédure distante.’, None, None)
  * Erreur de communication avec Excel : (-2147023174, ‘Le serveur RPC n’est pas disponible.’, None, None)

J’ai rajouté à la fin de ce document une partie Annexe où j’explique comment régler ce problème à l’équipe IT. Si vous n’avez pas encore résolu ce problème, fermez le script puis vérifiez si Excel est bien fermé dans le gestionnaire de tâche. Puis ressayer le script un peu plus tard à l’itération où cela s’est arrêté via le GUI.
Assurez-vous que le fichier Excel copié dans le chemin original n’est pas ouvert sur un autre ordinateur du réseau pendant l'exécution du programme, car cela peut entraîner des problèmes de verrouillage de fichier et cela pourra empêcher la copie du fichier sur le disque local « C:\ ».
Si vous rencontrez des erreurs de communication avec Excel, le programme attendra et réessaiera automatiquement. Cela peut être dû à des problèmes de connexion avec Excel ou à des problèmes de synchronisation.

**Annexe**  

1. Ouvrez le pare-feu Windows.
   Autorisez l'application ou la fonction à travers le pare-feu Windows.
   Activez le privilège de domaine pour Windows Management Instrumentation (WMI).
   Vous pouvez également vérifier d'autres éléments.

2. L'ordinateur distant est bloqué par le pare-feu.
   Solution : Ouvrez le snap-in Éditeur d'objets de stratégie de groupe (gpedit.msc) pour modifier l'objet de stratégie de groupe (GPO) utilisé pour gérer les paramètres du pare-feu Windows dans votre organisation. Ouvrez Configuration de l'ordinateur, Modèles d'administration, Réseau, Connexions réseau, Pare-feu Windows, puis ouvrez Profil de domaine ou Profil standard, selon le profil que vous souhaitez configurer. Activez les exceptions suivantes : Autoriser l'exception d'administration à distance et Autoriser l'exception de partage de fichiers et d'imprimantes.

3. Le nom d'hôte ou l'adresse IP est erroné ou l'ordinateur distant est éteint.
   Solution : Vérifiez que le nom d'hôte ou l'adresse IP est correct.

4. Le service "TCP/IP NetBIOS Helper" ne fonctionne pas.
   Solution : Vérifiez que le service "TCP/IP NetBIOS Helper" est en cours d'exécution et qu'il est configuré pour démarrer automatiquement après le redémarrage.

5. Le service "Remote Procedure Call (RPC)" ne fonctionne pas sur l'ordinateur distant.
   Solution : Ouvrez le fichier services.msc à l'aide de la fonction Exécuter de Windows. Dans Services Windows, vérifiez que le service "Remote Procedure Call (RPC)" est en cours d'exécution et qu'il est configuré pour démarrer automatiquement après le redémarrage.

6. Le service "Windows Management Instrumentation" ne fonctionne pas sur l'ordinateur distant.
   Solution : ouvrez le fichier services.msc à l'aide de la commande Exécuter de Windows : Ouvrez services.msc à l'aide de Windows Run. Vérifiez que le service Windows Management Instrumentation est en cours d'exécution et qu'il est configuré pour démarrer automatiquement après le redémarrage.

