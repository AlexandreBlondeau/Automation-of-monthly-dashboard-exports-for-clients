import os
import time
import win32com.client
import psutil
import shutil
import pywintypes
import sys
import pythoncom
import tkinter as tk
from tkinter import scrolledtext
import threading

def get_values_to_skip(worksheet):
    # Initialiser une liste vide pour stocker les valeurs à exclure
    values_to_skip = []
    # Initialiser l'index de la cellule à 2 (pour commencer à la cellule O2, cette cellule est modifiable dans le GUI
    i = 2
    # Boucler indéfiniment jusqu'à ce qu'une condition d'arrêt soit rencontrée
    while True:
        cell_value = worksheet.Range(f"O{i}").Value
        # Vérifier si la valeur de la cellule est None ou vide, si c'est le cas, sortir de la boucle
        if cell_value is None or cell_value == "":
            break
        # Ajouter la valeur de la cellule à la liste des valeurs à exclure
        values_to_skip.append(cell_value)
        # Augmenter l'index de la cellule pour passer à la cellule suivante
        i += 1
    # Retourner la liste des valeurs à exclure
    return values_to_skip

# Attend que toutes les requêtes asynchrones d'Excel soient terminées
def wait_for_sheets_to_refresh(excel):
    excel.Application.CalculateUntilAsyncQueriesDone()


def attendre_excel(excel):
    while True:
        try:
            # Vérifier si Excel est interactif (c'est-à-dire prêt à être utilisé)
            if excel.Interactive is False:
                time.sleep(1)  # Attendre 1 seconde avant de vérifier à nouveau
            else:
                break # Quitte la boucle si Excel est interactif (c'est-à-dire prêt à être utilisé)
        # Attraper les erreurs de communication avec Excel
        except pywintypes.com_error as e:
            print(f"Erreur de communication avec Excel: {e}")
            time.sleep(1)  # Attendre 1 seconde avant de réessayer

# Définition de la fonction `fermer_excel`
def fermer_excel():
    # Initialise une variable pour suivre si Excel a été fermé ou non
    excel_ferme = False

    # Parcourir la liste des processus en cours d'exécution
    for process in psutil.process_iter():
        try:
            # Récupère les informations du processus (ID et nom)
            process_info = process.as_dict(attrs=['pid', 'name'])

            # Vérifie si le processus est Excel
            if process_info['name'] == 'EXCEL.EXE':
                # Termine le processus Excel
                os.kill(process_info['pid'], 9)
                # Met à jour la variable pour indiquer qu'Excel a été fermé
                excel_ferme = True
                # Quitte la boucle
                break

        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            # Ignore les exceptions et continue la boucle pour vérifier les autres processus
            pass

    if not excel_ferme:
        # Retourne False si Excel n'a pas été fermé
        return False

    # Retourne True si Excel a été fermé
    return excel_ferme

class TextRedirector:
    # Initialise l'objet TextRedirector avec un widget et des attributs pour gérer les messages répétés
    def __init__(self, widget):
        self.widget = widget
        self.last_text = None
        self.repeat_count = 1

    def write(self, text):
        # Écrit le texte donné dans le widget en regroupant les messages répétés
        text = text.strip()

        # Ajout d'une condition pour détecter les messages d'erreur spécifiques
        is_error_message = "Erreur de communication avec Excel" in text

        if text == self.last_text or is_error_message:
            self.repeat_count += 1
            # Supprimer le dernier message affiché
            self.widget.after_idle(lambda: self.widget.delete("end-2l", "end-1c"))
            # Afficher le message mis à jour avec le compteur
            self.widget.after_idle(lambda: self.widget.insert(tk.END, f"{text} x{self.repeat_count}\n"))
        else:
            self.repeat_count = 1
            self.widget.after_idle(lambda: self.widget.insert(tk.END, text + "\n"))
            self.widget.after_idle(lambda: self.widget.see(tk.END))

        self.last_text = text

    def flush(self):
        # Cette méthode est présente pour la compatibilité avec sys.stdout mais n'a pas d'effet réel
        pass

    def insert_colored_text(self, text, color, bold=False):
        # Insère du texte coloré dans le widget avec la possibilité de le rendre gras
        tag_name = color + ("_bold" if bold else "")
        self.widget.tag_configure(tag_name, foreground=color, font=("TkDefaultFont", 10, "bold" if bold else "normal"))
        self.widget.after_idle(lambda: self.widget.insert(tk.END, text, (tag_name,)))
        self.widget.after_idle(lambda: self.widget.see(tk.END))

# Première partie: le GUI
def create_gui():
    def execute_second_part():
        # Récupération de la valeur entrée par l'utilisateur
        user_input = entry.get()

        # Conversion de la valeur entrée en entier
        user_iteration = int(user_input)

        # Création d'un thread pour exécuter la deuxième partie du script
        second_part_thread = threading.Thread(target=second_part, args=(
        user_iteration,))
        second_part_thread.start()

    def check_entry_value(value):
        # Vérifie si la valeur est vide
        if not value:
            # Désactive le bouton "Démarrer le script"
            start_button.configure(state="disabled")
            # Autorise la modification de la valeur dans le champ d'entrée
            return True

        try:
            # Tente de convertir la valeur en entier
            number = int(value)
            # Vérifie si la valeur est supérieure ou égale à 2
            if number >= 2:
                # Active le bouton "Démarrer le script"
                start_button.configure(state="normal")
            else:
                # Désactive le bouton "Démarrer le script"
                start_button.configure(state="disabled")
            # Autorise la modification de la valeur dans le champ d'entrée
            return True
        # Attrape l'exception si la conversion en entier échoue
        except ValueError:
            # Désactive le bouton "Démarrer le script"
            start_button.configure(state="disabled")
            # Interdit la modification de la valeur dans le champ d'entrée
            return False

    # Création du GUI (qui est responsive)
    # Crée un nouvel objet Tk, qui représente la fenêtre principale de l'application
    root = tk.Tk()
    # Définit le titre de la fenêtre principale
    root.title("Automatisation Export TBM EQ Excel")

    # Configure la largeur de la première colonne pour qu'elle s'adapte à la fenêtre
    root.columnconfigure(0, weight=1)
    # Configure la hauteur de la première ligne pour qu'elle s'adapte à la fenêtre
    root.rowconfigure(0, weight=1)

    # Crée un nouveau cadre (Frame) pour organiser les widgets à l'intérieur de la fenêtre principale
    frame = tk.Frame(root)
    # Positionne le cadre dans la fenêtre principale
    frame.grid(row=0, column=0, sticky='nsew')
    # Configure la largeur de la deuxième colonne du cadre pour qu'elle s'adapte à la fenêtre
    frame.columnconfigure(1, weight=1)
    # Configure la hauteur de la première ligne du cadre pour qu'elle s'adapte à la fenêtre
    frame.rowconfigure(0, weight=1)

    global text
    # Crée un widget de texte déroulant pour afficher les informations
    text = scrolledtext.ScrolledText(frame, wrap="word")
    # Positionne le widget de texte déroulant dans le cadre
    text.grid(row=0, column=0, columnspan=3, rowspan=3, sticky='nsew', padx=5, pady=5)

    # Rediriger les sorties stdout vers le widget text
    stdout_redirector = TextRedirector(text)
    sys.stdout = stdout_redirector

    # Définir le texte à afficher sur le label
    label_text = "Cellule où la première itération commence (Default=2 pour la cellule M2)"
    # Créer un label avec le texte défini précédemment et une largeur maximale de 300 pixels pour le texte
    iteration_label = tk.Label(frame, text=label_text, wraplength=300)
    # Positionner le label dans la grille du conteneur (frame) à la ligne 1 et colonne 3, avec un padding de 5 pixels autour
    iteration_label.grid(row=1, column=3, padx=5, pady=5, sticky='n')

    # Création d'un champ de texte (Entry) pour l'utilisateur
    entry = tk.Entry(frame)
    entry.grid(row=2, column=3, padx=5, pady=5, sticky='n')

    # Création d'un bouton "Démarrer le script" pour exécuter la deuxième partie
    start_button = tk.Button(frame, text="Démarrer le script", command=execute_second_part)
    start_button.grid(row=0, column=3, padx=5, pady=5, sticky='n')

    # Créez un validatecommand et liez-le à la fonction check_entry_value
    validate_cmd = frame.register(check_entry_value)

    # Création d'un champ de texte (Entry) pour l'utilisateur
    entry = tk.Entry(frame, validate="key", validatecommand=(validate_cmd, '%P'))
    entry.grid(row=2, column=3, padx=5, pady=5, sticky='n')

    # Ajoutez la valeur par défaut 2 dans le champ d'entrée
    entry.insert(0, "2")
    entry.focus_set()  # Met le focus sur l'entrée pour éviter de devoir cliquer dessus avant de modifier la valeur

    # Activez le bouton "start_button" par défaut, car la valeur par défaut est 2
    start_button = tk.Button(frame, text="Démarrer le script", command=execute_second_part, state="normal")
    start_button.grid(row=0, column=3, padx=5, pady=5, sticky='n')

    # Démarre la boucle principale d'événements de l'application
    root.mainloop()

# Seconde partie: le script appelé via le GUI
def second_part(start_iteration):
    pythoncom.CoInitialize()
    chemin_fichier = r"C:\TBM\NEW TBM EQ Exoé 2022_Template_avec_limites.xlsm"

    # Vérifier si le fichier existe déjà et le supprimer si c'est le cas
    if os.path.isfile(chemin_fichier):
        try:
            os.remove(chemin_fichier)
            print(f"Le fichier '{chemin_fichier}' a été supprimé.")
        except Exception as e:
            print(f"Erreur lors de la suppression du fichier '{chemin_fichier}': {e}")
    else:
        print(f"Le fichier '{chemin_fichier}' n'existe pas.")

    # Fermer Excel
    excel_ferme = False
    attempt = 0
    while not excel_ferme and attempt < 2:
        excel_ferme = fermer_excel()
        attempt += 1
        time.sleep(2)  # Attendre 5 secondes avant de vérifier à nouveau

    if not excel_ferme:
        print("Impossible de fermer Excel après 2 tentatives ou Excel n'était pas ouvert")

    file_path = r'X:\CLIENTS\DIVERS\Reportings Mensuels\NEW TBM EQ Exoé 2022_Template_avec_limites.xlsm'
    destination_folder = "C:\\TBM\\NEW TBM EQ Exoé 2022_Template_avec_limites.xlsm"


    i = start_iteration
    while True:  # Effectuer la tâche pour les cellules de M2 jusqu'à ce qu'il n'y ait plus de valeur alphanumérique

        # Définition chemin création dossier
        chemin = "C:\\TBM"

        # Vérifier si le dossier "TBM" existe
        if not os.path.exists(chemin):
            # Si le dossier n'existe pas, le créer
            os.makedirs(chemin)
        else:
            print("Le dossier 'TBM' existe déjà.")

        # Copier le fichier Excel dans le dossier "TBM"
        shutil.copy(file_path, destination_folder)
        print(f"Fichier copié à l'emplacement: {destination_folder}")

        # Lancer Excel et ouvrir le fichier
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False  # Désactiver les alertes Excel
        workbook = excel.Workbooks.Open(destination_folder, ReadOnly=False)

        # Sélectionner la feuille « Paramètres et Instructions »
        worksheet = workbook.Sheets("Paramètres et Instructions")

        # Récupérer les valeurs à exclure lors de la première itération
        if i == start_iteration:
            values_to_skip = get_values_to_skip(worksheet)

        # Copier la cellule Mi dans la cellule qui est fusionner « B1:C1 »
        cell_to_copy = f"M{i}"
        cell_value = worksheet.Range(cell_to_copy).Value

        # Vérifier si la valeur de la cellule en cours doit être exclue
        while cell_value in values_to_skip:
            i += 1
            cell_value = worksheet.Range(f"M{i}").Value

        worksheet.Range("B1:C1").Value = cell_value

        sys.stdout.insert_colored_text(f"Début de l'itération qui utilise la cellule M{i}\n", "black", bold=True)

        # Trouver la prochaine cellule non exclue ou vide dans la colonne M
        i += 1
        next_cell_value = worksheet.Range(f"M{i}").Value

        # Parcourir les cellules jusqu'à trouver une cellule non exclue ou vide
        while next_cell_value in values_to_skip:
            i += 1
            next_cell_value = worksheet.Range(f"M{i}").Value

        # Vérifier si la cellule trouvée est vide
        if next_cell_value is None or next_cell_value == "":
            last_iteration = True
            print("Dernière itération")
        else:
            last_iteration = False

        sys.stdout.insert_colored_text(f"Itération pour le client {cell_value}\n", "red", bold=True)

        time.sleep(1)

        try:
            print ("Le fichier Excel est en cours d'actualisation")
            # Actualiser toutes les données du fichier Excel
            workbook.RefreshAll()

            # Attendre que l'actualisation des données soit terminée avant de continuer
            wait_for_sheets_to_refresh(excel)
        except Exception as e:
            print(e)
            sys.stdout.insert_colored_text(f"Vérifiez le fichier PDF, il y a peut-être eu un souci dans "
                                           f"l'actualisation pour ce client.\n", "blue", bold=True)
        finally:
            print("The 'try except' is finished")
        # Exécuter la macro "Publication"
        try:
            print("La publication du PDF est en cours, votre ordinateur risque d'être ralenti")
            excel.Application.Run("Publication")
        except Exception as e:
            pass

        # Attendre que la publication du fichier pdf soit terminée avant de continuer
        attendre_excel(excel)

        # Fermer Excel
        excel_ferme = False
        attempt = 0
        while not excel_ferme and attempt < 2:
            excel_ferme = fermer_excel()
            attempt += 1
            time.sleep(5)  # Attendre 5 secondes avant de vérifier à nouveau

        if not excel_ferme:
            print("Impossible de fermer Excel après 2 tentatives ou Excel n'était pas ouvert")

        # Supprimer le fichier créé dans C:\\TBM
        try:
            os.remove(destination_folder)
            print(f"Fichier supprimé: {destination_folder}")
        except OSError as e:
            print(f"Erreur lors de la suppression du fichier: {destination_folder}")
            print(e)

        # Vérifier si c'est la dernière itération
        if last_iteration:
            sys.stdout.insert_colored_text(f"L'exécution du Script est terminée, vous pouvez fermer "
                                           f"l'interface GUI en cliquant sur la croix\n", "red", bold=True)
            break

            # Fermer Excel si c'était la dernière itération
        if last_iteration:
            workbook.Close(SaveChanges=False)  # Fermer le classeur sans enregistrer les modifications
            excel.Quit()  # Fermer Excel

if __name__ == "__main__":
    create_gui()  
    
