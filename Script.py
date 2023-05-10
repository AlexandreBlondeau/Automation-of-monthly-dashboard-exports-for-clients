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
    # Initialize an empty list to store the values to be excluded
    values_to_skip = []
    # Initialise the cell index to 2 (to start at cell O2, this cell is editable in the GUI)
    i = 2
    # Loop indefinitely until a stop condition is met
    while True:
        cell_value = worksheet.Range(f"O{i}").Value
        # Check if the cell value is None or empty, if so, exit the loop
        if cell_value is None or cell_value == "":
            break
        # Add the value of the cell to the list of values to exclude
        values_to_skip.append(cell_value)
        # Increase the cell index to move to the next cell
        i += 1
    # Return the list of values to exclude
    return values_to_skip

# Waits for all asynchronous Excel queries to finish
def wait_for_sheets_to_refresh(excel):
    excel.Application.CalculateUntilAsyncQueriesDone()


def attendre_excel(excel):
    while True:
        try:
            # Check if Excel is interactive (i.e. ready to use)
            if excel.Interactive is False:
                time.sleep(1)   # Wait 1 second before checking again
            else:
                break # Exits the loop if Excel is interactive (i.e. ready to be used)
        # Catching communication errors with Excel
        except pywintypes.com_error as e:
            print(f"Erreur de communication avec Excel: {e}")
            time.sleep(1)  # Wait 1 second before trying again

# Definition of the `close_excel` function
def fermer_excel():
    # Initializes a variable to track whether or not Excel has been closed
    excel_ferme = False

    # Browse the list of running processes
    for process in psutil.process_iter():
        try:
            # Retrieves the process information (ID and name)
            process_info = process.as_dict(attrs=['pid', 'name'])

            # Checks if the process is Excel
            if process_info['name'] == 'EXCEL.EXE':
                # Kill the Excel process
                os.kill(process_info['pid'], 9)
                # Updates the variable to indicate that Excel has been closed
                excel_ferme = True
                # Exit the loop
                break

        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            # Ignores exceptions and continues the loop to check other processes
            pass

    if not excel_ferme:
        # Returns False if Excel has not been closed
        return False

    # Returns True if Excel has been closed
    return excel_ferme

class TextRedirector:
    # Initializes the TextRedirector object with a widget and attributes to handle repeated messages
    def __init__(self, widget):
        self.widget = widget
        self.last_text = None
        self.repeat_count = 1

    def write(self, text):
        # Writes the text given in the widget by grouping repeated messages
        text = text.strip()

        # Add a condition to detect specific error messages
        is_error_message = "Erreur de communication avec Excel" in text

        if text == self.last_text or is_error_message:
            self.repeat_count += 1
            # Delete the last message displayed
            self.widget.after_idle(lambda: self.widget.delete("end-2l", "end-1c"))
            # Display the updated message with the counter
            self.widget.after_idle(lambda: self.widget.insert(tk.END, f"{text} x{self.repeat_count}\n"))
        else:
            self.repeat_count = 1
            self.widget.after_idle(lambda: self.widget.insert(tk.END, text + "\n"))
            self.widget.after_idle(lambda: self.widget.see(tk.END))

        self.last_text = text

    def flush(self):
        # This method is present for compatibility with sys.stdout but has no real effect
        pass

    def insert_colored_text(self, text, color, bold=False):
        # Inserts coloured text in the widget with the possibility to make it bold
        tag_name = color + ("_bold" if bold else "")
        self.widget.tag_configure(tag_name, foreground=color, font=("TkDefaultFont", 10, "bold" if bold else "normal"))
        self.widget.after_idle(lambda: self.widget.insert(tk.END, text, (tag_name,)))
        self.widget.after_idle(lambda: self.widget.see(tk.END))

# Part 1: the GUI
def create_gui():
    def execute_second_part():
        # Recovery of the value entered by the user
        user_input = entry.get()

        # Convert the input value to an integer
        user_iteration = int(user_input)

        # Creating a thread to execute the second part of the script
        second_part_thread = threading.Thread(target=second_part, args=(
        user_iteration,))
        second_part_thread.start()

    def check_entry_value(value):
        # Checks if the value is empty
        if not value:
            # Disables the "Start Script" button
            start_button.configure(state="disabled")
            # Allows the value in the input field to be changed
            return True

        try:
            # Tries to convert the value to an integer
            number = int(value)
            # Checks if the value is greater than or equal to 2
            if number >= 2:
                # Activates the "Start Script" button
                start_button.configure(state="normal")
            else:
                # Disables the "Start Script" button
                start_button.configure(state="disabled")
            # Allows the value in the input field to be changed
            return True
        # Catch the exception if the conversion to integer fails
        except ValueError:
            # Disables the "Start Script" button
            start_button.configure(state="disabled")
            # Prohibits the modification of the value in the input field
            return False

    # Creation of the GUI (which is responsive)
    # Creates a new Tk object, which represents the main window of the application
    root = tk.Tk()
    # Sets the title of the main window
    root.title("Automatisation Export TBM EQ Excel")

    # Sets the width of the first column to fit the window
    root.columnconfigure(0, weight=1)
    # Sets the height of the first line to fit the window
    root.rowconfigure(0, weight=1)

    # Creates a new frame to organise the widgets inside the main window
    frame = tk.Frame(root)
    # Positions the frame in the main window
    frame.grid(row=0, column=0, sticky='nsew')
    # Sets the width of the second column of the frame to fit the window
    frame.columnconfigure(1, weight=1)
    # Sets the height of the first line of the frame to fit the window
    frame.rowconfigure(0, weight=1)

    global text
    # Creates a drop-down text widget to display the information
    text = scrolledtext.ScrolledText(frame, wrap="word")
    # Positions the scrolling text widget in the frame
    text.grid(row=0, column=0, columnspan=3, rowspan=3, sticky='nsew', padx=5, pady=5)

    # Redirect stdout output to the text widget
    stdout_redirector = TextRedirector(text)
    sys.stdout = stdout_redirector

    # Define the text to be displayed on the label
    label_text = "Cellule où la première itération commence (Default=2 pour la cellule M2)"
    # Create a label with the previously defined text and a maximum width of 300 pixels for the text
    iteration_label = tk.Label(frame, text=label_text, wraplength=300)
    # Position the label in the frame grid at row 1 and column 3, with a padding of 5 pixels around it
    iteration_label.grid(row=1, column=3, padx=5, pady=5, sticky='n')

    # Creation of a text field (Entry) for the user
    entry = tk.Entry(frame)
    entry.grid(row=2, column=3, padx=5, pady=5, sticky='n')

    # Creating a "Start Script" button to run the second part
    start_button = tk.Button(frame, text="Démarrer le script", command=execute_second_part)
    start_button.grid(row=0, column=3, padx=5, pady=5, sticky='n')

    # Create a validatecommand and link it to the check_entry_value function
    validate_cmd = frame.register(check_entry_value)

    # Creation of a text field (Entry) for the user
    entry = tk.Entry(frame, validate="key", validatecommand=(validate_cmd, '%P'))
    entry.grid(row=2, column=3, padx=5, pady=5, sticky='n')

    # Add the default value 2 to the input field
    entry.insert(0, "2")
    entry.focus_set()  # Met le focus sur l'entrée pour éviter de devoir cliquer dessus avant de modifier la valeur

    # Enable the "start_button" by default, as the default value is 2
    start_button = tk.Button(frame, text="Démarrer le script", command=execute_second_part, state="normal")
    start_button.grid(row=0, column=3, padx=5, pady=5, sticky='n')

    # Starts the main event loop of the application
    root.mainloop()

# Second part: the script called via the GUI
def second_part(start_iteration):
    pythoncom.CoInitialize()
    chemin_fichier = r"C:\TBM\NEW TBM EQ Exoé 2022_Template_avec_limites.xlsm"

    # Check if the file already exists and delete it if it does
    if os.path.isfile(chemin_fichier):
        try:
            os.remove(chemin_fichier)
            print(f"Le fichier '{chemin_fichier}' a été supprimé.")
        except Exception as e:
            print(f"Erreur lors de la suppression du fichier '{chemin_fichier}': {e}")
    else:
        print(f"Le fichier '{chemin_fichier}' n'existe pas.")

    # Close Excel
    excel_ferme = False
    attempt = 0
    while not excel_ferme and attempt < 2:
        excel_ferme = fermer_excel()
        attempt += 1
        time.sleep(2)  # Wait 2 seconds before checking again

    if not excel_ferme:
        print("Could not close Excel after 2 attempts or Excel was not open")

    file_path = r'X:\CLIENTS\DIVERS\Reportings Mensuels\NEW TBM EQ Exoé 2022_Template_avec_limites.xlsm'
    destination_folder = "C:\\TBM\\NEW TBM EQ Exoé 2022_Template_avec_limites.xlsm"


    i = start_iteration
    while True:  # Perform the task for the cells in M2 until there are no more alphanumeric values

        # File creation path definition
        chemin = "C:\\TBM"

        # Check if the folder "TBM" exists
        if not os.path.exists(chemin):
            # Si le dossier n'existe pas, le créer
            os.makedirs(chemin)
        else:
            print("Le dossier 'TBM' existe déjà.")

        # Copy the Excel file to the "TBM" folder
        shutil.copy(file_path, destination_folder)
        print(f"Fichier copié à l'emplacement: {destination_folder}")

        # Launch Excel and open the file
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False  # Désactiver les alertes Excel
        workbook = excel.Workbooks.Open(destination_folder, ReadOnly=False)

        # Select the "Settings and Instructions" sheet
        worksheet = workbook.Sheets("Paramètres et Instructions")

        # Retrieve values to exclude in the first iteration
        if i == start_iteration:
            values_to_skip = get_values_to_skip(worksheet)

        # Copy cell Mi into the cell that is being merged "B1:C1"
        cell_to_copy = f"M{i}"
        cell_value = worksheet.Range(cell_to_copy).Value

        # Check if the value of the current cell should be excluded
        while cell_value in values_to_skip:
            i += 1
            cell_value = worksheet.Range(f"M{i}").Value

        worksheet.Range("B1:C1").Value = cell_value

        sys.stdout.insert_colored_text(f"Début de l'itération qui utilise la cellule M{i}\n", "black", bold=True)

        # Find the next non-excluded or empty cell in column M
        i += 1
        next_cell_value = worksheet.Range(f"M{i}").Value

        # Scroll through the cells until you find a non-excluded or empty cell
        while next_cell_value in values_to_skip:
            i += 1
            next_cell_value = worksheet.Range(f"M{i}").Value

        # Check if the cell found is empty
        if next_cell_value is None or next_cell_value == "":
            last_iteration = True
            print("Dernière itération")
        else:
            last_iteration = False

        sys.stdout.insert_colored_text(f"Itération pour le client {cell_value}\n", "red", bold=True)

        time.sleep(1)

        try:
            print ("Le fichier Excel est en cours d'actualisation")
            # Update all data in the Excel file
            workbook.RefreshAll()

            # Wait until the data refresh is complete before continuing
            wait_for_sheets_to_refresh(excel)
        except Exception as e:
            print(e)
            sys.stdout.insert_colored_text(f"Check the PDF file, there may have been a problem in "
                                           f"updating the data for this client.\n", "blue", bold=True)
        finally:
            print("The 'try except' is finished")
        # Run the "Publication" macro
        try:
            print("The PDF is being published, your computer may be slowed down")
            excel.Application.Run("Publication")
        except Exception as e:
            pass

        # Wait for the pdf file to be published before continuing
        attendre_excel(excel)

        # Close Excel
        excel_ferme = False
        attempt = 0
        while not excel_ferme and attempt < 2:
            excel_ferme = fermer_excel()
            attempt += 1
            time.sleep(5)  # Wait 5 seconds before checking again

        if not excel_ferme:
            print("Could not close Excel after 2 attempts or Excel was not open")

        # Delete the file created in C:\\TBM
        try:
            os.remove(destination_folder)
            print(f"File deleted: {destination_folder}")
        except OSError as e:
            print(f"Error when deleting the file: {destination_folder}")
            print(e)

        # Check if this is the last iteration
        if last_iteration:
            sys.stdout.insert_colored_text(f"The execution of the Script is finished, you can close"
                                           f"the software by clicking on the cross\n", "red", bold=True)
            break

            # Close Excel if this was the last iteration
        if last_iteration:
            workbook.Close(SaveChanges=False)  # Close the workbook without saving changes
            excel.Quit()  # Close Excel

if __name__ == "__main__":
    create_gui()  
    
