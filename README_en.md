# README - English version - [Link to the French version](README.md)

## **Project :** Automation of monthly dashboard exports for clients, with a responsive graphical user interface

### **Author :** Alexandre Blondeau - [alex991@live.fr](mailto:alex991@live.fr) - [GitHub](https://github.com/AlexandreBlondeau)  

This [script](Script.py) python performs the following actions:

#### **I.**  

1. Creates a main window for the application with specific dimensions and an appropriate title. Configure the window to be resizable, allowing for a responsive user interface.
2. Added widgets to the main window, such as:
    * A label to describe the purpose of the input field (for example, "Enter the starting cell of column M").
    * An input field to allow the user to specify the cell M from which to start the iterations.
    * A button to start the script with an associated event handler that triggers the script actions when the user clicks the button.
3. Configuration of the widget layout using an appropriate method (e.g. pack, grid or place), to make the interface easy to understand and use. Use column and row weights to ensure a balanced distribution of space when resizing the window.
4. Handling of errors during user input, such as ensuring that the specified cell is valid and displaying an error message if it is not.
5. Setting up an event loop to keep the user interface active and react to user actions, such as the mainloop.
6. When the user clicks the button to start the script, create and start a second thread to execute the second part of the script. This prevents the GUI from freezing while processing the script actions and ensures a smooth user experience.

#### **II.**  

The "second_part" function takes as input a start_iteration parameter, which determines from which iteration the process will start. The start_iteration value is defined by the user in the GUI. Here is a detailed summary of how this function works:

1. Initialize Python COM library with pythoncom.CoInitialize().
2. Defines the path of the Excel file to manipulate: file_path.
3. Checks if the Excel file already exists, and deletes it if it exists. If not, displays a message indicating that the file does not exist.
4. Closes the Excel application with a maximum of two attempts.
5. Defines the path of the source Excel file_path and the path of the destination folder_folder.
6. Starts a while loop to perform the operations on the cells in column M, starting with the iteration indicated by start_iteration.
7. Check if the folder "TBM" exists. If it does not exist, create it. Otherwise, displays a message indicating that the folder already exists.
8. Copies the source Excel file to the destination folder to perform the following actions locally on the "C:³" drive.
9. Launches the Excel application in invisible mode, disables Excel alerts and opens the copied file in read/write mode.
10. Select the "Settings and Instructions" sheet of the Excel file.
11. Retrieves the values to be excluded during the first iteration (which is defined by the user in the GUI).
12. Copy the value of cell Mi to the merged cell "B1:C1".
13. Checks if the value of the current cell should be excluded and moves to the next cell if it is.
14. Displays a message indicating the start of the iteration and the cell Mi used.
15. Increment the index i and read the value of the next cell in column M. Then browse the cells in column M until you find a non-excluded or empty cell.
16. Marks the iteration as the last iteration if the found cell is empty or None. Otherwise, continue with the next iteration.
17. Displays a message indicating the current iteration for the client concerned.
18. Refreshes all data in the Excel file and waits for the refresh to be completed.
19. Runs the "Publish" VBA macro and waits for the PDF file to be published.
20. Closes the Excel application with a maximum of two attempts.
21. Deletes the Excel file created in the "C:³BM" folder.
22. Checks if this is the last iteration. If it is, displays a message indicating the end of the iterations and exits the while loop.
23. Closes the Excel workbook without saving the changes and exits the Excel application if this was the last iteration.

**Dependencies**  

This program uses the following libraries:
- os
- time
- win32com.client
- psutil
- shutil
- pywintypes
- sys
- pythoncom
- tkinter

To install the missing dependencies, you can use the pip command. (No need if you use the .exe file that is provided).  

python -m pip install XXXX


**Use**  

If you have the executable file, run only the ".exe" and go to step 2.

1. Make sure all dependencies are installed.
Run the file Export_Client_.py to start the program.

2. The user interface will appear, indicating the cell where the first iteration starts. By default, this is cell M2.
Enter the desired starting cell and click "Start Script".
The program will read the values in column M and process them iteratively, excluding the values in column O, then generate and export PDF files.

**Known problems**  

- Do not change the name of the original Excel file located in the "X:³" drive, as the script will not be able to find it.
- The code makes a lot of requests to a remote server via "workbook.RefreshAll()", so the Firewall may block the requests. The code makes a lot of requests to a remote server via "workbook.RefreshAll()", so the firewall may block the requests after a certain number of iterations and the following errors may appear in the GUI:
  * Data refresh error (-2147023170, 'Failed to call remote procedure.', None, None)
  * Communication error with Excel: (-2147023174, 'The RPC server is not available.', None, None)

I added at the end of this document an appendix where I explain how to solve this problem to the IT team. If you haven't solved this problem yet, close the script and check if Excel is closed in the task manager. Then retry the script a little later at the iteration where it stopped via the GUI.  
Make sure that the Excel file copied to the original path is not open on another computer on the network while the program is running, as this can cause file locking problems and may prevent the file from being copied to the local "C:³" drive.
If you encounter communication errors with Excel, the program will wait and retry automatically. This may be due to connection problems with Excel or synchronization problems.

**Appendix**  

1. Open the Windows Firewall.
   Authorize the application or function through the Windows Firewall.
   Enable the domain privilege for Windows Management Instrumentation (WMI).
   You can also check other items.

2. The remote computer is blocked by the firewall.
   Solution: Open the Group Policy Object Editor snap-in (gpedit.msc) to edit the Group Policy Object (GPO) used to manage the Windows Firewall settings in your organization. Open Computer Configuration, Administrative Templates, Network, Network Connections, Windows Firewall, and then open Domain Profile or Standard Profile, depending on which profile you want to configure. Enable the following exceptions: Allow Remote Administration Exception and Allow File and Printer Sharing Exception.

3. The host name or IP address is wrong or the remote computer is turned off.
   Solution: Verify that the host name or IP address is correct.

4. The "TCP/IP NetBIOS Helper" service is not working.
   Solution: Make sure the "TCP/IP NetBIOS Helper" service is running and is configured to start automatically after reboot.

5. The "Remote Procedure Call (RPC)" service is not running on the remote computer.
   Solution: Open the services.msc file using the Windows Run function. In Windows Services, verify that the "Remote Procedure Call (RPC)" service is running and that it is configured to start automatically after reboot.

6. The "Windows Management Instrumentation" service is not running on the remote computer.
   Solution: Open the services.msc file using the Windows Run command: Open services.msc using Windows Run. Verify that the Windows Management Instrumentation service is running and that it is configured to start automatically after reboot.

