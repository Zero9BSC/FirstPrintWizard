# FirstPrintWizard

## Description
FirstPrintWizard is a simple application that allows you to select PDF files and print them to a specific printer on your system. The application is designed to be user-friendly and provide a quick solution for printing PDF documents.

## Technologies Used
- Python
- Tkinter (Graphical User Interface)
- PyPDF2 (PDF file manipulation)
- Pillow (Image processing)
- pywin32 (Interaction with Windows functions)

## Steps to Create the Executable

1. **Install Dependencies:**
   Make sure you have Python installed on your system. Then, create a virtual environment and run the following command to install the necessary dependencies:

   ```bash
   pip install -r requirements.txt

2. **Generate the Executable:**
    Ensure PyInstaller is installed. If not, you can install it with:

    ```bash
    pip install pyinstaller
    ```
    Use the following PyInstaller command to generate the executable:

    ```bash
    pyinstaller --name=FirstPrintWizard --onefile --noconsole --icon=file_icon.ico --add-data "github_icon.png;." --add-data "printer_icon.png;." --add-data "folder.png;." --add-data "file_icon.ico;." first_print_wizard.py
    ```
    This command will create a 'dist' folder containing the 'FirstPrintWizard.exe' executable file.

3. **Run the Application:**
    Go to the 'dist' folder and run 'FirstPrintWizard.exe'. The application should start without the need for additional installation.

## Contributions
Contributions are welcome! If you encounter any issues or have improvements, please create an issue or submit a pull request.

## Credits
This project was created by Franco Nicolas Jones and is available under the MIT license.