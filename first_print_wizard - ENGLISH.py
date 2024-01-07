import PyPDF2
import tkinter as tk
from tkinter import filedialog, simpledialog, Listbox, messagebox, PhotoImage, ttk
import os
import win32print
import win32api
from PIL import Image, ImageTk
import webbrowser
import sys
import pkg_resources

class PDFPrinterApp:
    def __init__(self, root):
        # Initialize the PDFPrinterApp class
        self.root = root
        self.root.title("FirstPrintWizard")  # Set the title of the application
        self.root.resizable(False, False)  # Disable window resizing
        self.create_ui()  # Create the user interface
        self.set_icon()  # Set the application icon

    def create_ui(self):
        # Create the user interface elements
        script_dir = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        self.root.geometry("400x430")  # Set the window size
        self.frame = tk.Frame(self.root)  # Create a frame within the window
        self.frame.pack(padx=5, pady=5)  # Pack the frame with padding

        # Labels, buttons, and entry widgets for printer selection and PDF file handling
        self.selected_printer_label = tk.Label(self.frame, text="Select a printer:")
        self.selected_printer_label.pack()

        self.printer_var = tk.StringVar()
        self.printer_combo = tk.OptionMenu(self.frame, self.printer_var, *self.get_printer_list())
        self.printer_combo.pack()

        self.select_pdf_button = tk.Button(self.frame, text="Select PDF files", command=self.select_pdf_files)
        self.select_pdf_button.pack()

        self.page_range_label = tk.Label(self.frame, text="Page range per document (e.g., 1-3, 5, 7-10, leave blank to print only the first page):", wraplength=300)
        self.page_range_label.pack()

        self.page_range_entry = tk.Entry(self.frame)
        self.page_range_entry.pack()

        self.selected_pdf_label = tk.Label(self.frame, text="Selected PDF files:")
        self.selected_pdf_label.pack()

        self.pdf_listbox = Listbox(self.frame, selectmode=tk.MULTIPLE, height=10, width=65)
        self.pdf_listbox.pack()

        self.delete_button = tk.Button(self.frame, text="Delete", command=self.delete_selected_pdfs)
        self.delete_button.pack()

        self.print_button = tk.Button(self.frame, text="Print", command=self.print_pdfs)
        self.print_button.pack()

        # GitHub icon and text
        github_icon_path = os.path.join(script_dir, "github_icon.png")
        github_image = Image.open(github_icon_path)
        github_resized = github_image.resize((30, 30), resample=Image.BILINEAR)
        github_icon = ImageTk.PhotoImage(github_resized)

        github_frame = tk.Frame(self.root)
        github_frame.pack(pady=3)

        github_image_label = tk.Label(github_frame, image=github_icon)
        github_image_label.image = github_icon
        github_image_label.pack(side=tk.LEFT)

        github_text_label = tk.Label(github_frame, text="Powered by Franco Jones", font=("Arial", 10, "bold"))
        github_text_label.pack(side=tk.LEFT)

        github_image_label.bind("<Button-1>", self.open_github)

        self.pdf_files = []

        # Icons for printer, folder, and print actions
        self.printer_icon = self.load_and_resize_icon(os.path.join(script_dir, "printer_icon.png"), 20, 20)
        self.folder_icon = self.load_and_resize_icon(os.path.join(script_dir, "folder.png"), 20, 20)
        self.print_icon = self.load_and_resize_icon(os.path.join(script_dir, "printer_icon.png"), 20, 20)

        # Set icons on corresponding buttons
        self.printer_combo.config(image=self.printer_icon, compound=tk.LEFT)
        self.select_pdf_button.config(image=self.folder_icon, compound=tk.LEFT)
        self.print_button.config(image=self.print_icon, compound=tk.LEFT)

    def get_printer_list(self):
        # Get a list of available printers
        printers = [printer[2] for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)]
        return printers

    def load_and_resize_icon(self, filename, width, height):
        # Load and resize an image to be used as an icon
        try:
            file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
            original_image = Image.open(file_path)
            resized_image = original_image.resize((width, height), resample=Image.BILINEAR)
            photo = ImageTk.PhotoImage(resized_image)
            return photo
        except Exception as e:
            print(f"Error loading image {filename}: {e}")
            return None

    def select_pdf_files(self):
        # Open a file dialog to select PDF files
        selected_files = filedialog.askopenfilenames(
            title="Select PDF files",
            filetypes=[("PDF files", "*.pdf")]
        )

        if selected_files:
            self.pdf_files.extend(selected_files)
            self.update_pdf_listbox()

    def update_pdf_listbox(self):
        # Update the listbox with selected PDF files
        self.pdf_listbox.delete(0, tk.END)
        for pdf_file in self.pdf_files:
            file_name = os.path.basename(pdf_file)
            self.pdf_listbox.insert(tk.END, file_name)

    def delete_selected_pdfs(self):
        # Delete selected PDF files from the listbox
        selected_indices = self.pdf_listbox.curselection()
        if not selected_indices:
            messagebox.showinfo("Delete PDFs", "Select PDF files to delete.")
            return

        selected_indices = list(selected_indices)
        selected_indices.reverse()

        for index in selected_indices:
            del self.pdf_files[index]
            self.pdf_listbox.delete(index)

    def print_pdfs(self):
        # Print selected PDF files
        selected_printer = self.printer_var.get()
        if not selected_printer:
            simpledialog.messagebox.showerror("Error", "Select a printer.")
            return

        if not self.pdf_files:
            simpledialog.messagebox.showerror("Error", "Select PDF files to print.")
            return

        try:
            temp_pdf_file = 'temp.pdf'
            pdf_writer = PyPDF2.PdfWriter()
            out_of_range_docs = []

            for pdf_file in self.pdf_files:
                with open(pdf_file, 'rb') as pdf_file:
                    pdf_reader = PyPDF2.PdfReader(pdf_file)

                    if pdf_reader.pages:
                        page_range_str = self.page_range_entry.get()
                        page_numbers = self.parse_page_range(page_range_str)

                        if not all(1 <= page <= len(pdf_reader.pages) for page in page_numbers):
                            out_of_range_docs.append(os.path.basename(pdf_file.name))
                            continue

                        for page_number in page_numbers:
                            pdf_page = pdf_reader.pages[page_number - 1]
                            pdf_writer.add_page(pdf_page)

            if out_of_range_docs:
                message = (
                    f"The following documents are out of page range and will not be printed:\n\n"
                    f"{', '.join(out_of_range_docs)}\n\nDo you want to continue?"
                )
                response = messagebox.askyesno("Warning", message)

                if not response:
                    return

            with open(temp_pdf_file, 'wb') as temp_file:
                pdf_writer.write(temp_file)

            win32print.SetDefaultPrinter(selected_printer)
            win32api.ShellExecute(0, 'print', temp_pdf_file, None, '.', 0)
        except Exception as e:
            simpledialog.messagebox.showerror("Error", f"Error while printing: {e}")

    def parse_page_range(self, page_range_str):
        # Parse the page range string into a list of page numbers
        if not page_range_str:
            return [1]

        page_numbers = []
        ranges = page_range_str.split(',')
        for r in ranges:
            r = r.strip()
            if '-' in r:
                start, end = map(int, r.split('-'))
                page_numbers.extend(range(start, end + 1))
            else:
                page_numbers.append(int(r))
        return page_numbers

    def open_github(self, event):
        # Open the GitHub page in the default web browser
        webbrowser.open_new("https://github.com/Zero9BSC")

    def set_icon(self):
        # Set the application icon
        try:
            script_dir = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
            icon_path = os.path.join(script_dir, "file_icon.ico")
            self.root.iconbitmap(default=icon_path)
        except tk.TclError:
            pass

if __name__ == "__main__":
    # Run the application
    root = tk.Tk()
    app = PDFPrinterApp(root)
    root.mainloop()
