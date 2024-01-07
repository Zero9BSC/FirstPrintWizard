import PyPDF2
import tkinter as tk
from tkinter import filedialog
from tkinter import simpledialog
from tkinter import Listbox
import os
import win32print
import win32api
from PIL import Image, ImageTk

class PDFPrinterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Impresión de PDF")
        self.create_ui()

    def create_ui(self):
        self.root.geometry("500x400")

        self.frame = tk.Frame(self.root)
        self.frame.pack(padx=20, pady=20)

        self.selected_printer_label = tk.Label(self.frame, text="Selecciona una impresora:")
        self.selected_printer_label.pack()

        self.printer_var = tk.StringVar()
        self.printer_combo = tk.OptionMenu(self.frame, self.printer_var, *self.get_printer_list())
        self.printer_combo.pack()

        self.select_pdf_button = tk.Button(self.frame, text="Seleccionar archivos PDF", command=self.select_pdf_files)
        self.select_pdf_button.pack()

        self.selected_pdf_label = tk.Label(self.frame, text="Archivos PDF seleccionados:")
        self.selected_pdf_label.pack()

        self.pdf_listbox = Listbox(self.frame, selectmode=tk.MULTIPLE, height=10, width=50)
        self.pdf_listbox.pack()

        self.print_button = tk.Button(self.frame, text="Imprimir", command=self.print_pdfs)
        self.print_button.pack()

        self.pdf_files = []

        # Cargar los iconos y redimensionarlos con subsample
        self.printer_icon = self.load_and_resize_icon("D:/programas VisualStudio/scripts_python/printer_icon.png", 20, 20)
        self.folder_icon = self.load_and_resize_icon("D:/programas VisualStudio/scripts_python/folder.png", 20, 20)
        self.print_icon = self.load_and_resize_icon("D:/programas VisualStudio/scripts_python/printer_icon.png", 20, 20)

        # Establecer los iconos en los botones
        self.printer_combo.config(image=self.printer_icon, compound=tk.LEFT)
        self.select_pdf_button.config(image=self.folder_icon, compound=tk.LEFT)
        self.print_button.config(image=self.print_icon, compound=tk.LEFT)

    def get_printer_list(self):
        printers = [printer[2] for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)]
        return printers

    def load_and_resize_icon(self, filename, width, height):
        original_image = Image.open(filename)
        resized_image = original_image.resize((width, height))
        photo = ImageTk.PhotoImage(resized_image)
        return photo

    def select_pdf_files(self):
        selected_files = filedialog.askopenfilenames(
            title="Selecciona archivos PDF",
            filetypes=[("Archivos PDF", "*.pdf")]
        )

        if selected_files:
            self.pdf_files = selected_files
            self.update_pdf_listbox()

    def update_pdf_listbox(self):
        self.pdf_listbox.delete(0, tk.END)
        for pdf_file in self.pdf_files:
            file_name = os.path.basename(pdf_file)
            self.pdf_listbox.insert(tk.END, file_name)

    def print_pdfs(self):
        selected_printer = self.printer_var.get()
        if not selected_printer:
            simpledialog.messagebox.showerror("Error", "Selecciona una impresora.")
            return

        if not self.pdf_files:
            simpledialog.messagebox.showerror("Error", "Selecciona archivos PDF para imprimir.")
            return

        try:
            temp_pdf_file = 'temp.pdf'
            pdf_writer = PyPDF2.PdfWriter()
            
            for pdf_file in self.pdf_files:
                with open(pdf_file, 'rb') as pdf_file:
                    pdf_reader = PyPDF2.PdfReader(pdf_file)
                    
                    if pdf_reader.pages:
                        pdf_page = pdf_reader.pages[0]
                        pdf_writer.add_page(pdf_page)

            with open(temp_pdf_file, 'wb') as temp_file:
                pdf_writer.write(temp_file)

            win32print.SetDefaultPrinter(selected_printer)
            win32api.ShellExecute(0, 'print', temp_pdf_file, None, '.', 0)
        except Exception as e:
            simpledialog.messagebox.showerror("Error", f"Error al imprimir: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFPrinterApp(root)
    root.mainloop()
