import PyPDF2
import tkinter as tk
from tkinter import filedialog, simpledialog, Listbox, messagebox, PhotoImage
import os
import win32print
import win32api
from PIL import Image, ImageTk
import webbrowser

class PDFPrinterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("FirstPrintWizard")
        self.create_ui()
        self.root.iconbitmap('D:/programas VisualStudio/scripts_python/FirstPrintWizard/file_icon.ico')

    def create_ui(self):
        # self.root.configure(bg='#e4e4e4')  # Define un gris oscuro como color de fondo
        self.root.geometry("400x370")

        self.frame = tk.Frame(self.root)
        self.frame.pack(padx=5, pady=5)

        self.selected_printer_label = tk.Label(self.frame, text="Selecciona una impresora:")
        self.selected_printer_label.pack()

        self.printer_var = tk.StringVar()
        self.printer_combo = tk.OptionMenu(self.frame, self.printer_var, *self.get_printer_list())
        self.printer_combo.pack()

        self.select_pdf_button = tk.Button(self.frame, text="Seleccionar archivos PDF", command=self.select_pdf_files)
        self.select_pdf_button.pack()

        self.selected_pdf_label = tk.Label(self.frame, text="Archivos PDF seleccionados:")
        self.selected_pdf_label.pack()

        self.pdf_listbox = Listbox(self.frame, selectmode=tk.MULTIPLE, height=10, width=65)
        self.pdf_listbox.pack()

        self.delete_button = tk.Button(self.frame, text="Eliminar", command=self.delete_selected_pdfs)
        self.delete_button.pack()

        self.print_button = tk.Button(self.frame, text="Imprimir", command=self.print_pdfs)
        self.print_button.pack()

        # Cargar la imagen de GitHub y redimensionarla
        github_icon_path = "D:/programas VisualStudio/scripts_python/FirstPrintWizard/github_icon.png"
        github_image = Image.open(github_icon_path)
        github_resized = github_image.resize((30, 30), resample=Image.BILINEAR)  # Utilizar resample=Image.BILINEAR
        github_icon = ImageTk.PhotoImage(github_resized)

        # Crear un marco para el icono y el texto de GitHub
        github_frame = tk.Frame(self.root)
        github_frame.pack(pady=3)

        # Crear una etiqueta para mostrar la imagen
        github_image_label = tk.Label(github_frame, image=github_icon)
        github_image_label.image = github_icon  # Asegurar que la imagen no sea eliminada por el recolector de basura
        github_image_label.pack(side=tk.LEFT)

        # Crear una etiqueta para mostrar el texto "Powered by Zero9BSC"
        github_text_label = tk.Label(github_frame, text="Powered by Franco Jones", font=("Arial", 10, "bold"))
        github_text_label.pack(side=tk.LEFT)

        # Enlazar la etiqueta de imagen a la función open_github al hacer clic
        github_image_label.bind("<Button-1>", self.open_github)

        self.pdf_files = []

        # Cargar los iconos y redimensionarlos con subsample
        self.printer_icon = self.load_and_resize_icon("D:/programas VisualStudio/scripts_python/FirstPrintWizard/printer_icon.png", 20, 20)
        self.folder_icon = self.load_and_resize_icon("D:/programas VisualStudio/scripts_python/FirstPrintWizard/folder.png", 20, 20)
        self.print_icon = self.load_and_resize_icon("D:/programas VisualStudio/scripts_python/FirstPrintWizard/printer_icon.png", 20, 20)

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
            self.pdf_files.extend(selected_files)
            self.update_pdf_listbox()

    def update_pdf_listbox(self):
        self.pdf_listbox.delete(0, tk.END)
        for pdf_file in self.pdf_files:
            file_name = os.path.basename(pdf_file)
            self.pdf_listbox.insert(tk.END, file_name)

    def delete_selected_pdfs(self):
        selected_indices = self.pdf_listbox.curselection()
        if not selected_indices:
            messagebox.showinfo("Eliminar PDFs", "Selecciona archivos PDF para eliminar.")
            return

        selected_indices = list(selected_indices)
        selected_indices.reverse()  # Reversar la lista para eliminar desde el final

        for index in selected_indices:
            del self.pdf_files[index]
            self.pdf_listbox.delete(index)

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

    def open_github(self, event):
        webbrowser.open_new("https://github.com/Zero9BSC")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFPrinterApp(root)
    root.mainloop()
