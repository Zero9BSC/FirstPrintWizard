import PyPDF2
import tkinter as tk
from tkinter import filedialog, simpledialog, Listbox, messagebox, ttk
import os
import win32print
import win32api
from PIL import Image, ImageTk
import webbrowser
import sys

class PDFPrinterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("FirstPrintWizard")
        self.root.resizable(False, False)
        self.create_ui()
        self.set_icon()

    def create_ui(self):
        # Configuración de la interfaz gráfica
        self.root.geometry("400x430")
        self.frame = tk.Frame(self.root)
        self.frame.pack(padx=5, pady=5)

        self.selected_printer_label = tk.Label(self.frame, text="Selecciona una impresora:")
        self.selected_printer_label.pack()

        self.printer_var = tk.StringVar()
        self.printer_combo = tk.OptionMenu(self.frame, self.printer_var, *self.get_printer_list())
        self.printer_combo.pack()

        self.select_pdf_button = tk.Button(self.frame, text="Seleccionar archivos PDF", command=self.select_pdf_files)
        self.select_pdf_button.pack()

        self.page_range_label = tk.Label(self.frame, text="Rango de páginas por documento (ej. 1-3, 5, 7-10, dejar en blanco imprime solo primera pagina):", wraplength=300)
        self.page_range_label.pack()

        self.page_range_entry = tk.Entry(self.frame)
        self.page_range_entry.pack()

        self.selected_pdf_label = tk.Label(self.frame, text="Archivos PDF seleccionados:")
        self.selected_pdf_label.pack()

        self.pdf_listbox = Listbox(self.frame, selectmode=tk.MULTIPLE, height=10, width=65)
        self.pdf_listbox.pack()

        self.delete_button = tk.Button(self.frame, text="Eliminar", command=self.delete_selected_pdfs)
        self.delete_button.pack()

        self.print_button = tk.Button(self.frame, text="Imprimir", command=self.print_pdfs)
        self.print_button.pack()

        # Cargar la imagen de GitHub y redimensionarla
        github_icon_path = self.resource_path("github_icon.png")
        github_image = Image.open(github_icon_path)
        github_resized = github_image.resize((30, 30), resample=Image.BILINEAR)
        github_icon = ImageTk.PhotoImage(github_resized)

        # Crear un marco para el icono y el texto de GitHub
        github_frame = tk.Frame(self.root)
        github_frame.pack(pady=3)

        # Crear una etiqueta para mostrar la imagen
        github_image_label = tk.Label(github_frame, image=github_icon)
        github_image_label.image = github_icon
        github_image_label.pack(side=tk.LEFT)

        # Crear una etiqueta para mostrar el texto "Powered by Franco Jones"
        github_text_label = tk.Label(github_frame, text="Powered by Franco Jones", font=("Arial", 10, "bold"))
        github_text_label.pack(side=tk.LEFT)

        # Enlazar la etiqueta de imagen a la función open_github al hacer clic
        github_image_label.bind("<Button-1>", self.open_github)

        self.pdf_files = []

        # Cargar los iconos y redimensionarlos con subsample
        self.printer_icon = self.load_and_resize_icon("printer_icon.png", 20, 20)
        self.folder_icon = self.load_and_resize_icon("folder.png", 20, 20)
        self.print_icon = self.load_and_resize_icon("printer_icon.png", 20, 20)

        # Establecer los iconos en los botones
        self.printer_combo.config(image=self.printer_icon, compound=tk.LEFT)
        self.select_pdf_button.config(image=self.folder_icon, compound=tk.LEFT)
        self.print_button.config(image=self.print_icon, compound=tk.LEFT)

    def set_icon(self):
        # Establecer el ícono de la aplicación
        icon_path = self.resource_path("file_icon.ico")
        if os.path.exists(icon_path):
            self.root.iconbitmap(default=icon_path)

    def get_printer_list(self):
        printers = [printer[2] for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)]
        return printers

    def load_and_resize_icon(self, filename, width, height):
        try:
            icon_path = self.resource_path(filename)
            original_image = Image.open(icon_path)
            resized_image = original_image.resize((width, height), resample=Image.BILINEAR)
            photo = ImageTk.PhotoImage(resized_image)
            return photo
        except Exception as e:
            print(f"Error al cargar la imagen {filename}: {e}")
            return None

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
        selected_indices.reverse()

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
            out_of_range_docs = []  # Almacenar documentos fuera de rango
            
            for pdf_file in self.pdf_files:
                with open(pdf_file, 'rb') as pdf_file:
                    pdf_reader = PyPDF2.PdfReader(pdf_file)
                    
                    if pdf_reader.pages:
                        page_range_str = self.page_range_entry.get()
                        page_numbers = self.parse_page_range(page_range_str)
                        
                        # Verificar si el documento está fuera de rango
                        if not all(1 <= page <= len(pdf_reader.pages) for page in page_numbers):
                            out_of_range_docs.append(os.path.basename(pdf_file.name))
                            continue

                        for page_number in page_numbers:
                            pdf_page = pdf_reader.pages[page_number - 1]
                            pdf_writer.add_page(pdf_page)

            if out_of_range_docs:
                # Mostrar advertencia y obtener confirmación del usuario
                message = (
                    f"Los siguientes documentos están fuera del rango de páginas y no se imprimirán:\n\n"
                    f"{', '.join(out_of_range_docs)}\n\n¿Deseas continuar?"
                )
                response = messagebox.askyesno("Advertencia", message)
                
                if not response:
                    return

            with open(temp_pdf_file, 'wb') as temp_file:
                pdf_writer.write(temp_file)

            win32print.SetDefaultPrinter(selected_printer)
            win32api.ShellExecute(0, 'print', temp_pdf_file, None, '.', 0)
        except Exception as e:
            simpledialog.messagebox.showerror("Error", f"Error al imprimir: {e}")

    def parse_page_range(self, page_range_str):
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
        webbrowser.open_new("https://github.com/Zero9BSC")

    def resource_path(self, relative_path):
        try:
            # PyInstaller crea un atributo 'sys._MEIPASS'
            base_path = sys._MEIPASS
        except Exception:
            # En caso contrario, utiliza la ruta del script
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFPrinterApp(root)
    root.mainloop()
