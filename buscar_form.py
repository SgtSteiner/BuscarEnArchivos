import os
import time
from datetime import datetime, timedelta
from idlelib.tooltip import Hovertip
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
from tkcalendar import DateEntry
import babel.numbers  # necesario para convertir a .exe
import PyPDF2
import textract


def buscar_palabras_en_archivos():
    folder_path = folder_var.get()
    palabras = palabras_var.get().split(",")

    files_text.delete('1.0', tk.END)

    if not folder_path:
        messagebox.showerror("Error", "Seleccione una carpeta.")
        return

    if not os.path.isdir(folder_path):
        messagebox.showerror("Error", "La ruta seleccionada no es una carpeta.")
        return

    # Selecciona los archivos que cumplen con los filtros
    archivos = [f for f in os.listdir(folder_path) if (
            os.path.isfile(os.path.join(folder_path, f)) and
            (f.endswith(".pdf") or
             f.endswith(".docx") or
             f.endswith(".doc")) and
            (time.mktime(cal_inicio.get_date().timetuple()) <=
             os.path.getmtime(os.path.join(folder_path, f)) <=
             time.mktime(cal_final.get_date().timetuple()))
    )]

    resultados = {}

    archivos_total = len(archivos)
    archivos_actual = 0
    progressbar.config(maximum=archivos_total)
    files_text.insert(tk.END, f"Buscando en {archivos_total} archivos...\n")

    for archivo in archivos:
        try:
            archivos_actual += 1
            progressbar["value"] = archivos_actual
            root.update()
            if archivo.endswith(".pdf"):  # archivos pdf
                with open(os.path.join(folder_path, archivo), 'rb') as f:
                    pdf_reader = PyPDF2.PdfReader(f)
                    for pagina in range(len(pdf_reader.pages)):
                        texto = pdf_reader.pages[pagina].extract_text()
                        for palabra in palabras:
                            if palabra.upper() in texto.upper():
                                if archivo not in resultados:
                                    resultados[archivo] = []
                                resultados[archivo].append(palabra)
            else:  # archivos docx y doc
                texto = textract.process(os.path.join(folder_path, archivo)).decode('utf-8')
                for palabra in palabras:
                    if palabra.upper() in texto.upper():
                        if archivo not in resultados:
                            resultados[archivo] = []
                        resultados[archivo].append(palabra)
        except Exception as e:
            print(f"Error al procesar el archivo {archivo}: {e}")

    if resultados:
        files_text.insert(tk.END, f"\nArchivos encontrados: {len(resultados)}\n")
        files_text.insert(tk.END, "\n".join(resultados))
    else:
        files_text.insert(tk.END, "No se han encontrado resultados")


def select_folder():
    folder_path = filedialog.askdirectory()
    folder_var.set(folder_path)

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Buscador de CV")
    root.geometry("750x690")

    # Carpeta
    folder_label = ttk.Label(root, text="Carpeta:")
    folder_label.place(x=5, y=9)

    folder_var = tk.StringVar()
    folder_entry = ttk.Entry(root, textvariable=folder_var, width=60)
    folder_entry.place(x=60, y=9)
    Hovertip(folder_entry, 'Carpeta donde se encuentran los archivos a buscar')

    folder_button = ttk.Button(root, text="Seleccionar carpeta", command=select_folder)
    folder_button.place(x=430, y=5)

    # Fecha inicio
    label_fecha_inicio = ttk.Label(root, text="Desde:")
    label_fecha_inicio.place(x=5, y=45)

    anyo_atras = datetime.now() - timedelta(days=365)  # un año atrás
    cal_inicio = DateEntry(
        root,
        date_pattern='dd/MM/yyyy',
        selectmode='day',
        locale="es_ES",
        year=anyo_atras.year,
        month=anyo_atras.month,
        day=anyo_atras.day
    )
    cal_inicio.place(x=60, y=43)
    Hovertip(cal_inicio, 'Fecha inicio de búsqueda')

    # Fecha final
    label_fecha_final = ttk.Label(root, text="Hasta:")
    label_fecha_final.place(x=180, y=45)

    cal_final = DateEntry(
        root,
        date_pattern='dd/MM/yyyy',
        selectmode='day',
        locale="es_ES"
    )
    cal_final.place(x=235, y=43)
    Hovertip(cal_final, 'Fecha fin de búsqueda')

    # Palabras
    palabras_label = ttk.Label(root, text="Palabras:")
    palabras_label.place(x=5, y=80)

    palabras_var = tk.StringVar()
    palabras_entry = ttk.Entry(root, textvariable=palabras_var, width=60)
    palabras_entry.place(x=60, y=80)
    Hovertip(palabras_entry, 'Palabras a buscar separadas por coma(",")')

    # Botón de búsqueda
    search_button = ttk.Button(root, text="Buscar", command=buscar_palabras_en_archivos)
    search_button.place(x=430, y=78)

    # Barra de progreso
    progressbar = ttk.Progressbar(root, mode="determinate")
    progressbar.place(x=5, y=110, width=710)

    # Resultados
    files_scrollbar_y = ttk.Scrollbar(root, orient=tk.VERTICAL)
    files_scrollbar_y.place(x=720, y=140, height=480)

    files_scrollbar_x = ttk.Scrollbar(root, orient=tk.HORIZONTAL)
    files_scrollbar_x.place(x=5, y=630, width=710)

    files_text = tk.Text(
        root,
        yscrollcommand=files_scrollbar_y.set,
        xscrollcommand=files_scrollbar_x.set,
        wrap=tk.NONE,
        width=88,
        height=30,
    )
    files_text.place(x=5, y=140)
    files_scrollbar_y.config(command=files_text.yview)
    files_scrollbar_x.config(command=files_text.xview)

    # Botón salir
    exit_button = ttk.Button(root, text="Salir", command=root.destroy)
    exit_button.place(x=330, y=650)

    root.mainloop()
