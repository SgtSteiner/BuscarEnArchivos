import os
import time
from datetime import datetime, timedelta
import streamlit as st
import tkinter as tk
from tkinter import filedialog
import PyPDF2
import textract


def select_folder():
   root = tk.Tk()
   root.withdraw()
   folder_path = filedialog.askdirectory(master=root)
   root.destroy()
   return folder_path

def buscar_palabras_en_archivos():
    progress_text = "Búsqueda en progreso. Espere, por favor."
    my_bar = st.progress(0, text=progress_text)
    palabras = palabras_a_buscar.split(",")

    if not selected_folder_path:
        st.error('Seleccione una carpeta', icon="🚨")
        return

    if not os.path.isdir(selected_folder_path):
        st.error('La ruta seleccionada no es una carpeta.', icon="🚨")
        return

    # Selecciona los archivos que cumplen con los filtros
    archivos = [f for f in os.listdir(selected_folder_path) if (
            os.path.isfile(os.path.join(selected_folder_path, f)) and
            (f.endswith(".pdf") or
             f.endswith(".docx") or
             f.endswith(".doc")) and
            (time.mktime(cal_inicio.timetuple()) <=
             os.path.getmtime(os.path.join(selected_folder_path, f)) <=
             time.mktime(cal_final.timetuple()))
    )]

    resultados = {}

    archivos_total = len(archivos)
    archivos_actual = 0
    st.write(f"Buscando en {archivos_total} archivos...\n")

    for archivo in archivos:
        try:
            archivos_actual += 1
            my_bar.progress(archivos_actual*100 // archivos_total, text=progress_text)
            if archivo.endswith(".pdf"):  # archivos pdf
                with open(os.path.join(selected_folder_path, archivo), 'rb') as f:
                    pdf_reader = PyPDF2.PdfReader(f)
                    for pagina in range(len(pdf_reader.pages)):
                        texto = pdf_reader.pages[pagina].extract_text()
                        for palabra in palabras:
                            if palabra.upper() in texto.upper():
                                if archivo not in resultados:
                                    resultados[archivo] = []
                                resultados[archivo].append(palabra)
            else:  # archivos docx y doc
                texto = textract.process(os.path.join(selected_folder_path, archivo)).decode('utf-8')
                for palabra in palabras:
                    if palabra.upper() in texto.upper():
                        if archivo not in resultados:
                            resultados[archivo] = []
                        resultados[archivo].append(palabra)
        except Exception as e:
            print(f"Error al procesar el archivo {archivo}: {e}")

    if resultados:
        result_text = f"Archivos encontrados: {len(resultados)}\n"
        result_text += "\n".join(resultados)
    else:
        result_text = "No se han encontrado resultados"
        
    my_results = st.text_area("Resultado de la búsqueda", result_text)
        
    my_bar.empty()
        
st.title("Buscador de CVs")

# Seleccionar carpeta
selected_folder_path = st.session_state.get("folder_path", None)
folder_select_button = st.button("Selecciona carpeta")
if folder_select_button:
  selected_folder_path = select_folder()
  st.session_state.folder_path = selected_folder_path
  
if selected_folder_path:
   st.write("Carpeta seleccionada:", selected_folder_path)

# Desde fecha archivo
cal_inicio = st.date_input(
    "Desde fecha archivo", 
    datetime.now() - timedelta(days=365),
    format="DD/MM/YYYY"
)

# Hasta fecha archivo
cal_final = st.date_input(
    "Desde fecha archivo", 
    datetime.now(),
    format="DD/MM/YYYY"
)

# Palabras a buscar
palabras_a_buscar = st.text_input(
    "Palabras",
    help="Palabras a buscar separadas por coma(',')"
)

# Buscar
if st.button('Buscar', type="primary"):
    buscar_palabras_en_archivos()
