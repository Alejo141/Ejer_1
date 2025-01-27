import streamlit as st
import pandas as pd
import os
from tkinter import Tk, filedialog

# Función para seleccionar carpeta con un cuadro de diálogo
def seleccionar_carpeta():
    root = Tk()
    root.withdraw()  # Oculta la ventana principal de Tkinter
    carpeta_seleccionada = filedialog.askdirectory()
    return carpeta_seleccionada

# Función para procesar los archivos Excel
def procesar_archivos(carpeta_origen, archivo_salida):
    datos_consolidados = []

    for archivo in os.listdir(carpeta_origen):
        if archivo.endswith('.xlsx') and not archivo.startswith('~$'):
            ruta_archivo = os.path.join(carpeta_origen, archivo)
            try:
                df = pd.read_excel(ruta_archivo, sheet_name=1)
                
                columna_1 = df.iloc[:, 0] # Identificacion
                columna_2 = df.iloc[:, 2].astype(str).str.replace("-", "") # Factura
                columna_3 = df.iloc[:, 6] # Centro costos
                columna_4 = df.iloc[:, 7] # Saldo cartera
                columna_5 = df.iloc[:, 9] # mes
                columna_6 = df.iloc[:, 10] # año

                for valor1, valor2, valor3, valor4, valor5, valor6 in zip(columna_1, columna_2, columna_3, columna_4, columna_5, columna_6):
                    datos_consolidados.append({
                        'Archivo': archivo,
                        'Identificacion': valor1,
                        'Factura': valor2,
                        'Centro de Costo': valor3,
                        'Saldo Factura': valor4,
                        'Mes': valor5,
                        'Año': valor6
                    })
            except Exception as e:
                st.warning(f"Error procesando {archivo}: {e}")

    if datos_consolidados:
        df_consolidado = pd.DataFrame(datos_consolidados)
        df_consolidado.to_excel(archivo_salida, index=False)
        return df_consolidado, archivo_salida
    else:
        return None, None

# Interfaz de Streamlit
st.title("Consolidación de Archivos Excel")

if st.button("Seleccionar Carpeta de Origen"):
    carpeta_origen = seleccionar_carpeta()
    if carpeta_origen:
        st.success(f"Carpeta seleccionada: {carpeta_origen}")
    else:
        st.error("No se seleccionó ninguna carpeta.")
else:
    carpeta_origen = None

archivo_salida = st.text_input("Ruta del archivo de salida:", "H:/Mi unidad/01. ZNI EXPERT/1. EMPRESAS/5. DISPOWER/5. Calculos 026/Cartera Dispower/conso_cartera.xlsx")

if carpeta_origen and st.button("Procesar Archivos"):
    if os.path.exists(carpeta_origen):
        with st.spinner("Procesando archivos..."):
            df_consolidado, archivo_salida_generado = procesar_archivos(carpeta_origen, archivo_salida)
            if df_consolidado is not None:
                st.success(f"Datos consolidados guardados en: {archivo_salida_generado}")
                st.dataframe(df_consolidado)
            else:
                st.error("No se encontraron datos válidos para consolidar.")
    else:
        st.error("La carpeta de origen no existe. Verifique la ruta ingresada.")
