import pandas as pd
import streamlit as st
import unicodedata
import re
from io import BytesIO
import difflib

# Mapeo de carreras
carrera_equivalencias = {
    "ingi": "industrial",
    "inge": "eléctrica",
    "ingo": "obras civiles",
    "ingc": "ciencia de la computación",
    "inga": "ambiental",
    "industrial": "industrial",
    "electrica": "eléctrica",
    "obras civiles": "obras civiles",
    "ciencia de la computacion": "ciencia de la computación",
    "ambiental": "ambiental"
}

def palabras_similares(palabra1, palabra2, umbral=0.8):
    """Compara dos palabras y retorna True si son suficientemente similares."""
    return difflib.SequenceMatcher(None, palabra1, palabra2).ratio() >= umbral

def normalizar_columnas(df):
    """Normaliza columnas de nombres y carreras en un DataFrame."""
    columnas = ["Nombre", "Email", "Título", "Carrera", "Guia_Interno", "Guia_Externo"]
    for col in columnas:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.lower()
            df[col] = df[col].apply(lambda x: unicodedata.normalize("NFKD", x).encode("ASCII", "ignore").decode("utf-8"))
            df[col] = df[col].str.replace(r"[\\/]", " ", regex=True)
            
    return df

def validar_columnas(df, columnas, nombre_archivo):
    """Valida que existan las columnas requeridas en el DataFrame."""
    for col in columnas:
        if col not in df.columns:
            st.error(f"La columna '{col}' no existe en el archivo {nombre_archivo}. Verifica el formato.")
            return False
    return True

def nombres_coinciden_por_palabras(nombre1, nombre2, min_palabras=3):
    """Compara dos nombres y retorna True si al menos min_palabras son similares."""
    if not isinstance(nombre1, str) or not isinstance(nombre2, str):
        return False
    if not nombre1 or not nombre2:
        return False
    palabras1 = nombre1.split()
    palabras2 = nombre2.split()
    usadas = set()
    coincidencias = 0
    for p1 in palabras1:
        for idx, p2 in enumerate(palabras2):
            if idx not in usadas and palabras_similares(p1, p2):
                coincidencias += 1
                usadas.add(idx)
                break
    return coincidencias >= min_palabras

def cargar_datos(file_saf, file_banner):
    """Carga y normaliza los datos de los archivos SAF y Banner."""
    try:
        df_saf = pd.read_excel(file_saf)
        df_banner = pd.read_excel(file_banner)
    except Exception as e:
        st.error(f"Error al leer los archivos: {e}")
        return None, None

    column_mapping_saf = {
        "RUT": "Rut",
        "Nombre alumno": "Nombre",
        "Email alumno": "Email",
        "Especialidad": "Carrera",
        "Tema memoria": "Título",
        "Guía Interno": "Guia_Interno",
        "Guía Externo": "Guia_Externo"
    }
    column_mapping_banner = {
        "Rut Alumno": "Rut",
        "Nombres": "Nombre",
        "Apellidos": "Apellido",
        "Correo": "Email",
        "Carrera": "Carrera",
        "Título": "Título",
        "Prof. Guía Interno": "Guia_Interno",
        "Prof. Guía Externo": "Guia_Externo"
    }

    df_saf.rename(columns=column_mapping_saf, inplace=True)
    df_banner.rename(columns=column_mapping_banner, inplace=True)
    if "Apellido" in df_banner.columns:
        df_banner["Nombre"] = df_banner["Nombre"] + " " + df_banner["Apellido"]
        df_banner.drop(columns=["Apellido"], inplace=True)

    columnas_requeridas = ["Rut", "Nombre", "Email", "Carrera", "Título", "Guia_Interno", "Guia_Externo"]
    if not validar_columnas(df_saf, columnas_requeridas, "SAF") or not validar_columnas(df_banner, columnas_requeridas, "Banner"):
        return None, None

    df_saf = normalizar_columnas(df_saf)
    df_banner = normalizar_columnas(df_banner)

    df_saf["Carrera"] = df_saf["Carrera"].map(carrera_equivalencias).fillna(df_saf["Carrera"])
    df_banner["Carrera"] = df_banner["Carrera"].map(carrera_equivalencias).fillna(df_banner["Carrera"])
    return df_saf, df_banner

def comparar_datos(df_saf, df_banner):
    if df_saf is None or df_banner is None:
        return None

    df_saf["Rut"] = df_saf["Rut"].astype(str).str.strip()
    df_banner["Rut"] = df_banner["Rut"].astype(str).str.strip()

    columnas_clave = ["Nombre", "Email", "Título", "Carrera", "Guia_Interno", "Guia_Externo"]
    for col in columnas_clave:
        if col not in df_saf.columns or col not in df_banner.columns:
            st.error(f"La columna '{col}' no existe en uno de los archivos. Verifica el formato.")
            return None
    
    # Merge para comparación
    df_comparacion = df_saf.merge(df_banner, on="Rut", how="outer", suffixes=("_SAF", "_Banner"))

    # Indicadores de coincidencia por campo
    condiciones = {
        "Nombre_Coincide": df_comparacion.apply(
            lambda row: nombres_coinciden_por_palabras(row["Nombre_SAF"], row["Nombre_Banner"]), axis=1),
        "Email_Coincide": df_comparacion["Email_SAF"] == df_comparacion["Email_Banner"],
        "Título_Coincide": df_comparacion.apply(
            lambda row: palabras_similares(str(row["Título_SAF"]), str(row["Título_Banner"])), axis=1),
        "Carrera_Coincide": df_comparacion["Carrera_SAF"] == df_comparacion["Carrera_Banner"],
        "Guia_Interno_Coincide": df_comparacion["Guia_Interno_SAF"] == df_comparacion["Guia_Interno_Banner"],
        "Guia_Externo_Coincide": df_comparacion["Guia_Externo_SAF"] == df_comparacion["Guia_Externo_Banner"]
    }

    for col, condition in condiciones.items():
        df_comparacion[col] = condition

    return df_comparacion

def exportar_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Discrepancias')
        workbook = writer.book
        worksheet = writer.sheets['Discrepancias']

        # Formato para las celdas con discrepancias
        formato_rojo = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
        
        for col_num, value in enumerate(df.columns):
            if 'Coincide' in value:
                worksheet.conditional_format(1, col_num, len(df), col_num,
                                             {'type': 'cell',
                                              'criteria': '==',
                                              'value': False,
                                              'format': formato_rojo})

    output.seek(0)
    return output

def main():
    st.title("Comparador de Reportes SAF y Banner")
    st.write("Sube los archivos de SAF y Banner para comparar los datos.")

    file_saf = st.file_uploader("Subir archivo SAF", type=["xls", "xlsx"])
    file_banner = st.file_uploader("Subir archivo Banner", type=["xls", "xlsx"])

    if file_saf and file_banner:
        df_saf, df_banner = cargar_datos(file_saf, file_banner)
        if df_saf is None or df_banner is None:
            st.error("Error al cargar los datos. Verifica el formato de los archivos.")
            return
        
        df_comparacion = comparar_datos(df_saf, df_banner)

        if df_comparacion is not None and not df_comparacion.empty:
            if "Celular" in df_comparacion.columns:
                df_comparacion["Celular"] = df_comparacion["Celular"].astype(str)
                
            st.write("### Resultado de la Comparación:")
            st.dataframe(df_comparacion)

            # Botón de descarga de Excel con formato
            excel_file = exportar_excel(df_comparacion)
            st.download_button(
                label="Descargar Discrepancias en Excel",
                data=excel_file,
                file_name="discrepancias_formateadas.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.success("No se encontraron discrepancias entre los reportes.")

if __name__ == "__main__":
    main()