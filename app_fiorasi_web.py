import streamlit as st
import pandas as pd
import io
from PyPDF2 import PdfReader
from docx import Document
from datetime import datetime
from PIL import Image

# --- Configuración inicial ---
st.set_page_config(page_title="Ford Fiorasi – Procesador de Antecedentes", page_icon="⚙️", layout="wide")

# --- Variables de color configurables ---
if "color_primario" not in st.session_state:
    st.session_state["color_primario"] = "#0047AB"  # Azul Ford
if "color_fondo" not in st.session_state:
    st.session_state["color_fondo"] = "#FFFFFF"  # Blanco

# --- Encabezado con logo ---
col1, col2, col3 = st.columns([1, 3, 1])
with col2:
    try:
        logo = Image.open("logo_fiorasi.png")
        st.image(logo, use_container_width=False)
    except Exception:
        st.markdown("### Ford Fiorasi – Procesador de Antecedentes Disciplinarios")

st.markdown(
    f"<div style='text-align:center; color:{st.session_state['color_primario']}; font-size:26px; font-weight:600;'>"
    "Procesador de Antecedentes Disciplinarios</div>",
    unsafe_allow_html=True,
)

# --- Menú de ajustes ---
with st.sidebar:
    st.markdown("## ⚙️ Ajustes de apariencia")
    st.session_state["color_primario"] = st.color_picker("Color institucional (Azul Ford)", st.session_state["color_primario"])
    st.session_state["color_fondo"] = st.color_picker("Color de fondo", st.session_state["color_fondo"])
    if st.button("🔄 Restaurar colores predeterminados"):
        st.session_state["color_primario"] = "#0047AB"
        st.session_state["color_fondo"] = "#FFFFFF"

# --- Carga de archivos ---
st.markdown("---")
st.markdown("### 📂 Cargar archivos de antecedentes (PDF o Word)")
archivos = st.file_uploader("Seleccionar múltiples archivos", type=["pdf", "docx"], accept_multiple_files=True)

# --- Procesamiento ---
def extraer_texto_pdf(archivo):
    texto = ""
    try:
        reader = PdfReader(archivo)
        for page in reader.pages:
            texto += page.extract_text() + "\n"
    except:
        texto = ""
    return texto.strip()

def extraer_texto_docx(archivo):
    texto = ""
    try:
        doc = Document(archivo)
        for p in doc.paragraphs:
            texto += p.text + "\n"
    except:
        texto = ""
    return texto.strip()

def procesar_archivo(nombre, contenido):
    texto = contenido.lower()
    data = {"Apellido y Nombre": "", "Fecha de Emisión": "", "Tipo de Antecedente": "", "Contestación": "No", "Resumen": ""}

    # --- Nombre ---
    for linea in texto.split("\n"):
        if "sr." in linea or "sra." in linea or "srta." in linea:
            data["Apellido y Nombre"] = linea.replace("sr.", "").replace("sra.", "").replace("srta.", "").strip().title()
            break

    # --- Fecha ---
    posibles_fechas = [p for p in texto.split() if "/" in p or "-" in p]
    if posibles_fechas:
        try:
            data["Fecha de Emisión"] = datetime.strptime(posibles_fechas[0], "%d/%m/%Y").strftime("%d/%m/%Y")
        except:
            data["Fecha de Emisión"] = posibles_fechas[0]

    # --- Tipo de antecedente ---
    if "llamado de atención" in texto:
        data["Tipo de Antecedente"] = "Llamado de Atención"
    elif "apercibimiento" in texto:
        data["Tipo de Antecedente"] = "Apercibimiento"
    elif "descargo" in texto:
        data["Tipo de Antecedente"] = "Solicitud de Descargo"
    else:
        data["Tipo de Antecedente"] = "Otro"

    # --- Contestación ---
    if "contesta" in texto or "descargo presentado" in texto or "responde" in texto:
        data["Contestación"] = "Sí"

    # --- Resumen ---
    resumen = " ".join(texto.split()[:40]) + "..."
    data["Resumen"] = resumen

    return data

if archivos:
    registros = []
    for archivo in archivos:
        if archivo.type == "application/pdf":
            contenido = extraer_texto_pdf(archivo)
        else:
            contenido = extraer_texto_docx(archivo)
        datos = procesar_archivo(archivo.name, contenido)
        registros.append(datos)

    # Crear DataFrame y ordenar
    df = pd.DataFrame(registros)
    df.sort_values(by="Apellido y Nombre", inplace=True)

    # Mostrar en pantalla
    st.success(f"✅ Se procesaron {len(df)} archivos correctamente.")
    st.dataframe(df)

    # Generar Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Base de Datos", index=False)
        resumen = df[["Apellido y Nombre", "Resumen"]]
        resumen.to_excel(writer, sheet_name="Resumen de Casos", index=False)

    st.download_button(
        label="📥 Descargar Excel procesado",
        data=output.getvalue(),
        file_name="FordFiorasi_Antecedentes.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Subí los archivos PDF o Word para comenzar el procesamiento.")
