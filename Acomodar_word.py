import streamlit as st
from docx import Document
from datetime import timedelta
import re
import io

def tc_to_timedelta(tc, fps):
    try:
        h, m, s, f = map(int, tc.strip().split(':'))
        total_seconds = h * 3600 + m * 60 + s + f / fps
        return timedelta(seconds=total_seconds)
    except Exception:
        raise ValueError("Formato de TC inválido. Usá hh:mm:ss:ff")

def timedelta_to_tc(td, fps):
    total_seconds = td.total_seconds()
    h = int(total_seconds // 3600)
    m = int((total_seconds % 3600) // 60)
    s = int(total_seconds % 60)
    f = int(round((total_seconds - int(total_seconds)) * fps))
    if f >= fps:
        f = 0
        s += 1
        if s >= 60:
            s = 0
            m += 1
            if m >= 60:
                m = 0
                h += 1
    return f"{h:02}:{m:02}:{s:02}:{f:02}"

def ajustar_tc(doc, delta, fps):
    tc_pattern = re.compile(r'\b\d{2}:\d{2}:\d{2}:\d{2}\b')
    
    for para in doc.paragraphs:
        matches = tc_pattern.findall(para.text)
        new_text = para.text
        for match in matches:
            updated_tc = timedelta_to_tc(tc_to_timedelta(match, fps) + delta, fps)
            new_text = new_text.replace(match, updated_tc)
        para.text = new_text

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    matches = tc_pattern.findall(para.text)
                    new_text = para.text
                    for match in matches:
                        updated_tc = timedelta_to_tc(tc_to_timedelta(match, fps) + delta, fps)
                        new_text = new_text.replace(match, updated_tc)
                    para.text = new_text

# Interfaz en Streamlit
st.title("Ajustador de Timecodes por TC de Referencia")

# Subir archivo
archivo = st.file_uploader("Seleccioná un archivo DOCX", type="docx")

# Entradas de TC original y nuevo
tc_original = st.text_input("TC original de referencia (ej. 01:00:00:00):")
tc_nuevo = st.text_input("Nuevo TC deseado (ej. 01:02:30:10):")

# FPS
fps = st.selectbox("Framerate (fps):", ["23.976", "24", "25", "29.97"])

# Procesar archivo
if st.button("Procesar"):
    if archivo and tc_original and tc_nuevo:
        try:
            fps = float(fps)
            td_original = tc_to_timedelta(tc_original, fps)
            td_nuevo = tc_to_timedelta(tc_nuevo, fps)
            delta = td_nuevo - td_original

            # Cargar el archivo DOCX desde el uploader de Streamlit
            doc = Document(io.BytesIO(archivo.read()))
            ajustar_tc(doc, delta, fps)

            # Guardar el archivo ajustado en memoria
            output = io.BytesIO()
            doc.save(output)
            output.seek(0)

            # Descargar archivo ajustado
            st.download_button(
                label="Descargar archivo ajustado",
                data=output,
                file_name=f"{archivo.name}_AJUSTADO.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        except Exception as e:
            st.error(f"Error: {e}")
    else:
        st.error("Por favor, completa todos los campos.")
