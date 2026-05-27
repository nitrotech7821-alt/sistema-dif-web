import streamlit as st
import imaplib
import email
import openpyxl
import io
import zipfile
import xml.etree.ElementTree as ET
import os

st.title("🤖 Robot Administrativo DIF")

def llenar_y_descargar():
    # 1. VERIFICACIÓN DE ARCHIVOS
    if not os.path.exists("cotizacion.xlsx") or not os.path.exists("requisicion.xlsx"):
        st.error("¡Error! No encuentro los archivos 'cotizacion.xlsx' o 'requisicion.xlsx' en la carpeta.")
        return

    # 2. CONEXIÓN (Simulada para prueba de descarga)
    # Si quieres que se conecte al correo, descomenta la parte de IMAP
    # Por ahora, probemos si genera el archivo con datos falsos para ver si descarga.
    datos = {"nombre": "PROVEEDOR PRUEBA", "total": "1000", "folio": "1721"}

    # 3. Llenar Excel
    wb_cot = openpyxl.load_workbook("cotizacion.xlsx")
    ws_cot = wb_cot.active
    ws_cot['H2'] = datos['folio']
    
    wb_req = openpyxl.load_workbook("requisicion.xlsx")
    ws_req = wb_req.active
    ws_req['I9'] = datos['folio']
    
    buf_cot = io.BytesIO()
    wb_cot.save(buf_cot)
    buf_req = io.BytesIO()
    wb_req.save(buf_req)
    
    # 4. Crear ZIP
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, 'w') as zf:
        zf.writestr("Cotizacion_Generada.xlsx", buf_cot.getvalue())
        zf.writestr("Requisicion_Generada.xlsx", buf_req.getvalue())
    
    # 5. BOTÓN DE DESCARGA
    st.download_button(
        label="📥 DESCARGAR EXPEDIENTE GENERADO",
        data=zip_buf.getvalue(),
        file_name="Expediente_Prueba.zip",
        mime="application/zip"
    )

if st.button("Probar Generación de Archivos"):
    llenar_y_descargar()
