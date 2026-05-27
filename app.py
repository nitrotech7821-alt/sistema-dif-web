import streamlit as st
import openpyxl
import io
import zipfile

st.title("🤖 Prueba de Generación")

if st.button("Generar Archivos de Prueba"):
    # 1. Cargamos archivos
    try:
        wb_cot = openpyxl.load_workbook("cotizacion.xlsx")
        ws_cot = wb_cot.active
        ws_cot['H2'] = "FOLIO-TEST-001" # Celda correcta para Cotización
        
        wb_req = openpyxl.load_workbook("requisicion.xlsx")
        ws_req = wb_req.active
        ws_req['I7'] = "FOLIO-TEST-001" # Celda correcta para Requisición
        
        # 2. Guardamos en memoria
        buf_cot = io.BytesIO()
        wb_cot.save(buf_cot)
        buf_req = io.BytesIO()
        wb_req.save(buf_req)
        
        # 3. Empaquetamos
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, 'w') as zf:
            zf.writestr("Cotizacion_Final.xlsx", buf_cot.getvalue())
            zf.writestr("Requisicion_Final.xlsx", buf_req.getvalue())
        
        # 4. Descarga
        st.download_button(
            label="📥 DESCARGAR AHORA",
            data=zip_buf.getvalue(),
            file_name="Expediente_Prueba.zip",
            mime="application/zip"
        )
        st.success("¡Archivos listos para descargar!")
        
    except Exception as e:
        st.error(f"Error: {e}. Asegúrate de que los archivos .xlsx existan en la carpeta.")
