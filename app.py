import streamlit as st
import openpyxl
import io
import zipfile
import os

st.title("🤖 Generador de Expedientes DIF")

# Función para manejar celdas combinadas
def escribir_en_celda(ws, celda, valor):
    for range_ in ws.merged_cells.ranges:
        if celda in range_:
            ws.unmerge_cells(str(range_))
            break
    ws[celda] = valor

if st.button("GENERAR EXPEDIENTE COMPLETO"):
    try:
        # 1. Definir nombres (simulando los datos que extraerás del XML)
        folio = "1721"
        # Aquí iría el contenido real de tu XML (puedes subirlo o extraerlo del correo)
        contenido_xml = b"<?xml version='1.0'?><factura>Contenido del XML</factura>" 
        
        # 2. Procesar Excel
        wb_cot = openpyxl.load_workbook("cotizacion.xlsx")
        ws_cot = wb_cot.active
        escribir_en_celda(ws_cot, 'H2', folio)
        
        wb_req = openpyxl.load_workbook("requisicion.xlsx")
        ws_req = wb_req.active
        escribir_en_celda(ws_req, 'I7', folio)
        
        # 3. Guardar en memoria
        buf_cot = io.BytesIO()
        wb_cot.save(buf_cot)
        buf_req = io.BytesIO()
        wb_req.save(buf_req)
        
        # 4. Empaquetar todo en el ZIP
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, 'w') as zf:
            # Aquí metemos el XML y los dos Excel
            zf.writestr(f"Factura_{folio}.xml", contenido_xml)
            zf.writestr(f"Cotizacion_{folio}.xlsx", buf_cot.getvalue())
            zf.writestr(f"Requisicion_{folio}.xlsx", buf_req.getvalue())
        
        # 5. Descarga
        st.download_button(
            label="📥 DESCARGAR EXPEDIENTE COMPLETO (ZIP)",
            data=zip_buf.getvalue(),
            file_name=f"Expediente_{folio}.zip",
            mime="application/zip"
        )
        st.success("¡Expediente listo! Contiene: XML, Cotización y Requisición.")
            
    except Exception as e:
        st.error(f"Error: {e}")
