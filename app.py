import streamlit as st
import openpyxl
import io
import zipfile
import os

st.title("🤖 Generador de Expedientes DIF")

def escribir_en_celda(ws, celda, valor):
    """Escribe en una celda, manejando celdas combinadas"""
    # Si la celda es parte de un rango combinado, obtenemos la celda principal
    for range_ in ws.merged_cells.ranges:
        if celda in range_:
            ws.unmerge_cells(str(range_)) # Descombinamos temporalmente
            break
    ws[celda] = valor

if st.button("GENERAR Y DESCARGAR"):
    try:
        # 1. Cargar archivos
        wb_cot = openpyxl.load_workbook("cotizacion.xlsx")
        ws_cot = wb_cot.active
        
        wb_req = openpyxl.load_workbook("requisicion.xlsx")
        ws_req = wb_req.active
        
        # 2. Escribir usando la función que maneja celdas combinadas
        escribir_en_celda(ws_cot, 'H2', "1721")
        escribir_en_celda(ws_req, 'I7', "1721")
        
        # 3. Guardar en memoria
        buf_cot = io.BytesIO()
        wb_cot.save(buf_cot)
        buf_req = io.BytesIO()
        wb_req.save(buf_req)
        
        # 4. Empaquetar
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, 'w') as zf:
            zf.writestr("Cotizacion_Llenada.xlsx", buf_cot.getvalue())
            zf.writestr("Requisicion_Llenada.xlsx", buf_req.getvalue())
        
        # 5. Descarga
        st.download_button(
            label="📥 DESCARGAR EXPEDIENTE ZIP",
            data=zip_buf.getvalue(),
            file_name="Expediente_Generado.zip",
            mime="application/zip"
        )
        st.success("¡Éxito! Archivos generados correctamente.")
            
    except Exception as e:
        st.error(f"Error técnico: {e}")
