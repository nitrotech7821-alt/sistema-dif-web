import streamlit as st
import openpyxl
import io
import zipfile
import os

st.title("🤖 Generador de Expedientes DIF")

# Datos de prueba (estos se obtendrán del XML después)
datos_prueba = {"folio": "1721"}

if st.button("GENERAR Y DESCARGAR"):
    # 1. Validación de existencia de archivos
    archivos_faltantes = []
    if not os.path.exists("cotizacion.xlsx"): archivos_faltantes.append("cotizacion.xlsx")
    if not os.path.exists("requisicion.xlsx"): archivos_faltantes.append("requisicion.xlsx")
    
    if archivos_faltantes:
        st.error(f"¡Error! No encuentro estos archivos en la carpeta del proyecto: {', '.join(archivos_faltantes)}")
    else:
        try:
            # 2. Procesamiento
            wb_cot = openpyxl.load_workbook("cotizacion.xlsx")
            ws_cot = wb_cot.active
            ws_cot['H2'] = datos_prueba['folio'] # Celda H2 para Cotización
            
            wb_req = openpyxl.load_workbook("requisicion.xlsx")
            ws_req = wb_req.active
            ws_req['I7'] = datos_prueba['folio'] # Celda I7 para Requisición
            
            buf_cot = io.BytesIO()
            wb_cot.save(buf_cot)
            buf_req = io.BytesIO()
            wb_req.save(buf_req)
            
            # 3. Empaquetado
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, 'w') as zf:
                zf.writestr("Cotizacion_Llenada.xlsx", buf_cot.getvalue())
                zf.writestr("Requisicion_Llenada.xlsx", buf_req.getvalue())
            
            # 4. Descarga
            st.download_button(
                label="📥 DESCARGAR EXPEDIENTE ZIP",
                data=zip_buf.getvalue(),
                file_name="Expediente_Generado.zip",
                mime="application/zip"
            )
            st.success("¡Expediente generado con éxito!")
            
        except Exception as e:
            st.error(f"Ocurrió un error al procesar el Excel: {e}")
