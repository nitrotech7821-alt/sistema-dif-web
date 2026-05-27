import streamlit as st
import openpyxl
import io
import zipfile
import xml.etree.ElementTree as ET

# Función para rellenar tus Excel
def rellenar_excel(plantilla_path, datos):
    wb = openpyxl.load_workbook(plantilla_path)
    ws = wb.active
    
    if "cotizacion" in plantilla_path:
        # Según tu archivo, los datos están en la columna B (después de las etiquetas en A)
        ws['B2'] = datos['folio']     # Suponiendo Folio en fila 2
        ws['B3'] = datos['nombre']    # Suponiendo Proveedor en fila 3
        ws['B4'] = datos['total']     # Suponiendo Total en fila 4
    elif "requisicion" in plantilla_path:
        ws['B2'] = datos['folio']
        ws['B3'] = datos['nombre']
        ws['B4'] = datos['total']
    
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()

# Interfaz Principal
st.title("📑 Automatizador de Documentos DIF")

if st.button("Procesar y Generar Expedientes"):
    # Estos son los datos que extraeremos del XML
    datos_factura = {"nombre": "PROVEEDOR EJEMPLO", "total": "$1,500.00", "folio": "A-123"}
    
    # Rellenamos basándonos en tus archivos
    cot_bytes = rellenar_excel("cotizacion.xlsx", datos_factura)
    req_bytes = rellenar_excel("requisicion.xlsx", datos_factura)
    
    # Empaquetamos en un ZIP
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w') as zf:
        zf.writestr(f"Cotizacion_{datos_factura['folio']}.xlsx", cot_bytes)
        zf.writestr(f"Requisicion_{datos_factura['folio']}.xlsx", req_bytes)
        
    st.download_button(
        label="📥 Descargar Expediente Completo (ZIP)",
        data=zip_buffer.getvalue(),
        file_name=f"Expediente_{datos_factura['folio']}.zip",
        mime="application/zip"
    )
    st.success("¡Expediente generado correctamente!")