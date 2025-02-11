import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# ------------------------------------------------------------------------ #
# FUNCIÓN DE PROCESAMIENTO DE EXCEL (ADAPTADA PARA STREAMLIT)
def get_litres_referencia(referencia, concepto):
    if not isinstance(referencia, str):
        return 0.0
    
    def extract_suffix(ref):
        match = re.search(r'(\d+)[CI]?$', ref)
        return match.group(1) if match else ""
    
    suffix = extract_suffix(referencia)
    
    mapping = {
        '10': 10.0,
        '20': 20.0,
        '30': 30.0,
        '44': 0.44,
        '33': 0.33,
        '33C': 0.33,
        '44C': 0.44,
        '37': 0.37,
    }
    
    if suffix in mapping:
        return mapping[suffix]
    
    if concepto and isinstance(concepto, str) and concepto.endswith('33C'):
        return 0.33
    
    return 0.0

def procesar_xls(df):
    try:
        columnas = {
            'Almacén': None,
            'Fecha': None,
            'Referencia': None,
            'Descripción': None,
            'Concepto': None,
            'Documento': None,
            'Cliente / Prov.': None,
            'Cantidad': None,
            'Precio': None,
        }
        
        header_identified = False
        header_row = -1
        for idx, row in df.iterrows():
            for col_idx, value in enumerate(row):
                if pd.notnull(value):
                    value_str = str(value).strip()
                    if columnas['Almacén'] is None and 'Almacén' in value_str:
                        columnas['Almacén'] = col_idx
                    if columnas['Fecha'] is None and 'Fecha' in value_str:
                        columnas['Fecha'] = col_idx
                    if columnas['Referencia'] is None and 'Referencia' in value_str:
                        columnas['Referencia'] = col_idx
                    if columnas['Descripción'] is None and 'Descripción' in value_str:
                        columnas['Descripción'] = col_idx
                    if columnas['Concepto'] is None and 'Concepto' in value_str:
                        columnas['Concepto'] = col_idx
                    if columnas['Documento'] is None and 'Documento' in value_str:
                        columnas['Documento'] = col_idx
                    if columnas['Cliente / Prov.'] is None and 'Cliente / Prov.' in value_str:
                        columnas['Cliente / Prov.'] = col_idx
                    if columnas['Cantidad'] is None and 'Cantidad' in value_str:
                        columnas['Cantidad'] = col_idx
                    if columnas['Precio'] is None and 'Precio' in value_str:
                        columnas['Precio'] = col_idx
            
            if all(col is not None for col in columnas.values()):
                header_row = idx
                header_identified = True
                break
        
        if not header_identified:
            raise ValueError("No se pudieron identificar todas las columnas necesarias.")
        
        filas_combinadas = []
        temp_fila = None
        lot = None
        
        header_pattern = re.compile(r'^(Movimientos:|Almacén|Página|Fecha|Referencia)', re.IGNORECASE)
        
        for index, row in df.iterrows():
            if pd.isnull(row[columnas['Referencia']]):
                if temp_fila is not None:
                    temp_fila['LOT'] = lot if lot else ''
                    filas_combinadas.append(temp_fila)
                    temp_fila = None
                    lot = None
            else:
                descripcion = str(row[columnas['Descripción']])
                try:
                    descripcion += " " + str(df.iat[index, columnas['Descripción'] + 1])
                except:
                    pass
                
                concepto = str(row[columnas['Concepto']]).strip()
                try:
                    concepto += " " + str(df.iat[index, columnas['Concepto'] + 1]).strip()
                except:
                    pass
                
                cliente = 0
                try:
                    cliente_str = str(row[columnas['Cliente / Prov.']]).replace("nan", "0").replace(",", ".")
                    cliente = int(float(cliente_str))
                except:
                    pass
                
                cantidad = 0.0
                try:
                    cantidad_str = str(row[columnas['Cantidad']]).replace(",", ".").strip()
                    cantidad = float(cantidad_str)
                except:
                    pass
                
                precio = 0.0
                try:
                    precio_str = str(row[columnas['Precio']]).replace(",", ".").strip()
                    precio = float(precio_str)
                except:
                    pass
                
                current_almacen = row[columnas['Almacén']]
                current_fecha = row[columnas['Fecha']]
                current_documento = row[columnas['Documento']]
                
                current_ref_col = columnas['Referencia']
                current_ref = row[current_ref_col].strip() if pd.notnull(row[current_ref_col]) else ''

                lot = None
                lot_found = False
                for next_index in range(index + 1, len(df)):
                    next_row = df.iloc[next_index]
                    
                    current_row_str = ''
                    for col in range(df.shape[1]):
                        value = next_row[col]
                        if pd.notnull(value):
                            current_row_str = str(value).strip()
                            break
                    
                    if header_pattern.search(current_row_str) or current_row_str == '':
                        continue
                    
                    next_ref_col = columnas['Referencia']
                    next_ref = next_row[next_ref_col].strip() if pd.notnull(next_row[next_ref_col]) else ''

                    if next_ref != '' and next_ref != current_ref and not lot_found:
                        break
                    
                    local_lot = None
                    for col in range(df.shape[1]):
                        lot_cell = str(next_row[col]).strip()
                        match = re.search(r'\d{2}\s*-\s*\d{3}', lot_cell)
                        if match:
                            local_lot = match.group().replace(" ", "")
                            break
                    
                    if local_lot:
                        lot = local_lot
                        lot_found = True
                        break

                temp_fila = {
                    'Almacén': current_almacen,
                    'Fecha': current_fecha,
                    'Referencia': current_ref,
                    'Descripción': descripcion.replace("nan", "").strip(),
                    'Concepto': concepto.replace("nan", "").strip(),
                    'Documento': current_documento,
                    'Cliente / Prov.': cliente,
                    'Cantidad': abs(cantidad),
                    'Precio': abs(precio),
                    'LOT': lot if lot else '',
                }
        
        df_final = pd.DataFrame(filas_combinadas)
        
        df_final.dropna(subset=['Concepto'], inplace=True)
        df_final['Concepto'] = df_final['Concepto'].str.strip()
        df_final = df_final[df_final['Concepto'].isin(['Salida por Factura', 'Entrada por abono en Factura'])]
        
        df_final = df_final[df_final['Descripción'].str.contains('ABV', case=False, na=False)]
        df_final = df_final[df_final['Referencia'].str.lower().str.startswith('e', na=False)]
        
        df_final['LITRES'] = (df_final.apply(lambda row: get_litres_referencia(row['Referencia'], row['Concepto']), axis=1) * df_final['Cantidad']).abs()
        df_final['VALOR'] = (df_final['Cantidad'] * df_final['Precio']).abs()
        
        return df_final
    
    except Exception as e:
        st.error(f"Error durante el procesamiento: {e}")
        return pd.DataFrame()

# ------------------------------------------------------------------------ #
# CONFIGURACIÓN DE LA VISTA EN STREAMLIT
def main():
    st.set_page_config(
        page_title="SILICE - Procesador de Excel",
        page_icon="📈",
    )

    st.title("Aplicación SILICE")
    st.subheader("Procesar Archivos Excel")
    st.write("---")

    uploaded_file = st.file_uploader("Sube un archivo Excel", type=["xlsx"])

    if uploaded_file:
        df = pd.read_excel(uploaded_file, engine="openpyxl")
        
        if st.button("Procesar Datos"):
            df_processed = procesar_xls(df)
            if not df_processed.empty:
                st.success("Datos procesados satisfactoriamente!")
                
                # Mostrar KPIs
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Total de Registros", df_processed.shape[0])
                with col2:
                    st.metric("Registros sin LOTE", (df_processed['LOT'] == '').sum())
                with col3:
                    st.metric("Total de Referencias Únicas", df_processed['Referencia'].nunique())
                
                st.subheader("Datos Procesados:")
                st.data_editor(df_processed, width=1000)
                
                st.subheader("Descarga el Archivo:")
                output = BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df_processed.to_excel(writer, sheet_name="Resultado", index=False)
                
                # Aplicar estilos con openpyxl
                wb = load_workbook(output)
                ws = wb.active
                ws.freeze_panes = "A2"
                ws.auto_filter.ref = ws.dimensions
                
                # Guardar y resetear el búfer
                output = BytesIO()
                wb.save(output)
                output.seek(0)
                
                st.download_button(
                    label="Descargar Excel Procesado",
                    data=output,
                    file_name="Processed_Output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("No se encontraron datos para procesar. Verifica los criterios y el formato del archivo.")

if __name__ == '__main__':
    main()