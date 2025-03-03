import pandas as pd
import re
import os
import sys
from openpyxl import load_workbook

# Define the function to get litres based on the reference and concept
def get_litres_referencia(referencia, concepto):
    if not isinstance(referencia, str):
        return 0.0
    
    # Convert reference to uppercase
    referencia = referencia.upper()
    
    def extract_suffix(ref):
        # Use regular expression to find the suffix
        match = re.search(r'(\d+)[CI]?$', ref.strip().replace(' ', ''))
        return match.group(1) if match else ""
    
    suffix = extract_suffix(referencia)
    
    # Define the mapping between suffixes and litres
    mapping = {
        '10': 10.0,
        '20': 20.0,
        '30': 30.0,
        '44': 0.44,
        '33': 0.33,
        '33C': 0.33,
        '44C': 0.44,
        '37': 0.37,
        '20I': 20.0,  # Specific case for 20I
        '30I': 30.0,  # Specific case for 30I
        '33C': 0.33,  # Specific case for 33C
    }
    
    # Return the corresponding litres or default to 0
    return mapping.get(suffix, 0.0) if suffix else 0.0

def procesar_xls(entrada, salida):
    try:
        # Read the input file into a DataFrame
        df = pd.read_excel(entrada, header=None)
        
        # Define the required columns
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
                # Check each cell for header keywords
                if pd.notnull(value):
                    value_str = str(value).strip().lower()
                    if columnas['Almacén'] is None and 'almacén' in value_str:
                        columnas['Almacén'] = col_idx
                    if columnas['Fecha'] is None and 'fecha' in value_str:
                        columnas['Fecha'] = col_idx
                    if columnas['Referencia'] is None and 'referencia' in value_str:
                        columnas['Referencia'] = col_idx
                    if columnas['Descripción'] is None and 'descripción' in value_str:
                        columnas['Descripción'] = col_idx
                    if columnas['Concepto'] is None and 'concepto' in value_str:
                        columnas['Concepto'] = col_idx
                    if columnas['Documento'] is None and 'documento' in value_str:
                        columnas['Documento'] = col_idx
                    if columnas['Cliente / Prov.'] is None and 'cliente' in value_str and 'prov' in value_str:
                        columnas['Cliente / Prov.'] = col_idx
                    if columnas['Cantidad'] is None and 'cantidad' in value_str:
                        columnas['Cantidad'] = col_idx
                    if columnas['Precio'] is None and 'precio' in value_str:
                        columnas['Precio'] = col_idx
            
            # Check if all columns have been identified
            if all(col is not None for col in columnas.values()):
                header_row = idx
                header_identified = True
                break
        
        if not header_identified:
            raise ValueError("No se pudieron identificar todas las columnas requeridas.")
        
        # Process each row to combine data and extract LOT
        filas_combinadas = []
        temp_fila = None
        lot = None
        lot_found = False
        header_pattern = re.compile(r'^(movimientos:|almacén|página|fecha|referencia)', re.IGNORECASE)
        
        for index, row in df.iterrows():
            # Check if the current row is part of a product's data
            is_product_row = pd.notnull(row[columnas['Referencia']]) if columnas['Referencia'] is not None else False
            
            # Process rows with a valid reference
            if is_product_row:
                # Extract values from current row
                cliente = 0
                try:
                    cliente_str = str(row[columnas['Cliente / Prov.']]).replace("nan", "0").replace(",", ".").strip()
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
                
                current_almacen = row[columnas['Almacén']] if pd.notnull(row[columnas['Almacén']]) else ''
                current_fecha = row[columnas['Fecha']] if pd.notnull(row[columnas['Fecha']]) else ''
                current_documento = row[columnas['Documento']] if pd.notnull(row[columnas['Documento']]) else ''
                
                # Extract reference and handle merged cells
                current_ref = str(row[columnas['Referencia']]).strip().upper() if pd.notnull(row[columnas['Referencia']]) else ''

                # Initialize temporary row data
                temp_fila = {
                    'Almacén': current_almacen,
                    'Fecha': current_fecha,
                    'Referencia': current_ref,
                    'Descripción': str(row[columnas['Descripción']]).strip().replace("nan", "").upper(),
                    'Concepto': str(row[columnas['Concepto']]).strip().replace("nan", "").upper(),
                    'Documento': current_documento,
                    'Cliente / Prov.': cliente,
                    'Cantidad': abs(cantidad),
                    'Precio': abs(precio),
                    'LOT': '',
                }
                
                # Search subsequent rows for the LOT number
                lot = None
                lot_found = False
                for next_index in range(index + 1, len(df)):
                    next_row = df.iloc[next_index]
                    
                    # Check if current row is a header or empty
                    current_row_str = ''
                    for col in range(df.shape[1]):
                        value = next_row[col]
                        if pd.notnull(value):
                            current_row_str = str(value).strip()
                            break
                    if header_pattern.search(current_row_str) or current_row_str == '':
                        continue
                    
                    # Extract next reference and check if it's a new product
                    next_ref = str(next_row[columnas['Referencia']]).strip().lower() if pd.notnull(next_row[columnas['Referencia']]) else ''
                    if next_ref and next_ref.startswith('e') and next_ref != current_ref.lower() and not lot_found:
                        break
                    
                    # Search for LOT pattern in all cells of the current row
                    for col in range(df.shape[1]):
                        lot_cell = str(next_row[col]).strip().upper()
                        match = re.search(r'(\d{2}[-\s]\d{3})', lot_cell)
                        if match:
                            lot = re.sub(r'\s+', '', match.group(1))
                            lot_found = True
                            break
                    if lot_found:
                        break
                
                # Update the temporary row with the found LOT
                temp_fila['LOT'] = lot if lot else ''
                filas_combinadas.append(temp_fila)
        
        # Create the final DataFrame
        df_final = pd.DataFrame(filas_combinadas)
        
        # Remove rows with missing concept and filter by specific conditions
        df_final.dropna(subset=['Concepto'], inplace=True)
        df_final['Concepto'] = df_final['Concepto'].str.strip()
        df_final = df_final[df_final['Concepto'].isin(['SALIDA POR FACTURA', 'ENTRADA POR ABONO EN FACTURA'])]
        df_final = df_final[df_final['Descripción'].str.contains('ABV', case=False, na=False)]
        df_final = df_final[df_final['Referencia'].str.lower().str.startswith('e', na=False)]
        
        # Calculate LITRES and VALOR columns
        df_final['LITRES'] = (df_final.apply(lambda row: get_litres_referencia(row['Referencia'], row['Concepto']), axis=1) * df_final['Cantidad']).abs()
        df_final['VALOR'] = (df_final['Cantidad'] * df_final['Precio']).abs()
        
        # Export to Excel
        df_final.to_excel(salida, index=False)
        
        # Add column filters and freeze header using openpyxl
        wb = load_workbook(salida)
        ws = wb.active
        
        # Freeze top row
        ws.freeze_panes = 'A2'
        
        # Apply auto-filter to all columns
        ws.auto_filter.ref = ws.dimensions
        
        # Save the modified workbook
        wb.save(salida)
        
        print(f"Procesamiento completado. Salida guardada en: {salida}")
    
    except Exception as e:
        print(f"Error durante el procesamiento: {e}")

# Main execution block
if __name__ == '__main__':
    # Define default filenames
    entrada_nombre = 'entrada'
    salida_nombre = 'salida'
    
    # Determine file extension based on available files
    if os.path.exists(f"{entrada_nombre}.xlsx"):
        entrada = f"{entrada_nombre}.xlsx"
        salida = f"{salida_nombre}.xlsx"
    elif os.path.exists(f"{entrada_nombre}.xls"):
        entrada = f"{entrada_nombre}.xls"
        salida = f"{salida_nombre}.xls"
    else:
        print("No se encontró el archivo de entrada (entrada.xls o entrada.xlsx).")
        sys.exit(1)
    
    # Execute the processing function
    procesar_xls(entrada, salida)