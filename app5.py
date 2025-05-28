import pandas as pd
import openpyxl
import os
from datetime import datetime

# Detalles de Articulos Activos - Rediseñado

def get_column_letter(col_num):
    """Convierte número de columna a letra de Excel"""
    letters = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL']
    return letters[col_num - 1] if col_num <= len(letters) else f"Col{col_num}"

def clean_value(value, data_type='string'):
    """Limpia y convierte valores según el tipo de dato especificado"""
    if pd.isna(value) or value == '' or str(value).strip() == '{Sin Definir}' or str(value).strip() == '':
        return "Dato no Definido"
    
    if data_type == 'integer':
        try:
            if isinstance(value, str):
                value = value.replace(',', '').replace(' ', '')
            return int(float(value))
        except (ValueError, TypeError):
            return "Dato no Definido"
    elif data_type == 'float':
        try:
            if isinstance(value, str):
                value = value.replace(',', '').replace(' ', '')
            return float(value)
        except (ValueError, TypeError):
            return "Dato no Definido"
    else:  # string
        return str(value).strip() if str(value).strip() != '' else "Dato no Definido"

def process_file(file_path):
    """Procesa el archivo de inventario y extrae la información de proveedores y artículos"""
    
    try:
        # Leer el archivo Excel
        workbook = openpyxl.load_workbook(file_path, data_only=True)
        sheet = workbook.active
        
        # Lista para almacenar todos los datos procesados
        processed_data = []
        
        # Recorrer todas las filas para buscar "Proveedor:"
        max_row = sheet.max_row
        current_row = 1
        proveedores_encontrados = 0
        
        print(f"Iniciando búsqueda en {max_row} filas...")
        
        while current_row <= max_row:
            # Buscar "Proveedor:" en la columna B
            cell_b = sheet[f'B{current_row}']
            
            if cell_b.value and 'Proveedor:' in str(cell_b.value):
                proveedores_encontrados += 1
                print(f"Encontrado proveedor #{proveedores_encontrados} en fila {current_row}")
                
                # Leer información del proveedor de la misma fila
                # Columna F: ID del proveedor (Integer)
                id_proveedor = clean_value(sheet[f'F{current_row}'].value, 'integer')
                # Columna M: Nombre del proveedor (String)
                nombre_proveedor = clean_value(sheet[f'M{current_row}'].value, 'string')
                
                print(f"ID Proveedor: {id_proveedor}, Nombre: {nombre_proveedor}")
                
                # Pasar a la siguiente fila para empezar a leer artículos
                current_row += 1
                
                # Leer artículos hasta encontrar otro proveedor o fin de archivo
                while current_row <= max_row:
                    # Verificar si en la columna B hay otro "Proveedor:" (fin de artículos de este proveedor)
                    cell_b_next = sheet[f'B{current_row}']
                    if cell_b_next.value and 'Proveedor:' in str(cell_b_next.value):
                        # Encontramos otro proveedor, salir del bucle de artículos
                        break
                    
                    # Leer información del artículo según las especificaciones:
                    # Columna B: ID del Articulo (String)
                    id_articulo = clean_value(sheet[f'B{current_row}'].value, 'string')
                    # Columna I: Nombre del articulo (string)
                    nombre_articulo = clean_value(sheet[f'I{current_row}'].value, 'string')
                    # Columna S: Stock Minimo (float)
                    stock_minimo = clean_value(sheet[f'S{current_row}'].value, 'float')
                    # Columna V: Estado del Producto (String)
                    estado_producto = clean_value(sheet[f'V{current_row}'].value, 'string')
                    # Columna Z: Importado (string)
                    importado = clean_value(sheet[f'Z{current_row}'].value, 'string')
                    # Columna AC: Codigo para proveedor (string)
                    codigo_proveedor = clean_value(sheet[f'AC{current_row}'].value, 'string')
                    
                    # Debug: mostrar valores de la fila actual
                    print(f"Fila {current_row}: ID='{id_articulo}', Nombre='{nombre_articulo}', Stock='{stock_minimo}'")
                    
                    # Solo agregar si hay información válida del artículo
                    # Verificar que al menos tenga ID o nombre del artículo
                    if (id_articulo != "Dato no Definido" or 
                        nombre_articulo != "Dato no Definido"):
                        
                        print(f"Artículo válido encontrado - ID: {id_articulo}, Nombre: {nombre_articulo}")
                        
                        articulo_data = {
                            'ID Proveedor': id_proveedor,
                            'Nombre Proveedor': nombre_proveedor,
                            'ID Articulo': id_articulo,
                            'Nombre Articulo': nombre_articulo,
                            'Stock Minimo': stock_minimo,
                            'Estado del Producto': estado_producto,
                            'Importado': importado,
                            'Codigo para Proveedor': codigo_proveedor
                        }
                        processed_data.append(articulo_data)
                    
                    current_row += 1
            else:
                current_row += 1
        
        if not processed_data:
            raise ValueError("No se encontraron proveedores con el formato esperado en el archivo")
        
        print(f"\nResumen del procesamiento:")
        print(f"- Proveedores encontrados: {proveedores_encontrados}")
        print(f"- Total de artículos procesados: {len(processed_data)}")
        
        # Crear DataFrame con los datos procesados
        df = pd.DataFrame(processed_data)
        
        # Reordenar columnas según especificación
        column_order = [
            'ID Proveedor',
            'Nombre Proveedor', 
            'ID Articulo',
            'Nombre Articulo',
            'Stock Minimo',
            'Estado del Producto',
            'Importado',
            'Codigo para Proveedor'
        ]
        df = df[column_order]
        
        # Generar nombre de archivo de salida
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"inventario_procesado_{timestamp}.xlsx"
        output_path = os.path.join(os.path.dirname(file_path), output_filename)
        
        # Guardar archivo Excel
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Inventario Procesado', index=False)
            
            # Obtener la hoja y aplicar formato
            worksheet = writer.sheets['Inventario Procesado']
            
            # Ajustar ancho de columnas
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
        
        print(f"Archivo procesado guardado en: {output_path}")
        print(f"Total de artículos procesados: {len(df)}")
        
        # Mostrar preview de los primeros registros
        print("\nPreview de los primeros 3 registros:")
        print(df.head(3).to_string())
        
        return output_path
        
    except Exception as e:
        print(f"Error procesando archivo: {str(e)}")
        raise Exception(f"Error al procesar el archivo de inventario: {str(e)}")

if __name__ == "__main__":
    # Para pruebas locales
    test_file = "test_inventario.xlsx"
    if os.path.exists(test_file):
        result = process_file(test_file)
        print(f"Archivo procesado: {result}")
    else:
        print("Archivo de prueba no encontrado")