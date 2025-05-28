import pandas as pd
import openpyxl
import os
from datetime import datetime

# Detalles de Articulos Activos

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
            # Buscar "Proveedor:" en la columna D
            cell_d = sheet[f'D{current_row}']
            
            if cell_d.value and 'Proveedor:' in str(cell_d.value):
                proveedores_encontrados += 1
                print(f"Encontrado proveedor #{proveedores_encontrados} en fila {current_row}")
                
                # Leer información del proveedor de la misma fila
                id_proveedor = clean_value(sheet[f'I{current_row}'].value, 'integer')
                nombre_proveedor = clean_value(sheet[f'N{current_row}'].value, 'string')
                
                print(f"ID Proveedor: {id_proveedor}, Nombre: {nombre_proveedor}")
                
                # Pasar a la siguiente fila para empezar a leer artículos
                current_row += 1
                
                # Variables para totales del proveedor
                total_unidades_proveedor = "Dato no Definido"
                total_proveedor = "Dato no Definido"
                
                # Lista temporal para almacenar artículos de este proveedor
                articulos_proveedor = []
                
                # Leer artículos hasta encontrar los totales
                while current_row <= max_row:
                    # Verificar si en la columna Z hay un float (indicador de fin de artículos)
                    cell_z = sheet[f'Z{current_row}']
                    cell_ag = sheet[f'AG{current_row}']
                    
                    # Si encontramos números en Z y AG, son los totales del proveedor
                    if cell_z.value is not None and cell_ag.value is not None:
                        try:
                            # Verificar si ambos son números (totales)
                            float(cell_z.value)
                            float(cell_ag.value)
                            # Verificar que no haya nombre de artículo en columna H (para distinguir entre artículo y totales)
                            cell_h = sheet[f'H{current_row}']
                            if cell_h.value is None or str(cell_h.value).strip() == '' or str(cell_h.value).strip() == 'Dato no Definido':
                                total_unidades_proveedor = clean_value(cell_z.value, 'float')
                                total_proveedor = clean_value(cell_ag.value, 'float')
                                print(f"Totales encontrados - Unidades: {total_unidades_proveedor}, Total: {total_proveedor}")
                                current_row += 1
                                break
                        except (ValueError, TypeError):
                            # No son números, continuar leyendo artículos
                            pass
                    
                    # Leer información del artículo
                    # Revisar varias columnas para encontrar el ID del artículo
                    id_articulo = "Dato no Definido"
                    cell_b = sheet[f'B{current_row}']
                    
                    # Verificar si hay ID en la columna B
                    if cell_b.value is not None and str(cell_b.value).strip() != '':
                        id_articulo = clean_value(cell_b.value, 'string')
                    
                    # El nombre del artículo está en la columna H
                    nombre_articulo = clean_value(sheet[f'H{current_row}'].value, 'string')
                    
                    cantidad_articulo = clean_value(sheet[f'Y{current_row}'].value, 'float')
                    precio_articulo = clean_value(sheet[f'AB{current_row}'].value, 'float')
                    total_articulo = clean_value(sheet[f'AI{current_row}'].value, 'float')
                    
                    # Debug: mostrar valores de la fila actual
                    print(f"Fila {current_row}: B='{cell_b.value}', H='{sheet[f'H{current_row}'].value}', Y='{sheet[f'Y{current_row}'].value}'")
                    
                    # Solo agregar si hay información válida del artículo (nombre o algún otro dato)
                    if (nombre_articulo != "Dato no Definido" or 
                        cantidad_articulo != "Dato no Definido" or 
                        precio_articulo != "Dato no Definido" or 
                        total_articulo != "Dato no Definido"):
                        
                        print(f"Artículo encontrado - ID: {id_articulo}, Nombre: {nombre_articulo}")
                        
                        articulo_data = {
                            'ID Proveedor': id_proveedor,
                            'Proveedor': nombre_proveedor,
                            'ID Articulo': id_articulo,
                            'Nombre Articulo': nombre_articulo,
                            'Cantidad Articulo': cantidad_articulo,
                            'Precio por Articulo': precio_articulo,
                            'Total por Articulo': total_articulo,
                            'Total de unidades por proveedor': "Dato no Definido",  # Se actualizará después
                            'Total por proveedor': "Dato no Definido"  # Se actualizará después
                        }
                        articulos_proveedor.append(articulo_data)
                    
                    current_row += 1
                
                # Agregar totales a todos los artículos de este proveedor
                for articulo in articulos_proveedor:
                    articulo['Total de unidades por proveedor'] = total_unidades_proveedor
                    articulo['Total por proveedor'] = total_proveedor
                    processed_data.append(articulo)
            else:
                current_row += 1
        
        if not processed_data:
            raise ValueError("No se encontraron proveedores con el formato esperado en el archivo")
        
        print(f"\nResumen del procesamiento:")
        print(f"- Proveedores encontrados: {proveedores_encontrados}")
        print(f"- Total de artículos procesados: {len(processed_data)}")
        
        # Crear DataFrame con los datos procesados
        df = pd.DataFrame(processed_data)
        
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