import pandas as pd
import openpyxl
import os
from datetime import datetime

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
        
        while current_row <= max_row:
            # Buscar "Proveedor:" en la columna D
            cell_d = sheet[f'D{current_row}']
            
            if cell_d.value and 'Proveedor:' in str(cell_d.value):
                print(f"Encontrado proveedor en fila {current_row}")
                
                # Leer información del proveedor de la misma fila
                id_proveedor = clean_value(sheet[f'I{current_row}'].value, 'integer')
                nombre_proveedor = clean_value(sheet[f'N{current_row}'].value, 'string')
                
                print(f"ID Proveedor: {id_proveedor}, Nombre: {nombre_proveedor}")
                
                # Pasar a la siguiente fila para empezar a leer artículos
                current_row += 1
                
                # Variables para totales del proveedor
                total_unidades_proveedor = "Dato no Definido"
                total_proveedor = "Dato no Definido"
                
                # Leer artículos hasta encontrar los totales
                while current_row <= max_row:
                    # Verificar si en la columna Z hay un float (indicador de fin de artículos)
                    cell_z = sheet[f'Z{current_row}']
                    
                    # Si encontramos un número en Z, son los totales del proveedor
                    if cell_z.value is not None:
                        try:
                            # Es un número, por lo tanto son los totales
                            total_unidades_proveedor = clean_value(cell_z.value, 'float')
                            total_proveedor = clean_value(sheet[f'AG{current_row}'].value, 'float')
                            print(f"Totales encontrados - Unidades: {total_unidades_proveedor}, Total: {total_proveedor}")
                            current_row += 1
                            break
                        except:
                            pass
                    
                    # Si no hay totales, leer información del artículo
                    id_articulo = clean_value(sheet[f'B{current_row}'].value, 'integer')
                    nombre_articulo = clean_value(sheet[f'I{current_row}'].value, 'string')
                    cantidad_articulo = clean_value(sheet[f'Y{current_row}'].value, 'float')
                    precio_articulo = clean_value(sheet[f'AB{current_row}'].value, 'float')
                    total_articulo = clean_value(sheet[f'AI{current_row}'].value, 'float')
                    
                    # Solo agregar si hay información válida del artículo (al menos ID o nombre)
                    if (id_articulo != "Dato no Definido" or nombre_articulo != "Dato no Definido"):
                        print(f"Artículo - ID: {id_articulo}, Nombre: {nombre_articulo}")
                        
                        processed_data.append({
                            'ID Proveedor': id_proveedor,
                            'Proveedor': nombre_proveedor,
                            'ID Articulo': id_articulo,
                            'Nombre Articulo': nombre_articulo,
                            'Cantidad Articulo': cantidad_articulo,
                            'Precio por Articulo': precio_articulo,
                            'Total por Articulo': total_articulo,
                            'Total de unidades por proveedor': total_unidades_proveedor,
                            'Total por proveedor': total_proveedor
                        })
                    
                    current_row += 1
                
                # Actualizar todos los artículos de este proveedor con los totales
                for item in processed_data:
                    if (item['ID Proveedor'] == id_proveedor and 
                        item['Total de unidades por proveedor'] == "Dato no Definido"):
                        item['Total de unidades por proveedor'] = total_unidades_proveedor
                        item['Total por proveedor'] = total_proveedor
            else:
                current_row += 1
        
        if not processed_data:
            raise ValueError("No se encontraron proveedores con el formato esperado en el archivo")
        
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