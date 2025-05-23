import os
import pandas as pd
import tempfile
from datetime import datetime
import re

def process_file(filepath):
    try:
        ext = os.path.splitext(filepath)[1].lower()
        if ext == ".csv":
            df = pd.read_csv(filepath, header=None)
        else:
            df = pd.read_excel(filepath, header=None)
        
        output_rows = []
        
        # Empezar desde el inicio del documento para capturar TODOS los tipos de documento
        i = 0
        
        while i < len(df):
            # Buscar "Vta.Cred.", "Vta.Cont." o "Nota Cred." en la columna A desde el principio
            if i >= len(df):
                break
                
            # Obtener el valor de la celda A de manera más robusta
            try:
                celda_a_raw = df.iloc[i, 0]
                if pd.isna(celda_a_raw):
                    celda_a = ""
                else:
                    celda_a = str(celda_a_raw).strip()
            except:
                celda_a = ""
            
            # Debug más detallado
            if i < 20 or celda_a in ["Vta.Cred.", "Nota Cred.", "Vta.Cont."]:
                print(f"Línea {i+1}: '{celda_a}' (raw: '{celda_a_raw}')")
            
            # Verificar si es un tipo de documento válido
            if celda_a not in ["Vta.Cred.", "Nota Cred.","Vta.Cont."]:
                i += 1
                continue
            
            print(f"Encontrado tipo de documento '{celda_a}' en línea {i+1}")  # Debug
            
            # Leer información del documento (línea actual)
            tipo_documento = celda_a  # Ya sabemos que es "Vta.Cred." o "Nota Cred."
            serie_documento = safe_get_string(df, i, 9)  # Columna J (índice 9)
            numero_documento = safe_get_integer(df, i, 12)  # Columna M (índice 12)
            fecha_documento = safe_get_date(df, i, 21)  # Columna V (índice 21)
            cliente = safe_get_string(df, i, 34)  # Columna AI (índice 34)
            descuento_porcentaje = safe_get_float(df, i, 52)  # Columna BA (índice 52)
            descuento_pesos = safe_get_float(df, i, 60)  # Columna BI (índice 60)
            total_pesos = safe_get_float(df, i, 67)  # Columna BP (índice 67)
            
            print(f"Tipo: {tipo_documento}, Cliente: {cliente}, Total: {total_pesos}")  # Debug
            
            # Leer información CAE (línea siguiente)
            i += 1
            if i >= len(df):
                break
                
            cae_nro = safe_get_integer(df, i, 7)  # Columna H (índice 7)
            cae_venc = safe_get_date(df, i, 26)  # Columna AA (índice 26)
            cae_serie = safe_get_string(df, i, 44)  # Columna AS (índice 44)  
            cae_numero_documento = safe_get_integer(df, i, 49)  # Columna AX (índice 49) - CORREGIDO
            cae_estado = safe_get_string(df, i, 64)  # Columna BM (índice 64)
            
            print(f"CAE Nro: {cae_nro}, CAE Estado: {cae_estado}")  # Debug
            
            # Leer artículos (líneas siguientes)
            i += 1
            articulos_encontrados = 0
            while i < len(df):
                # Verificar si hay código de artículo en columna B
                codigo_articulo = safe_get_string(df, i, 1)  # Columna B (índice 1)
                
                print(f"Línea {i+1}: Código artículo = '{codigo_articulo}'")  # Debug
                
                # Si no hay código de artículo, verificar si es el inicio de un nuevo documento
                if not codigo_articulo or codigo_articulo.strip() == "":
                    # Verificar si la siguiente línea tiene un tipo de documento
                    if i < len(df):
                        try:
                            next_celda_a_raw = df.iloc[i, 0]
                            if pd.isna(next_celda_a_raw):
                                next_celda_a = ""
                            else:
                                next_celda_a = str(next_celda_a_raw).strip()
                            
                            # Si encontramos un nuevo tipo de documento, salir sin incrementar i
                            if next_celda_a in ["Vta.Cred.", "Nota Cred.","Vta.Cont."]:
                                print(f"Nuevo documento encontrado en línea {i+1}: '{next_celda_a}'")
                                break
                        except:
                            pass
                    
                    # Si no es un nuevo documento, continuar buscando
                    i += 1
                    continue
                
                # Leer datos del artículo
                articulo = safe_get_string(df, i, 17)  # Columna R (índice 17)
                cantidad_articulo = safe_get_float(df, i, 41)  # Columna AP (índice 41)
                precio_unitario = safe_get_float(df, i, 47)  # Columna AV (índice 47)
                
                print(f"Artículo: {articulo}, Cantidad: {cantidad_articulo}, Precio: {precio_unitario}")  # Debug
                
                # Crear fila de resultado (sin Numero de Cliente)
                fila = [
                    cliente,
                    tipo_documento,
                    serie_documento,
                    numero_documento,
                    fecha_documento,
                    cae_nro,
                    cae_serie,
                    cae_numero_documento,
                    cae_estado,
                    codigo_articulo,
                    articulo,
                    cantidad_articulo,
                    precio_unitario,
                    total_pesos,
                    descuento_porcentaje,
                    descuento_pesos
                ]
                output_rows.append(fila)
                articulos_encontrados += 1
                
                i += 1
            
            print(f"Artículos procesados: {articulos_encontrados}")  # Debug
            
            # Continuar buscando el siguiente "Tipo de documento"
            # No incrementar i aquí porque ya se incrementó en el bucle de artículos
        
        print(f"Total de filas procesadas: {len(output_rows)}")  # Debug
        
        # Crear DataFrame resultado (sin columna Numero de Cliente)
        columnas = [
            "Cliente",
            "Tipo de Documento",
            "Serie del Documento",
            "Numero del documento",
            "Fecha del documento",
            "CAE Nro",
            "CAE Serie", 
            "CAE Numero de documento",
            "CAE Estado",
            "Codigo Articulo",
            "Articulo",
            "Cantidad articulo",
            "Precio unitario",
            "Total en pesos",
            "Descuento en %",
            "Descuento en pesos"
        ]
        
        df_resultado = pd.DataFrame(output_rows, columns=columnas)
        
        # Guardar archivo
        original_name = os.path.splitext(os.path.basename(filepath))[0]
        output_filename = f"{original_name}_PROCESADO.xlsx"
        output_path = os.path.join(tempfile.gettempdir(), output_filename)
        
        df_resultado.to_excel(output_path, index=False)
        
        return output_path
        
    except Exception as e:
        raise RuntimeError(f"Error procesando archivo de facturación: {str(e)}")

def safe_get_string(df, row, col):
    """Obtener string de manera segura"""
    try:
        if row < len(df) and col < len(df.columns):
            value = df.iloc[row, col]
            if pd.notna(value):
                return str(value).strip()
        return ""
    except:
        return ""

def safe_get_integer(df, row, col):
    """Obtener integer de manera segura"""
    try:
        if row < len(df) and col < len(df.columns):
            value = df.iloc[row, col]
            if pd.notna(value):
                # Limpiar el valor si es string con formato
                if isinstance(value, str):
                    clean_value = value.strip().replace(".", "").replace(",", ".")
                    return int(float(clean_value))
                return int(value)
        return 0
    except:
        return 0

def safe_get_float(df, row, col):
    """Obtener float de manera segura"""
    try:
        if row < len(df) and col < len(df.columns):
            value = df.iloc[row, col]
            if pd.notna(value):
                # Limpiar el valor si es string con formato argentino
                if isinstance(value, str):
                    clean_value = value.strip().replace(".", "").replace(",", ".")
                    return float(clean_value)
                return float(value)
        return 0.0
    except:
        return 0.0

def safe_get_date(df, row, col):
    """Obtener fecha de manera segura en formato dia-mes-año"""
    try:
        if row < len(df) and col < len(df.columns):
            value = df.iloc[row, col]
            if pd.notna(value):
                if isinstance(value, str):
                    # Intentar parsear fecha en formato dia-mes-año
                    try:
                        date_obj = pd.to_datetime(value, dayfirst=True)
                        return date_obj.strftime("%d/%m/%Y")
                    except:
                        return value.strip()
                else:
                    # Si es datetime, convertir a string
                    try:
                        return pd.to_datetime(value).strftime("%d/%m/%Y")
                    except:
                        return str(value)
        return ""
    except:
        return ""

def extract_client_number(cliente_string):
    """Extraer número de cliente del string cliente"""
    try:
        if cliente_string:
            # Buscar números al inicio del string
            match = re.match(r'^(\d+)', cliente_string.strip())
            if match:
                return int(match.group(1))
        return 0
    except:
        return 0

if __name__ == "__main__":
    print("Este script está diseñado para ser importado desde app.py")