import os
import pandas as pd
import tempfile
from datetime import datetime
import re
from dateutil.relativedelta import relativedelta

# Balance Resumido

def process_file(filepath):
    try:
        ext = os.path.splitext(filepath)[1].lower()
        if ext == ".csv":
            df = pd.read_csv(filepath, header=None)
        else:
            df = pd.read_excel(filepath, header=None)
        
        output_rows = []
        
        # Leer moneda y fecha base
        moneda = str(df.iloc[9, 5]).strip() if pd.notna(df.iloc[9, 5]) else ""
        fecha_base_raw = str(df.iloc[6, 10]).strip() if pd.notna(df.iloc[6, 10]) else ""
        
        try:
            fecha_base = pd.to_datetime(fecha_base_raw, dayfirst=True)
        except:
            fecha_base = datetime.now()
        
        fechas = [(fecha_base + relativedelta(months=i)).strftime("%d/%m/%Y") for i in range(6)]
        
        # Columnas específicas para las deudas: J=9, L=11, O=14, R=17, U=20, Y=24 (índices base 0)
        columnas_deudas = [9, 11, 14, 17, 20, 24]
        
        i = 10
        while i < len(df):
            # Verificar si la celda está vacía
            if pd.isna(df.iloc[i, 0]):
                i += 1
                continue
            
            cliente_cell = str(df.iloc[i, 0]).strip()
            
            # Verificar si empieza con números (cliente válido)
            if not re.match(r'^\d+', cliente_cell):
                i += 1
                continue
            
            # Extraer ID y nombre del cliente
            parts = cliente_cell.split(" ", 1)
            cliente_id = parts[0]
            cliente_nombre = parts[1] if len(parts) > 1 else ""
            
            # Leer las 6 deudas desde las columnas específicas
            deudas = []
            for col_index in columnas_deudas:
                if col_index < len(df.columns):
                    valor_celda = df.iloc[i, col_index]
                    
                    if pd.isna(valor_celda):
                        deudas.append(0)
                    else:
                        valor_str = str(valor_celda).strip()
                        
                        if valor_str == "-":
                            deudas.append(0)
                        elif re.match(r'^\d{1,3}(\.\d{3})*,\d{2}$', valor_str):
                            try:
                                monto = float(valor_str.replace(".", "").replace(",", "."))
                                deudas.append(monto)
                            except:
                                deudas.append(0)
                        else:
                            try:
                                monto = float(valor_str)
                                deudas.append(monto)
                            except:
                                deudas.append(0)
                else:
                    deudas.append(0)
            
            # Asegurar que tenemos exactamente 6 deudas
            while len(deudas) < 6:
                deudas.append(0)
            deudas = deudas[:6]  # Tomar solo las primeras 6
            
            # Leer saldo final (columna AD = índice 29)
            saldo_raw = df.iloc[i, 29] if 29 < len(df.columns) else None
            if pd.notna(saldo_raw):
                saldo_str = str(saldo_raw).strip()
                if re.match(r'^\d{1,3}(\.\d{3})*,\d{2}$', saldo_str):
                    saldo_final = float(saldo_str.replace(".", "").replace(",", "."))
                else:
                    try:
                        saldo_final = float(saldo_str)
                    except:
                        saldo_final = 0
            else:
                saldo_final = 0
            
            # Calcular suma y observación
            suma_deudas = sum(deudas)
            observacion = "OK"
            if abs(suma_deudas - saldo_final) > 0.1:
                observacion = f"⚠ Diferencia: {round(suma_deudas - saldo_final, 2)}"
            
            # Crear fila de resultado
            fila = [
                cliente_id,
                cliente_nombre,
                moneda,
                *deudas,
                saldo_final,
                observacion
            ]
            output_rows.append(fila)
            
            # Avanzar 3 filas (como en el código original)
            i += 3
        
        # Crear DataFrame resultado
        columnas = [
            "ID Cliente", "Nombre Cliente", "Moneda",
            *[f"Deuda al {fecha}" for fecha in fechas],
            "Saldo Final",
            "Observación"
        ]
        
        df_resultado = pd.DataFrame(output_rows, columns=columnas)
        
        # Guardar archivo
        original_name = os.path.splitext(os.path.basename(filepath))[0]
        output_filename = f"{original_name}_PROCESADO.xlsx"
        output_path = os.path.join(tempfile.gettempdir(), output_filename)
        
        df_resultado.to_excel(output_path, index=False)
        
        return output_path
        
    except Exception as e:
        raise RuntimeError(f"Error procesando archivo balance proyectado: {str(e)}")

if __name__ == "__main__":
    print("Este script está diseñado para ser importado desde app.py")