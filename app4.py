import os
import pandas as pd
import tempfile
from datetime import datetime

def process_file(filepath):
    try:
        ext = os.path.splitext(filepath)[1].lower()
        engine = 'xlrd' if ext == '.xls' else 'openpyxl'
        raw_df = pd.read_excel(filepath, header=None, engine=engine)
    except Exception as e:
        raise RuntimeError(f"Error reading Excel file: {e}")

    data = []
    current_client_id = None
    current_client_name = None

    for i, row in raw_df.iterrows():
        col_B = str(row[1]).strip().lower()
        col_G = str(row[6]).strip()

        if col_B == "cliente:":
            cliente_str = col_G.strip()
            if " " in cliente_str:
                current_client_id, current_client_name = cliente_str.split(" ", 1)
            else:
                current_client_id = cliente_str
                current_client_name = ""
        elif pd.notna(row[0]) and ("/" in str(row[0]) or "-" in str(row[0])):
            try:
                fecha_str = str(row[0]).strip()
                try:
                    fecha = pd.to_datetime(fecha_str, format="%d/%m/%Y %H:%M:%S", dayfirst=True)
                except:
                    try:
                        fecha = pd.to_datetime(fecha_str, format="%d/%m/%Y", dayfirst=True)
                    except:
                        fecha = pd.to_datetime(fecha_str, dayfirst=True)

                deposito = row[3]
                documento = row[11]
                serie = row[14]
                nro = int(float(row[17])) if pd.notna(row[17]) else None

                vencimiento = None
                if pd.notna(row[21]):
                    venc_str = str(row[21]).strip()
                    try:
                        vencimiento = pd.to_datetime(venc_str, format="%d/%m/%Y %H:%M:%S", dayfirst=True)
                    except:
                        try:
                            vencimiento = pd.to_datetime(venc_str, format="%d/%m/%Y", dayfirst=True)
                        except:
                            vencimiento = pd.to_datetime(venc_str, dayfirst=True)

                debe = str(row[26]).strip() if pd.notna(row[26]) and str(row[26]).strip() != "" else None
                haber = str(row[30]).strip() if pd.notna(row[30]) and str(row[30]).strip() != "" else None
                saldo = str(row[37]).strip() if pd.notna(row[37]) and str(row[37]).strip() != "" else None

                fecha_dmy = fecha.strftime('%-d-%-m-%y') if fecha.day < 10 else fecha.strftime('%d-%m-%y')
                try:
                    mes_ano = fecha.strftime('%B-%y').lower().replace("é", "e")
                except:
                    mes_nombres = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 
                                   'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre']
                    mes_nombre = mes_nombres[fecha.month-1]
                    mes_ano = f"{mes_nombre}-{str(fecha.year)[2:]}"

                fecha_formateada = fecha.strftime("%d/%m/%Y")
                venc_formateada = vencimiento.strftime("%d/%m/%Y") if vencimiento else None

                data.append([
                    fecha_formateada, deposito,
                    current_client_name, documento, serie, nro,
                    venc_formateada,
                    debe, haber, saldo,
                    current_client_id, fecha_dmy, mes_ano
                ])
            except:
                continue

    columns = [
        "Fecha", "Deposito", "Cliente:", "Documento", "SERIE", "Nro.", "Vencimiento",
        "Debe", "Haber", "Saldo", "ID CLIENTE", "Fecha d-m-a", "mes-año"
    ]
    output_df = pd.DataFrame(data, columns=columns)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = f"deudores_procesado_{timestamp}.xlsx"
    temp_dir = tempfile.gettempdir()
    output_path = os.path.join(temp_dir, output_filename)

    try:
        output_df.to_excel(output_path, index=False, engine='openpyxl')
    except Exception as e:
        csv_path = output_path.replace(".xlsx", ".csv")
        output_df.to_csv(csv_path, index=False, encoding="utf-8-sig")
        output_path = csv_path

    return output_path

if __name__ == "__main__":
    print("Este módulo está diseñado para ser importado desde la aplicación principal")
