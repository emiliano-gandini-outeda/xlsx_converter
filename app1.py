import os
import pandas as pd
import tempfile
from datetime import datetime
import re


def process_file(filepath):
    try:
        df = pd.read_excel(filepath, header=None)
        data = []

        current_fecha = None
        current_cliente_id = None
        current_cliente_nombre = None
        current_tipo_doc = None
        current_serie = None
        current_nro_doc = None

        for i, row in df.iterrows():
            row_values = [str(val).strip() if pd.notna(val) else "" for val in row.values]

            if re.match(r'\d{1,2}[/-]\d{1,2}[/-]\d{2,4}', row_values[1]) or pd.to_datetime(row_values[1], errors='coerce') is not pd.NaT:
                current_fecha = row_values[1]
                continue

            if row_values[0] and any(doc_type in row_values[1] for doc_type in ["Nota Cred.", "Vta.Cred."]):
                cliente_info = row_values[0].strip()
                parts = cliente_info.split(" ", 1)
                if len(parts) == 2:
                    current_cliente_id = parts[0]
                    current_cliente_nombre = parts[1]
                else:
                    current_cliente_id = cliente_info
                    current_cliente_nombre = ""
                current_tipo_doc = row_values[1]
                current_serie = row_values[2]
                current_nro_doc = row_values[3]
                continue

            if row_values[0] and row_values[1] and row_values[2] and current_cliente_id is not None:
                articulo_id = row_values[0]
                articulo_nombre = row_values[1]
                cantidad = row_values[2]
                precio_unitario = row_values[3]
                subtotal_iva = row_values[8] if len(row_values) > 8 else ""

                cliente_completo = f"{current_cliente_id} {current_cliente_nombre}".strip()

                data.append([
                    cliente_completo,
                    current_cliente_id,
                    current_cliente_nombre,
                    current_tipo_doc,
                    current_serie,
                    current_nro_doc,
                    articulo_id,
                    articulo_nombre,
                    cantidad,
                    precio_unitario,
                    subtotal_iva,
                    current_fecha
                ])

        columnas = [
            "Cliente Completo", "ID Cliente", "Cliente", "Tipo de Doc",
            "Serie", "Nro de Doc", "ID Articulo", "Articulo",
            "Cantidad", "Precio Unitario", "Subtotal con IVA", "Fecha"
        ]

        result_df = pd.DataFrame(data, columns=columnas)

        for col in ["Cantidad", "Precio Unitario", "Subtotal con IVA"]:
            result_df[col] = pd.to_numeric(result_df[col], errors='coerce')

        result_df["Fecha"] = pd.to_datetime(result_df["Fecha"], errors='coerce')
        result_df["Fecha"] = result_df["Fecha"].dt.strftime("%d/%m/%Y")
        result_df["Mes-Año"] = pd.to_datetime(result_df["Fecha"], dayfirst=True, errors='coerce').dt.strftime("%-m-%y")

        original_name = os.path.splitext(os.path.basename(filepath))[0]
        output_filename = f"{original_name}_Procesado.xlsx"
        temp_dir = tempfile.gettempdir()
        output_path = os.path.join(temp_dir, output_filename)

        result_df.to_excel(output_path, index=False)

        return output_path

    except Exception as e:
        raise RuntimeError(f"Error procesando el archivo de ventas: {str(e)}")


if __name__ == "__main__":
    print("Este módulo está diseñado para ser importado desde la aplicación principal")
