from flask import Flask, render_template, request, send_from_directory, redirect, url_for, flash
import os
import uuid
import shutil
from werkzeug.utils import secure_filename

# Importar procesadores
from app1 import process_file as process_ventas
from app4 import process_file as process_deudores
from app2 import process_file as process_balance

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
DOWNLOAD_FOLDER = 'downloads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER
app.secret_key = 'supersecretkey'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

PROCESSORS = {
    'ventas': {'module': 'app1', 'function': process_ventas},
    'deudores': {'module': 'app4', 'function': process_deudores},
    'balance-proyectado': {'module': 'app2', 'function': process_balance},
}

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/procesar', methods=['POST'])
def procesar():
    if 'file' not in request.files:
        flash('No se seleccionó ningún archivo', 'error')
        return redirect(url_for('index'))

    archivo = request.files['file']
    opcion = request.form.get('fileType')

    if archivo.filename == '' or not archivo:
        flash('Archivo inválido o vacío', 'error')
        return redirect(url_for('index'))

    if opcion not in PROCESSORS:
        flash('Opción no soportada actualmente', 'error')
        return redirect(url_for('index'))

    try:
        filename = secure_filename(archivo.filename)
        ruta_subida = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        archivo.save(ruta_subida)

        ruta_resultado = PROCESSORS[opcion]['function'](ruta_subida)

        nombre_descarga = os.path.basename(ruta_resultado)
        ruta_descarga = os.path.join(DOWNLOAD_FOLDER, nombre_descarga)
        shutil.move(ruta_resultado, ruta_descarga)

        return redirect(url_for('descargar_archivo', nombre_archivo=nombre_descarga))

    except Exception as e:
        flash(f"Error al procesar archivo: {str(e)}", 'error')
        return redirect(url_for('index'))

@app.route('/descargas/<nombre_archivo>')
def descargar_archivo(nombre_archivo):
    return send_from_directory(app.config['DOWNLOAD_FOLDER'], nombre_archivo, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
