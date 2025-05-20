from flask import Flask, render_template, request, send_from_directory, redirect, url_for, flash
import os
import uuid
from werkzeug.utils import secure_filename
import importlib.util
import sys
import shutil


app = Flask(__name__)

# Configuraciones de carpetas
UPLOAD_FOLDER = 'uploads'
DOWNLOAD_FOLDER = 'downloads'
TEMPLATE_FOLDER = 'templates'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER
app.config['TEMPLATE_FOLDER'] = TEMPLATE_FOLDER
app.secret_key = 'supersecretkey'

# Crear carpetas si no existen
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)
os.makedirs(os.path.join(TEMPLATE_FOLDER), exist_ok=True)

# Mapeo de tipos de archivo a módulos de procesamiento
PROCESSORS = {
    'ventas': {'module': 'app1', 'function': 'process_file'},
    'deudores': {'module': 'app4', 'function': 'process_file'},
    # Estos son placeholders para futuros tipos de archivos
    'precios-pesos': {'module': 'app1', 'function': 'process_file'},
    'precios-dolares': {'module': 'app1', 'function': 'process_file'}
}

def load_processor(processor_info):
    """
    Carga dinámicamente el módulo y función de procesamiento
    
    Args:
        processor_info: Diccionario con información sobre el módulo y función
        
    Returns:
        function: La función de procesamiento
    """
    module_name = processor_info['module']
    function_name = processor_info['function']
    
    # Importar el módulo dinámicamente
    spec = importlib.util.find_spec(module_name)
    if spec is None:
        raise ImportError(f"No se pudo encontrar el módulo {module_name}")
    
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    
    # Obtener la función de procesamiento
    if not hasattr(module, function_name):
        raise AttributeError(f"El módulo {module_name} no tiene la función {function_name}")
    
    return getattr(module, function_name)

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

        # Cargar dinámicamente el procesador según la opción seleccionada
        processor_info = PROCESSORS[opcion]
        process_function = load_processor(processor_info)
        
        # Ejecutar la función de procesamiento
        ruta_resultado = process_function(ruta_subida)

        # Generar nombre único para el archivo procesado
        nombre_descarga = os.path.basename(ruta_resultado)
        ruta_descarga = os.path.join(app.config['DOWNLOAD_FOLDER'], nombre_descarga)
        
        # Copiar o mover el archivo procesado a la carpeta de descargas
        shutil.move(ruta_resultado, ruta_descarga)

        # Redirigir a la página de descarga
        return send_from_directory(app.config['DOWNLOAD_FOLDER'], nombre_descarga, as_attachment=True)

    except Exception as e:
        flash(f"Error al procesar archivo: {str(e)}", 'error')
        return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)