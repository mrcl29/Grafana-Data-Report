#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# --------------------------------------------------------------------------------
#
#   Script para iniciar una API con Flask para poder ejecutar el script para
#   crear el informe haciendo una petición HTTP. La petición devuelve el archivo
#   excel creado del informe que se descarga en el dispositivo del usuario
#
#   Autor: Marc Llobera Villalonga
#
# --------------------------------------------------------------------------------

from flask import Flask, send_file
import os
import subprocess
import time
import config_flask

def buscar_archivo_en_subcarpetas(directorio, archivo):
    # Recorre todas las subcarpetas y archivos en el directorio
    for root, dirs, files in os.walk(directorio):  # Recorre todos los directorios y subdirectorios
        if archivo in files:  # Si el archivo se encuentra en el directorio actual
            return os.path.join(root, archivo)  # Retorna la ruta completa del archivo
    return None  # Si no se encuentra, retorna None

app = Flask(__name__)

@app.route('/grafana-data-report/<dashboard_id>')
def ejecutar_script(dashboard_id):
    informes_dir = config_flask.INFORMES_DIR
    informes_dict = config_flask.INFORMES_DICT

    print(f"Recibido dashboard_id: {dashboard_id}")
    
    name=""
    ruta_script = ""
    for k, v in informes_dict.items():
        if k == dashboard_id:
            name = k
            ruta_script = v
            print(f"Script encontrado: {ruta_script}")
    
    if ruta_script in [None, ""]:
        print("Dashboard no reconocido.")
        return "Dashboard no reconocido", 400
        
    # Ejecutar el script de forma asincrona
    print(f"Ejecutando script: {ruta_script}")
    process = subprocess.Popen(['python3', ruta_script], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    
    timeout = 300  # 5 minutos
    start_time = time.time()
    
    while time.time() - start_time < timeout:
        if process.poll() is not None:  # El proceso ha terminado
            print(f"El proceso ha terminado exitosamente.")
            break
        time.sleep(1)  # Esperar 1 segundo antes de verificar de nuevo
    
    # Verificar si el proceso termino correctamente
    if process.poll() is None:
        print("El script tardó demasiado en ejecutarse. Terminando proceso.")
        process.terminate()
        return "El script tardó demasiado en ejecutarse", 504
    
    # Obtener la salida del script
    stdout, stderr = process.communicate()
    print(f"Salida del script: {stdout}")
    print(f"Errores del script (si los hay): {stderr}")
    
    # Eliminar espacios en blanco y saltos de linea del nombre del archivo
    nombre_archivo = stdout.strip()
    print(f"Nombre de archivo generado: {nombre_archivo}")

    excel_path = buscar_archivo_en_subcarpetas(informes_dir, nombre_archivo)
    if excel_path is None:
        print(f"El archivo Excel no se ha generado correctamente: {nombre_archivo}")
        return f"El archivo Excel no se ha generado correctamente: {nombre_archivo}", 404
        
    # Enviar el archivo para descarga
    print(f"Enviando archivo {excel_path} para descarga.")
    return send_file(excel_path, 
                     as_attachment=True,
                     download_name=nombre_archivo,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    print("Iniciando servidor Flask...")
    app.run(host='0.0.0.0', port=5000)
