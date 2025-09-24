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
import sys
import os
import subprocess
import time
import config_flask
import importlib

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
    # process = subprocess.Popen(['python3', ruta_script],stdout=subprocess.PIPE,stderr=subprocess.STDOUT,text=True)
    directorio_script = os.path.dirname(ruta_script)
    nombre_modulo = os.path.splitext(os.path.basename(ruta_script))[0]
    # Agregar el directorio del script al sys.path si no está ya
    if directorio_script not in sys.path:
        sys.path.insert(0, directorio_script)
    spec = importlib.util.spec_from_file_location(nombre_modulo, ruta_script)
    modulo = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modulo)
    file_name = modulo.main()
    # print(f"aaaaaaaa")
    # timeout = 300  # 5 minutos
    # start_time = time.time()
    
    # while time.time() - start_time < timeout:
    #     print(str(time.time()))
    #     print(str(process.poll()))
    #     if process.poll() is not None:  # El proceso ha terminado
    #         print(f"El proceso ha terminado exitosamente.")
    #         break
    #     time.sleep(1)  # Esperar 1 segundo antes de verificar de nuevo
    
    # # Verificar si el proceso termino correctamente
    # if process.poll() is None:
    #     print("El script tarda demasiado en ejecutarse. Terminando proceso.")
    #     process.terminate()
    #     return "El script tarda demasiado en ejecutarse", 504
    
    # # Obtener la salida del script
    # for line in process.stdout:
    #   print(line, end='')  # Muestra los logs en tiempo real
    # process.wait()

    # print(f"Salida del script: {stdout}")
    # print(f"Errores del script (si los hay): {stderr}")
    
    # Eliminar espacios en blanco y saltos de linea del nombre del archivo
    # nombre_archivo = stdout.strip()
    nombre_archivo = file_name
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
