#!/usr/bin/env python3
# -*- coding: utf-8 -*-

#--------------------------------------------------------------------------------
#
#   Archivo donde se introduce la configuración pertinente y a partir del cual
#   se crea el informe. El informe se creará en el directorio establecido en
#   la variable 'INFORMES_DIR'.
#
#   Autor: Marc Llobera Villalonga
#
#--------------------------------------------------------------------------------

ACTIVAR_SELECCION_RANGO_DE_FECHAS = False
DEBUG_FINAL = False

# Informacion CT
TITULO = ""
INFORMES_DIR = "/[root]/Grafana-Data-Report/informes/"+TITULO+"/"

IMG = "/[root]/Grafana-Data-Report/scripts/assets/[sample].png"
DATA_DIR = "/[root]/Grafana-Data-Report/scripts/"+TITULO+"/data/"
# Informacion de autenticacion y URL base de Grafana
GRAFANA_URL = "http://[IP]:[PORT]"
API_KEY = ""

# Indica el rango de dias desde hoy para el que se quieren recoger datos de Grafana
DAYS = 7

# ¡¡¡IMPORTANTE MENOS DE 31 CARaCTERES CADA NOMBRE!!!
# CLAVE:[NOMBRE, TIPO(LINEAS, BARRAS), TAMAnO(PEQUEnO, MEDIANO, GRANDE), BINARIO(True, False), LEYENDA(True, False), EXTRA(MAXMIN, INFO, TABLA), EXTRA_INFO(mensaje, titulo tablas), EXTRA_INFO(unidad)]
# Informacion Dashboards
UIDS = ("")
DASHBOARDS = {("", ""): {0: ["","L","G",False,False,"",""]}}


