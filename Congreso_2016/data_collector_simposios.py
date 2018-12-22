# -*- coding: utf-8 -*-

# SCRIPT ARCHIVO PARA LEER TODOS LOS ARCHIVOS DE LOS SIMPOSIOS Y GUARDAR DATOS EN UNA SOLA Hoja1

import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import xlrd as read
import os

# LISTAS DE DATOS RELEVANTES

nombres = []
institucion = []
actividad = [] # si el participante es el coordinador del simposio, entra en esta lista
simposio = []

# STRINGS RELEVANTES

tit_sim = "Título del Simpsio" # para enviar a lista "simposio"
coor = "Coordinador" # para enviar a lista "actividad"
un_apoyo = "Universidad de apoyo" # para enviar a lista "institucion"
univ = "Afiliación" # para enviar a lista "institucion"
pon = "Ponente" # para enviar a lista "nombres"


# for archivo in os.listdir('Simposios 2016'):
# 	workbook = read.open_workbook(archivo)
# 	nombres_hojas = workbook.sheet_names()
#
# 	for i in nombres_hojas:
#         if i == 'Sheet1':
#             print(i.nrows())
#
# print(nombres)
