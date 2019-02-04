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
actividad = []
simposio = []

# STRINGS RELEVANTES

univ = "Afil" # para enviar a lista "institucion"
pon = "Ponente" # para enviar a lista "nombres"
tit = "ulo ponen" # para enviar a la lista "actividad"


for archivo in os.listdir('Simposios 2016'): # recorre todos los archivos
	if archivo.endswith('xlsx'): # si es un excel
		workbook = read.open_workbook(archivo) # abre el archivo
		nombres_hojas = workbook.sheet_names() # guarda los nombres de las hojas en una lista
		for i in nombres_hojas: # recorre la lista de hojas del archivo
			if i == 'Sheet1': # si la hoja se llama Sheet1
				hoja = workbook.sheet_by_name(i) # abre la hoja
				fila2 = hoja.row_values(1) # escoge la segunda fila de la hoja

				# PONENTES

				fila1 = hoja.row_values(0)

				aux = []

				for i in fila1:
					if pon in i:
						if fila2[fila1.index(i)] != '':
							nombres.append(fila2[fila1.index(i)])
							aux.append(fila2[fila1.index(i)])
					if univ in i:
						if fila2[fila1.index(i)] != '':
							institucion.append(fila2[fila1.index(i)])
					if tit in i:
						if fila2[fila1.index(i)] != '':
							actividad.append(fila2[fila1.index(i)])

				for i in range(len(aux)):
					simposio.append(fila2[0])

nom = pd.DataFrame({'1. Nombres': nombres, '2. Institución': institucion, '3. Título de la actividad': actividad, 'Temática': simposio})

doc1 = ExcelWriter('DATOS_SIMPOSIOS_2016.xlsx')
nom.to_excel(doc1, 'Hoja1', index=False)
doc1.save()
