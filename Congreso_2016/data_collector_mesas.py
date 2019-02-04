# -*- coding: utf-8 -*-

# SCRIPT PARA LEER LOS ARCHIVOS Y COPIAR TODOS LOS DATOS DE LAS MESAS TEMÁTICAS EN UNA SOLA HOJA

# diract = os.getcwd()

# print(diract)

# print(os.listdir('.'))

import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import xlrd as read
import os

# archivo = read.open_workbook('01_logica.xlsx')
#
# hoja = archivo.sheet_by_name('logica')

#nombres_hojas = []
nombres = []
apellidos = []
institucion = []
titulo = []
tema = []

#print(os.listdir('Mesas tematicas 2016'))

for archivo in os.listdir('Mesas tematicas 2016'):
	workbook = read.open_workbook(archivo)
	nombres_hojas = workbook.sheet_names()

	for i in nombres_hojas:
		if i != 'Datos':
			hoja = workbook.sheet_by_name(i)
			new_nombres = hoja.col_values(0, 1)
			new_apellidos = hoja.col_values(1, 1)
			new_institucion = hoja.col_values(6, 1)
			new_titulo = hoja.col_values(8, 1)
			new_tema = hoja.col_values(13, 1)

			nombres.extend(new_nombres)
			apellidos.extend(new_apellidos)
			institucion.extend(new_institucion)
			titulo.extend(new_titulo)
			tema.extend(new_tema)

columnas = pd.DataFrame({'1. Nombres': nombres, '2. Apellidos': apellidos, '3. Institución': institucion,
                        '4. Título de la actividad': titulo, '5. Temática': tema})

# columnas = pd.DataFrame({'1. Nombres': nombres, '2. Institución': institucion,
#                         '3. Título de la actividad': titulo, '4. Temática': tema})

writer = ExcelWriter('DATOS_MESAS_2016.xlsx')
columnas.to_excel(writer,'Hoja1',index=False)
writer.save()
