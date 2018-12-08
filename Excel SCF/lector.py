# -*- coding: utf-8 -*-

# ARCHIVO PARA LEER LOS ARCHIVOS Y COPIARLOS EN NUEVA HOJA

# import os

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
	#nombres_hojas.extend(workbook.sheet_names())
	print(nombres_hojas)
	hoja = workbook.sheet_by_name(nombres_hojas[1]) # escoge la hoja relevante
	
	# crea las listas de nombres, apellidos, instituciones, titulos y temas
	
	# new_nombres = hoja.col_values(0, 1)
	# new_apellidos = hoja.col_values(1, 1)
	# new_institucion = hoja.col_values(4, 1)
	# new_titulo = hoja.col_values(8, 1)
	# new_tema = hoja.col_values(13, 1)

	# nombres.extend(new_nombres)
	# apellidos.extend(new_apellidos)
	# institucion.extend(new_institucion)
	# titulo.extend(new_titulo)
	# tema.extend(new_tema)

# columnas = pd.DataFrame({'1. Nombres': nombres, '2. Apellidos': apellidos, '3. Institución': institucion,
#                         '4. Título de la actividad': titulo, '5. Temática': tema})

# writer = ExcelWriter('ExcelSCF.xlsx')
# columnas.to_excel(writer,'Hoja1',index=False)
# writer.save()
