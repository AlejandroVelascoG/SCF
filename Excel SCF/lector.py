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

nombres = []
apellidos = []
institucion = []
titulo = []
tema = []

for archivo in os.listdir('Mesas tematicas 2016'):

    w_archivo = read.open_workbook(archivo) # abre el archivo
    nombres_hojas = w_archivo.sheet_names() # crea lista de hojas
    hoja = w_archivo.sheet_by_name(nombres_hojas[0]) # escoge la primera hoja

    # crea las listas de nombres, apellidos, instituciones, titulos y temas

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

writer = ExcelWriter('ExcelSCF.xlsx')
columnas.to_excel(writer,'Hoja1',index=False)
writer.save()
