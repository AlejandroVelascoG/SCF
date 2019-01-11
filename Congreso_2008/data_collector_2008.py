# -*- coding: utf-8 -*-

import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import xlrd as read
import os


nombres = []
apellidos = []
institucion = []
region = []
genero = []
titulo = []
tema = []


workbook = read.open_workbook('Folleto-Listado Ponentes 2 CCF-2008_tema.xls')
nombres_hojas = workbook.sheet_names()

hoja = workbook.sheet_by_name('Hoja1')

new_nombres = hoja.col_values(0, 1)
new_apellidos = hoja.col_values(1, 1)
new_titulo = hoja.col_values(6, 1)
new_tema = hoja.col_values(5, 1)
new_genero = hoja.col_values(2, 1)
new_institucion = hoja.col_values(3, 1)
new_region = hoja.col_values(4, 1)





nombres.extend(new_nombres)
apellidos.extend(new_apellidos)
institucion.extend(new_institucion)
titulo.extend(new_titulo)
tema.extend(new_tema)
region.extend(new_region)
genero.extend(new_genero)

columnas = pd.DataFrame({'1. Nombres': nombres, '2. Apellidos': apellidos, '3. Institución': institucion,
                    '4. Título de la actividad': titulo, '5. Temática': tema, '6. Region': region, '7. Género': genero})

writer = ExcelWriter('Datos_2008.xlsx')
columnas.to_excel(writer,'Hoja1',index=False)
writer.save()
