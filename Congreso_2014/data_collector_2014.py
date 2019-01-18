# -*- coding: utf-8 -*-

import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import xlrd as read
import os


nombres = []
institucion = []
titulo = []
tema = []

workbook = read.open_workbook('Programación General V CCF Medellín 2014 mayo 22.xlsx')

plenarias = workbook.sheet_by_index(0)

p_nombres = plenarias.col_values(0, 1)
p_titulo = plenarias.col_values(2, 1)
p_tema = ['PLENARIA']*len(p_nombres)
p_institucion = ['X']*len(p_nombres)


simposios = workbook.sheet_by_index(1)

s_nombres = simposios.col_values(2, 1)
s_institucion =  simposios.col_values(4, 1)
s_titulo = simposios.col_values(5, 1)
s_tema = simposios.col_values(6, 1)


sesiones = workbook.sheet_by_index(2)

t_nombres = sesiones.col_values(0, 1)
t_institucion =  sesiones.col_values(2, 1)
t_titulo = sesiones.col_values(3, 1)
t_tema = sesiones.col_values(4, 1)


libros = workbook.sheet_by_index(3)

l_nombres = libros.col_values(0, 1)
l_institucion = libros.col_values(2, 1)
l_titulo = libros.col_values(1, 1)
l_tema = ['PRESENTACIÓN DE LIBRO']*len(l_nombres)

nombres.extend(p_nombres)
nombres.extend(s_nombres)
nombres.extend(t_nombres)
nombres.extend(l_nombres)

institucion.extend(p_institucion)
institucion.extend(s_institucion)
institucion.extend(t_institucion)
institucion.extend(l_institucion)

titulo.extend(p_titulo)
titulo.extend(s_titulo)
titulo.extend(t_titulo)
titulo.extend(l_titulo)

tema.extend(p_tema)
tema.extend(s_tema)
tema.extend(t_tema)
tema.extend(l_tema)

columnas = pd.DataFrame({'1. Nombres': nombres, '2. Institución': institucion,'3. Título de la actividad': titulo, '4. Temática': tema})

writer = ExcelWriter('Datos_filtrados_2014.xlsx')
columnas.to_excel(writer,'Hoja1',index=False)
writer.save()
