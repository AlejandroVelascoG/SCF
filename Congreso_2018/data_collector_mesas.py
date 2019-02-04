# -*- coding: utf-8 -*-

import xlrd as read
import pandas as pd

nombres = []
titulo = []
afiliacion = []
lista_celdas = []

workbook = read.open_workbook('Copia de Programación Mesas Temáticas FINAL.xlsx') # abre el archivo

lista_hojas = workbook.sheets() # lista de hojas en el libro

for i in range(len(lista_hojas)): # por cada hoja en el libro

    hoja = workbook.sheet_by_index(i) # selecciona la hoja

    for j in range(2, hoja.ncols): # recorre desde la tercera columna hasta la última

            columna = hoja.col_values(j,1) # crea una lista con lo que hay en la columna desde la segunda celda

            for k in range(hoja.nrows-1): # por cada celda de la columna

                celda = columna[k] # elige la celda
                #print(celda)
                str = '' # crea un string vacío
                contenido = [] # crea la lista que albergará los tres datos

                for c in celda: # por cada carácter en la celda
                    if(celda != 'Almuerzo' and celda != ''):
                        if c == celda[-1]: # si es el último carácter
                            contenido.append(str) # mete el string en la lista de contenido
                            lista_celdas.append(contenido) # mete la lista de contenido en la lista de todas las celdas
                            contenido = [] # vacía la lista de contenido
                            str = '' # vacía el string
                        if c != '\n': # si el carácter no es un linebreak
                            str += c # añade el carácter al string
                        elif c == '\n': # si el carácter es un linebreak
                            contenido.append(str) # añade el string a la lista de contenido
                            str = '' # vacía el string


print(lista_celdas)
# #lineas = celda.readlines()
# print(lista)
