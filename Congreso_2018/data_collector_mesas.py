# -*- coding: utf-8 -*-

import xlrd as read
import pandas as pd
from pandas import ExcelFile
from pandas import ExcelWriter

nombres = []
titulo = []
institucion = []
lista_celdas = []

workbook = read.open_workbook('Copia de Programación Mesas Temáticas FINAL.xlsx') # abre el archivo

lista_hojas = workbook.sheets() # lista de hojas en el libro

for i in range(len(lista_hojas)): # por cada hoja en el libro

    hoja = workbook.sheet_by_index(i) # selecciona la hoja

    for j in range(2, hoja.ncols): # recorre desde la tercera columna hasta la última

            columna = hoja.col_values(j,0) # crea una lista con lo que hay en la columna desde la segunda celda

            for k in range(hoja.nrows): # por cada celda de la columna

                celda = columna[k] # elige la celda
                dato = '' # crea un string vacío
                contenido = [] # crea la lista que albergará los tres datos
                i = 0
                alm = 'Almuerzo'
                aux = ['-----', '------', '------']


                for c in celda:
                    if(alm not in celda and celda != ''):
                        if i == len(celda)-1:
                            dato += c
                            contenido.append(dato) # mete el string en la lista de contenido
                            lista_celdas.append(contenido) # mete la lista de contenido en la lista de todas las celda
                            contenido = [] # vacía la lista de contenido
                            dato = ''
                            if(k==0 or alm in columna[k-1]):
                                lista_celdas.append(aux)
                        if c != '\n': # si el carácter no es un linebreak
                            dato += c # añade el carácter al string
                        elif c == '\n': # si el carácter es un linebreak
                            contenido.append(dato)
                            dato = '' # vacía el string
                        i+=1


#print(lista_celdas)

for i in lista_celdas:
    if len(i) == 4:
        cat_ins = i[2] + ' / ' + i[3]
        aux = [i[0], i[1], cat_ins]
        ind = lista_celdas.index(i)
        lista_celdas.remove(i)
        lista_celdas.insert(ind, aux)
    if len(i) < 3:
        i.append('empty')

for i in lista_celdas:
    if(len(i) == 2):
        i.append('empty')

for i in lista_celdas:
    titulo.append(i[0])
    nombres.append(i[1])
    institucion.append(i[2])

print(len(institucion))
print(len(nombres))
print(len(titulo))


columnas = pd.DataFrame({'1. Nombres': nombres, '2. Institución': institucion,
                        '3. Título de la actividad': titulo})

writer = ExcelWriter('DATOS_MESAS_2018.xlsx')
columnas.to_excel(writer,'Hoja1',index=False)
writer.save()
