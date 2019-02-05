# -*- coding: utf-8 -*-

import xlrd as read
import pandas as pd

nombres = []
titulo = []
institucion = []
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
                dato = '' # crea un string vacío
                contenido = [] # crea la lista que albergará los tres datos
                i = 0
                alm = 'Almuerzo'

            
                for c in celda: # por cada carácter en la celda
                    if(alm not in celda and celda != ''):

                    	##### SI ALM IN FILA-1 ?


                        if i == len(celda)-1: # si es el último carácte
                       		dato += c
                       		contenido.append(dato) # mete el string en la lista de contenido
                       		lista_celdas.append(contenido) # mete la lista de contenido en la lista de todas las celdas
                       		contenido = [] # vacía la lista de contenido
                       		dato = '' # vacía el string
                       	if c != '\n': # si el carácter no es un linebreak
                       		dato += c # añade el carácter al string
                       	elif c == '\n': # si el carácter es un linebreak
                       		contenido.append(dato)
                       		dato = '' # vacía el string
                    i+=1


#print(lista_celdas)

#for i in lista_celdas:
#	print len(i)

for i in lista_celdas:
	if len(i) == 4:
		cat_ins = i[2] + ' / ' + i[3]
		institucion.append(cat_ins)
	elif len(i) == 3:
		titulo.append(i[0])
		nombres.append(i[1])
		institucion.append(i[2])
print(institucion)


		
# #lineas = celda.readlines()
# print(lista)
