# -*- coding: utf-8 -*-
"""
Created on Mon Nov 22 10:53:17 2021

@author: AARON.RAMIREZ
"""
from sys import path
from openpyxl.worksheet.datavalidation import DataValidation
import pandas as pd
import string
import os

a=os.path.split(os.getcwd())[0]

catalogos = pd.read_csv("{}/mock/base_catalogos.csv".format(a))


def validar_catalogo(pregunta,dic,hoja):
    """
    

    Parameters
    ----------
    pregunta : int. Es el numero de fila en dataframe de pandas 
    dic : dict. Diccionario con las coordenadas y frases para clasificar
    hoja : hoja de openpyxl

    Returns
    -------
    None.

    """
    abc1 = list(string.ascii_uppercase)+['AA','AB','AC','AD']
    d = dic['catalogos']
    final = dic['final1'] - 1
    ind = 0
    for frase in d['frases']:
        try:
            frasem = frase.lower()
        except:
            frasem = "" #por error de catalogos no aptos para validar
        fila = 0
        for palabra in catalogos['palabras']:
            if palabra in frasem:
                lista = catalogos.iat[fila,1]
                cor = d['coordenadas'][ind]
                letra = abc1[cor[1]]
                filavalidar = pregunta+cor[0]+3
                celda = letra+str(filavalidar)
                celdas = emparejar_coordenadas([celda], final)
                tl = '"="",{}"'.format(lista)
                valid(hoja,celdas,tl)
            fila += 1
        ind += 1
    return

def valid(hoja,celdas,lista):
    """
    Hoja de openpy. Celdas debe ser una lista para iterar 
    con las lertras ya bien definidas
    """
    
    dv = DataValidation(type="list", formula1=lista, allow_blank=True)

    hoja.add_data_validation(dv)
    
    for celda in celdas:
        dv.add(celda)
    return

def emparejar_coordenadas(lista_inicio,numero_fin):
    """
    

    Parameters
    ----------
    lista_inicio : list. Coordenadas de inicio de tabla tipo excel
    numero_fin : int. numero de filas que contiene la tabla

    Returns
    -------
    coords : list. Lista con coordenadas de rangos inicio-fin tipo excel.
    Ejemplo:  'A1:A10'

    """
    
    coords = []
    
    for el in lista_inicio:
        numeroi = [val for val in el if val.isdigit()]
        a = '0'
        for i in numeroi:
            a += i 
        cifra_inicio = int(a)
        listaleras = el.split(str(cifra_inicio))
        letra_columna = listaleras[0]
        cifra_fin = cifra_inicio + numero_fin
        if cifra_fin < cifra_inicio:
            cifra_fin = cifra_inicio
        corrdenada_final = letra_columna + str(cifra_fin) 
        coords.append(el + ':' + corrdenada_final)                    

    return coords