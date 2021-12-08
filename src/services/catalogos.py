# -*- coding: utf-8 -*-
"""
Created on Mon Nov 22 10:53:17 2021

@author: AARON.RAMIREZ
"""
from openpyxl.worksheet.datavalidation import DataValidation
import pandas as pd
import string


catalogos = pd.read_csv(
    r"D:\OneDrive - INEGI\Documents\proyecto_validaciones\mock\base_catalogos.csv"
)


def validar_catalogo(pregunta, dic, hoja):
    abc1 = list(string.ascii_uppercase) + ["AA", "AB", "AC", "AD"]
    d = dic["catalogos"]
    ind = 0
    for frase in d["frases"]:
        try:
            frasem = frase.lower()
        except:
            frasem = ""  # por error de catalogos no aptos para validar
        fila = 0
        for palabra in catalogos["palabras"]:
            if palabra in frasem:
                lista = catalogos.iat[fila, 1]
                cor = d["coordenadas"][ind]
                letra = abc1[cor[1]]
                filavalidar = pregunta + cor[0] + 3
                celda = letra + str(filavalidar)
                tl = '"{}"'.format(lista)
                valid(hoja, [celda], tl)
            fila += 1
        ind += 1
    return


def valid(hoja, celdas, lista):
    "Hoja de openpy y celdas debe ser una lista para iterar con las lertras ya bien definidas"
    dv = DataValidation(type="list", formula1=lista, allow_blank=True)

    hoja.add_data_validation(dv)

    for celda in celdas:
        a1 = hoja[celda]
        a1.value = ""
        dv.add(a1)
    return
