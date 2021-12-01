# -*- coding: utf-8 -*-
"""
Created on Thu Oct 21 10:42:08 2021

@author: AARON.RAMIREZ
"""

import pandas as pd
import openpyxl as op
from principal_doc import procesarcoor


modulo = 1  

libro = '0{}_CNSIPEF_2022_M{}.xlsx'.format(modulo,modulo)   


book = op.load_workbook(libro)  

p = book.sheetnames
print (p)   

if modulo ==1:
    pags = p[4:13]  

if modulo ==2:
    pags = p[4:8] #modulo 2
print(pags)
"""
Antes de correr el ciclo, asegurarse de que pags tenga todas las hojas
en las que se van a escribir validaciones con este método.
""" 

for pa in pags:
    
    pagina = pa
    
    shi = book[pagina]
    shet = pd.read_excel(libro,sheet_name=pagina,engine='openpyxl')
    procesarcoor(shet, shi)
    
book.save('CNSIPEF_M{}_ver1.xlsx'.format(modulo))