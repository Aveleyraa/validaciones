# -*- coding: utf-8 -*-
"""
Created on Wed Dec  8 13:49:40 2021

@author: AARON.RAMIREZ
"""
import pandas as pd

class Extractor():
    
    def __init__(self):
        
        self.contenido = []
        self.anexo = []
    
    def extraer(self,lista):
        self.anexo += lista
        
    def reiniciar(self):
        self.contenido = []
        self.anexo = []
    
    def juntar(self):
        self.contenido += self.anexo
        a = [] + self.contenido
        self.contenido = []
        self.anexo = []
        return a

class Datos():

    def __init__(self,nombre):
        self.nombre = nombre
        self.db = {'pregunta':[],'totales':[]}
        self.dataframe = []
    
    def columna_vacia(self,nombre):
        self.db[nombre] = []
    
    def columna_convalor(self,lista,nombre_columna):
        self.db[nombre_columna] = lista

    def agregar_valor_acolumna(self,valor,nombre_columna):
        self.db[nombre_columna].append(valor)
    
    def ajustar_db(self):
        self.db = ajustar(self.db)
    
    def conjuntar_db(self):
        self.db = conjuntar(self.db)
        
    def crear_df(self):
        di = ajustar(self.db)
        try:
            self.dataframe = pd.DataFrame(di)
        except:
            print('Hay un problema con el diccionario para crear la base de datos')
        return self.dataframe        
    
    def guardar(self):
        if type(self.dataframe) != list:
            self.dataframe.to_csv(self.nombre+'.csv')
            print('frame guardado correctamente')
        else:
            print('no se guardó porque no hay dataframe')
    
        
def ajustar(diccionario):
    revisar = []
    for key in diccionario:
        revisar.append(len(diccionario[key]))
    variacion = list(set(revisar))
    
    if len(variacion) > 1:
        mayor = max(revisar)
        for key in diccionario:
            if len(diccionario[key]) != mayor:
                faltantes = mayor - len(diccionario[key])
                for valor in range(faltantes):
                    diccionario[key].append(0)
        return diccionario
    
    if len(variacion) <= 1:
        return diccionario

def conjuntar(diccionario):
    if len(diccionario['totales']) > len(diccionario['pregunta']):      
        nel = []
        for ele in diccionario['totales'][len(diccionario['pregunta'])-1:]:
            nel += ele 
            diccionario['totales'].remove(ele)
        diccionario['totales'].append(nel)
    return diccionario
# alfa = Datos('perez')
# alfa.agregar_valor_acolumna(5,'pregunta')
# alfa.columna_convalor(['jo',5,4,6], 'prueba')
# alfa.columna_vacia('vacia')
# t = alfa.crear_db()
# alfa.guardar()


