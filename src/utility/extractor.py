# -*- coding: utf-8 -*-
"""
Created on Wed Dec  8 13:49:40 2021

@author: AARON.RAMIREZ
"""

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

    

        


