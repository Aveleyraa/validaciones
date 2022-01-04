import pandas as pd
import openpyxl as op
from tkinter import filedialog
import tkinter as tk
import sys
<<<<<<< HEAD
import os
#sys.path.insert(0,'D:\OneDrive - INEGI\Documents\proyecto_validaciones') 
from services.principal_doc import procesarcoor
=======
sys.path.append('D:\Trabajo\Documentos\codigos_python\algoritmo_penitenciario2022')
from services.principal_doc import procesarcoor, p_especificas
>>>>>>> 3bd840a1df20cd4c40ff149f728647d4a622e84f

root= tk.Tk()
 
canvas1 = tk.Canvas(root, width = 800, height = 300)
canvas1.pack()
#imagen = tk.PhotoImage(file="D:\OneDrive - INEGI\Documents\proyecto_validaciones\mock\INEGI.png")
label1 = tk.Label(root, text='Importar Censo para validar' )
label1.config(font=('Arial', 20))
label1.place(x = 50, y = 0)
canvas1.create_window(400, 50, window=label1)


def main():

    """
    Esta funcion ejecuta el inicio del proceso

    Parameters
    ----------
    hopan : es un dataframe de pandas.

    Returns
    -------
    Regresa el excel validado

    """
    import_file_path = filedialog.askopenfilename()
    modulo = 2

    libro = import_file_path

    book = op.load_workbook(libro)

    p = book.sheetnames
    print(p)

    para_pags = ['Secc']
    pags = []
    for val in para_pags:
        pags1 = [pag for pag in p if val in pag]
        pags += pags1
    print(pags)  
    """
    Antes de correr el ciclo, asegurarse de que pags tenga todas las hojas
    en las que se van a escribir validaciones con este método.
    """

    for pa in pags:

        pagina = pa

        shi = book[pagina]
        shet = pd.read_excel(libro, sheet_name=pagina, engine="openpyxl")
        procesarcoor(shet, shi)
    
    nombre_archivo_salvado = "0{}_CNSIPEF_2022_M{}_validado.xlsx".format(modulo,modulo)
    directory = filedialog.askdirectory()
    book.save(directory + '/' + nombre_archivo_salvado)


#if __name__ == "__main__":
#    main()


browseButton_Excel = tk.Button(text='Cargar archivo...', command=main, bg='green', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(400, 180, window=browseButton_Excel)
 

button3 = tk.Button (root, text='Salir', command=root.destroy, bg='green', font=('helvetica', 11, 'bold'))
canvas1.create_window(400, 260, window=button3)
 
root.mainloop()