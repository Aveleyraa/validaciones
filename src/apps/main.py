import pandas as pd
import openpyxl as op
from tkinter import filedialog, messagebox
import tkinter as tk
from utility.common_utils import CommonUtils
from services.principal_doc import procesarcoor

root= tk.Tk()
 
canvas1 = tk.Canvas(root, width = 800, height = 300)
canvas1.pack()
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
    
    nombre_archivo_salvado = CommonUtils.path_leaf(import_file_path)
    nombre_archivo_salvado = 'Archivo_validado_' + nombre_archivo_salvado 
    directory = filedialog.askdirectory()
    book.save(directory + '/' + nombre_archivo_salvado)
    messagebox.showinfo('Aviso', 'Se ha terminado la validación del censo!')


#if __name__ == "__main__":
#    main()


browseButton_Excel = tk.Button(text='Cargar archivo...', command=main, bg='blue1', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(400, 180, window=browseButton_Excel)
 

button3 = tk.Button (root, text='Salir', command=root.destroy, bg='green', font=('helvetica', 11, 'bold'))
canvas1.create_window(400, 260, window=button3)
 
root.mainloop()