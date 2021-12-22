import pandas as pd
import openpyxl as op
from tkinter import filedialog
import sys
sys.path.append('D:\Trabajo\Documentos\codigos_python\algoritmo_penitenciario2022')
from services.principal_doc import procesarcoor, p_especificas


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
    modulo = 1

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

    book.save("{}_ver1.xlsx".format(modulo))


if __name__ == "__main__":
    main()
