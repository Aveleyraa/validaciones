from fileinput import filename
from multiprocessing.sharedctypes import Value
import pandas as pd
import openpyxl as op
from tkinter import filedialog, messagebox
import tkinter as tk
import customtkinter
from PIL import Image, ImageTk
from utility.common_utils import CommonUtils
from utility.common_utils_encontrar import CommonUtils_encontrar
from services.principal_doc import procesarcoor
import tkinter
import tkinter.messagebox
import customtkinter
import sys


customtkinter.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
#customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"


class App(customtkinter.CTk):

    WIDTH = 600
    HEIGHT = 350

    def __init__(self):
        super().__init__()

        self.title("Validación de censos de gobierno")
        self.geometry(f"{App.WIDTH}x{App.HEIGHT}")
        # self.minsize(App.WIDTH, App.HEIGHT)

        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        if sys.platform == "darwin":
            self.bind("<Command-q>", self.on_closing)
            self.bind("<Command-w>", self.on_closing)
            self.createcommand('tk::mac::Quit', self.on_closing)

        # ============ create frame ============

        # configure grid layout (1x2)
        self.grid_columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        self.frame_right = customtkinter.CTkFrame(master=self)
        self.frame_right.grid(row=0, column=1, sticky="nswe", padx=20, pady=20)

        # ============ frame_right ============

        # configure grid layout (3x7)
        #for i in [0, 1, 2, 3]:
            #self.frame_right.rowconfigure(i, weight=1)
        self.frame_right.rowconfigure(7, weight=10)
        self.frame_right.columnconfigure(0, weight=1)
        self.frame_right.columnconfigure(1, weight=1)
        self.frame_right.columnconfigure(2, weight=0)
        self.frame_info = customtkinter.CTkFrame(master=self.frame_right)
        self.frame_info.grid(row=1, column=0, columnspan=4, rowspan=4, pady=20, padx=20, sticky="nsew")


        self.switch_2 = customtkinter.CTkSwitch(master=self.frame_info,
                                                text="Modo oscuro",
                                                command=self.change_mode)
        self.switch_2.grid(row=8, column=0, pady=10, padx=20, sticky="w")
        # ============ frame_right -> frame_info ============

        self.frame_info.rowconfigure(1, weight=1)
        self.frame_info.columnconfigure(0, weight=2)

        self.label_info_1 = customtkinter.CTkLabel(master=self.frame_info,
                                                    text="                Sistema de Validación de Censos.    \n" +
                                                        "1. Seleccione un censo en blanco para su validación.\n" +
                                                        "2. Al terminar la validación aparecerá un mensaje exitoso." ,
                                                    height=100,
                                                    fg_color=("white", "gray38"),  # <- custom tuple-color
                                                    justify=tkinter.LEFT)
        self.label_info_1.grid(column=0, row=0, sticky="nwe", padx=15, pady=15)

        self.progressbar = customtkinter.CTkProgressBar(master=self.frame_info)
        self.progressbar.grid(row=1, column=0, sticky="ew", padx=15, pady=15)

        # ============ frame_right <- ============

        self.boton_importar = customtkinter.CTkButton(master=self.frame_right,
                                                        text="Importar archivo",
                                                        command=self.start_validation)
        self.boton_importar.grid(row=6, column=1, pady=20, padx=20, sticky="w")
        
        self.boton_importar_2 = customtkinter.CTkButton(master=self.frame_right,
                                                        text="Importar archivo a revisar",
                                                        command=self.start_revision)
        self.boton_importar_2.grid(row=6, column=0, pady=20, padx=20, sticky="w")

        self.button_5 = customtkinter.CTkButton(master=self.frame_right,
                                                text="Salir",
                                                command=self.destroy)
        self.button_5.grid(row=6, column=2, columnspan=1, pady=20, padx=20, sticky="we")

        # set default value
        self.progressbar.set(0.5)

    def button_event(self):
        print("Button pressed")

    def change_mode(self):
        if self.switch_2.get() == 1:
            customtkinter.set_appearance_mode("dark")
        else:
            customtkinter.set_appearance_mode("light")

    def on_closing(self, event=0):
        self.destroy()

    def start(self):
        self.mainloop()

    def start_validation(self):

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
            try:
                procesarcoor(shet, shi, pagina)
            except Exception as e:
                    messagebox.showerror(message='error: "{}"'.format(e))

        nombre_archivo_salvado = CommonUtils.path_leaf(import_file_path)
        nombre_archivo_salvado = 'Archivo_validado_' + nombre_archivo_salvado
        data = [('xlsx', '*.xlsx')] 
        filename = filedialog.asksaveasfile(filetypes=data, defaultextension=data,initialfile = nombre_archivo_salvado)
        book.save(filename.name)
        messagebox.showinfo('Aviso', 'Se ha terminado la validación del censo!')

    
    def start_revision(self):
        """
        Parameters
        ----------
        sec : self
            Es el nombre de la hoja de excel que se va a leer.
        datos : dataframe pandas
            es la hoja de excel a leer, expresada en un dataframe de pandas.

        Returns
        -------
        generar un archivo de salida llamado "PR_(nombre en variable libro).csv"
        Ese archivo generado será útil para los otros dos procesos siguientes. Se trata 
        del archivo guía.

        """

        libro = filedialog.askopenfilename()
        base = {}

        #Aqui se inicia el proceso de lectura para cada hoja del documento
        pags = pd.ExcelFile(libro).sheet_names

        for pag in pags:
            pagina = pag
            data = pd.read_excel(libro,sheet_name=pagina,engine='openpyxl')
            try:
                a = CommonUtils_encontrar.cordenadas(pag,data)
                r = CommonUtils_encontrar.nframe(a)
                base[pag] = r
            except Exception as e:
                messagebox.showerror(message='error: "{}"'.format(e))
        #sacar sumas y columnas a los diccionarios generados para cada hoja donde se encontró una marca

        for k in base:

            sua = []
            colum = []
            for ele in base[k]['ID']:
                if '+' in ele:
                    sua.append(ele)
                if ':' in ele:
                    colum.append(ele)
            sua1 = list(set(sua))

            if colum:
                try:
                    column = list(set(colum))
                    base[k] = CommonUtils_encontrar.columnas(column,base[k],k)
                except Exception as e:
                    messagebox.showerror(message='error: "{}"'.format(e))    
            
            for suma in sua1:
                seccion = k
                coord = []
                ide = ''.join(l for l in suma if l != '+')
                try:
                    op = CommonUtils_encontrar.clasif(suma)
                    c = 0
                except Exception as e:
                    messagebox.showerror(message='error: "{}"'.format(e))    
                
                for ele in base[k]['ID']:
                    if ele == suma:
                        coord.append(base[k]['coordenada'][c])
                    c += 1
                cor = coord[0]
                for i in coord[1:]:
                    cor += ','+i
                d = {'seccion':seccion,'coordenada':cor,
                    'comparacion':ide,'operacion':op,'ID':ide}
                if '.' in d['comparacion']:
                    iterar = d['comparacion'].split('.')
                    n = iterar[0]
                    d['comparacion'] = n
                    for ke in d:
                        base[k][ke].append(d[ke])
                else:
                    for ke in d:
                        base[k][ke].insert(0,d[ke])


        #Hacer el dataframe

        c = 0
        for k in base:
            if c == 0:
                original = pd.DataFrame(base[k])
            else:
                ad = pd.DataFrame(base[k])
                original = pd.concat([original,ad],ignore_index=True)
            c += 1

        #Hacer formulas   
        formulas = []
        fila = 0
        for element in original['comparacion']:
            if element == original['ID'][fila]:
                formulas.append('NA')

            else:
                c = original['coordenada'][fila]

                if ',' in c:
                    c = CommonUtils_encontrar.formulaS(c,'')
                a = CommonUtils_encontrar.determinar(original['operacion'][fila])
                sec = original['seccion'][fila]
                filac = 0
                for ele in original['ID']:
                    if ele == element:
                        b = original['coordenada'][filac]
                        sec1 = original['seccion'][filac]
                        if sec != sec1: 
                            b = sec1+'!'+b
                        if ',' in b:
                            if sec == sec1:
                                b = CommonUtils_encontrar.formulaS(b,'')
                            else: #referente de formula a otra hoja
                                b = CommonUtils_encontrar.formulaS(b,sec1)

                    else:
                        pass
                    filac += 1

                if a == 'posible mala referencia':
                    signos = ['<','>','=']
                    rt = 0
                    for signo in signos:
                        if signo in original['ID'][fila]: #esta comprobación es para las validaciones de columnas donde se usa el ":"
                            rt = 1
                    if original['operacion'][fila] == 'ref' and rt == 1:
                        formulas.append('NA')
                    else:
                        formulas.append(a)
                else:
                    try:
                        form = f'=IF(AND({c}{a}{b},OR(AND(ISNUMBER({b}),ISNUMBER({c})),AND(ISBLANK({b}),ISBLANK({c})),OR(AND(ISBLANK({b}),{c}=""),AND(ISBLANK({c}),{b}="")))),0,IF(OR(AND({c}="NS",{b}>0,ISNUMBER({b})),AND({c}="NS",{b}="NS"),OR(AND({b}="NA",{c}="NA"),AND({b}="NA",ISBLANK({c})))),0,1))'
                        formulas.append(form)
                    except:
                        form = 'No existe su referente'
                        formulas.append(form)
            fila += 1

        original['formulas'] = formulas

        #depurar formulas y dataframe

        fila = 0
        for element in original['ID']:
            if '+' in element:
                if '.' in element:
                    compa = ['>','<','=']
                    cumple = 'no'
                    for sig in compa:
                        if sig in element:
                            cumple = 'si'
                    if cumple == 'no':
                        formulas.append('Mala referencia')
                        original['formulas'][fila] = 'Mala referencia'
                    if cumple == 'si':
                        original['formulas'][fila] = 'Borrar'
                else:
                    original['formulas'][fila] = 'Borrar'
            if len(original['comparacion'][fila])>2:
                original['formulas'][fila] = 'NA'
            fila += 1
        borrar = original[original['formulas']=='Borrar'].index
        original = original.drop(borrar)

        #guardar
        nombre_archivo_salvado = CommonUtils.path_leaf(libro)
        nombre_archivo_salvado = 'PR_' + nombre_archivo_salvado
        data = [('csv', '*.csv')] 
        filename = filedialog.asksaveasfile(filetypes=data, defaultextension=data,initialfile = nombre_archivo_salvado)
        original.to_csv(filename.name,index=False)
        messagebox.showinfo('Aviso', 'Se ha completado el proceso')    


if __name__ == "__main__":
    app = App()
    app.start()