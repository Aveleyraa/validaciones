from fileinput import filename
import pandas as pd
import openpyxl as op
from tkinter import filedialog, messagebox
import tkinter as tk
import customtkinter
from PIL import Image, ImageTk
from utility.common_utils import CommonUtils
from services.principal_doc import procesarcoor
import tkinter
import tkinter.messagebox
import customtkinter
import sys


customtkinter.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"


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
            procesarcoor(shet, shi, pagina)

        nombre_archivo_salvado = CommonUtils.path_leaf(import_file_path)
        nombre_archivo_salvado = 'Archivo_validado_' + nombre_archivo_salvado
        data = [('xlsx', '*.xlsx')] 
        filename = filedialog.asksaveasfile(filetypes=data, defaultextension=data,initialfile = nombre_archivo_salvado)
        book.save(filename.name)
        messagebox.showinfo('Aviso', 'Se ha terminado la validación del censo!')    


if __name__ == "__main__":
    app = App()
    app.start()