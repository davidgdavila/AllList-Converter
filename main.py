import threading
import time
from kivy import Config
from kivy.core.window import Window
from kivy.lang import Builder
from kivy.properties import StringProperty
from kivy.uix.popup import Popup
from kivymd.app import MDApp
import tkinter as tk
from tkinter import filedialog as fd
import os
from kivymd.uix.boxlayout import MDBoxLayout
from openpyxl import load_workbook

class Acerca_de(MDBoxLayout):
    texto_etiqueta= StringProperty("Built on May 05th, 2022\n\n"
                                   "Powered by David González Software.\n"
                                   "david.gdavila08@gmail.com \n"
                                   "Runtime Version: 1.2.1\n\n"
                                   "Python 3.9, kivy 2.0.1, kivymd 0.104.2 \n\n"
                                   "AllListConverter was created to make K2 and \n"
                                   "versio system list in a more efficient way\n")
class AllListConverterApp(MDApp):

    directorio_inicial = os.path.expanduser("~") + "\\documents\\"
    bandera = True

    def __init__(self, **kwargs):
        super(AllListConverterApp, self).__init__(**kwargs)
        Config.set("kivy", "window_icon", "manzanaverde.ico")
        self.title = "All List-Converter 2.0"
        self.icon = "manzanaverde.ico"
        Window.bind(on_drop_file=self._on_file_drop)

    def _on_file_drop(self, window, file_path, *args):
        self.ruta_abrir = file_path.decode(encoding="utf-8")
        if self.root.ids.cbox_dyno.active or self.root.ids.cbox_versio.active:
            hilo_ventana_salvar = threading.Thread(name = "hilo_ventana_salvar", target=self.seleccionar_salvar_archivo, )
            hilo_ventana_salvar.start()
        else:
            self.root.ids.eti1.text = "Debe de seleccionar al menos una lista"

    def build(self):
        self.theme_cls.primary_palette = "BlueGray"
        self.root = Builder.load_file("All_List_Converter.kv")

    def modo_oscuro(self):
        if self.bandera:
            self.bandera=False
        else:
            self.bandera=True

        if self.bandera:
            self.theme_cls.theme_style = "Light"
            self.root.ids.cbox_dyno.selected_color = (0, 0, 0, 1)
            self.root.ids.cbox_versio.selected_color = (0, 0, 0, 1)

        else:
            self.theme_cls.theme_style = "Dark"
            self.root.ids.cbox_dyno.selected_color = (1, 1, 1, 1)
            self.root.ids.cbox_versio.selected_color = (1, 1, 1, 1)

    @staticmethod
    def texto_limpio(texto):
        reemplazar = (
            ("Á", "A"),
            ("É", "E"),
            ("Í", "I"),
            ("Ó", "O"),
            ("Ú", "U"),
            ("Ñ", "N"),
            (" ", "_"),
            (".", ""),
            (":", ""),
            (",", ""),
            ("¡", ""),
            ("!", ""),
            ("¿", ""),
            ("?", ""),
            ("…", ""),
            ("#", ""),
            ("\n", "_"),
            ("\t", "_"),
        )
        for a, b in reemplazar:
            texto = texto.replace(a, b)
        return texto

    def cargar_barra(self):
        self.bandera_dyno = True
        self.bandera_versio = True
        self.valor_barra = 0

        for i in range(0, 99):
            if self.bandera_versio and self.bandera_dyno:
                self.root.ids.eti1.text = "Convirtiendo pauta..." + self.nombre_pauta + " " + str(round(i)) + "%"
                self.root.ids.barra.value = i
                self.valor_barra = i
                time.sleep(0.04)
            else:
                break

    def crear_lista_dyno(self):
        texto = "<section> " + self.texto_limpio(str(self.hoja1["G3"].value)) + "\n"
        for i in range(1, self.longitudfilas+1):
            E = str(self.hoja1["E" + str(i)].value)
            E = E.replace(" ","")
            D = str(self.hoja1["D" + str(i)].value)
            F = str(self.hoja1["F" + str(i)].value)
            F = self.texto_limpio(F)
            if E == "None" or E == "PLAYLIST" or E == "":
                pass
            else:
                if D == "None" or "-" in D or "CAPSULA" in F or len(D) != 8:
                    texto += E + "\t" + "99:99:99:99" "\t" + "99:99:99:99" + "\n"
                else:
                    texto += "<section> " + F + "\n" + \
                             E + "\t" + "99:99:99:99" "\t" + "99:99:99:99" + "\n"
        d = self.ruta_guardar.replace(".txt", "")
        with open(d + "_dyno.txt", 'w') as stream:
            stream.write(texto)
        self.bandera_dyno = False
        if not self.bandera_versio:
            for i in range(self.valor_barra+1, 101):
                self.root.ids.eti1.text = "Convirtiendo pauta..." + self.nombre_pauta + " " + str(i) + "%"
                self.root.ids.barra.value = i
                time.sleep(0.02)
            self.root.ids.eti1.text = "¡Se ha convertido el archivo!\n" + self.ruta_guardar
        #MDApp.get_running_app().root.barra.value = 100

    def crear_lista_versio(self):
        texto = ""
        for i in range(1, self.longitudfilas+1):
            E = self.hoja1["E"+ str(i)].value
            F = self.hoja1["F" + str(i)].value
            G = self.hoja1["G" + str(i)].value
            H = self.hoja1["H" + str(i)].value
            J = self.hoja1["J" + str(i)].value
            L = self.hoja1["L" + str(i)].value
            texto += "\t\t\t\t"+str(E) + "\t" + str(F) + "\t" + str(G) + "\t"\
                     + str(H) + "\t\t" + str(J) + "\t\t" + str(L) + "\n"
            texto = texto.replace(str(None), "")
        d =self.ruta_guardar.replace(".txt","")
        with open(d +"_versio.txt", 'w') as stream:
            stream.write(texto)
        self.bandera_versio = False
        if not self.bandera_dyno:
            for i in range(self.valor_barra + 1, 101):
                self.root.ids.eti1.text = "Convirtiendo pauta..." + self.nombre_pauta + " " + str(i) + "%"
                self.root.ids.barra.value = i
                time.sleep(0.02)
            self.root.ids.eti1.text = "¡Se ha convertido el archivo!\n" + self.ruta_guardar

    def procesar_archivo(self):

        try:
            doc = load_workbook(filename=self.ruta_abrir, data_only=True)
            self.hoja1 = doc.active
            filas = tuple(self.hoja1.rows)
            self.longitudfilas = len(filas)
            if self.root.ids.cbox_dyno.active and self.root.ids.cbox_versio.active:
                hilo_dyno = threading.Thread(target=self.crear_lista_dyno, )
                hilo_versio = threading.Thread(target=self.crear_lista_versio, )
                hilo_dyno.start()
                hilo_versio.start()
                self.crear_lista_dyno()
                self.crear_lista_versio()
            elif self.root.ids.cbox_versio.active:
                self.bandera_dyno = False
                self.crear_lista_versio()
            elif self.root.ids.cbox_dyno.active:
                self.bandera_versio = False
                self.crear_lista_dyno()
            else:
                pass
        except IndexError:
            self.root.ids.eti1.text = "El formato no es correcto"

        except UnicodeDecodeError:
            self.root.ids.eti1.text = "Error: No abriste un archivo de excel"
        except Exception as e:

            self.root.ids.eti1.text = repr(e) + " Escribe a david.gdavila08@gmail.com"

    def seleccionar_salvar_archivo(self):
        root = tk.Tk()
        root.withdraw()
        directorio_pauta = os.path.dirname(self.ruta_abrir) #direccion de la pauta
        self.nombre_pauta = os.path.basename(self.ruta_abrir).replace(".xlsx", "")
        self.ruta_guardar = fd.asksaveasfilename(
            initialfile = self.nombre_pauta + ".txt",
            title = 'Salvar archivo.',
            initialdir = directorio_pauta,
            filetypes = (('text files', '*.txt'), ('All files', '*.*')))
        if self.ruta_guardar:
            root.destroy()
            hilo_1_procesar_archivo = threading.Thread(name = "hilo_procesar_archivo", target= self.procesar_archivo)
            hilo_1_procesar_archivo.start()
            hilo_2_cargar_barra = threading.Thread(name="cargar_barra", target=self.cargar_barra, )
            hilo_2_cargar_barra.start()
        else:
            root.destroy()

    def hilo_seleccionar_archivo(self):
        hilo_selecionar_archivo = threading.Thread(target = self.seleccionar_archivo, )
        hilo_selecionar_archivo.start()

    def seleccionar_archivo(self):
        if self.root.ids.cbox_dyno.active or self.root.ids.cbox_versio.active:
            root = tk.Tk()
            root.withdraw()
            self.ruta_abrir = fd.askopenfilename(
                title='Seleccionar Pauta...',
                initialdir=self.directorio_inicial,
                filetypes=(('excel files', '*.xlsx'), ('All files', '*.*')))
            if self.ruta_abrir:
                root.destroy()
                hilo_ventana_salvar = threading.Thread(target=self.seleccionar_salvar_archivo, )
                hilo_ventana_salvar.start()
            else:
                root.destroy()
        else:
            self.root.ids.eti1.text = "Debe de seleccionar al menos una lista"

    def show_acerca_de(self):
        contenido = Acerca_de()
        self.ventana_acercad = Popup(
            title = "Acerca de AllListConverter",
            content = contenido,
            size_hint =(.5,.7))
        self.ventana_acercad.open()

AllListConverterApp().run()
