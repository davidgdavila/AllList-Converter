import os
import threading
import time
from kivy import Config
from kivy.lang import Builder
from kivy.properties import BooleanProperty, StringProperty
from kivy.uix.popup import Popup
from kivymd.app import MDApp
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.screen import MDScreen
from openpyxl import load_workbook
from kivy.core.window import Window
import tkinter as tk
from tkinter import filedialog as fd

bandera = False
numero = 0
rutag= ""

class Nuevo_Hilo(object):

    def seleccionar_salvar_archivo(self, ruta_abrir):
        root = tk.Tk()
        root.withdraw()
        d = os.path.dirname(ruta_abrir)
        self.s = os.path.basename(ruta_abrir).replace(".xlsx", "")
        filename = fd.asksaveasfilename(
            initialfile = self.s + ".txt",
            title = 'Salvar Archivo',
            initialdir = d,
            filetypes = (('text files', '*.txt'), ('All files', '*.*')))
        ruta_guardar = filename
        if filename:
            hilo_procesar = threading.Thread(target = self.procesar_archivo,
                                             args=(ruta_abrir, ruta_guardar))
            hilo_procesar.start()
        else:
            MDApp.get_running_app().root.eti1.text = "Arrastra y suelta la pauta aquí."

    def cargar_barra(self,ruta_abrir):
        global numero,rutag
        s= os.path.basename(ruta_abrir).replace(".xlsx", "")
        rutag = s


        for i in range(0,99):
            if bandera:
                break
            else:
                valor =+ i
                time.sleep(0.02)
                MDApp.get_running_app().root.eti1.text = "Convirtiendo pauta..."+ s + " " + str(round(valor)) + "%"
                MDApp.get_running_app().root.barra.value = valor
                numero = valor


    def texto_limpio(self, texto):
        reemplazar = (
            ("Á", "A"),
            ("É", "E"),
            ("Í", "I"),
            ("Ó", "O"),
            ("Ú", "U"),
            (" ", "_"),
            (".", ""),
            (":", ""),
            (",", ""),
        )
        for a, b in reemplazar:
            texto = texto.replace(a, b)
        return texto

    def quitar_espacios(self, texto):
        texto.replace(" ","")
        return(texto)

    def crear_lista_dyno(self, ruta_guardar):
        texto = "<section> " + self.texto_limpio(str(self.hoja1["G3"].value)) + "\n"
        for i in range(1, self.longitudfilas+1):
            E = self.hoja1["E" + str(i)].value
            D = self.hoja1["D" + str(i)].value
            F = self.hoja1["F" + str(i)].value

            if D:
                if self.texto_limpio(str(F)).__contains__("SPOT") or self.texto_limpio(str(F)).__contains__("RTC"):
                    texto += str(E).replace(" ", "") + "\t" + "99:99:99:99" + "\t" "99:99:99:99" + "\n"
                else:
                    if self.texto_limpio(str(F)).__contains__("SERIE") or self.texto_limpio(str(F)).__contains__("CLASIFICACIONES"):
                        pass
                    else:
                        if E:
                            texto += "<section> " + self.texto_limpio(str(F)) + "\n" \
                                     + str(E).replace(" ", "") + "\t" + "99:99:99:99" + "\t" + "99:99:99:99" + "\n"
                        else:
                            pass
            else:
                if E:
                    texto += str(E).replace(" ", "") + "\t" + "99:99:99:99" + "\t" "99:99:99:99" + "\n"
                else:
                    pass
        d = ruta_guardar.replace(".txt", "")
        with open(d + "_dyno.txt", 'w') as stream:
            stream.write(texto)
        for valor in range(numero+1, 101):
            time.sleep(0.02)
            MDApp.get_running_app().root.eti1.text = "Convirtiendo pauta..." + rutag + " " + str(round(valor)) + "%"
            MDApp.get_running_app().root.barra.value = valor
        MDApp.get_running_app().root.eti1.text = "¡Se ha convertido el archivo!\n" + ruta_guardar
        #MDApp.get_running_app().root.barra.value = 100


    def crear_lista_versio(self,ruta_guardar):
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
        d =ruta_guardar.replace(".txt","")
        with open(d +"_versio.txt", 'w') as stream:
            stream.write(texto)
        MDApp.get_running_app().root.eti1.text = "¡Se ha convertido el archivo!\n" + ruta_guardar
        MDApp.get_running_app().root.barra.value = 100

    def procesar_archivo(self, ruta_abrir, ruta_guardar):
        texto = ""
        try:
            hilo2 = threading.Thread(target = self.cargar_barra, args= (ruta_abrir,))
            hilo2.start()
            doc = load_workbook(filename=ruta_abrir, data_only=True)
            self.hoja1 = doc.active
            self.filas = tuple(self.hoja1.rows)
            self.longitudfilas = len(self.filas)
            if  MDApp.get_running_app().root.cbox_dyno.active:
                self.crear_lista_dyno(ruta_guardar)
                if MDApp.get_running_app().root.cbox_versio.active:
                    self.crear_lista_versio(ruta_guardar)
                else:
                    pass
            elif MDApp.get_running_app().root.cbox_versio.active:
                self.crear_lista_versio(ruta_guardar)
            else:
                pass

        except IndexError:
            MDApp.get_running_app().root.eti1.text = "El formato no es correcto"

        except UnicodeDecodeError:
            MDApp.get_running_app().root.eti1.text = "Error: No abriste un archivo de excel"
        except:
            MDApp.get_running_app().root.eti1.text = "Error desconocido. Escribe a david.gdavila08@gmail.com"

class Principal(MDScreen):
    texto = ""
    ruta_abrir = ""
    ruta_guardar = ""
    directorio_inicial = os.path.expanduser("~") + "\\documents\\"
    estado_cbox_dyno= BooleanProperty()
    estado_cbox_versio = BooleanProperty()

    def __init__(self,**kwargs):
        super().__init__(**kwargs)
        Window.bind(on_drop_file=self._on_file_drop)
        self.estado_cbox_versio= True
        self.estado_cbox_dyno = True

    def _on_file_drop(self, window, file_path, *args):
        self.ruta_abrir = file_path.decode(encoding="utf-8")
        print("dyno:" + str(self.ids.cbox_dyno.active))
        print("versio:" + str(self.ids.cbox_versio.active))

        if self.ids.cbox_dyno.active:
            if self.ids.cbox_versio.active:
                self.hilo()
            else:
                self.hilo()
        else:
            if self.ids.cbox_versio.active:
                self.hilo()
            else:
                self.ids.eti1.text = "Debe de seleccionar al menos una lista"

    def hilo(self):
        hilo_ventana_salvar = threading.Thread(target= self.seleccionar_salvar_archivo,)
        hilo_ventana_salvar.start()

    def seleccionar_salvar_archivo(self):
        root = tk.Tk()
        root.withdraw()
        directorio_pauta=os.path.dirname(self.ruta_abrir) #direccion de la pauta
        nombre_pauta=os.path.basename(self.ruta_abrir).replace(".xlsx","")
        filename = fd.asksaveasfilename(
            initialfile= nombre_pauta+".txt",
            title='Salvar Archivo',
            initialdir=directorio_pauta,
            filetypes=(('text files', '*.txt'), ('All files', '*.*')))
        self.ruta_guardar=filename
        if filename:
            hilo1 = threading.Thread(target= Nuevo_Hilo().procesar_archivo, args=(self.ruta_abrir,self.ruta_guardar))
            hilo1.start()
        else:
            pass

    def seleccionar_archivo(self):
        if self.ids.cbox_dyno.active or self.ids.cbox_versio.active:
            root = tk.Tk()
            root.withdraw()
            filename = fd.askopenfilename(
                title='Seleccionar Pauta...',
                initialdir=self.directorio_inicial,
                filetypes=(('excel files', '*.xlsx'), ('All files', '*.*')))
            self.ruta_abrir = filename
            if filename:
                self.seleccionar_salvar_archivo()
            else:
                self.ids.eti1.text = "Arrastra y suelta la pauta aqui."
        else:
            self.ids.eti1.text = "Debe de seleccionar al menos una lista"

    def show_acerca_de(self):
        contenido = Acerca_de()
        self.ventana_acercad = Popup(
            title = "Acerca de AllListConverter",
            content = contenido,
            size_hint =(.6,.6))
        self.ventana_acercad.open()

class Acerca_de(MDBoxLayout):
    texto_etiqueta= StringProperty("Built on May 05th, 2022\n\n"
                                   "Powered by David González Software.\n"
                                   "david.gdavila08@gmail.com \n"
                                   "Runtime Version: 1.0.1\n\n"
                                   "Python 3.9, kivy 2.1.0, kivymd 0.104.2 \n\n"
                                   "AllListConverter was created to make K2 and \n"
                                   "versio system list in a more efficient way\n")


class AllListConverterApp(MDApp):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        Config.set("kivy", "window_icon", "manzanaverde.ico")
        self.title = "All List-Converter"
        self.icon = "manzanaverde.ico"

    def build(self):

        self.theme_cls.primary_palette = "BlueGray"
        #self.theme_cls.primary_hue = "200"
        #self.theme_cls.accent_palette = "Red"
        #self.theme_cls.theme_style = "Dark"
        self.root = Builder.load_file("All_List_Converter.kv")



AllListConverterApp().run()
