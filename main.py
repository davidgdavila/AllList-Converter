import os
import threading
import time
from kivy.lang import Builder
from kivy.properties import BooleanProperty
from kivymd.app import MDApp
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.button import MDFlatButton
from kivymd.uix.dialog import MDDialog
from kivymd.uix.screen import MDScreen
from openpyxl import load_workbook
from kivy.core.window import Window
import tkinter as tk
from tkinter import filedialog as fd

def seleccionar_salvar_archivo(ruta_abrir):
    root = tk.Tk()
    root.withdraw()
    d = os.path.dirname(ruta_abrir)
    s = os.path.basename(ruta_abrir).replace(".xlsx", "")
    filename = fd.asksaveasfilename(
        initialfile = s + ".txt",
        title = 'Salvar Archivo',
        initialdir = d,
        filetypes = (('text files', '*.txt'), ('All files', '*.*')))
    ruta_guardar = filename
    if filename:
        hilo_procesar = threading.Thread(target=procesar_archivo, args=(ruta_abrir, ruta_guardar))
        hilo_procesar.start()
    else:
        MDApp.get_running_app().root.eti1.text = "Arrastra y suelta la pauta aquí."

def cargar_barra():
    for i in range(0,100):
        valor =+ i
        time.sleep(0.04)
        MDApp.get_running_app().root.eti1.text = "Convirtiendo archivo..." + str(round(valor)) + "%"
        MDApp.get_running_app().root.barra.value = valor


def texto_limpio(titulo_serie):
    reemplazar = (
        ("Á", "A"),
        ("É", "E"),
        ("Í", "I"),
        ("Ó", "O"),
        ("Ú", "U"),
        (" ", "_"),
        (".", "" ),
        (":", "" ),
        (",", "" ),
    )
    for a, b in reemplazar:
        titulo_serie = titulo_serie.replace(a, b)
    return titulo_serie

def procesar_archivo(ruta_abrir, ruta_guardar):
    texto = ""
    try:
        hilo2 = threading.Thread(target=cargar_barra)
        hilo2.start()
        doc = load_workbook(filename=ruta_abrir, data_only=True)
        hoja1 = doc.active
        filas = tuple(hoja1.rows)
        longitudfilas = len(filas)
        #MDApp.get_running_app().root.eti1.text = "Convirtiendo archivo..."
        texto = "<section> " + texto_limpio(str(hoja1["G3"].value)) + "\n"
        for i in range(1, longitudfilas):
            E = hoja1["E" + str(i)].value
            D = hoja1["D" + str(i)].value
            F = hoja1["F" + str(i)].value

            if D:
                if texto_limpio(str(F)).__contains__("SPOT") or texto_limpio(str(F)).__contains__("RTC"):
                    texto += str(E).replace(" ", "") + "\t" + "99:99:99:99" + "\t" "99:99:99:99" + "\n"
                else:
                    if texto_limpio(str(F)).__contains__("SERIE"):
                        pass
                    else:
                        if E:
                            texto += "<section> " + texto_limpio(str(F)) + "\n" \
                                     + str(E).replace(" ", "") + "\t" + "99:99:99:99" + "\t" + "99:99:99:99" + "\n"
                        else:
                            pass
            else:
                if E:
                    texto += str(E).replace(" ", "") + "\t" + "99:99:99:99" + "\t" "99:99:99:99" + "\n"
                else:
                    pass
        d = ruta_guardar.replace(".txt", "")
        with open(d + "dyno.txt", 'w') as stream:
            stream.write(texto)

        MDApp.get_running_app().root.eti1.text = "¡Se ha convertido el archivo!\n" + ruta_guardar
    except IndexError:
        MDApp.get_running_app().root.eti1.text = "El formato no es correcto"
    except UnicodeDecodeError:
        MDApp.get_running_app().root.eti1.text = "Error:¡Checa que sea un txt!"
    except:
        MDApp.get_running_app().root.eti1.text = "No se realizó la conversion.Error desconocido"


class Acerca_de(MDScreen):
    pass

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
        print("dyno:" + str(self.ids.cbox_dyno.active))
        print("versio:" + str(self.ids.cbox_versio.active))

        if self.ids.cbox_dyno.active or self.ids.cbox_versio.active:
            self.ruta_abrir = file_path.decode(encoding="utf-8")
            self.hilo()


        else:
            self.ids.eti1.text = "Debe de seleccionar al menos una lista"

    def hilo(self):
        hilo_ventana_salvar = threading.Thread(target=seleccionar_salvar_archivo, args=(self.ruta_abrir,))
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
            hilo1 = threading.Thread(target=procesar_archivo, args=(self.ruta_abrir,self.ruta_guardar))
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

class Version(MDBoxLayout):
    pass

class AllListConverterApp(MDApp):
    dialog = None
    def build(self):

        self.theme_cls.primary_palette = "BlueGray"
        #self.theme_cls.primary_hue = "200"
        #self.theme_cls.accent_palette = "Red"
        #self.theme_cls.theme_style = "Dark"
        Builder.load_file("AllListConverterApp.kv")



AllListConverterApp().run()
