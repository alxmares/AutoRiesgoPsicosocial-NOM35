import customtkinter as ctk
from CTkMessagebox import CTkMessagebox
from PIL import Image
import numpy as np
import os
import sys

# Función para evitar problemas con los archivos en la conversión a .exe
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


class ScrollRadiobuttonFrame(ctk.CTkScrollableFrame):
    def __init__(self,master,questions, on_change_function,**kwargs):
        super().__init__(master, **kwargs)
        
        self.on_change_function = on_change_function
        self.radiovar_list = []
        self.radiobutton_list = []
        self.valores = ["Siempre", "Casi siempre", "Algunas veces", "Casi nunca", "Nunca"]
        for num,question in enumerate(questions):
            question = str(num+1) + ". " + question
            self.add_items(num,question)
        
        
    # Crear filas de preguntas
    def add_items(self, row, question):
        # Agregar variable
        radiovar = ctk.Variable()
        radiovar.trace_add("write", self.on_change_function)
        self.radiovar_list.append(radiovar)
        
        row*=2
        
        question_label = ctk.CTkLabel(self, text=question)
        question_label.grid(row=row,sticky="w", columnspan=5)
        
        for i in range(5):
            radiobutton = ctk.CTkRadioButton(self, text=self.valores[i], value=i, 
                                             variable=self.radiovar_list[int(row/2)])
            radiobutton.grid(row=row+1, column=i,padx=(0,40), pady=(0,50), sticky="w")
            self.radiobutton_list.append(radiobutton)
    
    def get_results(self):
        results = []
        for radiovar in self.radiovar_list:
            if radiovar:
                results.append(radiovar.get())
            else:
                results.append(None)
        return results
                
class App(ctk.CTk):
    def __init__(self, preguntas, callback_function1, callback_function2, callback_function3):
        super().__init__()
        self.callback_function1 = callback_function1
        self.callback_function2 = callback_function2
        self.callback_function3 = callback_function3
        
        self.preguntas=preguntas
        
        self.title("Analizador de Riesgo Psicosocial")
        self.geometry("800x620")
        # Configurar filas y columnas
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(2, weight=1)
        
        self.grid_rowconfigure(2, weight=1)
        self.grid_rowconfigure(3, weight=1)
        
        # Nombre,edad, Puesto, Empresa
        # Name Label
        self.nameLabel = ctk.CTkLabel(self,
                                      text="Nombre:")
        self.nameLabel.grid(row=0, column=0,
                            padx=10, pady=5,
                            sticky="ew")
        # Name Entry Field
        self.nameEntry = ctk.CTkEntry(self,
                         placeholder_text="Juan Domínguez Rosas")
        self.nameEntry.grid(row=0, column=1,
                            padx=10, columnspan=2,
                            pady=5, sticky="ew")
        self.nameEntry.bind("<Key>", command=self.on_change) # Seguimiento de evento
        # Age Label
        self.ageLabel = ctk.CTkLabel(self,
                                     text="Edad:")
        self.ageLabel.grid(row=0, column=3,
                           padx=5, pady=5,
                           sticky="ew")
        # Age Entry Field
        self.ageEntry = ctk.CTkEntry(self,
                            placeholder_text="20")
        self.ageEntry.grid(row=0, column=4,
                            padx=10, pady=5, sticky="ew")
        self.ageEntry.bind("<Key>", command=self.on_change)
        
        
        
        # Empresa Label
        self.empLabel = ctk.CTkLabel(self,
                                     text="Empresa:")
        self.empLabel.grid(row=1, column=0,
                           padx=20, pady=6,
                           sticky="ew")
        # Empresa Entry Field
        self.empEntry = ctk.CTkEntry(self,
                            placeholder_text="Empresa S.A. de C.V.")
        self.empEntry.grid(row=1, column=1,columnspan=2,
                            padx=10, pady=6, sticky="ew")
        self.empEntry.bind("<Key>", command=self.on_change)
        # Puesto Label
        self.puestoLabel = ctk.CTkLabel(self,
                                     text="Puesto: ")
        self.puestoLabel.grid(row=1, column=3,
                           padx=5, pady=6,
                           sticky="ew")
        # Puesto Entry Field
        self.puestoEntry = ctk.CTkEntry(self,
                            placeholder_text="Administrador")
        self.puestoEntry.grid(row=1, column=4,
                            padx=10, pady=6, sticky="ew")
        self.puestoEntry.bind("<Key>", command=self.on_change)
        
        
        
        # R E S U L T A D O S
        self.generateResultsButton = ctk.CTkButton(self,
                                         text="Generar Resultados",
                                         command=self.generate_results)
        self.generateResultsButton.grid(row=4, column=1,padx=(20,0),pady=4,
                                        columnspan=3, sticky="ew") 
        # G R A F I C A R 
        self.generate_graphs = ctk.CTkButton(self,
                                         text = "Ver Gráficas",
                                         command = self.graficar, state="disabled")
        self.generate_graphs.grid(row=5, column=1, padx=(20,0),pady=4,sticky="ew")
        # D O C U M E N T O S
        self.generate_document = ctk.CTkButton(self,
                                               text = "Generar Documento",
                                               command = self.generate_doc, state="disabled")
        self.generate_document.grid(row=5, column=2, columnspan=2, padx=(20,0),pady=4,sticky="ew")
        
        
        # A L E A T O R I O
        self.generarButton = ctk.CTkButton(self,
                                           text="Aleatorio",
                                           command=self.generar_aleatorio)
        self.generarButton.grid(row=6, column=1, columnspan=3, padx=(20,0), pady=4,sticky="ew")
        
        # R E I N I C I A R 
        self.resetButton = ctk.CTkButton(self,
                                         text="Reiniciar",
                                         command=self.reset_all)   
        self.resetButton.grid(row=7, column=1, columnspan=3, padx=(20,0),pady=4, sticky="ew")

        # I N F O R M A C I Ó N   D E   C O N T A C T O
        self.infoButton = ctk.CTkButton(self,
                                         text="Contacto",
                                         command=self.show_info)   
        self.infoButton.grid(row=7, column=4, padx=(20,20),pady=4, sticky="ns")
        
        # =======================
        # C U E S T I O N A R I O
        self.radiobuttons_frame = ScrollRadiobuttonFrame(self, questions=preguntas,on_change_function=self.on_change, 
                                                         corner_radius=10)
        self.radiobuttons_frame.grid(row=2, column=0, rowspan=2, columnspan=5,sticky="nsew", padx=(5,5))
        
    # Si hay algún en algún botón se deshabilitan los botones
    def on_change(self, _name=None, _index=None, _mode=None):
        if self.generate_graphs.cget("state") == "normal":
            self.generate_graphs.configure(state="disabled")
            self.generate_document.configure(state="disabled")
            
    def generate_doc(self):
        msg = CTkMessagebox(title="Logotipo", message="¿Buscar logotipo para añadir?",
                            icon=resource_path("question.png"), option_1="Sí", option_2="No")
        response = msg.get()
        
        if response == "Sí":
            op = True
        else:
            op = False
            
        self.callback_function3(op)
        
    def error_doc(self):
        CTkMessagebox(title="Error al Guardar", message="No se puede guardar mientras el archivo Word esté abierto",
                              icon=resource_path("cancel.png"))
    def saved_doc(self):
        CTkMessagebox(title="Archivo Guardado", message="Archivo Guardado Correctamente",
                   icon=resource_path("check.png"))
    
    def graficar(self):
        buff = self.callback_function2()
        graph_window = ctk.CTkToplevel()
        width = graph_window.winfo_screenwidth()
        height = graph_window.winfo_screenheight()
        
        center_x = int(width/2)
        center_y = int(height/2)
        graph_window.geometry(f"{width}x{height}+{0}+{0}")
        graph_window.wm_title("Gráficas de barras")
        
        # Mantener un margen y calcular el tamaño máximo de la imagen
        margen = 0  # Margen total
        max_ancho = width - margen
        max_alto = height - margen

        # Calcular el tamaño de la imagen manteniendo la relación de aspecto 15/12
        relacion_aspecto = 10 / 5
        if max_alto * relacion_aspecto <= max_ancho:
            imagen_alto = max_alto
            imagen_ancho = int(imagen_alto * relacion_aspecto) - 10
        else:
            imagen_ancho = max_ancho
            imagen_alto = int(imagen_ancho / relacion_aspecto) - 10
        
        image_tk = ctk.CTkImage(light_image=Image.open(buff),
                                  size=(imagen_ancho, imagen_alto))
        """
        pil_image = Image.open("temp_graph.png")
        pil_image_resized = pil_image.resize((imagen_ancho,imagen_alto))
        image_tk = ImageTk.PhotoImage(pil_image_resized)
        """
        
        label_image = ctk.CTkLabel(graph_window, image=image_tk, text="")
        label_image.pack(padx=20, pady=20)
        
        label_image.image = image_tk
    
    def show_info(self):
        CTkMessagebox(title="Sobre el Autor", message="Autor: Francisco Alex Mares Solano\nContacto: alxmares@outlook.com\n+52 2289887255",
                   icon=resource_path("check.png"))
        
    def generar_aleatorio(self):
        #print("Generación aleatoria")
        for radiovar in self.radiobuttons_frame.radiovar_list[1:64]:
            valor_aleatorio = np.random.randint(0,5)
            radiovar.set(valor_aleatorio)
        
    def reset_all(self):
        #print("Botón reiniciar")
        self.ageEntry.delete(0,'end')
        self.nameEntry.delete(0,'end')
        self.empEntry.delete(0,"end")
        self.puestoEntry.delete(0,"end")
        self.generate_graphs.configure(state="disabled")
        self.generate_document.configure(state="disabled")
        
        for radiovar in self.radiobuttons_frame.radiobutton_list:
            radiovar.deselect()
    
    # Revisar que todas las preguntas hayan sido respondidas
    def check_values(self):
        # MENSAJES DE ERROR
        if not self.nameEntry.get():
            #print("Sin nombre")
            CTkMessagebox(title="Información incompleta", message="Por favor, colocar el nombre",
                              icon=resource_path("cancel.png"))
            return False
        
        if not self.ageEntry.get():
            #print("Sin edad")
            CTkMessagebox(title="Información incompleta", message="Por favor, colocar la edad",
                              icon=resource_path("cancel.png"))
            return False
        
        try:
            edad = int(self.ageEntry.get())
            #print(edad)
        except:
            #print("Edad no válida")
            CTkMessagebox(title="Información incompleta", message="Edad no válida",
                              icon=resource_path("cancel.png"))
            return False
        
        if not self.empEntry.get():
            CTkMessagebox(title="Información incompleta", message="Por favor, colocar el nombre de la empresa",
                              icon=resource_path("cancel.png"))
            return False
        
        if not self.puestoEntry.get():
            CTkMessagebox(title="Información incompleta", message="Por favor, colocar el nombre del puesto",
                              icon=resource_path("cancel.png"))
            return False
        
        
        for radiovar in self.radiobuttons_frame.radiovar_list[:64]:
            # La variable es de tipo str si no se respondió
            if type(radiovar.get()) is str: 
                #print("Sin respuesta")
                CTkMessagebox(title="Información incompleta", message="Por favor, contestar todas las preguntas",
                              icon=resource_path("cancel.png"))
                return False
        
        return True
    
    def print_results(self, puntaje,calificacion):
        CTkMessagebox(title= "Resultados Generados Correctamente", message=f"Puntaje: {puntaje}\nCalificación: {calificacion}",
                      icon=resource_path("info.png"))
        
    def generate_results(self):
        # Revisar que todos las preguntas hayan sido seleccionadas
        if not self.check_values():
            return
        self.generate_graphs.configure(state="normal")
        self.generate_document.configure(state="normal")
        self.callback_function1(self.nameEntry.get(),int(self.ageEntry.get()), self.empEntry.get(), self.puestoEntry.get())
        #print(self.radiobuttons_frame.get_results())
    
    def start(self):
        #with open("testnom/preguntas.txt", encoding="utf8") as f:
        #    preguntas = f.readlines()  
    
        ctk.set_appearance_mode("light")
        #ctk.set_default_color_theme("blue")
        self.mainloop()

     