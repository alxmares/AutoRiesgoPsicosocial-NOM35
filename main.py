from window import App
from analisisv2 import Analisis
import os
import sys

class MainApplication():
    def __init__(self, data):
        self.analizador = Analisis(self.error_guardar, self.saved_doc)
        self.app = App(data, callback_function1=self.generate_results, callback_function2=self.show_graphs,
                       callback_function3=self.generate_doc)
        self.data = data
        self.name = ""
    
    def generate_results(self,nombre, edad,empresa,puesto):
        resultados = self.app.radiobuttons_frame.get_results()
        #rand_list = np.random.randint(0,5,72)
        puntaje, calificacion = self.analizador.analizar(nombre, edad, empresa, puesto,resultados)
        #analizador.graficar_categorias()
        #analizador.graficar_dominio()
        self.app.print_results(puntaje, calificacion)
        return
    
    def error_guardar(self):
        self.app.error_doc()
        
    def saved_doc(self):
        #print("Archivo guardado")
        self.app.saved_doc()
        
    def show_graphs(self):
        buff = self.analizador.graficar()
        #print("Mostrar gráficas")
        return buff
    
    def generate_doc(self, op):
        self.analizador.generar_documento(op)
        
    def start_window(self):
        self.app.start()
    

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

if __name__ == '__main__':
    with open(resource_path("preguntas.txt"), encoding="utf8") as f:
        preguntas = f.readlines()
    
    app = MainApplication(preguntas)
    app.start_window()
    