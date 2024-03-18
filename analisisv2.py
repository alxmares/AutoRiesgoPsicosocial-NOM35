import numpy as np
import matplotlib
import matplotlib.pyplot as plt
from tkinter import filedialog
import threading
from docx import Document
from docx.shared import RGBColor, Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

# Asegurarse que no utilice backend de CTkinter
matplotlib.use('Agg')

class Analisis():
    def __init__(self, callback_function1, callback_function2):
        self.callback_function1 = callback_function1
        self.callback_function2 = callback_function2
        
        self.categorias = {
            "index": {1:[1,2,3,4,5], 2:[6,7,8,9,10,11,12,13,14,15,16,23,24,25,26,27,28,29,30,35,36,65,66,67,68],
                      3:[17,18,19,20,21,22], 4:[31, 32, 33, 34, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 57, 58, 59, 60, 61, 62, 63, 64, 69, 70, 71, 72],
                      5:[47, 48, 49, 50, 51, 52, 53, 54, 55, 56]},
            "range":{1:[5,9,11,14],2:[15,30,45,60],3:[5,7,10,13],4:[14,29,42,58],5:[10,14,18,23]}
        }
        
        self.dominios = {
            "index": {1:[1,3,2,4,5],2:[6,12,7,8,9,10,11,65,66,67,68,13,14,15,16],
                      3:[23, 24, 29,30,35,36],4:[17,18],5:[19,20,21,22],6:[31, 32, 33, 34,37, 38, 39, 40, 41],
                      7:[42, 43, 44, 45, 46, 69, 70, 71, 72],8:[57, 58, 59, 60, 61, 62, 63, 64],
                      9:[47, 48, 49, 50, 51, 52],10:[55, 56, 53, 54]},
            "range":{1:[5,9,11,14],2:[15,21,27,37],3:[11,16,21,25],
                     4:[1,2,4,6],5:[4,6,8,10],6:[9,12,16,20],7:[10,13,17,21],
                     8:[7,10,13,14,16],9:[6,10,14,18],10:[4,6,8,10]}
        }
        
        self.rango_final = [(0,50),(50,75),(75,99),(99,140),(140,300)]
        
        self.int_calificaciones = [1,2,3,4,5]
        self.calificaciones = ["Nulo", "Bajo", "Medio", "Alto", "Muy Alto"]
        self.calif_der = [1,4,23,24,25,26,27,28,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,55,56,57]
        self.calif_izq = [2,3,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,29,54,58,59,60,61,62,63,64,65,66,67,68,69,70,71,72]
        
        self.colores_por_calificacion = ['#1f77b4', '#2ca02c', '#ffd700', '#ff7f0e', '#d62728']
        
        self.necesidad_accion = {
            "Muy Alto": "Es preciso realizar el análisis de cada categoría y dominio para establecer las acciones de intervención apropiadas, mediante un Programa de intervención intensivo que deberá incluir campañas de sensibilización, aplicar la política de prevención de riesgos psicosociales e implementar programas para la prevención de los factores de riesgo psicosocial, la promoción de un entorno organizacional favorable y la prevención de la violencia laboral, así como garantizar su aplicación y difusión.",
            "Alto":"Es preciso realizar un análisis de cada categoría y dominio, de manera que se puedan determinar las acciones de intervención concretas a través de un Programa de intervención, que podrá incluir una evaluación específica y deberá incluir una campaña de sensibilización, implementar la política de prevención de riesgos psicosociales y programas para la prevención de los factores de riesgo psicosocial, la promoción de un entorno organizacional favorable y la prevención de la violencia laboral, así como garantizar su aplicación y difusión.",
            "Medio":"Se requiere reforzar la política de prevención de riesgos psicosociales y programas para la prevención de los factores de riesgo psicosocial, la promoción de un entorno organizacional favorable y la promoción de un entorno laboral saludable, así como garantizar su aplicación y difusión, mediante un Programa de intervención.",
            "Bajo":"Es necesario una mayor difusión de la política de prevención de riesgos psicosociales y programas para: la prevención de los factores de riesgo psicosocial, la promoción de un entorno organizacional favorable y la prevención de la violencia laboral.",
            "Nulo":"El riesgo resulta poco significativo por lo que no se requiere medidas adicionales.",  
        }
        self.cat_nombres = ['Ambiente de trabajo', 'Factores propios de actividad', 
                      'Organización tiempo de trabajo', 'Liderazgo relaciones en trabajo',
                      'Entorno organizacional']
        self.dom_nombres = ['Condiciones ambiente de trabajo', 'Cargo de trabajo', 'Falta de control trabajo', 
                      'Jornada de trabajo','Interferencia relación trabajo-familia','Liderazgo',
                      'Relaciones en trabajo','Violencia','Reconocimiento desempeño',
                      'Insuficiente sentido pertenecencia inestabilidad']
        
    def analizar(self, nombre, edad, empresa, puesto, respuestas):
        self.nombre = nombre
        self.edad = edad
        self.empresa = empresa
        self.puesto = puesto
        self.puntajes = respuestas
        
        if len(self.calificaciones)>5:
            self.calificaciones.pop(0)
            
        self.resultado_cat = {"punt":{1:0,2:0,3:0,4:0,5:0},
                         "cal":{1:0,2:0,3:0,4:0,5:0}}
        
        self.resultado_dom = {"punt":{1:0,2:0,3:0,4:0,5:0,6:0,7:0,8:0,9:0,10:0},
                         "cal":{1:"",2:"",3:"",4:"",5:"",6:"",7:"",8:"",9:"",10:""}}
        self.punt_final = 0
        # categoría
        for idx,respuesta in enumerate(respuestas):
            idx=idx+1 # índices comienzan en 1
            if respuesta == "":continue # De las 65 a 72 no importa si responde o no 
            
            # Buscar categoria
            for i in range(1,6):
                if idx in self.categorias["index"][i]:
                    categoria = i
            
            # Buscar dominio
            for i in range(1,11):
                if idx in self.dominios["index"][i]:
                    dominio = i 
            
            # Obtener puntuación
            if idx in self.calif_der:
                punt = respuesta
            else:
                punt = 4-respuesta
            
            # Sumar calificaciones
            self.punt_final += punt
            self.resultado_cat["punt"][categoria] += punt
            self.resultado_dom["punt"][dominio] += punt
        
        # Obtener calificación de categorias
        for i in range(1,6):
            puntuacion = self.resultado_cat["punt"][i]
            rangos = self.categorias["range"][i]
            if puntuacion < rangos[0]:
                calificacion = self.int_calificaciones[0]
            elif puntuacion < rangos[1]:
                calificacion = self.int_calificaciones[1]
            elif puntuacion < rangos[2]:
                calificacion = self.int_calificaciones[2]
            elif puntuacion < rangos[3]:
                calificacion = self.int_calificaciones[3]
            else:
                calificacion = self.int_calificaciones[4]
            
            self.resultado_cat["cal"][i]=calificacion
            
        # Obtener calificación de dominio
        for i in range(1,11):
            puntuacion = self.resultado_dom["punt"][i]
            rangos = self.dominios["range"][i]
            if puntuacion < rangos[0]:
                calificacion = self.int_calificaciones[0]
            elif puntuacion < rangos[1]:
                calificacion = self.int_calificaciones[1]
            elif puntuacion < rangos[2]:
                calificacion = self.int_calificaciones[2]
            elif puntuacion < rangos[3]:
                calificacion = self.int_calificaciones[3]
            else:
                calificacion = self.int_calificaciones[4]
            
            self.resultado_dom["cal"][i]=calificacion
            
        # Obtener calificación final
        for i,(rango,calf) in enumerate(zip(self.rango_final, self.calificaciones)):
            if rango[0] <= self.punt_final < rango[1]:
                self.cal_final = calf
                self.color_final = self.colores_por_calificacion[i]
                #print(self.cal_final)
                #print(self.color_final) 
                break
        
        self.accion = self.necesidad_accion[self.cal_final]
        #print(f'Puntaje_final: {self.punt_final}, Calificación: {self.cal_final}')
        #print(self.resultado_cat)
        #print(self.resultado_dom)
        
        return self.punt_final,self.cal_final
    
    def graficar(self):
        import matplotlib.image as mpimg
        
        nombre = self.nombre
        puntaje_final = self.punt_final
        calificacion_final = self.cal_final
        
        buffer1 = self.graficar_categorias()
        buffer2 = self.graficar_dominios()
        
        plot1 = mpimg.imread(buffer1, format="PNG")
        plot2 = mpimg.imread(buffer2, format="PNG")
        
        fig,axs = plt.subplots(1,2, figsize=(30, 15))
        
        # Configuración de la figura
        fig.suptitle(f"Nombre: {nombre}\nCalificación final: {puntaje_final} | {calificacion_final}", fontsize=16)
        
        # Mostrar cada imagen en un subplot
        axs[0].imshow(plot1)
        axs[0].axis('off')  # Opcional: quitar los ejes
        axs[1].imshow(plot2)
        axs[1].axis('off')  # Opcional: quitar los ejes
    
        # Guardar en buffer
        #plt.savefig("temp_graph.png")
        fig_buf = io.BytesIO()
        plt.savefig(fig_buf, format='png', bbox_inches='tight')
        fig_buf.seek(0)
        return fig_buf
    
    
    def graficar_categorias(self):
        # Definimos las calificaciones para cada categoría
        calificaciones_por_categoria = list(self.resultado_cat["cal"].values())
        
        # Asignamos un color a cada calificación
        colores = [self.colores_por_calificacion[calificacion-1] for calificacion in calificaciones_por_categoria]

        # Crear figura para gráfico 3D
        fig = plt.figure(figsize=(10, 10))
        ax = fig.add_subplot(111, projection='3d')

        # Configuración de las barras
        x_pos = range(len(calificaciones_por_categoria))
        y_pos = np.zeros(len(calificaciones_por_categoria))
        z_pos = np.zeros(len(calificaciones_por_categoria))
        dx = np.ones(len(calificaciones_por_categoria))  # Ancho constante para todas las barras
        dy = np.ones(len(calificaciones_por_categoria))  # Profundidad constante (no significa nada)
        dz = calificaciones_por_categoria  # Altura de las barras (calificaciones)


        # Configuración de las barras con profundidad ajustada
        dy = np.full(len(calificaciones_por_categoria), 0.6)  # Profundidad ajustada a 0.4 para todas las barras
        dx = np.full(len(calificaciones_por_categoria), 0.8)  # Profundidad ajustada a 0.4 para todas las barras

        # Crear las barras con la profundidad ajustada
        ax.bar3d(x_pos, y_pos, z_pos, dx, dy, dz, color=colores)

        # Configurar etiquetas y títulos
        ax.set_title('Calificaciones por Categoría', fontsize=15)
        ax.set_zticks(np.arange(len(self.cat_nombres)+1))

        if len(self.calificaciones) == 5:
            self.calificaciones.insert(0,"")
        ax.set_zticklabels(self.calificaciones, fontsize=13)

        # Ajustar los límites del eje y para mantener la escala hasta 1
        ax.set_ylim([-2,0.8])
        ax.set_yticks([])


        # Cambiar las etiquetas del eje x para que coincidan con las categorías
        ax.set_xticks(x_pos)
        ax.set_xticklabels([label.replace(' ', '\n') for label in self.cat_nombres],rotation=30, ha='left', fontsize=11)

        ax.view_init(elev=25, azim=-110)
        fig_buf = io.BytesIO()
        plt.savefig(fig_buf, format='png', bbox_inches='tight')
        fig_buf.seek(0)
        return fig_buf
    
    def graficar_dominios(self):
        # Definimos las calificaciones para cada dominio
        calificaciones_por_dominio = list(self.resultado_dom["cal"].values())

        # Asignamos un color a cada calificación
        colores = [self.colores_por_calificacion[calificacion - 1] for calificacion in calificaciones_por_dominio]
        
        # Crear figura para gráfico 3D
        fig = plt.figure(figsize=(30, 10))
        ax = fig.add_subplot(111, projection='3d')

        # Configuración de las barras
        x_pos = range(len(calificaciones_por_dominio))
        y_pos = np.zeros(len(calificaciones_por_dominio))
        z_pos = np.zeros(len(calificaciones_por_dominio))
        dx = np.ones(len(calificaciones_por_dominio))  # Ancho constante para todas las barras
        dy = np.ones(len(calificaciones_por_dominio))  # Profundidad constante (no significa nada)
        dz = calificaciones_por_dominio  # Altura de las barras (calificaciones)


        # Configuración de las barras con profundidad ajustada
        dy = np.full(len(calificaciones_por_dominio), 0.2)  # Profundidad ajustada a 0.4 para todas las barras
        dx = np.full(len(calificaciones_por_dominio), 0.8)  # Profundidad ajustada a 0.4 para todas las barras

        # Crear las barras con la profundidad ajustada
        ax.bar3d(x_pos, y_pos, z_pos, dx, dy, dz, color=colores)

        # Configurar etiquetas y títulos
        ax.set_title('Calificaciones por Dominio', fontsize=15)
        ax.set_zticks(np.arange(len(self.cat_nombres)+1))

        if len(self.calificaciones) == 5:
            self.calificaciones.insert(0,"")
        ax.set_zticklabels(self.calificaciones, fontsize=13)

        # Ajustar los límites del eje y para mantener la escala hasta 1
        ax.set_ylim([-1,0.5])
        ax.set_yticks([])


        # Cambiar las etiquetas del eje x para que coincidan con las categorías
        ax.set_xticks(x_pos)
        x_labels = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"]
        ax.set_xticklabels(x_labels)

        ax.view_init(elev=25, azim=-110)

        # Colocar texto de etiquetas en la parte inferior del gráfico
        etiquetas_texto = ', '.join(f"{x_labels[i]}: {nombre}" for i, nombre in enumerate(self.dom_nombres))

        # Ajustar la posición y el alineamiento del texto
        ax.text(-3.5, 0.2, -4,etiquetas_texto,wrap=True, fontsize=11)
        
        # Guardar en buffer
        fig_buf = io.BytesIO()
        plt.savefig(fig_buf, format='png', bbox_inches='tight')
        fig_buf.seek(0)
        return fig_buf
    
    def hex_to_rgb(self, hex_color):
        hex_color = hex_color.lstrip('#')
        return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
    
    def save_doc(self):
        """
        Guarda el archivo .docx. En caso de no hacerlo retorna None
        """
        # Guardar documento
        dirname = self.nombre.replace(" ", "_")
        filename = "resultados_"+dirname+".docx"
        
        opciones = {
        'defaultextension': '.docx',
        'filetypes': [('Documentos de Word', '.docx'), ('Todos los archivos', '.*')],
        'title': 'Guardar como...',
        "initialfile": filename
        }
    
        filename = filedialog.asksaveasfilename(**opciones)
        if filename:
            #print(f"Archivo guardado en {filename}")
            return filename,dirname
        
        return None, None
    
    def search_logo(self):
        filepath = filedialog.askopenfilename(
          title="Selecciona una imagen para el encabezado",
        filetypes=[("Archivos de imagen", "*.png;*.jpg;*.jpeg;*.gif")]
        )

        # Verificar si se seleccionó un archivo
        if not filepath:
            #print("No se seleccionó ningún archivo.")
            return
        
        return filepath
    
    def hilo_word(self, logo_fp):
        """
        Se crea el archivo de Word en un hilo.
        """
        filename,dirname = self.save_doc() # Obtener locación del archivo, no lo guarda
        
        # Significa que canceló la acción
        if not filename or not dirname:
            return 
        
        info = {
            "Cal": self.cal_final,
            "Nombre": self.nombre,
            "Edad": self.edad,
            "Empresa": self.empresa,
            "Puesto": self.puesto,
            "Puntaje Final": self.punt_final,
            "Calificación final": self.cal_final,
            "Necesidad de acción": self.accion
            }
        
        doc = Document()
        # Acceder a la sección actual del documento para ajustar los márgenes
        section = doc.sections[0]
        section.top_margin = Inches(1)  # Margen superior de 1 pulgada
        section.bottom_margin = Inches(1)  # Margen inferior de 1 pulgada
        section.left_margin = Inches(1)  # Margen izquierdo de 1 pulgada
        section.right_margin = Inches(1)  # Margen derecho de 1 pulgada
        
        # Establecer la fuente predeterminada del documento
        style = doc.styles['Normal']
        style.font.name = 'Calibri'
        style.font.size = Pt(11)
        
        # Añadir logotipo en caso de que se haya declarado
        if logo_fp:
            # Acceder al encabezado de la sección por defecto del documento
            section = doc.sections[0]
            header = section.header

            # Añadir un párrafo al encabezado y alinear el texto a la derecha
            paragraph = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
            paragraph.alignment = 2  # 2 es el valor para la alineación a la derecha

            # Añadir la imagen al párrafo del encabezado
            run = paragraph.add_run()
            run.add_picture(logo_fp, width=Inches(1))  # Ajusta el tamaño según sea necesario

        # = = =  H E A D I N G  = = =
        heading = doc.add_heading(level=0)
        h = heading.add_run("Reporte de Evaluación: ")
        h.font.size = Pt(22)
        h.font.name = "Calibri"
        h = heading.add_run(self.cal_final)
        r,g,b = self.hex_to_rgb(self.color_final)
        h.font.color.rgb = RGBColor(r,g,b)
        h.font.size = Pt(22)
        h.font.name = "Calibri"
        # ===========================

        # Añadir un párrafo al documento
        p = doc.add_paragraph()
        # Establecer el espaciado de línea
        p.paragraph_format.line_spacing = 1.5  # Espaciado de línea 1.5

        # Establecer el espacio después del párrafo (espacio entre líneas)
        p.paragraph_format.space_after = Pt(6)
        # Indicar la clave específica donde quieres agregar un salto de línea
        salto_de_linea = ["Nombre"]

        # Iterar sobre los elementos del diccionario y agregarlos al párrafo
        for key, value in info.items():
            if key == "Cal":
                continue  # Ignorar esta clave
            
            p.add_run(f"{key}: ").bold=True
            p.add_run(f"{value}")
            
            if key == "Puntaje Final":
                break
            # Comprobar si es la clave para insertar un salto de línea
            if key in salto_de_linea:
                # Agregar una tabulación, se asume que se quiere doble tabulación
                p.add_run('\t\t')
            else:
                p.add_run().add_break()  # Insertar salto de línea

        p = doc.add_paragraph()
        p.alignment= WD_ALIGN_PARAGRAPH.JUSTIFY
        p.add_run("Necesidad de acción: ").bold=True
        p.add_run(info["Necesidad de acción"])
        
        
        img_stream = self.graficar_categorias()
        p = doc.add_paragraph()
        p.alignment = 1
        run = p.add_run()
        run.add_picture(img_stream, width=Inches(4.8))
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        subtitulo = p.add_run("Gráfica de calificación por categoría")
        subtitulo.bold=True
        subtitulo.font.size = Pt(9)

        img_stream = self.graficar_dominios()
        p = doc.add_paragraph()
        p.alignment = 1
        run = p.add_run()
        run.add_picture(img_stream, width=Inches(5.2))
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        subtitulo = p.add_run("Gráfica de calificación por Dominio")
        subtitulo.bold=True
        subtitulo.font.size = Pt(9)

        # Añadir espacio para la firma
        p = doc.add_paragraph()
        run = p.add_run("\n\n\nFirma de conformidad:\n\n\n").bold=True

        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER  # Añade una línea para la firma
        run = p.add_run("_" * 30).bold=True

        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER  # Espacio para escribir el nombre
        run = p.add_run("Nombre, fecha y firma").bold=True
        
        # Guardar el documento
        try:
            doc.save(filename)
            self.callback_function2()
        except:
            self.callback_function1()
    
    def generar_documento(self, op):
        # Buscar logotipo
        if op:
            logo_fp = self.search_logo()
        else: 
            logo_fp = False
            
        word_thread = threading.Thread(target=self.hilo_word, args=(logo_fp,))
        word_thread.daemon = True
        word_thread.start()