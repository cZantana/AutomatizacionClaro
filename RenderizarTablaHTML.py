import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget
from PyQt5.QtWebEngineWidgets import QWebEngineView
from PyQt5.QtCore import QEventLoop
import os
import cv2

class HTMLViewer(QMainWindow):
    """Ventana que muestra una tabla HTML usando QWebEngineView."""
    def __init__(self, html_content, event_loop,tabla_actual):
        super().__init__()
        self.setWindowTitle(tabla_actual)
        self.setGeometry(100, 100, 800, 600)

        # Crear el visor web
        self.browser = QWebEngineView()
        self.browser.setHtml(html_content)

        # Configurar el layout
        container = QWidget()
        layout = QVBoxLayout()
        layout.addWidget(self.browser)
        container.setLayout(layout)
        self.setCentralWidget(container)

        # Guardar referencia del event loop
        self.event_loop = event_loop

    def closeEvent(self, event):
        """Cierra la ventana y detiene el event loop."""
        self.event_loop.quit()
        event.accept()

def generar_html_tabla(tabla):
    """Genera el HTML de la tabla respetando rowspan y colspan."""
    html = "<html><body>"
    html += "<table border='1' style='border-collapse: collapse; text-align: center; width: 100%;'>\n"

    for fila in tabla:
        html += "  <tr>\n"
        for celda in fila:
            if celda["rowspan"] > 0 and celda["colspan"] > 0:
                html += f"    <td rowspan='{celda['rowspan']}' colspan='{celda['colspan']}'>{celda['contenido']}</td>\n"
        html += "  </tr>\n"

    html += "</table>"
    html += "</body></html>"
    
    return html

def mostrar_html_pyqt(tabla,tabla_actual):
    """Muestra la tabla en una ventana PyQt y espera a que el usuario la cierre antes de continuar."""
    global app  # Reutilizar una sola instancia de QApplication

    html_content = generar_html_tabla(tabla)
    event_loop = QEventLoop()  # Crear un EventLoop para pausar la ejecución
    viewer = HTMLViewer(html_content, event_loop,tabla_actual)
    viewer.show()

    # Esperar hasta que el usuario cierre la ventana antes de continuar
    event_loop.exec_()


import DetectarCentroidesDeCeldas as dcdc
import ExtraerEstructuraDeTabla as eedt

def image_to_HTML(path_image,tabla_actual):

    # Detectar celdas y obtener coordenadas directamente
    imagenes_celdas, coordenadas_celdas, centros_celdas, imagen_width, imagen_height = dcdc.detectar_celdas(path_image)

    print("\nCanitdad de celdas encontrdadas:\n",len(imagenes_celdas),"\n")

    # Generar estructura de la tabla con rowspan y colspan
    # Generar la malla
    lineas_x, lineas_y, max_filas, max_columnas, imagen_malla, umbral_x, umbral_y = eedt.generar_malla(coordenadas_celdas, imagen_width, imagen_height)

    # Crear la estructura de la cuadrícula
    cuadricula = [[{"x": x, "y": y, "w": 0, "h": 0} for x in lineas_x] for y in lineas_y]

    # Generar la estructura de la tabla con celdas unidas aplicando el umbral
    tabla_generada = eedt.generar_estructura_tabla(coordenadas_celdas, cuadricula, max_filas, max_columnas, imagen_width, imagen_height, umbral_x, umbral_y)


    # Mostrar la tabla generada
    print("tabla generada:")
    for fila in tabla_generada:
        print(fila)

    # Ejecutar la ventana con la tabla renderizada
    mostrar_html_pyqt(tabla_generada,tabla_actual)


def obtener_imagenes_con_ruta(ruta_carpeta):
    """Obtiene todas las imágenes de una carpeta y devuelve una lista de rutas completas."""
    extensiones_validas = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tiff"}
    imagenes_con_ruta = []

    if not os.path.exists(ruta_carpeta):
        print(f"Error: La carpeta '{ruta_carpeta}' no existe.")
        return []

    for archivo in os.listdir(ruta_carpeta):
        ruta_completa = os.path.join(ruta_carpeta, archivo)
        if os.path.isfile(ruta_completa) and os.path.splitext(archivo)[1].lower() in extensiones_validas:
            imagenes_con_ruta.append(ruta_completa)

    return imagenes_con_ruta

def get_images_from_path(ruta_carpeta):
    # Inicializar QApplication una sola vez al inicio
    app = QApplication(sys.argv)

    # Definir la carpeta con imágenes
    # ruta_carpeta = "imagenes_De_Prueba"
    lista_rutas_imagenes = obtener_imagenes_con_ruta(ruta_carpeta)

    print("Rutas completas de las imágenes encontradas:")
    for ruta in lista_rutas_imagenes:
        print(ruta)
        
        # Mostrar imagen con OpenCV (si es necesario, pero sin bloquear)
        # imagen = cv2.imread(ruta)
        # cv2.imshow("Imagen", imagen)
        
        # Llamar a la función para procesar la imagen y mostrar HTML
        image_to_HTML(ruta)
        
        # Cerrar la imagen de OpenCV antes de continuar con la siguiente
        cv2.destroyAllWindows()

    # Salir de la aplicación cuando termine el bucle
    app.quit()