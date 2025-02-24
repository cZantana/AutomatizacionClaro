# Extraer Tablas de PDF en formato HTML

Este proyecto tiene como objetivo extraer tablas de archivos PDF, eliminando el texto y guardando las tablas como imágenes de alta calidad, y generando una estructura HTML de las celdas detectadas.

## Estructura del Proyecto

.
├── [ExtraerTablasSinTextoPDF.py](./ExtraerTablasSinTextoPDF.py)
├── [DetectarCentroidesDeCeldas.py](./DetectarCentroidesDeCeldas.py)
├── [VerificarTablaCerrada.py](./VerificarTablaCerrada.py)
├── [DibujarContornosCuadrados.py](./DibujarContornosCuadrados.py)
├── [ExtraerEstructuraDeTabla.py](./ExtraerEstructuraDeTabla.py)
└── [RenderizarTablaHTML.py](./RenderizarTablaHTML.py)


## Descripción de los Archivos

- **[ExtraerTablasSinTextoPDF.py](./ExtraerTablasSinTextoPDF.py)**: Script principal que coordina la eliminación de texto del PDF, la detección de tablas y la generación de imágenes y HTML.
- **[DetectarCentroidesDeCeldas.py](./DetectarCentroidesDeCeldas.py)**: Detecta las celdas en una imagen de tabla y retorna las celdas ordenadas junto con sus centroides.
- **[VerificarTablaCerrada.py](./VerificarTablaCerrada.py)**: Verifica si la tabla está cerrada detectando bordes oscuros en la imagen.
- **[DibujarContornosCuadrados.py](./DibujarContornosCuadrados.py)**: Dibuja contornos cuadrados en las celdas detectadas y elimina vértices alineados.
- **[ExtraerEstructuraDeTabla.py](./ExtraerEstructuraDeTabla.py)**: Genera la estructura de la tabla con rowspan y colspan, identificando las celdas fusionadas.
- **[RenderizarTablaHTML.py](./RenderizarTablaHTML.py)**: Genera el HTML de la tabla respetando rowspan y colspan y muestra la tabla en una ventana PyQt.

## Requisitos

- Python 3.x
- Librerías: `pdfplumber`, `fitz` (PyMuPDF), `pikepdf`, `re`, `matplotlib`, `numpy`, `PIL` (Pillow), `colorama`, `tabulate`, `PyQt5`, `PyQtWebEngine`, `opencv-python`

Puedes instalar las librerías necesarias usando pip:

```sh
pip install pdfplumber pymupdf pikepdf matplotlib numpy pillow colorama tabulate PyQt5 opencv-python
```

## Uso

1. Coloca el archivo PDF en el mismo directorio que los scripts.
2. Cambia la **ruta_pdf** en la funcion **main** y ejecuta el script principal `ExtraerTablasSinTextoPDF.py`:

```sh
python ExtraerTablasSinTextoPDF.py
```

El script eliminará el texto del PDF, detectará las tablas, generará imágenes de estas, generara la estructura HTML de las tablas detectadas, y asignara el texto correctamente a cada celda.

## Funciones Principales

### ExtraerTablasSinTextoPDF.py

- `eliminar_texto_preciso(pdf_path, output_path, altura_pagina)`: Genera un nuevo PDF sin texto en la parte superior.

- `crop_and_save_image(original_pdf, page_number, coords, output_path, tabla_actual, lista_tablas)`: Recorta la tabla y la guarda como imagen de alta calidad.

- `convertir_coordenadas_imagen_a_pdf(coordenadas_celdas, x0, top, x1, bottom, dimensiones_tabla, left_margin, top_margin)`: Convierte las coordenadas de la imagen a coordenadas del PDF.

- `asignar_texto_a_estructura(tabla_estructura, coordenadas_celdas_convertidas, textos_pdf)`: Asigna el contenido a cada celda en la estructura de la tabla.

- `show_pdfplumber_tables_with_buttons(pdf_path)`: Muestra las tablas y las guarda como imágenes recortadas desde el PDF sin texto.

### DetectarCentroidesDeCeldas.py

- `detectar_celdas(path_imagen)`: Detecta las celdas en una imagen de tabla y retorna las celdas ordenadas junto con sus centroides.

### DibujarContornosCuadrados.py

- `eliminar_vertices_alineados(contour, threshold_angle, min_distance)`: Elimina vértices innecesarios de un contorno si están en una línea recta.

- `cargar_imagen(image)`: Carga una imagen, encuentra y filtra los contornos.

### ExtraerEstructuraDeTabla.py

- `generar_malla(coordenadas_celdas, imagen_width, imagen_height)`: Genera una malla basada en las coordenadas de las celdas y la dibuja en una imagen en blanco.

- `generar_estructura_tabla(coordenadas_celdas, cuadricula, max_filas, max_columnas, imagen_width, imagen_height, umbral_x, umbral_y)`: Genera la estructura de la tabla con rowspan y colspan.

### RenderizarTablaHTML.py

- `generar_html_tabla(tabla)`: Genera el HTML de la tabla respetando rowspan y colspan.

- `mostrar_html_pyqt(tabla, tabla_actual)`: Muestra la tabla en una ventana PyQt y espera a que el usuario la cierre antes de continuar.

### VerificarTablaCerrada.py

- `verificar_cierre(image)`: Aplica la detección de bordes oscuros en cada lado de la imagen.

## Contribuciones

Las contribuciones son bienvenidas. Por favor, abre un issue o envía un pull request.

## Licencia

Este proyecto está bajo la Licencia MIT.
