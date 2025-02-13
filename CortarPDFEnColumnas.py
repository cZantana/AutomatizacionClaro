import pdfplumber
import matplotlib.pyplot as plt
from matplotlib.widgets import RectangleSelector, Button
import fitz  # PyMuPDF
import re

# Ruta del PDF
pdf_path = "PCAM 1006 ENERO.pdf"  # Reemplaza esto con la ruta a tu archivo PDF

# Variables para almacenar las coordenadas
rectangles = {
    'Encabezado': {'coords': None, 'color': 'r'},
    'Pie de página': {'coords': None, 'color': 'g'},
    'Columna izquierda': {'coords': None, 'color': 'b'},
    'Columna derecha': {'coords': None, 'color': 'm'}
}

# Variables para la navegación entre páginas
current_page_index = 0
current_selector_key = 'Encabezado'

def detect_footer_and_page_number(pdf):
    first_page = pdf.pages[0]
    text = first_page.extract_text(x_tolerance=3, y_tolerance=3)
    footer_phrase = "Clasificación: Uso Interno. Documento Claro Colombia"
    words = footer_phrase.split()
    found_words = []
    footer_y_top = None
    words_data = first_page.extract_words()
    
    if text:
        for word_data in words_data:
            word = word_data['text']
            if word == words[len(found_words)]:  # Verifica la secuencia de palabras
                found_words.append(word_data)
                if len(found_words) == len(words):  # Si encontró todas en orden
                    footer_y_top = min(w['top'] for w in found_words)
                    break
            else:
                found_words = []  # Reiniciar si se rompe la secuencia
    
    # Buscar palabra inmediatamente arriba del pie de página
    closest_word = None
    closest_y = None
    for word_data in words_data:
        if footer_y_top and word_data['top'] < footer_y_top:
            if closest_y is None or word_data['top'] > closest_y:
                closest_y = word_data['top']
                closest_word = word_data
    
    # Validar si la palabra encontrada es un número o "Página # de #"
    page_number_y = footer_y_top
    if closest_word:
        leftmost = True  # Verificar si es la primera palabra en su línea
        for word_data in words_data:
            if word_data['top'] == closest_word['top'] and word_data['x0'] < closest_word['x0']:
                leftmost = False
                break
        
        if leftmost:
            if re.match(r'^\d+$', closest_word['text']):  # Si es un número solo
                page_number_y = closest_word['top']
            else:
                possible_page_label = "".join(
                    w['text'] for w in words_data 
                    if w['top'] == closest_word['top'] and w['x0'] >= closest_word['x0']
                )
                if re.match(r'^(P|p)ágina \d+ de \d+$', possible_page_label):
                    page_number_y = closest_word['top']
    
    return (0, page_number_y - 5, first_page.width, first_page.height)


# Función para recortar y agregar la región de una página al nuevo PDF
def crop_and_add_to_pdf(doc, original_pdf, page_number, coords):
    left, top, right, bottom = coords
    rect = fitz.Rect(left, top, right, bottom)
    page = original_pdf[page_number]
    new_page = doc.new_page(width=rect.width, height=rect.height)
    new_page.show_pdf_page(new_page.rect, original_pdf, page_number, clip=rect)

# Función para dibujar los recuadros en la página
def draw_rectangles(ax):
    for key, value in rectangles.items():
        coords = value['coords']
        color = value['color']
        if coords:
            rect = plt.Rectangle((coords[0], coords[1]), coords[2] - coords[0], coords[3] - coords[1],
                                 linewidth=2, edgecolor=color, facecolor='none')
            ax.add_patch(rect)

# Función para mostrar la página actual
def show_page():
    global current_page_index
    ax.clear()
    page = pdf.pages[current_page_index]
    im = page.to_image()
    ax.imshow(im.original)
    draw_rectangles(ax)
    fig.canvas.draw()

# Función para manejar la selección de recuadros
def onselect(eclick, erelease):
    global current_selector_key

    x1, y1 = eclick.xdata, eclick.ydata
    x2, y2 = erelease.xdata, erelease.ydata
    coords = (min(x1, x2), min(y1, y2), max(x1, x2), max(y1, y2))

    rectangles[current_selector_key]['coords'] = coords
    print(f"{current_selector_key} definido en: {coords}")
    show_page()

# Cambiar el recuadro a editar
def switch_rectangle(event):
    global current_selector_key
    keys = list(rectangles.keys())
    current_index = keys.index(current_selector_key)
    current_selector_key = keys[(current_index + 1) % len(keys)]
    print(f"Ahora editando: {current_selector_key}")

# Funciones para la navegación de páginas
def next_page(event):
    global current_page_index
    if current_page_index < len(pdf.pages) - 1:
        current_page_index += 1
        show_page()

def prev_page(event):
    global current_page_index
    if current_page_index > 0:
        current_page_index -= 1
        show_page()

# Función para procesar el PDF final
def confirm_and_process(event):
    if all(rect['coords'] for rect in rectangles.values()):
        process_pdf()

# Procesamiento del PDF
def process_pdf():
    new_doc = fitz.open()

    with fitz.open(pdf_path) as original_pdf:
        crop_and_add_to_pdf(new_doc, original_pdf, 0, rectangles['Encabezado']['coords'])

        for page_number in range(len(original_pdf)):
            crop_and_add_to_pdf(new_doc, original_pdf, page_number, rectangles['Columna izquierda']['coords'])
            crop_and_add_to_pdf(new_doc, original_pdf, page_number, rectangles['Columna derecha']['coords'])
            crop_and_add_to_pdf(new_doc, original_pdf, page_number, rectangles['Pie de página']['coords'])

    new_doc.save("nuevo_documento.pdf")
    print("PDF generado exitosamente como 'nuevo_documento.pdf'")

# Mostrar la primera página para definir áreas
with pdfplumber.open(pdf_path) as pdf:
    # Detectar pie de página solo una vez
    rectangles['Pie de página']['coords'] = detect_footer_and_page_number(pdf)
    
    fig, ax = plt.subplots()

    # Botones de navegación
    axprev = plt.axes([0.1, 0.9, 0.1, 0.05])
    axnext = plt.axes([0.21, 0.9, 0.1, 0.05])
    axconfirm = plt.axes([0.75, 0.9, 0.15, 0.05])
    axswitch = plt.axes([0.45, 0.9, 0.15, 0.05])

    bprev = Button(axprev, 'Anterior')
    bprev.on_clicked(prev_page)

    bnext = Button(axnext, 'Siguiente')
    bnext.on_clicked(next_page)

    bconfirm = Button(axconfirm, 'Confirmar')
    bconfirm.on_clicked(confirm_and_process)

    bswitch = Button(axswitch, 'Cambiar Área')
    bswitch.on_clicked(switch_rectangle)

    toggle_selector = RectangleSelector(
        ax, onselect, useblit=True,
        button=[1],  # Botón izquierdo del mouse
        minspanx=5, minspany=5, spancoords='pixels', interactive=True
    )

    show_page()
    plt.show()
