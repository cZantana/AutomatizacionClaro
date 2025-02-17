import pdfplumber
import fitz  # PyMuPDF
import pikepdf
import re
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle
from matplotlib.widgets import Button
import numpy as np
import os
from tabulate import tabulate  # Para mostrar tablas en consola de forma bonita
from PIL import Image

import RenderizarTablaHTML as RtHTML

# ============================
# 1. FUNCION PARA ELIMINAR TEXTO DEL PDF
# ============================
def eliminar_texto_preciso(pdf_path, output_path, altura_pagina):
    """Genera un nuevo PDF sin texto en la parte superior."""
    pdf = pikepdf.open(pdf_path)

    for page in pdf.pages:
        if "/Contents" not in page:
            continue  # Saltar páginas sin contenido

        contenido_obj = page["/Contents"]
        if isinstance(contenido_obj, pikepdf.Array):  
            contenido_completo = b"".join(p.read_bytes() for p in contenido_obj)
        else:
            contenido_completo = contenido_obj.read_bytes()

        contenido_decodificado = contenido_completo.decode("latin1", errors="ignore")

        # Expresión regular para encontrar comandos de texto en el PDF
        patron_texto = re.compile(r"\((.*?)\)\s*Tj|\[(.*?)\]\s*TJ")
        patron_coordenadas = re.compile(r"([\d\.\-]+) ([\d\.\-]+) Td|([\d\.\-]+) ([\d\.\-]+) Tm")

        lineas_modificadas = []
        eliminar_siguiente_texto = False  

        for line in contenido_decodificado.split("\n"):
            match = patron_coordenadas.search(line)
            if match:
                try:
                    y_pos = float(match.group(2) or match.group(4))  

                    if y_pos > (altura_pagina * (0/3)):  
                        eliminar_siguiente_texto = True
                    else:
                        eliminar_siguiente_texto = False
                except ValueError:
                    pass

            if eliminar_siguiente_texto:
                line = re.sub(patron_texto, "", line)

            lineas_modificadas.append(line)

        nuevo_contenido = "\n".join(lineas_modificadas).encode("latin1")

        if isinstance(contenido_obj, pikepdf.Array):
            for obj in contenido_obj:
                obj.write(nuevo_contenido)
        else:
            contenido_obj.write(nuevo_contenido)

    pdf.save(output_path)
    print(f"PDF sin texto guardado en: {output_path}")

# ============================
# 2. FUNCION PARA RECORTAR TABLAS Y GUARDAR COMO IMAGEN DE ALTA CALIDAD
# ============================
def crop_and_save_image(original_pdf, page_number, coords, output_path,tabla_actual):
    """ Recorta la tabla y la guarda como imagen de alta calidad. """
    left, top, right, bottom = coords
    rect = fitz.Rect(left-3, top-4, right+6, bottom+4)
    
    page = original_pdf[page_number]  
    zoom_x, zoom_y = 4.0, 4.0  # Ajusta la calidad (factor de escala)
    mat = fitz.Matrix(zoom_x, zoom_y)  # Transformación para mejorar resolución

    pix = page.get_pixmap(matrix=mat, clip=rect, alpha=True)  # Renderizar imagen con transparencia
    img = Image.frombytes("RGBA", [pix.width, pix.height], pix.samples)

    img.save(output_path, format="PNG", dpi=(300, 300))  # Guardar con alta resolución
    print(f"Imagen de tabla guardada en: {output_path}")

    RtHTML.image_to_HTML(output_path,tabla_actual)

# ============================
# 3. FUNCION PRINCIPAL: OBTENER TABLAS USANDO EL PDF ORIGINAL
# ============================
def show_pdfplumber_tables_with_buttons(pdf_path):
    """Muestra las tablas y las guarda como imágenes recortadas desde el PDF sin texto."""
    # Crear PDF sin texto solo para recortar las tablas
    pdf_sin_texto_path = "pdf_temporal_sin_texto.pdf"
    eliminar_texto_preciso(pdf_path, pdf_sin_texto_path, 800)  

    # Abrir ambos PDFs
    pdf_original = pdfplumber.open(pdf_path)  # Se usa para visualizar e imprimir texto
    pdf_sin_texto = fitz.open(pdf_sin_texto_path)  # Se usa solo para recortes
    total_pages = len(pdf_original.pages)
    current_page_idx = 0  

    output_folder = f"tablas_recortadas {pdf_path[:12]}"
    os.makedirs(output_folder, exist_ok=True)  

    fig, ax = plt.subplots(figsize=(11, 8))
    plt.subplots_adjust(bottom=0.15)
    ax.axis("off")

    def display_page(page_idx):
        ax.clear()
        ax.axis("off")

        page = pdf_original.pages[page_idx]  
        page_image = page.to_image(resolution=72)
        pil_img = page_image.original  
        img_array = np.array(pil_img)
        ax.imshow(img_array)

        tables = page.find_tables()  

        print(f"\n=== Página {page_idx + 1} ===")
        if not tables:
            print("  No se han encontrado tablas en esta página.")
        else:
            for table_idx, table in enumerate(tables):
                x0, top, x1, bottom = table.bbox
                print(f"  - Tabla {table_idx + 1} | BBox = ({x0:.2f}, {top:.2f}) - ({x1:.2f}, {bottom:.2f})")

                data = table.extract()
                if not data:
                    print("    (Tabla vacía o no se pudo extraer contenido)")
                else:
                    headers = data[0]
                    rows = data[1:] if len(data) > 1 else []
                    tabla_formateada = tabulate(rows, headers=headers, tablefmt="fancy_grid")
                    print("    Contenido de la tabla:\n", tabla_formateada)

                words = page.extract_words()
                
                print("\n    Texto dentro de la tabla por celda:")

                celda_texto_centros = []  # Lista para almacenar datos de las celdas con coordenadas promedio

                for row_idx, row in enumerate(table.rows):  
                    for col_idx, cell in enumerate(row.cells):
                        if cell is None:
                            continue
                        
                        # Extraer coordenadas de la celda
                        cell_x0, cell_top, cell_x1, cell_bottom = cell  

                        # Filtrar palabras dentro de esta celda
                        words_in_cell = [
                            word for word in words
                            if cell_x0 <= float(word['x0']) <= cell_x1 and cell_top <= float(word['top']) <= cell_bottom
                        ]

                        print(f"      Celda ({row_idx}, {col_idx}) | BBox: ({cell_x0:.2f}, {cell_top:.2f}) - ({cell_x1:.2f}, {cell_bottom:.2f})")
                        
                        if words_in_cell:
                            # Calcular promedios de coordenadas
                            avg_x0 = sum(float(word['x0']) for word in words_in_cell) / len(words_in_cell)
                            avg_y0 = sum(float(word['top']) for word in words_in_cell) / len(words_in_cell)
                            avg_x1 = sum(float(word['x1']) for word in words_in_cell) / len(words_in_cell)
                            avg_y1 = sum(float(word['bottom']) for word in words_in_cell) / len(words_in_cell)

                            # Obtener el centro de la celda basado en las coordenadas promedio
                            centro_x = (avg_x0 + avg_x1) / 2
                            centro_y = (avg_y0 + avg_y1) / 2

                            print(f"        - Centro promedio del texto: ({centro_x:.2f}, {centro_y:.2f})")

                            for word in words_in_cell:
                                print(f"        - '{word['text']}' en ({word['x0']}, {word['top']}) - ({word['x1']}, {word['bottom']})")

                            # Guardar información de la celda con su centroide promedio
                            celda_texto_centros.append({
                                "fila": row_idx,
                                "columna": col_idx,
                                "bbox": (cell_x0, cell_top, cell_x1, cell_bottom),
                                "centro_texto": (centro_x, centro_y),
                                "contenido": " ".join(word["text"] for word in words_in_cell)
                            })

                        else:
                            print("        (Celda vacía)")

                output_img_path = "imagenTemporal.png"
                tabla_actual = os.path.join(output_folder, f"tabla_{page_idx + 1}_{table_idx + 1}.png")
                crop_and_save_image(pdf_sin_texto, page_idx, (x0, top, x1, bottom), output_img_path,tabla_actual)



        for t in tables:
            x0, top, x1, bottom = t.bbox
            rect_w, rect_h = x1 - x0, bottom - top
            rect = Rectangle((x0, top), rect_w, rect_h, edgecolor="red", facecolor="none", linewidth=2)
            ax.add_patch(rect)

        ax.set_xlim([0, page.width])
        ax.set_ylim([page.height, 0])
        ax.set_title(f"Página {page_idx + 1} / {total_pages}")
        plt.draw()

    def next_page(event):
        nonlocal current_page_idx
        if current_page_idx < total_pages - 1:
            current_page_idx += 1
            display_page(current_page_idx)

    def prev_page(event):
        nonlocal current_page_idx
        if current_page_idx > 0:
            current_page_idx -= 1
            display_page(current_page_idx)

    ax_prev = plt.axes([0.3, 0.02, 0.15, 0.07])
    ax_next = plt.axes([0.55, 0.02, 0.15, 0.07])
    btn_prev = Button(ax_prev, 'Anterior')
    btn_next = Button(ax_next, 'Siguiente')

    btn_prev.on_clicked(prev_page)
    btn_next.on_clicked(next_page)

    display_page(current_page_idx)
    plt.show()
    pdf_original.close()
    pdf_sin_texto.close()

if __name__ == "__main__":
    ruta_pdf = "PTAR 5071 Tarifa Esp_NuevosCamposdeJuego_NTC_TC_V21_0225.pdf"
    show_pdfplumber_tables_with_buttons(ruta_pdf)
