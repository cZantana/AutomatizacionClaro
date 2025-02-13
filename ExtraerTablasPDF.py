import pdfplumber
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle
from matplotlib.widgets import Button
import numpy as np
from PIL import Image
from tabulate import tabulate  # Para mostrar tablas en consola de forma bonita

def show_pdfplumber_tables_with_buttons(pdf_path):
    pdf = pdfplumber.open(pdf_path)
    total_pages = len(pdf.pages)
    current_page_idx = 0  # índice de página (0-based)

    fig, ax = plt.subplots(figsize=(11, 8))
    plt.subplots_adjust(bottom=0.15)
    ax.axis("off")

    def display_page(page_idx):
        ax.clear()
        ax.axis("off")

        page = pdf.pages[page_idx]

        # Convertimos la página a imagen con resolución = 72
        page_image = page.to_image(resolution=72)

        # En algunas versiones antiguas de pdfplumber, el atributo
        # para obtener la imagen PIL puede ser distinto:
        # page_image.original, page_image.image, o page_image.to_pil().
        # Ajusta según tu versión. Intentemos .original primero:
        pil_img = page_image.original  
        # Si falla, prueba:
        # pil_img = page_image.image
        # pil_img = page_image.to_pil()

        # Convertimos a array de NumPy para imshow
        img_array = np.array(pil_img)
        ax.imshow(img_array)

        # Detectar tablas con find_tables()
        tables = page.find_tables()

        # Imprimir info en consola:
        print(f"\n=== Página {page_idx + 1} ===")
        if not tables:
            print("  No se han encontrado tablas en esta página.")
        else:
            for table_idx, table in enumerate(tables):
                x0, top, x1, bottom = table.bbox
                print(f"  - Tabla {table_idx + 1} | BBox = ({x0:.2f}, {top:.2f}) - ({x1:.2f}, {bottom:.2f})")

                data = table.extract()  # lista de filas -> lista de celdas

                if not data:
                    print("    (Tabla vacía o no se pudo extraer contenido)")
                else:
                    # Primera fila como headers (si corresponde)
                    headers = data[0]
                    rows = data[1:] if len(data) > 1 else []

                    print("    Contenido de la tabla (formateado con tabulate):")
                    tabla_formateada = tabulate(rows, headers=headers, tablefmt="fancy_grid")
                    print(tabla_formateada)

                # Esperamos Enter para continuar (opcional)
                # input("\nPresiona Enter para continuar...")

        # Ahora dibujamos los recuadros en la imagen
        for t in tables:
            x0, top, x1, bottom = t.bbox
            rect_w = x1 - x0
            rect_h = bottom - top

            rect = Rectangle(
                (x0, top),
                rect_w,
                rect_h,
                edgecolor=(1, 0, 0),      # Rojo
                facecolor=(1, 0, 0, 0.2), # Semitransparente
                linewidth=2
            )
            ax.add_patch(rect)

        # Ajustamos límites
        ax.set_xlim([0, page.width])
        ax.set_ylim([page.height, 0])  # invertimos el eje Y
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

    # Botones
    ax_prev = plt.axes([0.3, 0.02, 0.15, 0.07])
    ax_next = plt.axes([0.55, 0.02, 0.15, 0.07])
    btn_prev = Button(ax_prev, 'Anterior')
    btn_next = Button(ax_next, 'Siguiente')

    btn_prev.on_clicked(prev_page)
    btn_next.on_clicked(next_page)

    # Mostramos la primera página
    display_page(current_page_idx)
    plt.show()
    pdf.close()

if __name__ == "__main__":
    ruta_pdf = "PTAR 5071 Tarifa Esp_NuevosCamposdeJuego_NTC_TC_V21_0225.pdf"
    show_pdfplumber_tables_with_buttons(ruta_pdf)