import fitz  # PyMuPDF
from PIL import Image
import numpy as np
import os
import io

# Ruta al archivo PDF
pdf_path = "PTAR 5071 Tarifa Esp_NuevosCamposdeJuego_NTC_TC_V21_0225.pdf"

# Carpeta de salida para las imágenes extraídas
output_folder = "imagenes_extraidas2"
os.makedirs(output_folder, exist_ok=True)

# Abrir el documento PDF
pdf_document = fitz.open(pdf_path)

# Iterar sobre cada página
for page_number in range(len(pdf_document)):
    page = pdf_document.load_page(page_number)
    image_list = page.get_images(full=True)

    # Iterar sobre cada imagen de la página
    for image_index, img in enumerate(image_list, start=1):
        xref = img[0]
        base_image = pdf_document.extract_image(xref)
        image_bytes = base_image["image"]

        # Verificar si la imagen tiene una máscara de transparencia
        smask = base_image.get("smask")
        if smask:
            # Extraer la máscara de transparencia
            mask_image = pdf_document.extract_image(smask)
            mask_bytes = mask_image["image"]

            # Cargar la imagen base y la máscara en PIL
            with Image.open(io.BytesIO(image_bytes)) as base_img:
                with Image.open(io.BytesIO(mask_bytes)) as mask_img:
                    # Asegurarse de que la máscara esté en modo 'L' (escala de grises)
                    if mask_img.mode != 'L':
                        mask_img = mask_img.convert('L')

                    # Combinar la imagen base y la máscara para conservar la transparencia
                    base_img.putalpha(mask_img)

                    # Guardar la imagen resultante con transparencia
                    image_filename = f"pagina_{page_number + 1}_imagen_{image_index}.png"
                    image_path = os.path.join(output_folder, image_filename)
                    base_img.save(image_path)
                    print(f"Imagen guardada en: {image_path}")
        else:
            # Si no hay máscara, guardar la imagen base directamente
            with Image.open(io.BytesIO(image_bytes)) as base_img:
                image_filename = f"pagina_{page_number + 1}_imagen_{image_index}.png"
                image_path = os.path.join(output_folder, image_filename)
                base_img.save(image_path)
                print(f"Imagen guardada en: {image_path}")
