import pikepdf
import re

pdf_path = "PTAR 1006 Tarifa especial Todo Claro DTH_WTTH_V110_0225.pdf"
output_path = "documento_modificado.pdf"

def eliminar_texto_preciso(pdf_path, output_path, altura_pagina):
    pdf = pikepdf.open(pdf_path)

    for page in pdf.pages:
        if "/Contents" not in page:
            continue  # Saltar páginas sin contenido

        # Extraer el contenido del flujo de datos
        contenido_obj = page["/Contents"]
        if isinstance(contenido_obj, pikepdf.Array):  
            contenido_completo = b"".join(p.read_bytes() for p in contenido_obj)
        else:
            contenido_completo = contenido_obj.read_bytes()

        contenido_decodificado = contenido_completo.decode("latin1", errors="ignore")

        # Expresión regular para encontrar comandos de texto en el PDF
        patron_texto = re.compile(r"\((.*?)\)\s*Tj|\[(.*?)\]\s*TJ")

        # Expresión regular para encontrar coordenadas de texto (Tm o Td)
        patron_coordenadas = re.compile(r"([\d\.\-]+) ([\d\.\-]+) Td|([\d\.\-]+) ([\d\.\-]+) Tm")

        lineas_modificadas = []
        eliminar_siguiente_texto = False  # Bandera para determinar si se debe eliminar la próxima cadena de texto

        for line in contenido_decodificado.split("\n"):
            match = patron_coordenadas.search(line)
            if match:
                try:
                    y_pos = float(match.group(2) or match.group(4))  # Obtener coordenada Y

                    if y_pos > (altura_pagina * (0/3)):  # Si está en el primer tercio, marcar para eliminar
                        eliminar_siguiente_texto = True
                    else:
                        eliminar_siguiente_texto = False

                except ValueError:
                    pass  # Evita errores si la línea no tiene valores numéricos

            if eliminar_siguiente_texto:
                # Buscar y eliminar solo el contenido de texto en esta línea, sin eliminar la estructura
                line = re.sub(patron_texto, "", line)

            lineas_modificadas.append(line)

        nuevo_contenido = "\n".join(lineas_modificadas).encode("latin1")

        # Reemplazar el contenido en el PDF
        if isinstance(contenido_obj, pikepdf.Array):
            for obj in contenido_obj:
                obj.write(nuevo_contenido)  # Modificar el contenido del flujo
        else:
            contenido_obj.write(nuevo_contenido)

    pdf.save(output_path)
    print(f"PDF modificado guardado en: {output_path}")

eliminar_texto_preciso(pdf_path, output_path, 800)  # Ajusta la altura de la página
