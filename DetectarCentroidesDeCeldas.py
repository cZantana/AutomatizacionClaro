import cv2
import numpy as np
import VerificarTablaCerrada as vtc

def calcular_angulo(v1, v2):
    """Calcula el ángulo entre dos vectores en grados."""
    v1 = v1.flatten()
    v2 = v2.flatten()
    cos_angulo = np.clip(np.dot(v1, v2) / (np.linalg.norm(v1) * np.linalg.norm(v2)), -1.0, 1.0)
    return np.degrees(np.arccos(cos_angulo))

def sort_cropped_images(cropped_images, tolerance=4):
    """Ordena las imágenes recortadas en una sola lista basada en sus centroides."""
    cropped_images.sort(key=lambda x: (x["centroide"][1], x["centroide"][0]))  # Ordenar por Y y luego por X
    return cropped_images

def limpiar_imagen(imagen):
    """Convierte el fondo transparente a blanco y limpia la imagen."""
    if imagen.shape[2] == 4:
        bgr = imagen[:, :, :3]  # Canales de color
        alpha = imagen[:, :, 3]  # Canal alfa
        fondo_blanco = np.full_like(bgr, 255, dtype=np.uint8)
        alpha = alpha.astype(float) / 255.0
        alpha = alpha[:, :, np.newaxis]
        imagen_sin_transparencia = (bgr * alpha + fondo_blanco * (1 - alpha)).astype(np.uint8)

        # Convertir a HSV y eliminar colores no oscuros
        hsv = cv2.cvtColor(imagen_sin_transparencia, cv2.COLOR_BGR2HSV)
        lower_dark = np.array([0, 0, 0])
        upper_dark = np.array([180, 255, 115])
        mascara_negro = cv2.inRange(hsv, lower_dark, upper_dark)

        # Crear imagen blanca con solo los elementos oscuros conservados
        resultado = np.full_like(imagen_sin_transparencia, 255, dtype=np.uint8)
        resultado[mascara_negro > 0] = imagen_sin_transparencia[mascara_negro > 0]

        return resultado
    else:
        return imagen

def detectar_celdas(path_imagen):
    """Detecta las celdas en una imagen de tabla y retorna las celdas ordenadas junto con sus centroides."""
    # Cargar la imagen
    image = cv2.imread(path_imagen, cv2.IMREAD_UNCHANGED)
    if image is None:
        raise FileNotFoundError(f"No se pudo cargar la imagen en: {path_imagen}")

    # Limpiar imagen
    clean_image = limpiar_imagen(image)

    clean_image = vtc.verificar_cierre(clean_image)


    # Convertir a escala de grises y binarizar
    gray = cv2.cvtColor(clean_image, cv2.COLOR_BGR2GRAY)
    _, binary = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)

    # Encontrar contornos
    contours, _ = cv2.findContours(binary, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)

    # Filtrar contornos grandes (evitar bordes de la imagen)
    height, width = gray.shape
    image_area = height * width
    area_chica_threshold = image_area*0.0005
    area_grande_threshold = image_area*0.5
    contours_to_keep = [c for c in contours if cv2.contourArea(c) < 0.75 * image_area]

    # Lista para almacenar las celdas detectadas
    cropped_images = []

    for contour in contours_to_keep:
        area = cv2.contourArea(contour)
        if area > area_chica_threshold and area < area_grande_threshold:
            perimeter = cv2.arcLength(contour, True)
            approx = cv2.approxPolyDP(contour, 0.015 * perimeter, True)

            # clean_image_copy = clean_image.copy()

            # cv2.drawContours(clean_image_copy, [approx], -1, (0, 255, 0), 1)

            # cv2.imshow("contornos a mantener",clean_image_copy)

            # cv2.waitKey(0)

            if len(approx) == 4:  # Verificar si tiene 4 lados
                es_cuadrado = all(80 <= calcular_angulo(approx[i] - approx[(i - 1) % 4], 
                                                        approx[(i + 1) % 4] - approx[i]) <= 100 
                                  for i in range(4))

                if es_cuadrado:
                    M = cv2.moments(contour)
                    if M["m00"] != 0:
                        cX, cY = int(M["m10"] / M["m00"]), int(M["m01"] / M["m00"])
                        x, y, w, h = cv2.boundingRect(approx)

                        # Guardar información de la celda
                        cropped = clean_image[y:y+h, x:x+w]

                        cropped_images.append({
                            "imagen": cropped,
                            "coordenadas": (x, y, w, h),  # Guardar coordenadas exactas
                            "centroide": (cX, cY)  # Guardar centro de la celda
                        })

    # Ordenar celdas en una sola lista
    sorted_images = sort_cropped_images(cropped_images, tolerance=10)

    # Extraer las imágenes y los centroides en listas separadas
    imagenes_celdas = [img_data["imagen"] for img_data in sorted_images]

    # cv2.imshow(path_imagen,clean_image)

    # for i in imagenes_celdas:
    #     cv2.imshow("celda encontrada",i)
    #     cv2.waitKey(0)



    coordenadas_celdas = [img_data["coordenadas"] for img_data in sorted_images]
    coordenadas_centros = [img_data["centroide"] for img_data in sorted_images]
    imagen_height, imagen_width, _ = image.shape

        # **Dibujar las celdas detectadas en la imagen original**
    image_copy = clean_image.copy()
    
    for (x, y, w, h), (cX, cY) in zip(coordenadas_celdas, coordenadas_centros):
        cv2.rectangle(image_copy, (x, y), (x + w, y + h), (0, 255, 0), 2)  # Rectángulo verde
        cv2.circle(image_copy, (cX, cY), 5, (0, 0, 255), -1)  # Centroide rojo

    # Mostrar la imagen con celdas resaltadas
    # cv2.imshow("Celdas Detectadas", image_copy)
    # cv2.waitKey(0)
    # cv2.destroyAllWindows()

    return imagenes_celdas, coordenadas_celdas, coordenadas_centros, imagen_width, imagen_height
