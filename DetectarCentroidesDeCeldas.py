import cv2
import numpy as np

# Función para calcular el ángulo entre dos vectores
def calcular_angulo(v1, v2):
    v1 = v1.flatten()
    v2 = v2.flatten()
    cos_angulo = np.clip(np.dot(v1, v2) / (np.linalg.norm(v1) * np.linalg.norm(v2)), -1.0, 1.0)
    return np.degrees(np.arccos(cos_angulo))

# Función para ordenar las imágenes recortadas por filas y columnas
def sort_cropped_images_by_grid(cropped_images, tolerance=4):
    cropped_images.sort(key=lambda x: x["centroide"][1])
    sorted_images = []
    current_row = []
    current_y = cropped_images[0]["centroide"][1]

    for img_data in cropped_images:
        cX, cY = img_data["centroide"]
        if abs(cY - current_y) <= tolerance:
            current_row.append(img_data)
        else:
            current_row.sort(key=lambda x: x["centroide"][0])
            sorted_images.append(current_row)
            current_row = [img_data]
            current_y = cY

    if current_row:
        current_row.sort(key=lambda x: x["centroide"][0])
        sorted_images.append(current_row)

    return sorted_images


def blaquear_imagen(imagen):
    # Verificar si la imagen tiene 4 canales (RGBA)
    if imagen.shape[2] == 4:
        # Separar canales BGR y Alfa
        bgr = imagen[:, :, :3]  # Canales de color
        alpha = imagen[:, :, 3]  # Canal alfa

        # Crear un fondo blanco para eliminar la transparencia
        fondo_blanco = np.full_like(bgr, 255, dtype=np.uint8)

        # Normalizar el canal alfa para manejar la transparencia
        alpha = alpha.astype(float) / 255.0
        alpha = alpha[:, :, np.newaxis]  # Convertir a 3D para operar con BGR

        # Combinar la imagen con el fondo blanco respetando la transparencia
        imagen_sin_transparencia = (bgr * alpha + fondo_blanco * (1 - alpha)).astype(np.uint8)

        # Convertir la imagen de BGR a HSV
        hsv = cv2.cvtColor(imagen_sin_transparencia, cv2.COLOR_BGR2HSV)

        # Definir rangos para detectar tonos oscuros (sin incluir el rojo brillante)
        # Valores aproximados para tonos oscuros:
        lower_dark = np.array([0, 0, 0])  # Mínimo tono, saturación y valor
        upper_dark = np.array([180, 255, 145])  # Máximo tono, saturación y valor

        # Crear máscara para conservar solo los colores oscuros
        mascara_negro = cv2.inRange(hsv, lower_dark, upper_dark)

        # Crear una imagen completamente blanca
        resultado = np.full_like(imagen_sin_transparencia, 255, dtype=np.uint8)

        # Aplicar la máscara: conservar solo los colores oscuros en la imagen original
        resultado[mascara_negro > 0] = imagen_sin_transparencia[mascara_negro > 0]

        # Mostrar y guardar la imagen final
        # cv2.imshow("Imagen con solo negro conservado", resultado)
        # cv2.imwrite(f"img_limpia {path_imagen}", resultado)
        # cv2.waitKey(0)
        # cv2.destroyAllWindows()

        return resultado
    else:
        print("La imagen no tiene canal alfa (transparencia)")


# Cargar la imagen con el canal alfa
path_imagen = "tabla_7_5.png"

image = cv2.imread(path_imagen, cv2.IMREAD_UNCHANGED)


# Verificar si la imagen se cargó correctamente
if image is None:
    print("Error: No se pudo cargar la imagen. Verifique la ruta y el archivo.")
    exit()

# Cargar imagen
clean_image = blaquear_imagen(image)
cv2.imshow("Limpia",clean_image)
# Convertir a escala de grises y binarizar
gray = cv2.cvtColor(clean_image, cv2.COLOR_BGR2GRAY)
_, binary = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)

# Encontrar contornos
contours, _ = cv2.findContours(binary, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)

# Filtrar contornos grandes (evitar bordes de la imagen)
height, width = gray.shape
image_area = height * width
area_threshold = 500
contours_to_keep = [c for c in contours if cv2.contourArea(c) < 0.75 * image_area]

# Lista para almacenar las celdas detectadas
cropped_images = []

for contour in contours_to_keep:
    area = cv2.contourArea(contour)
    if area > area_threshold:
        perimeter = cv2.arcLength(contour, True)
        approx = cv2.approxPolyDP(contour, 0.015 * perimeter, True)

        if len(approx) == 4:  # Verificar si tiene 4 lados
            es_cuadrado = all(80 <= calcular_angulo(approx[i] - approx[(i - 1) % 4], 
                                                    approx[(i + 1) % 4] - approx[i]) <= 100 
                              for i in range(4))

            if es_cuadrado:
                M = cv2.moments(contour)
                if M["m00"] != 0:
                    cX, cY = int(M["m10"] / M["m00"]), int(M["m01"] / M["m00"])
                    x, y, w, h = cv2.boundingRect(approx)
                    
                    # Dibujar el rectángulo en la imagen original
                    cv2.rectangle(image, (x, y), (x + w, y + h), (0, 255, 0), 2)
                    
                    # Dibujar el centroide
                    cv2.circle(image, (cX, cY), 5, (0, 0, 255), -1)

                    # Guardar información de la celda
                    cropped = clean_image[y:y+h, x:x+w]
                    cropped_images.append({"imagen": cropped, "centroide": (cX, cY)})

# Ordenar celdas por posición en la cuadrícula
sorted_images = sort_cropped_images_by_grid(cropped_images, tolerance=10)

# Mostrar las imágenes ordenadas
# for row in sorted_images:
#     for img_data in row:
#         cv2.imshow("Celda detectada", img_data["imagen"])
#         cv2.waitKey(0)

# Mostrar la imagen con celdas detectadas y sus centros
cv2.imshow("Celdas detectadas con centros", image)
# cv2.imwrite("celdas_detectadas.png", clean_image)  # Guardar resultado
cv2.waitKey(0)
cv2.destroyAllWindows()
