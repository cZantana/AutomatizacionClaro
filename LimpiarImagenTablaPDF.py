import cv2
import numpy as np

# Cargar la imagen con el canal alfa
path_imagen = "tabla_12_4.png"
imagen = cv2.imread(path_imagen, cv2.IMREAD_UNCHANGED)

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
    cv2.imshow("Imagen con solo negro conservado", resultado)
    cv2.imwrite(f"img_limpia {path_imagen}", resultado)
    cv2.waitKey(0)
    cv2.destroyAllWindows()
else:
    print("La imagen no tiene canal alfa (transparencia)")
