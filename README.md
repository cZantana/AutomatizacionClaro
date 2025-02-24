# Listado Automático de Claro Hogar

## Descripción

Este proyecto automatiza la generación de un listado de documentos de Claro Hogar correspondientes al mes en curso. El listado incluirá los siguientes campos:

- **Nombre Archivo**: Nombre del archivo PDF descargado.
- **Código Archivo**: Código de la política obtenido dentro del archivo.
- **Link de descarga**: Enlace para descargar el archivo PDF.
- **Link Conectados**: Enlace a la página de Conectados de Claro referente al archivo.
- **Vigencia**: Período de validez de la política obtenida dentro del archivo.
- **Versión**: Versión de la política.

## Requisitos previos

- **Python 3.x**
- Librerías necesarias:
  - `pandas`
  - `openpyxl`
  - `requests`
  - `PyMuPDF`

Pueden instalarse con:

```bash
pip install pandas openpyxl requests pymupdf
```

## Archivos necesarios

Para la correcta generación del listado, se deben incluir en la misma carpeta los siguientes archivos:

1. **Exportación de SharePoint**: Un archivo Excel generado a partir de la exportación de la carpeta de archivos de Claro del mes.
2. **MapeoArbolGenealogico.py**: Script para procesar y extraer los archivos y enlaces del SharePoint.
3. **Excel con archivos Claro Hogar**: Un archivo generado a partir del mapeo, que contiene exclusivamente los archivos de Claro Hogar.
4. **Listado consolidado brindado por Claro**: Archivo Excel con la referencia de documentos de Claro Hogar, prepago y postpago.
5. **Cookies.JSON**: Archivo con las cookies necesarias para la descarga de archivos de SharePoint.
6. **ListadoAutomaticoClaroHogar.py**: Script para la generación automática del listado final.

## Pasos para la generación del listado

### 1. Exportar archivos de SharePoint

- Exportar la carpeta del mes desde SharePoint a un archivo `query.iqy`.
- Ejecutar el archivo `.iqy` en Excel para generar un archivo Excel con la información de los archivos y subcarpetas.
- Guardar este archivo Excel junto con `MapeoArbolGenealogico.py` y actualizar el nombre en la variable `excel_file_path` dentro del script.

![Carpeta de archivos SharePoint](ruta_de_imagen_1)

### 2. Generar el mapeo de archivos

Ejecutar:

```bash
python MapeoArbolGenealogico.py
```

Esto generará un nuevo archivo Excel con todos los archivos de SharePoint y sus enlaces.

![Query de archivos SharePoint](ruta_de_imagen_2)

### 3. Filtrar archivos de Claro Hogar

Crear un nuevo archivo Excel que contenga solo los archivos de Claro Hogar. Este archivo es necesario para el funcionamiento del script de generación del listado.

![Mapeo de archivos SharePoint](ruta_de_imagen_3)
![Archivos Claro Hogar aislados](ruta_de_imagen_4)

### 4. Configurar `Cookies.JSON`

- Instalar la extensión "Cookie-Editor" en el navegador.
- Iniciar sesión en SharePoint y acceder a la carpeta de archivos de Claro Hogar.
- Usar la extensión para extraer las cookies `FedAuth` y `rtFa`.
- Guardar estas cookies en `Cookies.JSON`.

![Extensión Cookie-Editor](ruta_de_imagen_5)
![Extensión en el navegador](ruta_de_imagen_6)
![FedAuth Cookie](ruta_de_imagen_7)
![rtFa Cookie](ruta_de_imagen_8)

### 5. Ejecutar `ListadoAutomaticoClaroHogar.py`

Asegurar que los siguientes parámetros estén correctamente configurados dentro del script:

```python
TEMP_PDF_NAME = 'archivo_temporal.pdf'
EXCEL_GENERADO = 'Claro Hogar (Mes_Año).xlsx'
NombreListadoReferencia = '(Nombre del listado consolidado brindado por Claro)'
NombreMapaArchivosADescargar = '(Nombre del mapa de archivos aislados de Claro Hogar)'
ColumnaComparar = 'B'
ColumnaConectados = 'D'
ColumnaCategoria = 'A'
NombreHojaReferencia = 'HOGAR 2025'
```

Ejecutar el script con:

```bash
python ListadoAutomaticoClaroHogar.py
```

Si los pasos anteriores se realizaron correctamente, el listado de Claro Hogar se generará automáticamente.

## Contribución

Si deseas contribuir, puedes hacer un fork del repositorio y enviar un pull request con mejoras.

## Licencia

Este proyecto está bajo la licencia MIT.
