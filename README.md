# Cerebro de Proyección Multimedia - Iglesia App ⛪🎥

Este es un sistema de proyección multimedia elegante, ligero y rápido diseñado específicamente para pequeñas iglesias. Funciona de manera similar a sistemas profesionales (como ProPresenter o EasyWorship), permitiendo gestionar canciones (cantos), pasajes bíblicos, imágenes y presentaciones de PowerPoint para proyectarlos en un segundo monitor, proyector o capturarlos directamente desde OBS Studio.

---

## 🛠️ Stack Tecnológico Utilizado
- **Lenguaje Base**: Python (3.13+)
- **Interfaz Gráfica (UI)**: PyQt6
- **Estilos Visuales**: QSS (Qt Style Sheets) - Hoja de estilo en modo oscuro premium
- **Base de Datos**: SQLite (almacenamiento local para cantos, Biblia, y referencias multimedia)
- **Procesamiento de Presentaciones**: `python-pptx` + Integración COM de Windows (`win32com`) para conversión nativa de PowerPoint
- **Manejo de Imágenes**: Pillow (PIL)

---

## 📂 Estructura del Proyecto
```text
iglesia-app/
├── main.py              # Código principal y diseño de la interfaz de usuario en PyQt6
├── database.py          # Inicialización y manejo de la base de datos SQLite (iglesia.db)
├── pptx_helper.py       # Helper para convertir archivos .pptx a imágenes de proyección
├── styles.qss           # Hoja de estilos moderna y oscura para la aplicación
├── iglesia.db           # Base de datos local (se genera automáticamente)
├── imported_images/     # Carpeta para almacenar las imágenes importadas
└── imported_pptx/       # Carpeta para almacenar las diapositivas extraídas de PowerPoint
```

---

## 🚀 Cómo Ejecutar la Aplicación

1. Asegúrate de tener Python instalado.
2. Ejecuta el archivo principal en la terminal o consola de comandos:
   ```bash
   python main.py
   ```

---

## ✨ Características Principales

### 1. Panel de Vista Previa (Izquierda)
- Es **estático e inamovible**. Muestra el contenido detallado de lo que tienes seleccionado en la biblioteca.
- Si seleccionas un canto, muestra las estrofas.
- Si seleccionas un PowerPoint, lista sus diapositivas.
- Si seleccionas la Biblia, muestra los versículos del capítulo seleccionado.
- **Flujo en Vivo**: Al hacer clic sobre cualquier elemento de esta lista, se envía y proyecta inmediatamente a la pantalla.

### 2. Biblioteca y Cerebro de Control (Arriba a la Derecha)
- **Cantos**: Buscador predictivo en tiempo real. Botón para crear/agregar canciones completas a la base de datos local de SQLite de forma permanente.
- **Biblias**: Filtros desplegables para seleccionar rápidamente `Libro` -> `Capítulo` -> `Versículo`, buscador de palabras clave y opción para añadir capítulos completos al guion del culto.
- **Imágenes**: Importa cualquier imagen local (`.png`, `.jpg`, `.jpeg`). La aplicación copiará la imagen a la carpeta `imported_images/` para que esté disponible de forma permanente.
- **PPTX**: Importador de presentaciones PowerPoint. 
  - Si el equipo tiene PowerPoint instalado, exportará las diapositivas con la calidad nativa original.
  - Si no tiene PowerPoint, extraerá el texto de cada diapositiva y generará diapositivas oscuras limpias de forma automática.
- **Guiones**: Cuatro columnas paralelas correspondientes a `CANTOS`, `BIBLIA`, `IMÁGENES` y `PPTX` añadidas para el servicio actual. 
  - El guion inicia vacío cada vez que abres el sistema.
  - Al hacer clic en un elemento del guion, este se carga en el panel izquierdo (Preview) listo para ser proyectado en el culto.

### 3. Ventana de Proyección y Controles (Abajo a la Derecha)
- Muestra una miniatura en tiempo real de lo que está viendo el público.
- **Activar Ventana de Proyección**: Abre una ventana independiente sin bordes. Arrástrala a tu proyector o segundo monitor y haz doble clic para ponerla en pantalla completa (o configúrala como entrada en tu OBS).
- **Negro Total (Black)**: Pone la pantalla de proyección completamente en negro inmediatamente.
- **Limpiar Texto**: Oculta la letra proyectada pero mantiene la imagen de fondo si existía.
- **Restaurar**: Vuelve a mostrar el texto/diapositiva ocultada al instante.
