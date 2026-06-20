"""
pdf_helper.py
Convierte cada página de un PDF en una imagen PNG, usando PyMuPDF (fitz).
100% Python puro, sin dependencias externas del sistema.
"""

import os
import fitz  # PyMuPDF


def convert_pdf_to_images(pdf_path, output_dir, dpi=150):
    """
    Convierte un archivo PDF en una serie de imágenes PNG, una por página.
    Retorna una lista de rutas de archivos de imágenes en orden.
    """
    os.makedirs(output_dir, exist_ok=True)
    images = []

    try:
        doc = fitz.open(pdf_path)
        zoom = dpi / 72  # 72 dpi es la base de PDF
        matrix = fitz.Matrix(zoom, zoom)

        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            pix = page.get_pixmap(matrix=matrix)

            img_path = os.path.join(output_dir, f"page_{page_num + 1}.png")
            pix.save(img_path)
            images.append(img_path)

        doc.close()
        print(f"PDF convertido exitosamente: {len(images)} páginas.")
        return images

    except Exception as e:
        print(f"Error al convertir PDF: {e}")
        return []