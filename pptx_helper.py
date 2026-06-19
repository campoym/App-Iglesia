import os
import shutil
from pptx import Presentation
from PIL import Image, ImageDraw, ImageFont

# Intentar importar win32com para conversión nativa usando PowerPoint en Windows
try:
    import win32com.client
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False

def convert_pptx_to_images(pptx_path, output_dir):
    """
    Convierte un archivo PPTX en una serie de imágenes PNG.
    Retorna una lista de rutas de archivos de imágenes.
    """
    os.makedirs(output_dir, exist_ok=True)
    images = []
    
    # Método 1: Intentar conversión nativa con MS PowerPoint
    if HAS_WIN32COM:
        powerpoint = None
        presentation = None
        try:
            # Inicializar PowerPoint en segundo plano
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            
            # Abrir la presentación (WithWindow=False para no interrumpir al usuario)
            abs_pptx_path = os.path.abspath(pptx_path)
            abs_output_dir = os.path.abspath(output_dir)
            
            presentation = powerpoint.Presentations.Open(abs_pptx_path, WithWindow=False)
            
            # Exportar la presentación completa como carpeta de imágenes
            # 17 representa formato PNG en la enumeración de PowerPoint
            presentation.SaveAs(abs_output_dir, 17)
            
            # Cerrar presentación y PowerPoint
            presentation.Close()
            powerpoint.Quit()
            
            # PowerPoint crea archivos llamados Slide1.PNG, Slide2.PNG, etc.
            files = os.listdir(output_dir)
            png_files = [f for f in files if f.lower().endswith(".png")]
            
            # Ordenar numéricamente (Slide1, Slide2, Slide10...)
            def extract_number(filename):
                num_str = "".join([c for c in filename if c.isdigit()])
                return int(num_str) if num_str else 0
            
            png_files.sort(key=extract_number)
            
            for png in png_files:
                images.append(os.path.join(output_dir, png))
                
            if images:
                print(f"PowerPoint convertido exitosamente vía win32com: {len(images)} diapositivas.")
                return images
        except Exception as e:
            print(f"Error al convertir PPTX con win32com: {e}. Intentando fallback...")
            # Limpiar recursos si quedaron abiertos
            try:
                if presentation:
                    presentation.Close()
            except:
                pass
            try:
                if powerpoint:
                    powerpoint.Quit()
            except:
                pass
            
    # Método 2: Fallback (python-pptx + Pillow) - Extrae texto y dibuja diapositivas oscuras limpias
    try:
        prs = Presentation(pptx_path)
        for i, slide in enumerate(prs.slides):
            # Recolectar todo el texto de la diapositiva
            slide_text = []
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text.strip():
                    slide_text.append(shape.text.strip())
            
            text_content = "\n".join(slide_text)
            
            # Crear una imagen negra/oscura de resolución FHD (1920x1080)
            # Usamos una proporción de 16:9 estándar
            width, height = 1920, 1080
            img = Image.new("RGB", (width, height), color="#121214")
            draw = ImageDraw.Draw(img)
            
            # Intentar cargar una fuente bonita
            font = None
            try:
                font = ImageFont.truetype("arial.ttf", 48)
            except IOError:
                font = ImageFont.load_default()
            
            # Dividir texto por líneas y dibujar centrado
            lines = [l.strip() for l in text_content.split("\n") if l.strip()]
            
            # Calcular la altura total del bloque de texto
            y_offset = height // 4
            
            for line in lines[:8]:  # Limitar a 8 líneas por diapositiva para no desbordar
                try:
                    # Pillow moderno
                    bbox = draw.textbbox((0, 0), line, font=font)
                    text_w = bbox[2] - bbox[0]
                    text_h = bbox[3] - bbox[1]
                except AttributeError:
                    # Pillow antiguo
                    text_w, text_h = draw.textsize(line, font=font)
                
                # Centrar X
                x = (width - text_w) // 2
                draw.text((x, y_offset), line, font=font, fill="#ffffff")
                y_offset += text_h + 30
                
            img_path = os.path.join(output_dir, f"slide_{i+1}.png")
            img.save(img_path)
            images.append(img_path)
            
        print(f"PowerPoint procesado con python-pptx (fallback): {len(images)} diapositivas.")
        return images
    except Exception as e:
        print(f"Error crítico en pptx_helper fallback: {e}")
        return []
