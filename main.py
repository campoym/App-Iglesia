import os
import sys
import shutil
import sqlite3
import re
import unicodedata
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QListWidget, QListWidgetItem, QPushButton, QLineEdit, QLabel,
    QFileDialog, QMessageBox, QSplitter, QFrame, QGridLayout, QStackedWidget,
    QDialog, QTextEdit, QScrollArea, QComboBox, QSizePolicy
)
from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QPixmap, QColor

import database
import pptx_helper
import pdf_helper

# Configurar carpetas de almacenamiento para recursos
PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
database.init_db()


def remove_accents(text):
    """
    Normaliza el texto y elimina las tildes/acentos en español (á -> a, ü -> u, etc.)
    Manteniendo caracteres propios como la ñ.
    """
    if not text:
        return ""
    # Normalizar en forma NFD (separa letras de sus diacríticos)
    normalized = unicodedata.normalize('NFD', text)
    # Filtrar diacríticos (marcas de acento)
    return "".join(
        c for c in normalized
        if unicodedata.category(c) != 'Mn'
    ).lower()


class PreviewListWidget(QListWidget):
    """QListWidget personalizado para soportar navegación por teclado con proyección inmediata."""

    def keyPressEvent(self, event):
        super().keyPressEvent(event)
        if event.key() in (Qt.Key.Key_Up, Qt.Key.Key_Down, Qt.Key.Key_PageUp, Qt.Key.Key_PageDown, Qt.Key.Key_Home, Qt.Key.Key_End):
            current = self.currentItem()
            if current:
                # Obtener la ventana principal para disparar la proyección
                main_window = self.window()
                if hasattr(main_window, "on_preview_card_clicked"):
                    main_window.on_preview_card_clicked(current)


class PreviewImageCardWidget(QFrame):
    """Tarjeta de previsualización individual para imágenes en el panel izquierdo."""

    def __init__(self, header, image_path, parent=None):
        super().__init__(parent)
        self.setObjectName("previewCard")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 60, 15, 60)
        layout.setSpacing(60)

        self.header_label = QLabel(header)
        self.header_label.setObjectName("cardHeader")
        layout.addWidget(self.header_label)

        # Etiqueta para la miniatura de la imagen
        self.thumb_label = QLabel()
        # Relación de aspecto típica 16:9
        self.thumb_label.setFixedSize(160, 90)
        self.thumb_label.setStyleSheet(
            "background-color: #1a1a1e; border-radius: 4px;")
        self.thumb_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # Cargar y escalar la imagen
        pixmap = QPixmap(image_path)
        if not pixmap.isNull():
            scaled = pixmap.scaled(self.thumb_label.size(
            ), Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            self.thumb_label.setPixmap(scaled)
        layout.addWidget(self.thumb_label,
                         alignment=Qt.AlignmentFlag.AlignCenter)

        self.name_label = QLabel(os.path.basename(image_path))
        self.name_label.setObjectName("cardLyrics")  # Reusar estilo de lyrics
        self.name_label.setWordWrap(True)
        self.name_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.name_label)

        self.set_selected(False)

    def set_selected(self, is_selected):
        if is_selected:
            self.setStyleSheet("""
                QFrame#previewCard {
                    background-color: #eceeff;
                    border: 2px solid #6366f1;
                    border-radius: 8px;
                }
            """)
            self.header_label.setStyleSheet("""
                color: #4f46e5;
                font-size: 11px;
                font-weight: bold;
                background: transparent;
            """)
            self.name_label.setStyleSheet("""
                color: #1e1b4b;
                font-size: 13px;
                font-weight: bold;
                background: transparent;
            """)
        else:
            self.setStyleSheet("""
                QFrame#previewCard {
                    background-color: #202024;
                    border: 1px solid #3a3a3c;
                    border-radius: 8px;
                }
                QFrame#previewCard:hover {
                    background-color: #2e2e33;
                    border: 1px solid #52525b;
                }
            """)
            self.header_label.setStyleSheet("""
                color: #71717a;
                font-size: 11px;
                font-weight: bold;
                background: transparent;
            """)
            self.name_label.setStyleSheet("""
                color: #e5e7eb;
                font-size: 13px;
                font-weight: 600;
                background: transparent;
            """)


class LibraryImageRowWidget(QWidget):
    """Fila de imagen para la biblioteca. Clickeable directamente."""

    def __init__(self, title, file_path, on_add_clicked, parent=None):
        super().__init__(parent)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(15, 10, 15, 10)
        layout.setSpacing(12)

        # Miniatura de la imagen
        self.thumb_lbl = QLabel()
        self.thumb_lbl.setFixedSize(48, 36)
        self.thumb_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.thumb_lbl.setStyleSheet("""
            background-color: #1a1a1e; 
            border-radius: 4px;
        """)

        # Cargar miniatura
        pixmap = QPixmap(file_path)
        if not pixmap.isNull():
            scaled = pixmap.scaled(self.thumb_lbl.size(
            ), Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            self.thumb_lbl.setPixmap(scaled)
        layout.addWidget(self.thumb_lbl)

        # Textos de título
        text_layout = QVBoxLayout()
        text_layout.setSpacing(2)

        self.title_lbl = QLabel(title)
        self.title_lbl.setStyleSheet(
            "color: #fafafa; font-weight: bold; font-size: 14px; background: transparent;")
        text_layout.addWidget(self.title_lbl)

        self.sub_lbl = QLabel("Imagen Local")
        self.sub_lbl.setStyleSheet(
            "color: #a1a1aa; font-size: 11px; background: transparent;")
        text_layout.addWidget(self.sub_lbl)

        layout.addLayout(text_layout)
        layout.addStretch()

        # Hacer transparentes para eventos del ratón
        self.thumb_lbl.setAttribute(
            Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        self.title_lbl.setAttribute(
            Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        self.sub_lbl.setAttribute(
            Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)

        # Botón para añadir al guión
        self.add_btn = QPushButton("+ Guión")
        self.add_btn.setObjectName("addToGuionBtn")
        self.add_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.add_btn.clicked.connect(on_add_clicked)
        layout.addWidget(self.add_btn)


class LibraryPptxRowWidget(QWidget):
    """Fila de presentación (PPTX o PDF) para la biblioteca. Clickeable directamente."""

    def __init__(self, title, on_add_clicked, on_delete_clicked, file_type="pptx", parent=None):
        super().__init__(parent)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(15, 10, 15, 10)
        layout.setSpacing(12)

        is_pdf = file_type == "pdf"

        # Icono circular: rojo "PDF" o naranja "P" de PowerPoint
        self.icon_lbl = QLabel("PDF" if is_pdf else "P")
        self.icon_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.icon_lbl.setFixedSize(36, 36)
        if is_pdf:
            self.icon_lbl.setStyleSheet("""
                background-color: #7f1d1d; 
                color: #fca5a5; 
                border-radius: 18px; 
                font-size: 10px;
                font-weight: bold;
            """)
        else:
            self.icon_lbl.setStyleSheet("""
                background-color: #7c2d12; 
                color: #ff7a59; 
                border-radius: 18px; 
                font-size: 15px;
                font-weight: bold;
            """)
        layout.addWidget(self.icon_lbl)

        # Textos de título
        text_layout = QVBoxLayout()
        text_layout.setSpacing(2)

        self.title_lbl = QLabel(title)
        self.title_lbl.setStyleSheet(
            "color: #fafafa; font-weight: bold; font-size: 14px; background: transparent;")
        text_layout.addWidget(self.title_lbl)

        self.sub_lbl = QLabel(
            "Documento PDF" if is_pdf else "Presentación PowerPoint")
        self.sub_lbl.setStyleSheet(
            "color: #a1a1aa; font-size: 11px; background: transparent;")
        text_layout.addWidget(self.sub_lbl)

        layout.addLayout(text_layout)
        layout.addStretch()

        # Hacer que los clicks en los textos e icono se transfieran al QListWidget
        self.icon_lbl.setAttribute(
            Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        self.title_lbl.setAttribute(
            Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        self.sub_lbl.setAttribute(
            Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)

        # Botón Eliminar
        self.del_btn = QPushButton("🗑")
        self.del_btn.setFixedSize(30, 30)
        self.del_btn.setToolTip("Eliminar presentación")
        self.del_btn.setStyleSheet("""
            QPushButton {
                background-color: #27272a;
                color: #a1a1aa;
                border: 1px solid #3f3f46;
                border-radius: 6px;
                font-size: 13px;
            }
            QPushButton:hover {
                background-color: #7f1d1d;
                color: #fca5a5;
                border-color: #7f1d1d;
            }
        """)
        self.del_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.del_btn.clicked.connect(on_delete_clicked)
        layout.addWidget(self.del_btn)

        # Botón para añadir al guión
        self.add_btn = QPushButton("+ Guión")
        self.add_btn.setObjectName("addToGuionBtn")
        self.add_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.add_btn.clicked.connect(on_add_clicked)
        layout.addWidget(self.add_btn)


class PreviewCardWidget(QFrame):
    """Tarjeta de previsualización individual para las estrofas/diapositivas en el panel izquierdo."""

    def __init__(self, header, lyrics, parent=None):
        super().__init__(parent)
        self.setObjectName("previewCard")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 12, 15, 12)
        layout.setSpacing(6)

        self.header_label = QLabel(header)
        self.header_label.setObjectName("cardHeader")
        layout.addWidget(self.header_label)

        self.lyrics_label = QLabel(lyrics)
        self.lyrics_label.setObjectName("cardLyrics")
        self.lyrics_label.setWordWrap(True)
        layout.addWidget(self.lyrics_label)

        self.set_selected(False)

    def set_selected(self, is_selected):
        if is_selected:
            self.setStyleSheet("""
                QFrame#previewCard {
                    background-color: #eceeff;
                    border: 2px solid #6366f1;
                    border-radius: 8px;
                }
            """)
            self.header_label.setStyleSheet("""
                color: #4f46e5;
                font-size: 11px;
                font-weight: bold;
                background: transparent;
            """)
            self.lyrics_label.setStyleSheet("""
                color: #1e1b4b;
                font-size: 13px;
                font-weight: bold;
                background: transparent;
            """)
        else:
            self.setStyleSheet("""
                QFrame#previewCard {
                    background-color: #202024;
                    border: 1px solid #3a3a3c;
                    border-radius: 8px;
                }
                QFrame#previewCard:hover {
                    background-color: #2e2e33;
                    border: 1px solid #52525b;
                }
            """)
            self.header_label.setStyleSheet("""
                color: #71717a;
                font-size: 11px;
                font-weight: bold;
                background: transparent;
            """)
            self.lyrics_label.setStyleSheet("""
                color: #e5e7eb;
                font-size: 13px;
                font-weight: 600;
                background: transparent;
            """)


class LibrarySongRowWidget(QWidget):
    """Fila de canto para la biblioteca. Sin botón Proyectar. Clickeable directamente."""

    def __init__(self, title, category, on_add_clicked, on_edit_clicked, on_delete_clicked, parent=None):
        super().__init__(parent)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(15, 10, 15, 10)
        layout.setSpacing(12)

        # Icono circular de música
        self.icon_lbl = QLabel("♫")
        self.icon_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.icon_lbl.setFixedSize(36, 36)
        self.icon_lbl.setStyleSheet("""
            background-color: #1e1b4b; 
            color: #818cf8; 
            border-radius: 18px; 
            font-size: 15px;
            font-weight: bold;
        """)
        layout.addWidget(self.icon_lbl)

        # Textos de título y categoría
        text_layout = QVBoxLayout()
        text_layout.setSpacing(2)

        self.title_lbl = QLabel(title)
        self.title_lbl.setStyleSheet(
            "color: #fafafa; font-weight: bold; font-size: 14px; background: transparent;")
        text_layout.addWidget(self.title_lbl)

        self.category_lbl = QLabel(category)
        self.category_lbl.setStyleSheet(
            "color: #a1a1aa; font-size: 11px; background: transparent;")
        text_layout.addWidget(self.category_lbl)

        layout.addLayout(text_layout)
        layout.addStretch()

        # Transparentes para mouse events
        self.icon_lbl.setAttribute(
            Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        self.title_lbl.setAttribute(
            Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        self.category_lbl.setAttribute(
            Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)

        # Botón Editar
        self.edit_btn = QPushButton("✏")
        self.edit_btn.setFixedSize(30, 30)
        self.edit_btn.setToolTip("Editar canto")
        self.edit_btn.setStyleSheet("""
            QPushButton {
                background-color: #27272a;
                color: #a1a1aa;
                border: 1px solid #3f3f46;
                border-radius: 6px;
                font-size: 13px;
            }
            QPushButton:hover {
                background-color: #3f3f46;
                color: #fafafa;
            }
        """)
        self.edit_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.edit_btn.clicked.connect(on_edit_clicked)
        layout.addWidget(self.edit_btn)

        # Botón Eliminar
        self.del_btn = QPushButton("🗑")
        self.del_btn.setFixedSize(30, 30)
        self.del_btn.setToolTip("Eliminar canto")
        self.del_btn.setStyleSheet("""
            QPushButton {
                background-color: #27272a;
                color: #a1a1aa;
                border: 1px solid #3f3f46;
                border-radius: 6px;
                font-size: 13px;
            }
            QPushButton:hover {
                background-color: #7f1d1d;
                color: #fca5a5;
                border-color: #7f1d1d;
            }
        """)
        self.del_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.del_btn.clicked.connect(on_delete_clicked)
        layout.addWidget(self.del_btn)

        # Botón + Guión
        self.add_btn = QPushButton("+ Guión")
        self.add_btn.setObjectName("addToGuionBtn")
        self.add_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.add_btn.clicked.connect(on_add_clicked)
        layout.addWidget(self.add_btn)


class CleanProjectionWidget(QFrame):
    """Widget de proyección OBS limpio que renderiza el contenido."""

    def __init__(self, is_external=False, parent=None):
        super().__init__(parent)
        self.setObjectName("obsContainerFrame")
        # Establecer fondo transparente para que nuestro paintEvent controle el dibujado completo
        self.setStyleSheet(
            "background-color: transparent; border: none; border-radius: 0px;")

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 15, 20, 15)

        # Línea superior (Etiqueta OBS)
        top_layout = QHBoxLayout()
        self.obs_title = QLabel("VENTANA OBS")
        self.obs_title.setObjectName("obsLabelTitle")
        top_layout.addWidget(self.obs_title)
        top_layout.addStretch()
        main_layout.addLayout(top_layout)

        main_layout.addStretch()

        # Texto central
        self.text_label = QLabel(self)
        self.text_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.text_label.setWordWrap(True)
        self.text_label.setStyleSheet("""
            color: #ffffff; 
            font-size: 38px; 
            font-weight: bold; 
            background: transparent;
        """)
        main_layout.addWidget(self.text_label)

        main_layout.addStretch()

        # Línea inferior (Pie de página y ayuda)
        bottom_layout = QHBoxLayout()
        bottom_layout.addStretch()

        self.footer_label = QLabel("")
        self.footer_label.setObjectName("obsFooterTitle")
        bottom_layout.addWidget(self.footer_label)

        bottom_layout.addStretch()

        self.help_btn = QPushButton("?")
        self.help_btn.setObjectName("helpRoundBtn")
        bottom_layout.addWidget(self.help_btn)

        main_layout.addLayout(bottom_layout)

        # Ocultar controles administrativos en ventana externa
        if is_external:
            self.obs_title.hide()
            self.help_btn.hide()

        self.raw_pixmap = None
        self.general_bg_pixmap = None
        self.is_black_screen = True
        self.current_song_title = ""
        self.current_slide_name = ""

    def display_text(self, text, slide_name="", song_title=""):
        self.is_black_screen = False
        self.raw_pixmap = None
        self.text_label.setText(text)
        self.text_label.show()
        self.update()

        if song_title:
            self.current_song_title = song_title
        if slide_name:
            self.current_slide_name = slide_name

        if self.current_slide_name and self.current_song_title:
            self.footer_label.setText(
                f"{self.current_slide_name} · {self.current_song_title}")
        else:
            self.footer_label.setText(
                self.current_song_title or self.current_slide_name)

    def display_image(self, pixmap_path, slide_name="", song_title=""):
        self.is_black_screen = False
        self.text_label.hide()
        # Cargar con el device pixel ratio de la pantalla para evitar borrosidad
        pixmap = QPixmap(pixmap_path)
        dpr = self.devicePixelRatioF()
        if dpr != 1.0:
            pixmap.setDevicePixelRatio(dpr)
        self.raw_pixmap = pixmap
        self.update()

        if song_title:
            self.current_song_title = song_title
        if slide_name:
            self.current_slide_name = slide_name

        if self.current_slide_name and self.current_song_title:
            self.footer_label.setText(
                f"{self.current_slide_name} · {self.current_song_title}")
        else:
            self.footer_label.setText(
                self.current_song_title or self.current_slide_name)

    def set_black_screen(self):
        self.is_black_screen = True
        self.text_label.hide()
        self.raw_pixmap = None
        self.footer_label.setText("")
        self.update()

    def clear_text_only(self):
        self.text_label.setText("")
        self.footer_label.setText("")
        self.update()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.update()

    def paintEvent(self, event):
        from PyQt6.QtGui import QPainter
        painter = QPainter(self)

        rect = self.contentsRect()
        # Dibujar fondo negro sólido siempre
        painter.fillRect(rect, Qt.GlobalColor.black)

        # Si hay imagen de fondo general (para letras), pintarla primero
        if not getattr(self, "is_black_screen", False) and self.general_bg_pixmap and not self.general_bg_pixmap.isNull() and self.raw_pixmap is None:
            scaled = self.general_bg_pixmap.scaled(
                rect.size(),
                Qt.AspectRatioMode.KeepAspectRatioByExpanding,
                Qt.TransformationMode.SmoothTransformation
            )
            x = rect.x() + (rect.width() - scaled.width()) // 2
            y = rect.y() + (rect.height() - scaled.height()) // 2
            painter.drawPixmap(x, y, scaled)

        # Si hay una imagen/diapositiva cargada, dibujarla cubriendo toda la pantalla
        if self.raw_pixmap and not self.raw_pixmap.isNull():
            scaled = self.raw_pixmap.scaled(
                rect.size(),
                Qt.AspectRatioMode.KeepAspectRatioByExpanding,
                Qt.TransformationMode.SmoothTransformation
            )
            x = rect.x() + (rect.width() - scaled.width()) // 2
            y = rect.y() + (rect.height() - scaled.height()) // 2
            painter.drawPixmap(x, y, scaled)

        painter.end()
        super().paintEvent(event)


class ProjectionWindow(QWidget):
    """Ventana externa independiente (Proyector/Pantalla Secundaria)."""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Salida de Proyección (Pantalla Completa)")
        self.setStyleSheet("background-color: #000000;")
        self.resize(800, 600)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self.projection_widget = CleanProjectionWidget(is_external=True)
        layout.addWidget(self.projection_widget)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Escape:
            if self.isFullScreen():
                self.showNormal()


class SectionBlockWidget(QFrame):
    """Bloque individual de sección (Título, Estrofa, Coro) dentro del editor."""
    COLORS = {
        "Título":   {"border": "#6366f1", "header_bg": "#1e1b4b", "header_fg": "#818cf8", "label": "#818cf8"},
        "Estrofa":  {"border": "#3b82f6", "header_bg": "#172554", "header_fg": "#93c5fd", "label": "#93c5fd"},
        "Coro":     {"border": "#a855f7", "header_bg": "#2e1065", "header_fg": "#d8b4fe", "label": "#d8b4fe"},
        "Puente":   {"border": "#ec4899", "header_bg": "#4a044e", "header_fg": "#f9a8d4", "label": "#f9a8d4"},
    }

    def __init__(self, section_type="Estrofa", number=1, content="", on_delete=None, parent=None):
        super().__init__(parent)
        self.on_delete = on_delete
        self.section_type = section_type
        self.number = number

        colors = self.COLORS.get(section_type, self.COLORS["Estrofa"])

        self.setStyleSheet(f"""
            QFrame {{
                background-color: #202024;
                border: 1px solid {colors['border']};
                border-radius: 8px;
            }}
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(6)

        # Header row
        header_row = QHBoxLayout()

        self.type_combo = QComboBox()
        self.type_combo.addItems(["Título", "Estrofa", "Coro", "Puente"])
        self.type_combo.setCurrentText(section_type)
        self.type_combo.setFixedWidth(90)
        self.type_combo.setStyleSheet(f"""
            QComboBox {{
                background-color: {colors['header_bg']};
                color: {colors['header_fg']};
                border: none;
                border-radius: 4px;
                padding: 3px 8px;
                font-size: 11px;
                font-weight: bold;
            }}
            QComboBox::drop-down {{ border: none; }}
            QComboBox QAbstractItemView {{
                background-color: #27272a;
                color: #f4f4f5;
                selection-background-color: #3f3f46;
            }}
        """)
        self.type_combo.currentTextChanged.connect(self._on_type_changed)
        header_row.addWidget(self.type_combo)

        self.num_combo = QComboBox()
        self.num_combo.addItems([str(i) for i in range(1, 9)])
        self.num_combo.setCurrentText(str(number))
        self.num_combo.setFixedWidth(50)
        self.num_combo.setStyleSheet(f"""
            QComboBox {{
                background-color: #27272a;
                color: {colors['label']};
                border: 1px solid #3f3f46;
                border-radius: 4px;
                padding: 3px 6px;
                font-size: 11px;
                font-weight: bold;
            }}
            QComboBox::drop-down {{ border: none; }}
            QComboBox QAbstractItemView {{
                background-color: #27272a;
                color: #f4f4f5;
                selection-background-color: #3f3f46;
            }}
        """)
        header_row.addWidget(self.num_combo)
        header_row.addStretch()

        del_btn = QPushButton("✕")
        del_btn.setFixedSize(24, 24)
        del_btn.setStyleSheet("""
            QPushButton {
                background-color: #3f3f46;
                color: #a1a1aa;
                border: none;
                border-radius: 12px;
                font-size: 11px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #7f1d1d;
                color: #fca5a5;
            }
        """)
        del_btn.clicked.connect(self._delete_self)
        header_row.addWidget(del_btn)

        layout.addLayout(header_row)

        # Text area
        self.text_edit = QTextEdit()
        self.text_edit.setPlaceholderText("Escribe la letra aquí...")
        self.text_edit.setPlainText(content)
        self.text_edit.setMinimumHeight(90)
        self.text_edit.setMaximumHeight(160)
        self.text_edit.setStyleSheet("""
            QTextEdit {
                background-color: #18181b;
                color: #f4f4f5;
                border: none;
                border-radius: 4px;
                padding: 8px;
                font-size: 13px;
            }
        """)
        layout.addWidget(self.text_edit)

    def _on_type_changed(self, new_type):
        self.section_type = new_type
        colors = self.COLORS.get(new_type, self.COLORS["Estrofa"])
        self.setStyleSheet(f"""
            QFrame {{
                background-color: #202024;
                border: 1px solid {colors['border']};
                border-radius: 8px;
            }}
        """)
        self.type_combo.setStyleSheet(f"""
            QComboBox {{
                background-color: {colors['header_bg']};
                color: {colors['header_fg']};
                border: none;
                border-radius: 4px;
                padding: 3px 8px;
                font-size: 11px;
                font-weight: bold;
            }}
            QComboBox::drop-down {{ border: none; }}
            QComboBox QAbstractItemView {{
                background-color: #27272a;
                color: #f4f4f5;
                selection-background-color: #3f3f46;
            }}
        """)
        self.num_combo.setStyleSheet(f"""
            QComboBox {{
                background-color: #27272a;
                color: {colors['label']};
                border: 1px solid #3f3f46;
                border-radius: 4px;
                padding: 3px 6px;
                font-size: 11px;
                font-weight: bold;
            }}
            QComboBox::drop-down {{ border: none; }}
            QComboBox QAbstractItemView {{
                background-color: #27272a;
                color: #f4f4f5;
                selection-background-color: #3f3f46;
            }}
        """)
        self.num_combo.setVisible(new_type != "Título")

    def _delete_self(self):
        if self.on_delete:
            self.on_delete(self)

    def get_data(self):
        sec_type = self.type_combo.currentText()
        num = self.num_combo.currentText() if sec_type != "Título" else ""
        label = f"{sec_type} {num}".strip() if num else sec_type
        content = self.text_edit.toPlainText().strip()
        return label, content


class SongEditorDialog(QDialog):
    """Diálogo para crear o editar un canto."""

    def __init__(self, parent=None, song_id=None):
        super().__init__(parent)
        self.song_id = song_id
        self.section_widgets = []

        self.setWindowTitle("Nuevo Canto" if not song_id else "Editar Canto")
        self.setMinimumSize(620, 700)
        self.setStyleSheet("""
            QDialog {
                background-color: #0f0f11;
            }
            QLabel {
                color: #a1a1aa;
                font-size: 11px;
                font-weight: bold;
                background: transparent;
            }
            QLineEdit {
                background-color: #27272a;
                border: 1px solid #3f3f46;
                border-radius: 8px;
                color: #ffffff;
                padding: 10px 14px;
                font-size: 14px;
            }
            QLineEdit:focus {
                border: 1px solid #6366f1;
            }
            QScrollArea {
                border: none;
                background: transparent;
            }
            QScrollBar:vertical {
                border: none;
                background: #18181b;
                width: 6px;
                border-radius: 3px;
            }
            QScrollBar::handle:vertical {
                background: #3f3f46;
                border-radius: 3px;
            }
        """)

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(14)

        # Título del canto
        lbl_title = QLabel("TÍTULO DEL CANTO")
        main_layout.addWidget(lbl_title)

        self.title_input = QLineEdit()
        self.title_input.setPlaceholderText("Ej. A Dios sea la Gloria")
        main_layout.addWidget(self.title_input)

        # Fila: Categoría y Tono
        meta_row = QHBoxLayout()
        meta_row.setSpacing(10)

        left_col = QVBoxLayout()
        lbl_cat = QLabel("CATEGORÍA")
        self.category_input = QLineEdit()
        self.category_input.setPlaceholderText(
            "Ej. Alabanza, Adoración, Himno")
        left_col.addWidget(lbl_cat)
        left_col.addWidget(self.category_input)
        meta_row.addLayout(left_col)

        right_col = QVBoxLayout()
        lbl_tone = QLabel("TONO")
        self.tone_input = QLineEdit()
        self.tone_input.setPlaceholderText("Ej. G, Am, F#")
        self.tone_input.setMaximumWidth(120)
        right_col.addWidget(lbl_tone)
        right_col.addWidget(self.tone_input)
        meta_row.addLayout(right_col)

        main_layout.addLayout(meta_row)

        # Separador
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("background-color: #27272a; max-height: 1px;")
        main_layout.addWidget(sep)

        # Label secciones
        sections_header = QHBoxLayout()
        lbl_sections = QLabel("SECCIONES")
        sections_header.addWidget(lbl_sections)
        sections_header.addStretch()
        main_layout.addLayout(sections_header)

        # Scroll area para los bloques
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        self.sections_container = QWidget()
        self.sections_container.setStyleSheet("background: transparent;")
        self.sections_layout = QVBoxLayout(self.sections_container)
        self.sections_layout.setContentsMargins(0, 0, 8, 0)
        self.sections_layout.setSpacing(8)
        self.sections_layout.addStretch()

        self.scroll_area.setWidget(self.sections_container)
        main_layout.addWidget(self.scroll_area, stretch=1)

        # Botones para agregar secciones
        add_row = QHBoxLayout()
        add_row.setSpacing(8)

        for label, sec_type in [("+ Título", "Título"), ("+ Estrofa", "Estrofa"), ("+ Coro", "Coro"), ("+ Puente", "Puente")]:
            colors = SectionBlockWidget.COLORS.get(
                sec_type, SectionBlockWidget.COLORS["Estrofa"])
            btn = QPushButton(label)
            btn.setStyleSheet(f"""
                QPushButton {{
                    background-color: {colors['header_bg']};
                    color: {colors['header_fg']};
                    border: 1px solid {colors['border']};
                    border-radius: 14px;
                    padding: 6px 14px;
                    font-size: 11px;
                    font-weight: bold;
                }}
                QPushButton:hover {{
                    background-color: {colors['border']};
                    color: #ffffff;
                }}
            """)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.clicked.connect(
                lambda checked, st=sec_type: self.add_section(st))
            add_row.addWidget(btn)

        add_row.addStretch()
        main_layout.addLayout(add_row)

        # Botones guardar/cancelar
        bottom_row = QHBoxLayout()
        bottom_row.setSpacing(10)
        bottom_row.addStretch()

        cancel_btn = QPushButton("Cancelar")
        cancel_btn.setStyleSheet("""
            QPushButton {
                background-color: #27272a;
                color: #a1a1aa;
                border: 1px solid #3f3f46;
                border-radius: 8px;
                padding: 10px 24px;
                font-weight: bold;
            }
            QPushButton:hover { background-color: #3f3f46; color: #fff; }
        """)
        cancel_btn.clicked.connect(self.reject)
        bottom_row.addWidget(cancel_btn)

        save_btn = QPushButton("💾  Guardar Canto")
        save_btn.setStyleSheet("""
            QPushButton {
                background-color: #4f46e5;
                color: #ffffff;
                border: none;
                border-radius: 8px;
                padding: 10px 24px;
                font-weight: bold;
                font-size: 13px;
            }
            QPushButton:hover { background-color: #6366f1; }
        """)
        save_btn.clicked.connect(self.save_song)
        bottom_row.addWidget(save_btn)

        main_layout.addLayout(bottom_row)

        # Si es edición, cargar datos
        if self.song_id:
            self._load_existing_song()
        else:
            # Bloques por defecto
            self.add_section("Título")
            self.add_section("Estrofa")
            self.add_section("Coro")

    def add_section(self, sec_type="Estrofa", number=1, content=""):
        existing = [
            w for w in self.section_widgets if w.type_combo.currentText() == sec_type]
        auto_num = len(existing) + 1 if sec_type != "Título" else 1

        block = SectionBlockWidget(
            section_type=sec_type,
            number=number if number > 1 else auto_num,
            content=content,
            on_delete=self.remove_section
        )
        block.num_combo.setVisible(sec_type != "Título")

        idx = self.sections_layout.count() - 1
        self.sections_layout.insertWidget(idx, block)
        self.section_widgets.append(block)

        from PyQt6.QtCore import QTimer
        QTimer.singleShot(50, lambda: self.scroll_area.verticalScrollBar().setValue(
            self.scroll_area.verticalScrollBar().maximum()
        ))

    def remove_section(self, block):
        self.sections_layout.removeWidget(block)
        self.section_widgets.remove(block)
        block.deleteLater()

    def _load_existing_song(self):
        conn = sqlite3.connect(database.DB_PATH)
        cursor = conn.cursor()
        cursor.execute(
            "SELECT title, category, tone, lyrics FROM songs WHERE id = ?", (self.song_id,))
        row = cursor.fetchone()
        conn.close()
        if not row:
            return
        title, category, tone, lyrics = row
        self.title_input.setText(title or "")
        self.category_input.setText(category or "")
        self.tone_input.setText(tone or "")

        slides = [s.strip() for s in lyrics.split("\n\n") if s.strip()]
        for slide in slides:
            lines = slide.split("\n")
            if lines and lines[0].startswith("[") and lines[0].endswith("]"):
                header = lines[0][1:-1]
                content = "\n".join(lines[1:])
            else:
                header = "Estrofa"
                content = slide

            parts = header.rsplit(" ", 1)
            if len(parts) == 2 and parts[1].isdigit():
                sec_type, num = parts[0], int(parts[1])
            else:
                sec_type, num = header, 1

            self.add_section(sec_type, num, content)

    def save_song(self):
        title = self.title_input.text().strip()
        if not title:
            QMessageBox.warning(self, "Campo requerido",
                                "El título del canto es obligatorio.")
            return
        if not self.section_widgets:
            QMessageBox.warning(self, "Sin secciones",
                                "Agrega al menos una sección al canto.")
            return

        category = self.category_input.text().strip()
        tone = self.tone_input.text().strip()

        parts = []
        for w in self.section_widgets:
            label, content = w.get_data()
            if content:
                parts.append(f"[{label}]\n{content}")
        lyrics = "\n\n".join(parts)

        conn = sqlite3.connect(database.DB_PATH)
        cursor = conn.cursor()

        # Migración segura: agregar columna tone si no existe
        try:
            cursor.execute("ALTER TABLE songs ADD COLUMN tone TEXT")
            conn.commit()
        except Exception:
            pass  # Ya existe

        if self.song_id:
            cursor.execute(
                "UPDATE songs SET title=?, category=?, tone=?, lyrics=? WHERE id=?",
                (title, category, tone, lyrics, self.song_id)
            )
        else:
            cursor.execute(
                "INSERT INTO songs (title, category, tone, lyrics) VALUES (?, ?, ?, ?)",
                (title, category, tone, lyrics)
            )
        conn.commit()
        conn.close()
        self.accept()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Cerebro de Proyección Multimedia - Iglesia App")
        self.resize(1200, 750)

        self.load_styles()

        # Datos del Guion (en memoria para la sesión)
        self.guion = {
            "songs": [],
            "bible": [],
            "images": [],
            "pptx": []
        }

        # Estado de la proyección
        self.active_item_data = None
        self.projection_window = None

        self.current_projection_mode = "black"
        self.last_projected_mode = None
        self.last_projected_text = ""
        self.last_projected_header = ""
        self.last_projected_song_title = ""
        self.last_projected_image_path = ""

        self.setup_ui()
        self.load_songs_library()
        self.load_images_library()
        self.load_pptx_library()
        self.load_persisted_lyrics_bg()

    def load_styles(self):
        qss_path = os.path.join(PROJECT_DIR, "styles.qss")
        if os.path.exists(qss_path):
            with open(qss_path, "r", encoding="utf-8") as f:
                self.setStyleSheet(f.read())

    def setup_ui(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)

        # Layout vertical: Barra superior + Cuerpo principal
        window_layout = QVBoxLayout(main_widget)
        window_layout.setContentsMargins(0, 0, 0, 0)
        window_layout.setSpacing(0)

        # ---------------- BARRA DE NAVEGACIÓN SUPERIOR (OSCURA) ----------------
        self.top_bar = QFrame()
        self.top_bar.setObjectName("topNavBar")
        top_bar_layout = QHBoxLayout(self.top_bar)
        top_bar_layout.setContentsMargins(15, 0, 15, 0)
        top_bar_layout.setSpacing(10)

        self.nav_buttons = {}
        nav_items = [
            ("Cantos", "♫"),
            ("Biblias", "📖"),
            ("Imágenes", "🖼"),
            ("PPTX", "🖥"),
            ("Guión", "📋")
        ]

        for name, icon in nav_items:
            btn = QPushButton(f"{icon} {name}")
            btn.setObjectName("navBtn")
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.clicked.connect(self.make_nav_callback(name))
            top_bar_layout.addWidget(btn)
            self.nav_buttons[name] = btn

        top_bar_layout.addStretch()

        # Buscador contextual en el topbar — visible solo en Biblias
        self.topbar_bible_search = QLineEdit()
        self.topbar_bible_search.setObjectName("searchBar")
        self.topbar_bible_search.setPlaceholderText(
            "🔎 Buscar cita (ej. Juan 3 16) y Enter...")
        self.topbar_bible_search.setMinimumWidth(320)
        self.topbar_bible_search.setMaximumWidth(480)
        self.topbar_bible_search.setMaximumHeight(34)
        self.topbar_bible_search.setVisible(False)
        self.topbar_bible_search.returnPressed.connect(
            self._on_topbar_bible_search)
        self.topbar_bible_search.textChanged.connect(
            self._on_topbar_bible_search_changed)
        top_bar_layout.addWidget(self.topbar_bible_search)

        window_layout.addWidget(self.top_bar)

        # ---------------- CUERPO PRINCIPAL (CON SPLITTER) ----------------
        body_widget = QWidget()
        body_layout = QHBoxLayout(body_widget)
        body_layout.setContentsMargins(15, 15, 15, 15)
        body_layout.setSpacing(15)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        body_layout.addWidget(splitter)
        window_layout.addWidget(body_widget)

        # ---- Panel Izquierdo: Vista Previa ----
        preview_container = QFrame()
        preview_container.setObjectName("previewPanel")
        preview_layout = QVBoxLayout(preview_container)
        preview_layout.setContentsMargins(15, 15, 15, 15)

        # Cabecera superior con texto descriptivo como el mockup
        header_row = QHBoxLayout()
        preview_header_lbl = QLabel("PREVIEW")
        preview_header_lbl.setObjectName("previewHeaderTitle")
        header_row.addWidget(preview_header_lbl)

        click_instruction_lbl = QLabel("click bloque → proyecta")
        click_instruction_lbl.setStyleSheet(
            "color: #71717a; font-size: 11px; font-weight: normal; background: transparent;")
        header_row.addWidget(click_instruction_lbl,
                             alignment=Qt.AlignmentFlag.AlignRight)

        preview_layout.addLayout(header_row)

        self.active_item_title = QLabel("Ningún canto seleccionado")
        self.active_item_title.setObjectName("previewSongTitle")
        preview_layout.addWidget(self.active_item_title)

        # Lista de tarjetas
        self.preview_list = PreviewListWidget()
        self.preview_list.setObjectName("previewList")
        self.preview_list.itemClicked.connect(self.on_preview_card_clicked)
        preview_layout.addWidget(self.preview_list)

        # Controles del operador
        divider = QFrame()
        divider.setFrameShape(QFrame.Shape.HLine)
        divider.setFrameShadow(QFrame.Shadow.Sunken)
        divider.setStyleSheet(
            "background-color: #27272a; max-height: 1px; margin: 10px 0px;")
        preview_layout.addWidget(divider)

        controls_layout = QHBoxLayout()
        controls_layout.setSpacing(6)

        self.btn_clear_black = QPushButton("Negro (Black)")
        self.btn_clear_black.setObjectName("subtleCtrlBtnDanger")
        self.btn_clear_black.clicked.connect(self.project_black)
        controls_layout.addWidget(self.btn_clear_black)

        self.btn_clear_text = QPushButton("Limpiar Letra")
        self.btn_clear_text.setObjectName("subtleCtrlBtn")
        self.btn_clear_text.clicked.connect(self.project_clear_text)
        controls_layout.addWidget(self.btn_clear_text)

        self.btn_restore = QPushButton("Restaurar")
        self.btn_restore.setObjectName("subtleCtrlBtn")
        self.btn_restore.clicked.connect(self.project_restore)
        controls_layout.addWidget(self.btn_restore)

        self.btn_toggle_projector = QPushButton("V. Externa")
        self.btn_toggle_projector.setObjectName("subtleCtrlBtn")
        self.btn_toggle_projector.clicked.connect(
            self.toggle_projection_window)
        controls_layout.addWidget(self.btn_toggle_projector)

        self.btn_lyrics_bg = QPushButton("Fondo Letra")
        self.btn_lyrics_bg.setObjectName("subtleCtrlBtn")
        self.btn_lyrics_bg.clicked.connect(self.change_lyrics_background)
        controls_layout.addWidget(self.btn_lyrics_bg)

        preview_layout.addLayout(controls_layout)
        splitter.addWidget(preview_container)

        # ---- Panel Derecho: Biblioteca Stacked + Ventana OBS ----
        right_container = QWidget()
        right_layout = QVBoxLayout(right_container)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(15)

        # Stacked Widget para la biblioteca (Cerebro)
        self.library_stack = QStackedWidget()

        self.setup_songs_page()
        self.setup_bible_page()
        self.setup_images_page()
        self.setup_pptx_page()
        self.setup_guion_page()

        right_layout.addWidget(self.library_stack, stretch=2)

        # Panel Inferior: Ventana OBS Limpia y Más Grande
        self.obs_container = QFrame()
        self.obs_container.setObjectName("obsContainerFrame")
        self.obs_container.setStyleSheet("background-color: #000000;")
        self.obs_container.setMinimumHeight(320)

        obs_layout = QVBoxLayout(self.obs_container)
        obs_layout.setContentsMargins(0, 0, 0, 0)
        obs_layout.setSpacing(0)

        self.local_projection_widget = CleanProjectionWidget(is_external=False)
        obs_layout.addWidget(self.local_projection_widget)

        right_layout.addWidget(self.obs_container, stretch=4)

        right_widget_wrap = QWidget()
        right_widget_wrap.setLayout(right_layout)
        splitter.addWidget(right_widget_wrap)

        splitter.setSizes([420, 780])

        self.set_active_nav_button("Cantos")

    # =========================================================================
    # PAGINAS DE LA BIBLIOTECA
    # =========================================================================

    def setup_songs_page(self):
        container = QFrame()
        container.setObjectName("libraryPanel")
        layout = QVBoxLayout(container)
        layout.setContentsMargins(15, 15, 15, 15)

        # Fila superior: buscador + botón nuevo
        top_row = QHBoxLayout()
        top_row.setSpacing(10)

        self.song_search = QLineEdit()
        self.song_search.setObjectName("searchBar")
        self.song_search.setPlaceholderText("🔎 Buscar cantos...")
        self.song_search.textChanged.connect(self.filter_songs)
        top_row.addWidget(self.song_search)

        new_song_btn = QPushButton("+ Nuevo Canto")
        new_song_btn.setObjectName("importBtn")
        new_song_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        new_song_btn.clicked.connect(self.open_new_song_dialog)
        top_row.addWidget(new_song_btn)

        layout.addLayout(top_row)

        # Lista de canciones
        self.songs_list = QListWidget()
        self.songs_list.setObjectName("libraryList")
        self.songs_list.itemClicked.connect(self.on_library_song_clicked)
        self.songs_list.itemDoubleClicked.connect(
            self.on_library_song_double_clicked)
        layout.addWidget(self.songs_list)

        self.library_stack.addWidget(container)

    def setup_placeholder_page(self, name):
        container = QFrame()
        container.setObjectName("libraryPanel")
        layout = QVBoxLayout(container)
        layout.setContentsMargins(15, 15, 15, 15)

        lbl = QLabel(f"📖 Sección de {name} (Próximamente)")
        lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl.setStyleSheet(
            "color: #71717a; font-size: 14px; font-weight: bold;")
        layout.addWidget(lbl)

        self.library_stack.addWidget(container)

    # =========================================================================
    # MÓDULO DE BIBLIA
    # =========================================================================

    # Libros agrupados por testamento, en orden canónico (coincide con import_bible.py)
    BIBLE_BOOKS_OT = [
        "Génesis", "Éxodo", "Levítico", "Números", "Deuteronomio",
        "Josué", "Jueces", "Rut", "1 Samuel", "2 Samuel",
        "1 Reyes", "2 Reyes", "1 Crónicas", "2 Crónicas",
        "Esdras", "Nehemías", "Ester", "Job", "Salmos",
        "Proverbios", "Eclesiastés", "Cantares", "Isaías",
        "Jeremías", "Lamentaciones", "Ezequiel", "Daniel",
        "Oseas", "Joel", "Amós", "Abdías", "Jonás",
        "Miqueas", "Nahúm", "Habacuc", "Sofonías", "Hageo",
        "Zacarías", "Malaquías",
    ]
    BIBLE_BOOKS_NT = [
        "Mateo", "Marcos", "Lucas", "Juan", "Hechos",
        "Romanos", "1 Corintios", "2 Corintios", "Gálatas",
        "Efesios", "Filipenses", "Colosenses", "1 Tesalonicenses",
        "2 Tesalonicenses", "1 Timoteo", "2 Timoteo", "Tito",
        "Filemón", "Hebreos", "Santiago", "1 Pedro", "2 Pedro",
        "1 Juan", "2 Juan", "3 Juan", "Judas", "Apocalipsis",
    ]

    def setup_bible_page(self):
        container = QFrame()
        container.setObjectName("libraryPanel")
        layout = QVBoxLayout(container)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(6)

        # Estado interno del módulo biblia
        self.bible_current_book = "Juan"
        self.bible_current_version = "RVR60"
        self.bible_search_mode = False  # False = capítulos, True = sugerencias

        # ---- Fila: selector de libro + versiones ----
        selector_row = QHBoxLayout()
        selector_row.setSpacing(8)

        self.bible_book_combo = QComboBox()
        self.bible_book_combo.setObjectName("bibleBookCombo")
        self.bible_book_combo.setStyleSheet("""
            QComboBox {
                background-color: #27272a;
                color: #ffffff;
                border: 1px solid #3f3f46;
                border-radius: 8px;
                padding: 8px 12px;
                font-size: 13px;
                font-weight: bold;
            }
            QComboBox::drop-down { border: none; }
            QComboBox QAbstractItemView {
                background-color: #27272a;
                color: #f4f4f5;
                selection-background-color: #3f3f46;
            }
        """)
        self.bible_book_combo.setMinimumWidth(160)
        self._populate_book_combo()
        self.bible_book_combo.currentTextChanged.connect(
            self.on_bible_book_changed)
        selector_row.addWidget(self.bible_book_combo)

        selector_row.addStretch()

        # Pastillas de versión
        self.bible_version_buttons = {}
        for version in ["RVR60", "NVI", "LBLA", "DHH"]:
            btn = QPushButton(version)
            btn.setCheckable(True)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.setStyleSheet(self._version_btn_style(False))
            btn.clicked.connect(
                lambda checked, v=version: self.on_bible_version_changed(v))
            selector_row.addWidget(btn)
            self.bible_version_buttons[version] = btn

        self.bible_version_buttons["RVR60"].setChecked(True)
        self.bible_version_buttons["RVR60"].setStyleSheet(
            self._version_btn_style(True))
        layout.addLayout(selector_row)

        # ---- Stack interno: alterna entre sugerencias y capítulos ----
        self.bible_inner_stack = QStackedWidget()
        layout.addWidget(self.bible_inner_stack, stretch=1)

        # --- Página 0: Panel de sugerencias ---
        suggestions_page = QWidget()
        suggestions_page.setStyleSheet("background: transparent;")
        sug_layout = QVBoxLayout(suggestions_page)
        sug_layout.setContentsMargins(0, 4, 0, 0)
        sug_layout.setSpacing(4)

        self.bible_suggestions_label = QLabel("SUGERENCIAS")
        self.bible_suggestions_label.setStyleSheet(
            "color: #71717a; font-size: 10px; font-weight: bold; "
            "letter-spacing: 0.5px; background: transparent;")
        sug_layout.addWidget(self.bible_suggestions_label)

        self.bible_suggestions_list = QListWidget()
        self.bible_suggestions_list.setObjectName("libraryList")
        self.bible_suggestions_list.setSpacing(2)
        self.bible_suggestions_list.itemClicked.connect(
            self._on_suggestion_clicked)
        sug_layout.addWidget(self.bible_suggestions_list)

        self.bible_inner_stack.addWidget(suggestions_page)  # index 0

        # --- Página 1: Grid de capítulos ---
        chapters_page = QWidget()
        chapters_page.setStyleSheet("background: transparent;")
        ch_layout = QVBoxLayout(chapters_page)
        ch_layout.setContentsMargins(0, 0, 0, 0)
        ch_layout.setSpacing(4)

        self.bible_chapters_label = QLabel("Capítulos · Juan")
        self.bible_chapters_label.setStyleSheet(
            "color: #71717a; font-size: 10px; font-weight: bold; "
            "letter-spacing: 0.5px; background: transparent;")
        ch_layout.addWidget(self.bible_chapters_label)

        self.bible_chapters_scroll = QScrollArea()
        self.bible_chapters_scroll.setWidgetResizable(True)
        self.bible_chapters_scroll.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.bible_chapters_scroll.setStyleSheet(
            "background: transparent; border: none;")

        self.bible_chapters_container = QWidget()
        self.bible_chapters_container.setStyleSheet("background: transparent;")
        self.bible_chapters_grid = QGridLayout(self.bible_chapters_container)
        self.bible_chapters_grid.setSpacing(4)
        self.bible_chapters_grid.setContentsMargins(0, 0, 0, 0)

        self.bible_chapters_scroll.setWidget(self.bible_chapters_container)
        ch_layout.addWidget(self.bible_chapters_scroll, stretch=1)

        self.bible_inner_stack.addWidget(chapters_page)  # index 1
        self.bible_inner_stack.setCurrentIndex(
            1)  # mostrar capítulos por defecto

        self.library_stack.addWidget(container)

        # Cargar capítulos del libro por defecto
        self.load_bible_chapters(self.bible_current_book)

    def _version_btn_style(self, active):
        if active:
            return """
                QPushButton {
                    background-color: #0F6E56;
                    color: #ffffff;
                    border: 1px solid #0F6E56;
                    border-radius: 14px;
                    padding: 5px 12px;
                    font-size: 11px;
                    font-weight: bold;
                }
            """
        return """
            QPushButton {
                background-color: #27272a;
                color: #a1a1aa;
                border: 1px solid #3f3f46;
                border-radius: 14px;
                padding: 5px 12px;
                font-size: 11px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #3f3f46;
                color: #ffffff;
            }
        """

    def _populate_book_combo(self):
        self.bible_book_combo.blockSignals(True)
        self.bible_book_combo.clear()
        self.bible_book_combo.addItem("── Antiguo Testamento ──")
        self.bible_book_combo.model().item(0).setEnabled(False)
        for book in self.BIBLE_BOOKS_OT:
            self.bible_book_combo.addItem(book)
        self.bible_book_combo.addItem("── Nuevo Testamento ──")
        self.bible_book_combo.model().item(
            self.bible_book_combo.count() - 1).setEnabled(False)
        for book in self.BIBLE_BOOKS_NT:
            self.bible_book_combo.addItem(book)
        self.bible_book_combo.setCurrentText(self.bible_current_book)
        self.bible_book_combo.blockSignals(False)

    def on_bible_version_changed(self, version):
        self.bible_current_version = version
        for v, btn in self.bible_version_buttons.items():
            is_active = (v == version)
            btn.setChecked(is_active)
            btn.setStyleSheet(self._version_btn_style(is_active))
        # Recargar capítulos por si la versión no tiene ese libro completo
        self.load_bible_chapters(self.bible_current_book)

    def on_bible_book_changed(self, book_name):
        if book_name.startswith("──"):
            return
        self.bible_current_book = book_name
        self.load_bible_chapters(book_name)

    def load_bible_chapters(self, book_name, highlight_chapter=None):
        """Carga el grid de capítulos disponibles para el libro/versión actual.
        Si highlight_chapter se especifica, ese botón queda resaltado en verde."""
        self.bible_chapters_label.setText(f"Capítulos · {book_name}")

        # Limpiar grid existente
        while self.bible_chapters_grid.count():
            child = self.bible_chapters_grid.takeAt(0)
            if child.widget():
                child.widget().deleteLater()

        conn = sqlite3.connect(database.DB_PATH)
        cursor = conn.cursor()
        cursor.execute(
            "SELECT DISTINCT chapter FROM bible WHERE book = ? AND version = ? ORDER BY chapter",
            (book_name, self.bible_current_version)
        )
        chapters = [row[0] for row in cursor.fetchall()]
        conn.close()

        if not chapters:
            empty_lbl = QLabel(
                f"Sin datos para «{book_name}» en {self.bible_current_version}")
            empty_lbl.setStyleSheet(
                "color: #71717a; font-size: 12px; background: transparent;")
            self.bible_chapters_grid.addWidget(empty_lbl, 0, 0)
            return

        normal_style = """
            QPushButton {
                background-color: #27272a;
                color: #e4e4e7;
                border: 1px solid #3f3f46;
                border-radius: 8px;
                font-size: 13px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #3f3f46;
                color: #ffffff;
            }
        """
        active_style = """
            QPushButton {
                background-color: #0F6E56;
                color: #ffffff;
                border: 1px solid #0F6E56;
                border-radius: 8px;
                font-size: 13px;
                font-weight: bold;
            }
        """

        cols = 8
        for idx, chapter_num in enumerate(chapters):
            row, col = divmod(idx, cols)
            btn = QPushButton(str(chapter_num))
            btn.setFixedSize(38, 34)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.setStyleSheet(active_style if chapter_num ==
                              highlight_chapter else normal_style)
            btn.clicked.connect(lambda checked, b=book_name,
                                c=chapter_num: self.load_bible_chapter_to_preview(b, c))
            self.bible_chapters_grid.addWidget(btn, row, col)

    def add_bible_to_guion(self, book, chapter, verse, text):
        title = f"{book} {chapter}:{verse}"
        if any(b["book"] == book and b["chapter"] == chapter and b["verse"] == verse
               for b in self.guion["bible"]):
            QMessageBox.information(
                self, "Guión", f"«{title}» ya está en el Guión.")
            return
        self.guion["bible"].append({
            "book": book, "chapter": chapter, "verse": verse,
            "text": text, "title": title
        })
        self.refresh_guion_list()

    def on_bible_search_enter(self):
        """Detecta si el texto es una cita válida ('Juan 3 16' o 'Juan 3:16') y navega directo."""
        query = self.topbar_bible_search.text().strip()
        if not query:
            return

        parsed = self._parse_bible_reference(query)
        if not parsed:
            QMessageBox.information(
                self, "Cita no reconocida",
                "Escribe una cita como: Juan 3 16  ó  Juan 3:16"
            )
            return

        book_name, chapter, verse = parsed
        self.load_bible_chapter_to_preview(
            book_name, chapter, focus_verse=verse)

    def _parse_bible_reference(self, query):
        """Intenta extraer (libro, capitulo, versiculo) de un texto libre."""
        text = query.strip()
        # Normalizar separadores : y , a espacios
        text = text.replace(":", " ").replace(",", " ")
        match = re.match(r"^(.*?)\s+(\d+)\s*(\d+)?$", text)
        if not match:
            return None

        raw_book = match.group(1).strip()
        chapter = int(match.group(2))
        verse = int(match.group(3)) if match.group(3) else 1

        # Buscar el libro de forma flexible (sin acentos, case-insensitive)
        target = remove_accents(raw_book)
        all_books = self.BIBLE_BOOKS_OT + self.BIBLE_BOOKS_NT
        for book in all_books:
            if remove_accents(book) == target or remove_accents(book).startswith(target):
                return book, chapter, verse

        return None

    def load_bible_chapter_to_preview(self, book_name, chapter_num, focus_verse=None):
        """Carga todos los versículos de un capítulo en la Preview izquierda.
        También sincroniza el selector de libro y el grid de capítulos visualmente,
        sin importar si la llamada vino de un click manual o del buscador."""
        conn = sqlite3.connect(database.DB_PATH)
        cursor = conn.cursor()
        cursor.execute(
            "SELECT verse, text FROM bible WHERE book = ? AND chapter = ? AND version = ? ORDER BY verse",
            (book_name, chapter_num, self.bible_current_version)
        )
        verses = cursor.fetchall()
        conn.close()

        if not verses:
            QMessageBox.warning(
                self, "Sin contenido", f"No hay versículos para {book_name} {chapter_num} en {self.bible_current_version}.")
            return

        # Sincronizar estado interno + combo de libro (sin disparar su señal para evitar loops)
        self.bible_current_book = book_name
        if self.bible_book_combo.currentText() != book_name:
            self.bible_book_combo.blockSignals(True)
            self.bible_book_combo.setCurrentText(book_name)
            self.bible_book_combo.blockSignals(False)

        # Refrescar grid de capítulos con el capítulo actual resaltado en verde
        self.load_bible_chapters(book_name, highlight_chapter=chapter_num)

        chapter_title = f"{book_name} {chapter_num}"
        self.active_item_title.setText(chapter_title)
        self.preview_list.clear()

        focus_item = None
        for verse_num, verse_text in verses:
            header = f"Versículo {verse_num}"
            reference = f"{book_name} {chapter_num}:{verse_num}"

            list_item = QListWidgetItem(self.preview_list)
            list_item.setData(Qt.ItemDataRole.UserRole, {
                "mode": "text",
                "text": verse_text,
                "header": reference,
                "song_title": f"{book_name} {chapter_num} · {self.bible_current_version}"
            })

            # Contenedor: tarjeta + botón "+ Guión" a la derecha
            wrapper = QWidget()
            wrapper.setStyleSheet("background: transparent;")
            h = QHBoxLayout(wrapper)
            h.setContentsMargins(0, 0, 0, 0)
            h.setSpacing(6)

            card_widget = PreviewCardWidget(header=header, lyrics=verse_text)
            h.addWidget(card_widget, stretch=1)

            wrapper.set_selected = card_widget.set_selected

            guion_btn = QPushButton("+ Guión")
            guion_btn.setObjectName("addToGuionBtn")
            guion_btn.setFixedWidth(70)
            guion_btn.setCursor(Qt.CursorShape.PointingHandCursor)
            guion_btn.clicked.connect(
                lambda _, b=book_name, c=chapter_num, v=verse_num, t=verse_text:
                self.add_bible_to_guion(b, c, v, t)
            )
            h.addWidget(guion_btn, alignment=Qt.AlignmentFlag.AlignVCenter)

            list_item.setSizeHint(wrapper.minimumSizeHint())
            self.preview_list.setItemWidget(list_item, wrapper)

            if focus_verse and verse_num == focus_verse:
                focus_item = list_item

        # Seleccionar y proyectar: el versículo buscado, o el primero por defecto
        target_item = focus_item if focus_item else self.preview_list.item(0)
        if target_item:
            self.preview_list.setCurrentItem(target_item)
            self.preview_list.scrollToItem(target_item)
            self.on_preview_card_clicked(target_item)

    # =========================================================================
    # MÓDULO DE GUIÓN
    # =========================================================================

    # Paleta visual por tipo de item
    GUION_COLORS = {
        "song":   {"bg": "#1e1b4b", "border": "#6366f1", "icon_bg": "#4f46e5", "icon_fg": "#ffffff", "icon": "♫",  "label": "Canto"},
        "bible":  {"bg": "#052e16", "border": "#16a34a", "icon_bg": "#15803d", "icon_fg": "#ffffff", "icon": "📖", "label": "Biblia"},
        "image":  {"bg": "#431407", "border": "#ea580c", "icon_bg": "#c2410c", "icon_fg": "#ffffff", "icon": "🖼", "label": "Imagen"},
        "pptx":   {"bg": "#3b0764", "border": "#a855f7", "icon_bg": "#9333ea", "icon_fg": "#ffffff", "icon": "🖥", "label": "Presentación"},
    }

    def setup_guion_page(self):
        container = QFrame()
        container.setObjectName("libraryPanel")
        layout = QVBoxLayout(container)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(10)

        # Header
        header_row = QHBoxLayout()
        title_lbl = QLabel("GUIÓN DEL SERVICIO")
        title_lbl.setStyleSheet(
            "color: #71717a; font-size: 11px; font-weight: bold; letter-spacing: 0.5px; background: transparent;")
        header_row.addWidget(title_lbl)
        header_row.addStretch()

        self.guion_count_lbl = QLabel("0 ítems")
        self.guion_count_lbl.setStyleSheet(
            "color: #52525b; font-size: 11px; background: transparent;")
        header_row.addWidget(self.guion_count_lbl)

        clear_btn = QPushButton("🗑 Limpiar todo")
        clear_btn.setStyleSheet("""
            QPushButton {
                background-color: #27272a;
                color: #a1a1aa;
                border: 1px solid #3f3f46;
                border-radius: 6px;
                padding: 4px 10px;
                font-size: 11px;
            }
            QPushButton:hover { background-color: #7f1d1d; color: #fca5a5; border-color: #7f1d1d; }
        """)
        clear_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        clear_btn.clicked.connect(self.guion_clear_all)
        header_row.addWidget(clear_btn)

        layout.addLayout(header_row)

        # Lista del guión
        self.guion_list = QListWidget()
        self.guion_list.setObjectName("libraryList")
        self.guion_list.setSpacing(3)
        self.guion_list.itemClicked.connect(self.on_guion_item_clicked)
        self.guion_list.itemDoubleClicked.connect(
            self.on_guion_item_double_clicked)
        layout.addWidget(self.guion_list)

        # Estado vacío
        self.guion_empty_lbl = QLabel(
            "📋  Agrega cantos, versículos, imágenes o\npresentaciones desde las otras pestañas")
        self.guion_empty_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.guion_empty_lbl.setStyleSheet(
            "color: #52525b; font-size: 13px; background: transparent; line-height: 1.6;")
        layout.addWidget(self.guion_empty_lbl)

        self.library_stack.addWidget(container)

    def refresh_guion_list(self):
        """Reconstruye la lista visual del Guión desde self.guion."""
        self.guion_list.clear()

        all_items = (
            [("song",  item) for item in self.guion["songs"]] +
            [("bible", item) for item in self.guion["bible"]] +
            [("image", item) for item in self.guion["images"]] +
            [("pptx",  item) for item in self.guion["pptx"]]
        )

        total = len(all_items)
        self.guion_count_lbl.setText(
            f"{total} ítem{'s' if total != 1 else ''}")
        self.guion_empty_lbl.setVisible(total == 0)
        self.guion_list.setVisible(total > 0)

        for item_type, item_data in all_items:
            colors = self.GUION_COLORS[item_type]
            list_item = QListWidgetItem(self.guion_list)
            list_item.setData(Qt.ItemDataRole.UserRole, {
                              "type": item_type, "data": item_data})

            widget = self._make_guion_row(item_type, item_data, colors)
            list_item.setSizeHint(widget.minimumSizeHint())
            self.guion_list.setItemWidget(list_item, widget)

    def _make_guion_row(self, item_type, item_data, colors):
        """Crea el widget visual de una fila del Guión."""
        row = QFrame()
        row.setStyleSheet(f"""
            QFrame {{
                background-color: {colors['bg']};
                border: 1px solid {colors['border']};
                border-radius: 8px;
            }}
        """)
        layout = QHBoxLayout(row)
        layout.setContentsMargins(10, 8, 10, 8)
        layout.setSpacing(10)

        # Icono circular
        icon_lbl = QLabel(colors["icon"])
        icon_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon_lbl.setFixedSize(34, 34)
        icon_lbl.setStyleSheet(f"""
            background-color: {colors['icon_bg']};
            color: {colors['icon_fg']};
            border-radius: 17px;
            font-size: 14px;
        """)
        icon_lbl.setAttribute(
            Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        layout.addWidget(icon_lbl)

        # Textos
        text_col = QVBoxLayout()
        text_col.setSpacing(2)

        name_lbl = QLabel(item_data.get(
            "title", item_data.get("name", "Sin título")))
        name_lbl.setStyleSheet(
            "color: #f4f4f5; font-weight: bold; font-size: 13px; background: transparent;")
        name_lbl.setAttribute(
            Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        text_col.addWidget(name_lbl)

        sub_lbl = QLabel(colors["label"])
        sub_lbl.setStyleSheet(
            f"color: {colors['border']}; font-size: 10px; font-weight: bold; background: transparent;")
        sub_lbl.setAttribute(
            Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        text_col.addWidget(sub_lbl)

        layout.addLayout(text_col)
        layout.addStretch()

        # Botón eliminar
        del_btn = QPushButton("✕")
        del_btn.setFixedSize(26, 26)
        del_btn.setStyleSheet("""
            QPushButton {
                background-color: rgba(255,255,255,0.07);
                color: #a1a1aa;
                border: none;
                border-radius: 13px;
                font-size: 11px;
                font-weight: bold;
            }
            QPushButton:hover { background-color: #7f1d1d; color: #fca5a5; }
        """)
        del_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        del_btn.clicked.connect(
            lambda: self.guion_remove_item(item_type, item_data))
        layout.addWidget(del_btn)

        return row

    def guion_remove_item(self, item_type, item_data):
        reply = QMessageBox.question(
            self, "Quitar del Guión",
            f"¿Quitar «{item_data.get('title', item_data.get('name', ''))}» del Guión?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Cancel
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        key_map = {"song": "songs", "bible": "bible",
                   "image": "images", "pptx": "pptx"}
        key = key_map[item_type]
        if item_data in self.guion[key]:
            self.guion[key].remove(item_data)
        self.refresh_guion_list()
        self._update_guion_buttons()

    def guion_clear_all(self):
        total = sum(len(v) for v in self.guion.values())
        if total == 0:
            return
        reply = QMessageBox.question(
            self, "Limpiar Guión",
            "¿Limpiar todos los ítems del Guión?\nEl programa no los guardará al cerrarse.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Cancel
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        self.guion = {"songs": [], "bible": [], "images": [], "pptx": []}
        self.refresh_guion_list()
        self._update_guion_buttons()

    def _update_guion_buttons(self):
        """Actualiza el estado visual de los botones + Guión en todas las pestañas."""
        # Para cantos
        for i in range(self.songs_list.count()):
            item = self.songs_list.item(i)
            widget = self.songs_list.itemWidget(item)
            if widget:
                data = item.data(Qt.ItemDataRole.UserRole)
                if data:
                    in_guion = any(s["id"] == data["id"]
                                   for s in self.guion["songs"])
                    if in_guion:
                        widget.add_btn.setText("✓ En Guión")
                        widget.add_btn.setStyleSheet("""
                            QPushButton {
                                background-color: #052e16;
                                color: #4ade80;
                                border: 1px solid #16a34a;
                                border-radius: 14px;
                                padding: 6px 14px;
                                font-size: 11px;
                                font-weight: bold;
                            }
                        """)
                    else:
                        widget.add_btn.setText("+ Guión")
                        widget.add_btn.setObjectName("addToGuionBtn")
                        widget.add_btn.setStyleSheet("")
                        widget.add_btn.style().unpolish(widget.add_btn)
                        widget.add_btn.style().polish(widget.add_btn)

    def on_guion_item_clicked(self, list_item):
        """Click simple → carga en preview."""
        data = list_item.data(Qt.ItemDataRole.UserRole)
        if not data:
            return
        self._load_guion_item_to_preview(data["type"], data["data"])

    def on_guion_item_double_clicked(self, list_item):
        """Doble click → carga en preview y proyecta directo."""
        data = list_item.data(Qt.ItemDataRole.UserRole)
        if not data:
            return
        self._load_guion_item_to_preview(
            data["type"], data["data"], auto_project=True)

    def _load_guion_item_to_preview(self, item_type, item_data, auto_project=False):
        """Carga un ítem del Guión en la preview, según su tipo."""
        if item_type == "song":
            self.load_song_to_preview_by_id(item_data["id"])
            if auto_project and self.preview_list.count() > 0:
                self.preview_list.setCurrentRow(0)
                self.on_preview_card_clicked(self.preview_list.item(0))

        elif item_type == "bible":
            self.load_bible_chapter_to_preview(
                item_data["book"], item_data["chapter"],
                focus_verse=item_data.get("verse")
            )

        elif item_type == "image":
            self.load_image_to_preview(
                item_data["name"], item_data["file_path"])
            if auto_project and self.preview_list.count() > 0:
                self.preview_list.setCurrentRow(0)
                self.on_preview_card_clicked(self.preview_list.item(0))

        elif item_type == "pptx":
            self.load_pptx_to_preview(item_data["id"], item_data["name"])
            if auto_project and self.preview_list.count() > 0:
                self.preview_list.setCurrentRow(0)
                self.on_preview_card_clicked(self.preview_list.item(0))

    def setup_images_page(self):
        container = QFrame()
        container.setObjectName("libraryPanel")
        layout = QVBoxLayout(container)
        layout.setContentsMargins(15, 15, 15, 15)

        # Barra superior con buscador y botón de importar
        top_layout = QHBoxLayout()
        top_layout.setSpacing(10)

        # Buscador de imágenes
        self.image_search = QLineEdit()
        self.image_search.setObjectName("searchBar")
        self.image_search.setPlaceholderText("🔎 Buscar imágenes...")
        self.image_search.textChanged.connect(self.filter_images)
        top_layout.addWidget(self.image_search)

        # Botón de importar
        self.import_btn = QPushButton("Importar desde PC")
        self.import_btn.setObjectName("importBtn")
        self.import_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.import_btn.clicked.connect(self.import_image_from_pc)
        top_layout.addWidget(self.import_btn)

        layout.addLayout(top_layout)

        # Lista de imágenes
        self.images_list = QListWidget()
        self.images_list.setObjectName("libraryList")
        self.images_list.itemClicked.connect(self.on_library_image_clicked)
        self.images_list.itemDoubleClicked.connect(
            self.on_library_image_double_clicked)
        layout.addWidget(self.images_list)

        self.library_stack.addWidget(container)

    def load_images_library(self, filter_text=""):
        self.images_list.clear()
        conn = sqlite3.connect(database.DB_PATH)
        conn.create_function("remove_accents", 1, remove_accents)
        cursor = conn.cursor()

        if filter_text:
            normalized_search = remove_accents(filter_text)
            cursor.execute(
                "SELECT id, name, file_path FROM images "
                "WHERE remove_accents(name) LIKE ? "
                "ORDER BY name",
                (f"%{normalized_search}%",)
            )
        else:
            cursor.execute(
                "SELECT id, name, file_path FROM images ORDER BY name")

        rows = cursor.fetchall()
        for row in rows:
            image_id, name, file_path = row

            abs_path = file_path
            if not os.path.isabs(file_path):
                abs_path = os.path.join(PROJECT_DIR, file_path)

            item = QListWidgetItem(self.images_list)
            item.setData(Qt.ItemDataRole.UserRole, {
                         "id": image_id, "name": name, "file_path": abs_path})

            def make_add_callback(iid=image_id, iname=name, ipath=abs_path):
                return lambda: self.add_image_to_guion(iid, iname, ipath)

            widget = LibraryImageRowWidget(
                title=name,
                file_path=abs_path,
                on_add_clicked=make_add_callback()
            )
            item.setSizeHint(widget.minimumSizeHint())
            self.images_list.setItemWidget(item, widget)

        conn.close()

    def filter_images(self):
        self.load_images_library(self.image_search.text())

    def import_image_from_pc(self):
        file_dialog = QFileDialog(self)
        file_dialog.setWindowTitle("Seleccionar Imágenes")
        file_dialog.setFileMode(QFileDialog.FileMode.ExistingFiles)
        file_dialog.setNameFilter("Imágenes (*.png *.jpg *.jpeg *.bmp *.webp)")

        if file_dialog.exec():
            selected_files = file_dialog.selectedFiles()
            if not selected_files:
                return

            imported_dir = os.path.join(PROJECT_DIR, "imported_images")
            if not os.path.exists(imported_dir):
                os.makedirs(imported_dir)

            conn = sqlite3.connect(database.DB_PATH)
            cursor = conn.cursor()

            imported_count = 0
            for file_path in selected_files:
                if not os.path.exists(file_path):
                    continue

                base_name = os.path.basename(file_path)
                dest_path = os.path.join(imported_dir, base_name)

                # Evitar sobreescribir renombrando
                if os.path.exists(dest_path):
                    name_part, ext_part = os.path.splitext(base_name)
                    counter = 1
                    while os.path.exists(os.path.join(imported_dir, f"{name_part}_{counter}{ext_part}")):
                        counter += 1
                    base_name = f"{name_part}_{counter}{ext_part}"
                    dest_path = os.path.join(imported_dir, base_name)

                try:
                    shutil.copy2(file_path, dest_path)
                    relative_path = os.path.join("imported_images", base_name)
                    cursor.execute(
                        "INSERT INTO images (name, file_path) VALUES (?, ?)", (base_name, relative_path))
                    imported_count += 1
                except Exception as e:
                    QMessageBox.warning(
                        self, "Error de Importación", f"No se pudo copiar {base_name}: {str(e)}")

            conn.commit()
            conn.close()

            if imported_count > 0:
                self.load_images_library()
                QMessageBox.information(
                    self, "Importación", f"Se importaron {imported_count} imágenes exitosamente.")

    def on_library_image_clicked(self, item):
        data = item.data(Qt.ItemDataRole.UserRole)
        if data:
            self.load_image_to_preview(data["name"], data["file_path"])

    def on_library_image_double_clicked(self, item):
        data = item.data(Qt.ItemDataRole.UserRole)
        if data:
            self.load_image_to_preview(data["name"], data["file_path"])
            if self.preview_list.count() > 0:
                first_item = self.preview_list.item(0)
                self.preview_list.setCurrentItem(first_item)
                self.on_preview_card_clicked(first_item)

    def load_image_to_preview(self, name, file_path):
        self.active_item_title.setText(name)
        self.preview_list.clear()

        self.active_item_data = {
            "type": "image",
            "title": name,
            "file_path": file_path
        }

        # Insertar una única tarjeta de visualización de la imagen
        list_item = QListWidgetItem(self.preview_list)
        list_item.setData(Qt.ItemDataRole.UserRole, {
            "mode": "image",
            "file_path": file_path,
            "header": "Imagen",
            "name": name
        })

        card_widget = PreviewImageCardWidget(
            header="Imagen", image_path=file_path)
        list_item.setSizeHint(card_widget.minimumSizeHint())
        self.preview_list.setItemWidget(list_item, card_widget)

        # Seleccionar visualmente y enfocar para la navegación por teclado
        if self.preview_list.count() > 0:
            self.preview_list.setCurrentRow(0)
            self.preview_list.setFocus()

    def add_image_to_guion(self, image_id, name, file_path):
        if any(i["id"] == image_id for i in self.guion["images"]):
            QMessageBox.information(
                self, "Guión", f"«{name}» ya está en el Guión.")
            return
        self.guion["images"].append(
            {"id": image_id, "name": name, "file_path": file_path})
        self.refresh_guion_list()

    def setup_pptx_page(self):
        container = QFrame()
        container.setObjectName("libraryPanel")
        layout = QVBoxLayout(container)
        layout.setContentsMargins(15, 15, 15, 15)

        # Barra superior con buscador y botón de importar
        top_layout = QHBoxLayout()
        top_layout.setSpacing(10)

        # Buscador de presentaciones
        self.pptx_search = QLineEdit()
        self.pptx_search.setObjectName("searchBar")
        self.pptx_search.setPlaceholderText("🔎 Buscar presentaciones...")
        self.pptx_search.textChanged.connect(self.filter_pptx)
        top_layout.addWidget(self.pptx_search)

        # Botón de importar PPTX o PDF
        self.import_pptx_btn = QPushButton("Importar PPTX / PDF")
        self.import_pptx_btn.setObjectName("importBtn")
        self.import_pptx_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.import_pptx_btn.clicked.connect(self.import_pptx_from_pc)
        top_layout.addWidget(self.import_pptx_btn)

        layout.addLayout(top_layout)

        # Lista de PPTX
        self.pptx_list = QListWidget()
        self.pptx_list.setObjectName("libraryList")
        self.pptx_list.itemClicked.connect(self.on_library_pptx_clicked)
        self.pptx_list.itemDoubleClicked.connect(
            self.on_library_pptx_double_clicked)
        layout.addWidget(self.pptx_list)

        self.library_stack.addWidget(container)

    def load_pptx_library(self, filter_text=""):
        self.pptx_list.clear()
        conn = sqlite3.connect(database.DB_PATH)
        conn.create_function("remove_accents", 1, remove_accents)
        cursor = conn.cursor()

        if filter_text:
            normalized_search = remove_accents(filter_text)
            cursor.execute(
                "SELECT id, name, file_path FROM pptx_files "
                "WHERE remove_accents(name) LIKE ? "
                "ORDER BY name",
                (f"%{normalized_search}%",)
            )
        else:
            cursor.execute(
                "SELECT id, name, file_path FROM pptx_files ORDER BY name")

        rows = cursor.fetchall()
        for row in rows:
            pptx_id, name, file_path = row

            item = QListWidgetItem(self.pptx_list)
            item.setData(Qt.ItemDataRole.UserRole, {
                         "id": pptx_id, "name": name, "file_path": file_path})

            def make_add_callback(pid=pptx_id, pname=name):
                return lambda: self.add_pptx_to_guion(pid, pname)

            def make_delete_callback(pid=pptx_id, pname=name):
                return lambda: self.delete_pptx(pid, pname)

            file_type = "pdf" if name.lower().endswith(".pdf") else "pptx"

            widget = LibraryPptxRowWidget(
                title=name,
                on_add_clicked=make_add_callback(),
                on_delete_clicked=make_delete_callback(),
                file_type=file_type
            )
            item.setSizeHint(widget.minimumSizeHint())
            self.pptx_list.setItemWidget(item, widget)

        conn.close()

    def filter_pptx(self):
        self.load_pptx_library(self.pptx_search.text())

    def delete_pptx(self, pptx_id, name):
        reply = QMessageBox.question(
            self,
            "Eliminar presentación",
            f"¿Estás seguro que quieres eliminar «{name}»?\nEsta acción no se puede deshacer.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Cancel
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        conn = sqlite3.connect(database.DB_PATH)
        cursor = conn.cursor()
        cursor.execute(
            "SELECT file_path FROM pptx_files WHERE id = ?", (pptx_id,))
        row = cursor.fetchone()
        cursor.execute("DELETE FROM pptx_files WHERE id = ?", (pptx_id,))
        conn.commit()
        conn.close()

        # Borrar también el archivo original y la carpeta de diapositivas generadas en disco
        if row:
            file_path = row[0]
            if file_path:
                abs_file_path = os.path.join(PROJECT_DIR, file_path)
                if os.path.exists(abs_file_path):
                    try:
                        os.remove(abs_file_path)
                    except Exception:
                        pass

            slides_dir = os.path.join(
                PROJECT_DIR, "imported_pptx", f"pptx_{pptx_id}_slides")
            if os.path.exists(slides_dir):
                try:
                    shutil.rmtree(slides_dir)
                except Exception:
                    pass

        self.preview_list.clear()
        self.load_pptx_library(self.pptx_search.text())

    def import_pptx_from_pc(self):
        file_dialog = QFileDialog(self)
        file_dialog.setWindowTitle(
            "Seleccionar Presentaciones (PowerPoint o PDF)")
        file_dialog.setFileMode(QFileDialog.FileMode.ExistingFiles)
        file_dialog.setNameFilter("Presentaciones (*.pptx *.ppt *.pdf)")

        if file_dialog.exec():
            selected_files = file_dialog.selectedFiles()
            if not selected_files:
                return

            imported_dir = os.path.join(PROJECT_DIR, "imported_pptx")
            if not os.path.exists(imported_dir):
                os.makedirs(imported_dir)

            conn = sqlite3.connect(database.DB_PATH)
            cursor = conn.cursor()

            imported_count = 0
            failed_files = []

            for file_path in selected_files:
                if not os.path.exists(file_path):
                    continue

                base_name = os.path.basename(file_path)
                file_ext = os.path.splitext(base_name)[1].lower()

                # Registrar primero en la DB para obtener el ID único
                cursor.execute(
                    "INSERT INTO pptx_files (name, file_path) VALUES (?, ?)", (base_name, ""))
                pptx_id = cursor.lastrowid

                # Crear ruta de destino con el ID para evitar colisiones
                dest_name = f"pptx_{pptx_id}{file_ext}"
                dest_path = os.path.join(imported_dir, dest_name)

                try:
                    # Copiar archivo original (pptx, ppt o pdf)
                    shutil.copy2(file_path, dest_path)

                    # Guardar la ruta relativa en la base de datos
                    relative_path = os.path.join("imported_pptx", dest_name)
                    cursor.execute(
                        "UPDATE pptx_files SET file_path = ? WHERE id = ?", (relative_path, pptx_id))

                    # Convertir a imágenes en una carpeta dedicada, según el tipo de archivo
                    slides_dir = os.path.join(
                        imported_dir, f"pptx_{pptx_id}_slides")

                    if file_ext == ".pdf":
                        slide_images = pdf_helper.convert_pdf_to_images(
                            dest_path, slides_dir)
                    else:
                        slide_images = pptx_helper.convert_pptx_to_images(
                            dest_path, slides_dir)

                    if not slide_images:
                        failed_files.append(base_name)
                    else:
                        imported_count += 1

                except Exception as e:
                    failed_files.append(base_name)
                    QMessageBox.warning(
                        self, "Error de Importación", f"No se pudo procesar {base_name}: {str(e)}")

            conn.commit()
            conn.close()

            if imported_count > 0:
                self.load_pptx_library()
                QMessageBox.information(
                    self, "Importación", f"Se importaron y convirtieron {imported_count} presentaciones exitosamente.")

            if failed_files:
                names = "\n".join(failed_files)
                QMessageBox.warning(
                    self, "Algunos archivos no se pudieron convertir",
                    f"No se generaron diapositivas para:\n{names}\n\n"
                    f"Tip: los archivos .ppt antiguos requieren Microsoft PowerPoint instalado. "
                    f"Prueba exportando a .pptx o .pdf."
                )

    def on_library_pptx_clicked(self, item):
        data = item.data(Qt.ItemDataRole.UserRole)
        if data:
            self.load_pptx_to_preview(data["id"], data["name"])

    def on_library_pptx_double_clicked(self, item):
        data = item.data(Qt.ItemDataRole.UserRole)
        if data:
            self.load_pptx_to_preview(data["id"], data["name"])
            if self.preview_list.count() > 0:
                first_item = self.preview_list.item(0)
                self.preview_list.setCurrentItem(first_item)
                self.on_preview_card_clicked(first_item)

    def load_pptx_to_preview(self, pptx_id, name):
        self.active_item_title.setText(name)
        self.preview_list.clear()

        slides_dir = os.path.join(
            PROJECT_DIR, "imported_pptx", f"pptx_{pptx_id}_slides")
        if not os.path.exists(slides_dir):
            return

        # Buscar todas las imágenes de diapositiva en el directorio
        try:
            files = os.listdir(slides_dir)
        except Exception:
            return

        png_files = [f for f in files if f.lower().endswith(".png")]

        # Ordenar numéricamente (Slide1, Slide2...)
        def extract_number(filename):
            num_str = "".join([c for c in filename if c.isdigit()])
            return int(num_str) if num_str else 0

        png_files.sort(key=extract_number)

        self.active_item_data = {
            "type": "pptx",
            "title": name,
            "slides": [os.path.join(slides_dir, f) for f in png_files]
        }

        for i, png in enumerate(png_files):
            slide_path = os.path.join(slides_dir, png)
            header = f"Diapositiva {i+1}"

            list_item = QListWidgetItem(self.preview_list)
            list_item.setData(Qt.ItemDataRole.UserRole, {
                "mode": "image",
                "file_path": slide_path,
                "header": header,
                "name": name
            })

            card_widget = PreviewImageCardWidget(
                header=header, image_path=slide_path)
            list_item.setSizeHint(card_widget.minimumSizeHint())
            self.preview_list.setItemWidget(list_item, card_widget)

        # Seleccionar visualmente y enfocar para la navegación por teclado
        if self.preview_list.count() > 0:
            self.preview_list.setCurrentRow(0)
            self.preview_list.setFocus()

    def add_pptx_to_guion(self, pptx_id, name):
        if any(p["id"] == pptx_id for p in self.guion["pptx"]):
            QMessageBox.information(
                self, "Guión", f"«{name}» ya está en el Guión.")
            return
        self.guion["pptx"].append({"id": pptx_id, "name": name})
        self.refresh_guion_list()

    def change_lyrics_background(self):
        # Preguntar si desea cambiar o quitar el fondo
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("Fondo General para Letras")
        msg_box.setText(
            "Selecciona una opción para el fondo general de las letras:")

        select_btn = msg_box.addButton(
            "Seleccionar Imagen", QMessageBox.ButtonRole.AcceptRole)
        remove_btn = msg_box.addButton(
            "Quitar Fondo", QMessageBox.ButtonRole.DestructiveRole)
        cancel_btn = msg_box.addButton(
            "Cancelar", QMessageBox.ButtonRole.RejectRole)

        msg_box.exec()

        clicked_btn = msg_box.clickedButton()

        if clicked_btn == select_btn:
            file_dialog = QFileDialog(self)
            file_dialog.setWindowTitle("Seleccionar Imagen de Fondo")
            file_dialog.setNameFilter(
                "Imágenes (*.png *.jpg *.jpeg *.bmp *.webp)")
            if file_dialog.exec():
                selected_files = file_dialog.selectedFiles()
                if selected_files:
                    file_path = selected_files[0]
                    # Copiar a imported_images para hacerlo permanente y portable
                    imported_dir = os.path.join(PROJECT_DIR, "imported_images")
                    if not os.path.exists(imported_dir):
                        os.makedirs(imported_dir)

                    base_name = "general_lyrics_bg" + \
                        os.path.splitext(file_path)[1]
                    dest_path = os.path.join(imported_dir, base_name)

                    try:
                        shutil.copy2(file_path, dest_path)
                        # Guardar ruta absoluta/relativa
                        abs_dest_path = os.path.abspath(dest_path)

                        # Guardar configuración en archivo de texto
                        bg_config_path = os.path.join(
                            PROJECT_DIR, "lyrics_bg_path.txt")
                        with open(bg_config_path, "w", encoding="utf-8") as f:
                            f.write(abs_dest_path)

                        # Cargar el pixmap con DPI correcto para evitar borrosidad
                        pixmap = QPixmap(abs_dest_path)
                        dpr = self.local_projection_widget.devicePixelRatioF()
                        if dpr != 1.0:
                            pixmap.setDevicePixelRatio(dpr)
                        self.local_projection_widget.general_bg_pixmap = pixmap
                        if self.projection_window:
                            self.projection_window.projection_widget.general_bg_pixmap = pixmap

                        self.local_projection_widget.update()
                        if self.projection_window:
                            self.projection_window.projection_widget.update()

                        QMessageBox.information(
                            self, "Fondo Guardado", "Se ha establecido el fondo general para las letras.")
                    except Exception as e:
                        QMessageBox.warning(
                            self, "Error", f"No se pudo copiar la imagen de fondo: {str(e)}")

        elif clicked_btn == remove_btn:
            # Quitar fondo
            bg_config_path = os.path.join(PROJECT_DIR, "lyrics_bg_path.txt")
            if os.path.exists(bg_config_path):
                try:
                    os.remove(bg_config_path)
                except Exception:
                    pass

            self.local_projection_widget.general_bg_pixmap = None
            if self.projection_window:
                self.projection_window.projection_widget.general_bg_pixmap = None

            self.local_projection_widget.update()
            if self.projection_window:
                self.projection_window.projection_widget.update()

            QMessageBox.information(
                self, "Fondo Quitado", "Se ha quitado el fondo general para las letras (ahora es negro sólido).")

    def load_persisted_lyrics_bg(self):
        bg_config_path = os.path.join(PROJECT_DIR, "lyrics_bg_path.txt")
        if os.path.exists(bg_config_path):
            try:
                with open(bg_config_path, "r", encoding="utf-8") as f:
                    path = f.read().strip()
                if path and os.path.exists(path):
                    pixmap = QPixmap(path)
                    if not pixmap.isNull():
                        dpr = self.local_projection_widget.devicePixelRatioF()
                        if dpr != 1.0:
                            pixmap.setDevicePixelRatio(dpr)
                        self.local_projection_widget.general_bg_pixmap = pixmap
                        if self.projection_window:
                            self.projection_window.projection_widget.general_bg_pixmap = pixmap
                        self.local_projection_widget.update()
            except Exception as e:
                print(f"Error cargando fondo de letras persistido: {e}")

    # =========================================================================
    # CARGAR CANTOS A LA BIBLIOTECA (CON NORMALIZACIÓN DE ACENTOS)
    # =========================================================================

    def load_songs_library(self, filter_text=""):
        self.songs_list.clear()
        conn = sqlite3.connect(database.DB_PATH)
        # Registrar la función de normalización en la conexión de SQLite
        conn.create_function("remove_accents", 1, remove_accents)
        cursor = conn.cursor()

        if filter_text:
            normalized_search = remove_accents(filter_text)
            cursor.execute(
                "SELECT id, title, category, lyrics FROM songs "
                "WHERE remove_accents(title) LIKE ? OR remove_accents(lyrics) LIKE ? "
                "ORDER BY title",
                (f"%{normalized_search}%", f"%{normalized_search}%")
            )
        else:
            cursor.execute(
                "SELECT id, title, category, lyrics FROM songs ORDER BY title")

        rows = cursor.fetchall()
        for row in rows:
            song_id, title, category, lyrics = row

            item = QListWidgetItem(self.songs_list)
            item.setData(Qt.ItemDataRole.UserRole, {
                         "id": song_id, "title": title, "category": category})

            def make_add_callback(sid=song_id, stitle=title):
                return lambda: self.add_song_to_guion(sid, stitle)

            def make_edit_callback(sid=song_id):
                return lambda: self.open_new_song_dialog(song_id=sid)

            def make_delete_callback(sid=song_id, stitle=title):
                return lambda: self.delete_song(sid, stitle)

            widget = LibrarySongRowWidget(
                title=title,
                category=category,
                on_add_clicked=make_add_callback(),
                on_edit_clicked=make_edit_callback(),
                on_delete_clicked=make_delete_callback()
            )
            item.setSizeHint(widget.minimumSizeHint())
            self.songs_list.setItemWidget(item, widget)

        conn.close()

    def filter_songs(self):
        self.load_songs_library(self.song_search.text())

    def open_new_song_dialog(self, song_id=None):
        dialog = SongEditorDialog(parent=self, song_id=song_id)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.load_songs_library(self.song_search.text())

    def delete_song(self, song_id, title):
        reply = QMessageBox.question(
            self,
            "Eliminar canto",
            f"¿Estás seguro que quieres eliminar «{title}»?\nEsta acción no se puede deshacer.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Cancel
        )
        if reply == QMessageBox.StandardButton.Yes:
            conn = sqlite3.connect(database.DB_PATH)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM songs WHERE id = ?", (song_id,))
            conn.commit()
            conn.close()
            self.preview_list.clear()
            self.load_songs_library(self.song_search.text())

    # =========================================================================
    # NAVEGACIÓN
    # =========================================================================

    def make_nav_callback(self, name):
        def callback():
            self.set_active_nav_button(name)
            pages = ["Cantos", "Biblias", "Imágenes", "PPTX", "Guión"]
            if name in pages:
                self.library_stack.setCurrentIndex(pages.index(name))
            # Mostrar el buscador del topbar solo en Biblias
            self.topbar_bible_search.setVisible(name == "Biblias")
            if name == "Biblias":
                self.topbar_bible_search.setFocus()
        return callback

    def _on_topbar_bible_search(self):
        """Se llama al dar Enter en el topbar — navega a la cita si es válida."""
        query = self.topbar_bible_search.text().strip()
        if not query:
            self._bible_show_chapters()
            return

        # Intentar como cita completa primero
        parsed = self._parse_bible_reference(query)
        if parsed:
            book_name, chapter, verse = parsed
            self.load_bible_chapter_to_preview(
                book_name, chapter, focus_verse=verse)
            self._bible_show_chapters()
            return

        # Si no es cita completa, actualizar sugerencias
        self._update_bible_suggestions(query)

    def _on_topbar_bible_search_changed(self, text):
        """Se llama en tiempo real conforme escribe — muestra sugerencias."""
        query = text.strip()
        if not query:
            # Sin texto → volver al modo capítulos
            self._bible_show_chapters()
            return
        self._update_bible_suggestions(query)

    def _update_bible_suggestions(self, query):
        """Genera sugerencias de libros, capítulos o versículos según lo escrito."""
        self.bible_suggestions_list.clear()
        query_clean = remove_accents(query.lower())
        all_books = self.BIBLE_BOOKS_OT + self.BIBLE_BOOKS_NT

        suggestions = []

        # Detectar si hay número al final tipo "juan 3" o "juan 3:16" o "juan 3 16"
        # Patrón: texto_libro + número_capítulo + (opcional: número_versículo)
        m = re.match(r"^(.+?)\s+(\d+)(?:[:|\s]+(\d+))?$", query.strip())

        if m:
            raw_book = m.group(1).strip()
            chapter = int(m.group(2))
            verse = int(m.group(3)) if m.group(3) else None
            book_clean = remove_accents(raw_book.lower())

            for book in all_books:
                book_norm = remove_accents(book.lower())
                if book_norm.startswith(book_clean) or book_clean in book_norm:
                    if verse:
                        # Sugerencia con versículo específico
                        label = f"{book}  {chapter}:{verse}"
                        suggestions.append({
                            "type": "verse",
                            "book": book, "chapter": chapter, "verse": verse,
                            "label": label, "sub": f"Ir al versículo · {self.bible_current_version}"
                        })
                    else:
                        # Sugerencia de capítulo dentro del libro
                        label = f"{book}  Capítulo {chapter}"
                        suggestions.append({
                            "type": "chapter",
                            "book": book, "chapter": chapter,
                            "label": label, "sub": f"Ver capítulo · {self.bible_current_version}"
                        })
        else:
            # Solo texto → sugerir libros que coincidan
            for book in all_books:
                book_norm = remove_accents(book.lower())
                if book_norm.startswith(query_clean) or query_clean in book_norm:
                    suggestions.append({
                        "type": "book",
                        "book": book,
                        "label": book,
                        "sub": "Seleccionar libro"
                    })

        if not suggestions:
            item = QListWidgetItem("Sin resultados para «" + query + "»")
            item.setForeground(QColor("#52525b"))
            item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEnabled)
            self.bible_suggestions_list.addItem(item)
        else:
            for sug in suggestions[:12]:  # máximo 12 sugerencias
                self._add_suggestion_row(sug)

        # Mostrar panel de sugerencias
        self.bible_inner_stack.setCurrentIndex(0)

    def _add_suggestion_row(self, sug):
        """Agrega una fila visual de sugerencia a la lista."""
        item = QListWidgetItem(self.bible_suggestions_list)
        item.setData(Qt.ItemDataRole.UserRole, sug)

        row = QWidget()
        row.setStyleSheet("background: transparent;")
        layout = QHBoxLayout(row)
        layout.setContentsMargins(10, 8, 10, 8)
        layout.setSpacing(10)

        # Icono según tipo
        icons = {"book": "📖", "chapter": "📄", "verse": "✦"}
        colors = {"book": "#16a34a", "chapter": "#6366f1", "verse": "#0F6E56"}
        t = sug["type"]

        icon_lbl = QLabel(icons.get(t, "📖"))
        icon_lbl.setFixedSize(30, 30)
        icon_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon_lbl.setStyleSheet(f"""
            background-color: {colors.get(t, '#27272a')};
            border-radius: 15px; font-size: 13px;
        """)
        icon_lbl.setAttribute(
            Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        layout.addWidget(icon_lbl)

        text_col = QVBoxLayout()
        text_col.setSpacing(1)
        lbl = QLabel(sug["label"])
        lbl.setStyleSheet(
            "color: #f4f4f5; font-size: 13px; font-weight: bold; background: transparent;")
        lbl.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        sub = QLabel(sug["sub"])
        sub.setStyleSheet(
            "color: #71717a; font-size: 10px; background: transparent;")
        sub.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        text_col.addWidget(lbl)
        text_col.addWidget(sub)
        layout.addLayout(text_col)
        layout.addStretch()

        row.setAttribute(
            Qt.WidgetAttribute.WA_TransparentForMouseEvents, False)
        item.setSizeHint(row.sizeHint())
        self.bible_suggestions_list.setItemWidget(item, row)

    def _on_suggestion_clicked(self, list_item):
        """Procesa el click en una sugerencia."""
        sug = list_item.data(Qt.ItemDataRole.UserRole)
        if not sug:
            return

        if sug["type"] == "book":
            # Seleccionar el libro y mostrar capítulos
            self.topbar_bible_search.clear()
            self.load_bible_chapter_to_preview(sug["book"], 1)
            self._bible_show_chapters()

        elif sug["type"] == "chapter":
            # Navegar al capítulo
            self.topbar_bible_search.clear()
            self.load_bible_chapter_to_preview(sug["book"], sug["chapter"])
            self._bible_show_chapters()

        elif sug["type"] == "verse":
            # Navegar directo al versículo
            self.topbar_bible_search.clear()
            self.load_bible_chapter_to_preview(
                sug["book"], sug["chapter"], focus_verse=sug["verse"])
            self._bible_show_chapters()

    def _bible_show_chapters(self):
        """Vuelve al modo grid de capítulos."""
        self.bible_inner_stack.setCurrentIndex(1)

    def set_active_nav_button(self, active_name):
        for name, btn in self.nav_buttons.items():
            is_active = (name == active_name)
            btn.setProperty("active", "true" if is_active else "false")
            btn.style().polish(btn)

    # =========================================================================
    # INTERACCIÓN CON LA BIBLIOTECA (UN CLIC Y DOBLE CLIC)
    # =========================================================================

    def on_library_song_clicked(self, item):
        data = item.data(Qt.ItemDataRole.UserRole)
        if data:
            self.load_song_to_preview_by_id(data["id"])

    def on_library_song_double_clicked(self, item):
        data = item.data(Qt.ItemDataRole.UserRole)
        if data:
            # Cargar
            self.load_song_to_preview_by_id(data["id"])
            # Proyectar el título primero automáticamente (primer elemento en la lista)
            if self.preview_list.count() > 0:
                first_item = self.preview_list.item(0)
                self.preview_list.setCurrentItem(first_item)
                self.on_preview_card_clicked(first_item)

    # =========================================================================
    # CARGA A VISTA PREVIA (IZQUIERDA) CON TARJETA DE TÍTULO
    # =========================================================================

    def load_song_to_preview_by_id(self, song_id):
        conn = sqlite3.connect(database.DB_PATH)
        cursor = conn.cursor()
        cursor.execute(
            "SELECT title, lyrics, category FROM songs WHERE id = ?", (song_id,))
        res = cursor.fetchone()
        conn.close()

        if not res:
            return

        title, lyrics, category = res
        self.active_item_title.setText(title)
        self.preview_list.clear()

        slides = [s.strip() for s in lyrics.split("\n\n") if s.strip()]

        self.active_item_data = {
            "type": "song",
            "title": title,
            "category": category,
            "slides": [title] + slides
        }

        # Insertar las secciones (el bloque Título ya viene en lyrics desde el editor)
        for i, slide in enumerate(slides):
            header = f"Estrofa {i+1}"
            first_line = slide.split("\n")[0]
            if first_line.lower().startswith("[") and first_line.lower().endswith("]"):
                header = first_line[1:-1]
                display_content = "\n".join(slide.split("\n")[1:])
            else:
                display_content = slide

            list_item = QListWidgetItem(self.preview_list)
            list_item.setData(Qt.ItemDataRole.UserRole, {
                "mode": "text",
                "text": display_content,
                "header": header,
                "song_title": title
            })

            card_widget = PreviewCardWidget(
                header=header, lyrics=display_content)
            list_item.setSizeHint(card_widget.minimumSizeHint())
            self.preview_list.setItemWidget(list_item, card_widget)

        # Seleccionar visualmente y enfocar para la navegación por teclado
        if self.preview_list.count() > 0:
            self.preview_list.setCurrentRow(0)
            self.preview_list.setFocus()

    def on_preview_card_clicked(self, item):
        for i in range(self.preview_list.count()):
            list_item = self.preview_list.item(i)
            widget = self.preview_list.itemWidget(list_item)
            if hasattr(widget, "set_selected"):
                widget.set_selected(False)

        widget = self.preview_list.itemWidget(item)
        if hasattr(widget, "set_selected"):
            widget.set_selected(True)

        self.project_selected_card(item)

    # =========================================================================
    # ACCIONES DE PROYECCIÓN ESTABLE
    # =========================================================================

    def project_selected_card(self, item):
        data = item.data(Qt.ItemDataRole.UserRole)
        if not data:
            return

        mode = data["mode"]
        if mode == "text":
            text = data["text"]
            header = data.get("header", "")
            song_title = data.get("song_title", "")

            self.current_projection_mode = "text"
            self.last_projected_mode = "text"
            self.last_projected_text = text
            self.last_projected_header = header
            self.last_projected_song_title = song_title

            # Local
            self.local_projection_widget.display_text(text, header, song_title)

            # Externo
            if self.projection_window and self.projection_window.isVisible():
                self.projection_window.projection_widget.display_text(
                    text, header, song_title)
        elif mode == "image":
            file_path = data["file_path"]
            name = data.get("name", "")
            header = data.get("header", "Imagen")

            self.current_projection_mode = "image"
            self.last_projected_mode = "image"
            self.last_projected_image_path = file_path
            self.last_projected_header = header
            self.last_projected_song_title = name

            # Local
            self.local_projection_widget.display_image(file_path, header, name)

            # Externo
            if self.projection_window and self.projection_window.isVisible():
                self.projection_window.projection_widget.display_image(
                    file_path, header, name)

    def project_black(self):
        self.current_projection_mode = "black"
        self.local_projection_widget.set_black_screen()
        if self.projection_window and self.projection_window.isVisible():
            self.projection_window.projection_widget.set_black_screen()

    def project_clear_text(self):
        if self.current_projection_mode == "text":
            self.local_projection_widget.clear_text_only()
            if self.projection_window and self.projection_window.isVisible():
                self.projection_window.projection_widget.clear_text_only()

    def project_restore(self):
        if self.last_projected_mode == "text":
            self.current_projection_mode = "text"
            self.local_projection_widget.display_text(
                self.last_projected_text, self.last_projected_header, self.last_projected_song_title
            )
            if self.projection_window and self.projection_window.isVisible():
                self.projection_window.projection_widget.display_text(
                    self.last_projected_text, self.last_projected_header, self.last_projected_song_title
                )
        elif self.last_projected_mode == "image":
            self.current_projection_mode = "image"
            self.local_projection_widget.display_image(
                self.last_projected_image_path, self.last_projected_header, self.last_projected_song_title
            )
            if self.projection_window and self.projection_window.isVisible():
                self.projection_window.projection_widget.display_image(
                    self.last_projected_image_path, self.last_projected_header, self.last_projected_song_title
                )

    def toggle_projection_window(self):
        if not self.projection_window:
            self.projection_window = ProjectionWindow()
            # Copiar la imagen de fondo actual a la nueva ventana externa
            self.projection_window.projection_widget.general_bg_pixmap = self.local_projection_widget.general_bg_pixmap

        if self.projection_window.isVisible():
            self.projection_window.hide()
            self.btn_toggle_projector.setText("V. Externa")
        else:
            self.projection_window.show()
            self.btn_toggle_projector.setText("Ocultar V. Ext")

            # Sincronizar
            self.projection_window.projection_widget.is_black_screen = self.local_projection_widget.is_black_screen
            if self.current_projection_mode == "text":
                if not self.local_projection_widget.text_label.text():
                    self.projection_window.projection_widget.display_text(
                        "", self.last_projected_header, self.last_projected_song_title
                    )
                else:
                    self.projection_window.projection_widget.display_text(
                        self.last_projected_text, self.last_projected_header, self.last_projected_song_title
                    )
            elif self.current_projection_mode == "image":
                self.projection_window.projection_widget.display_image(
                    self.last_projected_image_path, self.last_projected_header, self.last_projected_song_title
                )
            elif self.current_projection_mode == "black":
                self.projection_window.projection_widget.set_black_screen()

    # =========================================================================
    # GESTIÓN DEL GUION (TEMPORAL)
    # =========================================================================

    def add_song_to_guion(self, song_id, title):
        if any(s["id"] == song_id for s in self.guion["songs"]):
            QMessageBox.information(
                self, "Guión", f"«{title}» ya está en el Guión.")
            return
        self.guion["songs"].append({"id": song_id, "title": title})
        self.refresh_guion_list()
        self._update_guion_buttons()

    def closeEvent(self, event):
        if self.projection_window:
            self.projection_window.close()
        event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)

    palette = app.palette()
    palette.setColor(palette.ColorGroup.All,
                     palette.ColorRole.Window, QColor("#0f0f11"))
    palette.setColor(palette.ColorGroup.All,
                     palette.ColorRole.WindowText, QColor("#f4f4f5"))
    app.setPalette(palette)

    window = MainWindow()
    window.showMaximized()
    sys.exit(app.exec())
