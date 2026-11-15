# ---------------------------------------------------------------------------
# App bootstrap and project setup
# ---------------------------------------------------------------------------
import os
import sys
import shutil
import sqlite3
import re
import unicodedata
from html import escape as escape_html

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QListWidget, QListWidgetItem, QPushButton, QLineEdit, QLabel,
    QFileDialog, QMessageBox, QSplitter, QFrame, QGridLayout, QStackedWidget,
    QDialog, QTextEdit, QScrollArea, QComboBox, QSizePolicy, QMenu, QCheckBox,
    QToolButton, QStyle
)
from PyQt6.QtCore import Qt, QSize, QTimer, QPoint
from PyQt6.QtGui import QPixmap, QColor, QShortcut, QKeySequence, QIcon

import database
import pptx_helper
import pdf_helper


def is_frozen_app():
    return bool(getattr(sys, "frozen", False))


def get_project_dir():
    if is_frozen_app():
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


PROJECT_DIR = get_project_dir()
DB_PATH = os.path.join(PROJECT_DIR, "iglesia.db")
LOGO_PATH = os.path.join(PROJECT_DIR, "station_selah.svg")
if not os.path.exists(LOGO_PATH) and hasattr(sys, "_MEIPASS"):
    LOGO_PATH = os.path.join(sys._MEIPASS, "station_selah.svg")
database.DB_PATH = DB_PATH


def init_project_resources():
    """Ensures local resource folders exist and copies packaged assets when needed."""
    for folder_name in ("imported_images", "imported_pptx"):
        target_dir = os.path.join(PROJECT_DIR, folder_name)
        os.makedirs(target_dir, exist_ok=True)

        if is_frozen_app() and hasattr(sys, "_MEIPASS"):
            source_dir = os.path.join(sys._MEIPASS, folder_name)
            if os.path.exists(source_dir):
                for item in os.listdir(source_dir):
                    source_path = os.path.join(source_dir, item)
                    destination_path = os.path.join(target_dir, item)
                    if os.path.exists(destination_path):
                        continue
                    try:
                        if os.path.isdir(source_path):
                            shutil.copytree(source_path, destination_path)
                        else:
                            shutil.copy2(source_path, destination_path)
                    except Exception as exc:
                        print(f"Error copiando recurso empaquetado {item}: {exc}")


database.init_db()
init_project_resources()


# ---------------------------------------------------------------------------
# Text utilities
# ---------------------------------------------------------------------------
def remove_accents(text):
    """Normalizes text and removes Spanish accents while keeping characters like ñ."""
    if not text:
        return ""
    normalized = unicodedata.normalize("NFD", text)
    return "".join(
        char for char in normalized
        if unicodedata.category(char) != "Mn"
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
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(8)

        self.header_label = QLabel(header)
        self.header_label.setObjectName("cardHeader")
        layout.addWidget(self.header_label)

        # Miniatura grande para que se vea bien en el panel
        self.thumb_label = QLabel()
        self.thumb_label.setFixedSize(220, 124)
        self.thumb_label.setStyleSheet(
            "background-color: #1a1a1e; border-radius: 4px;")
        self.thumb_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        pixmap = QPixmap(image_path)
        if not pixmap.isNull():
            scaled = pixmap.scaled(
                self.thumb_label.size(),
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation)
            self.thumb_label.setPixmap(scaled)
        layout.addWidget(self.thumb_label,
                         alignment=Qt.AlignmentFlag.AlignCenter)

        self.name_label = QLabel(os.path.basename(image_path))
        self.name_label.setObjectName("cardLyrics")
        self.name_label.setWordWrap(True)
        self.name_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.name_label)

        self.set_selected(False)

    def set_selected(self, is_selected):
        if is_selected:
            self.setStyleSheet("""
                QFrame#previewCard {
                    background-color: #1b2b25;
                    border: 2px solid #65c5a6;
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


class LibraryImageRowWidget(QFrame):
    """Fila de imagen para la biblioteca. Clickeable directamente."""

    def __init__(self, title, file_path, on_add_clicked, on_delete_clicked, parent=None):
        super().__init__(parent)
        self.setObjectName("libraryImageRow")

        layout = QHBoxLayout(self)
        layout.setContentsMargins(12, 8, 12, 8)
        layout.setSpacing(12)

        # Miniatura de la imagen
        self.thumb_lbl = QLabel()
        self.thumb_lbl.setFixedSize(176, 99)
        self.thumb_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.thumb_lbl.setStyleSheet("""
            background-color: #1a1a1e; 
            border: 1px solid #27272a;
            border-radius: 4px;
        """)

        # Cargar miniatura
        pixmap = QPixmap(file_path)
        if not pixmap.isNull():
            self.image_dimensions = pixmap.size()
            scaled = pixmap.scaled(
                self.thumb_lbl.size(),
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation,
            )
            self.thumb_lbl.setPixmap(scaled)
            self.thumb_lbl.setToolTip(
                f"{self.image_dimensions.width()} × {self.image_dimensions.height()} px")
            image_details = (
                f"{self.image_dimensions.width()} × "
                f"{self.image_dimensions.height()} px · Imagen local"
            )
        else:
            self.image_dimensions = QSize()
            image_details = "Imagen local · Archivo no disponible"
        layout.addWidget(self.thumb_lbl)

        # Textos de título
        text_layout = QVBoxLayout()
        text_layout.setSpacing(2)

        self.title_lbl = QLabel(title)
        self.title_lbl.setStyleSheet(
            "color: #fafafa; font-weight: bold; font-size: 14px; background: transparent;")
        text_layout.addWidget(self.title_lbl)

        self.sub_lbl = QLabel(image_details)
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

        # Botón eliminar
        self.del_btn = QPushButton("×")
        self.del_btn.setFixedSize(30, 30)
        self.del_btn.setToolTip("Eliminar imagen")
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
                border-color: #991b1b;
            }
        """)
        self.del_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.del_btn.clicked.connect(on_delete_clicked)
        layout.addWidget(self.del_btn)

        # Botón para añadir al guión
        self.add_btn = QPushButton("+ Guión")
        self.add_btn.setObjectName("addToGuionBtn")
        self.add_btn.setText("Agregar al Guion")
        self.add_btn.setMinimumWidth(132)
        self.add_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.add_btn.clicked.connect(on_add_clicked)
        layout.addWidget(self.add_btn)

        self.set_in_guion(False)

    def set_in_guion(self, in_guion):
        self.is_in_guion = in_guion
        if in_guion:
            self.setStyleSheet("""
                QFrame#libraryImageRow {
                    background-color: rgba(99, 102, 241, 0.12);
                    border: 1.5px solid #6366f1;
                    border-radius: 8px;
                }
                QFrame#libraryImageRow:hover {
                    background-color: rgba(99, 102, 241, 0.20);
                    border: 1.5px solid #818cf8;
                }
            """)
            self.add_btn.setText("✓ En Guión")
            self.add_btn.setStyleSheet("""
                QPushButton {
                    background-color: #1e1b4b;
                    color: #a5b4fc;
                    border: 1px solid #6366f1;
                    border-radius: 14px;
                    padding: 6px 14px;
                    font-size: 11px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #312e81;
                    color: #ffffff;
                    border: 1px solid #818cf8;
                }
            """)
        else:
            self.setStyleSheet("""
                QFrame#libraryImageRow {
                    background-color: transparent;
                    border: 1px solid transparent;
                    border-radius: 8px;
                }
                QFrame#libraryImageRow:hover {
                    background-color: #27272a;
                    border: 1px solid #3f3f46;
                }
            """)
            self.add_btn.setText("Agregar al Guion")
            self.add_btn.setStyleSheet("""
                QPushButton {
                    background-color: #2563eb;
                    color: #ffffff;
                    border: 1px solid #3b82f6;
                    border-radius: 14px;
                    padding: 6px 14px;
                    font-size: 11px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #1d4ed8;
                    border: 1px solid #60a5fa;
                }
            """)


class LibraryPptxRowWidget(QFrame):
    """Fila de presentación (PPTX o PDF) para la biblioteca. Clickeable directamente."""

    def __init__(self, title, on_add_clicked, on_delete_clicked, file_type="pptx", parent=None):
        super().__init__(parent)
        self.setObjectName("libraryPptxRow")

        layout = QHBoxLayout(self)
        layout.setContentsMargins(12, 8, 12, 8)
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
                border: 1px solid #991b1b;
                border-radius: 18px; 
                font-size: 10px;
                font-weight: bold;
            """)
        else:
            self.icon_lbl.setStyleSheet("""
                background-color: #7c2d12; 
                color: #ff7a59; 
                border: 1px solid #9a3412;
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
        self.del_btn = QPushButton("×")
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
                border-color: #991b1b;
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

        self.set_in_guion(False)

    def set_in_guion(self, in_guion):
        self.is_in_guion = in_guion
        if in_guion:
            self.setStyleSheet("""
                QFrame#libraryPptxRow {
                    background-color: rgba(99, 102, 241, 0.12);
                    border: 1.5px solid #6366f1;
                    border-radius: 8px;
                }
                QFrame#libraryPptxRow:hover {
                    background-color: rgba(99, 102, 241, 0.20);
                    border: 1.5px solid #818cf8;
                }
            """)
            self.add_btn.setText("✓ En Guión")
            self.add_btn.setStyleSheet("""
                QPushButton {
                    background-color: #1e1b4b;
                    color: #a5b4fc;
                    border: 1px solid #6366f1;
                    border-radius: 14px;
                    padding: 6px 14px;
                    font-size: 11px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #312e81;
                    color: #ffffff;
                    border: 1px solid #818cf8;
                }
            """)
        else:
            self.setStyleSheet("""
                QFrame#libraryPptxRow {
                    background-color: transparent;
                    border: 1px solid transparent;
                    border-radius: 8px;
                }
                QFrame#libraryPptxRow:hover {
                    background-color: #27272a;
                    border: 1px solid #3f3f46;
                }
            """)
            self.add_btn.setText("+ Guión")
            self.add_btn.setStyleSheet("""
                QPushButton {
                    background-color: #2563eb;
                    color: #ffffff;
                    border: 1px solid #3b82f6;
                    border-radius: 14px;
                    padding: 6px 14px;
                    font-size: 11px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #1d4ed8;
                    border: 1px solid #60a5fa;
                }
            """)


class PreviewCardWidget(QFrame):
    """Tarjeta de previsualización individual para las estrofas/diapositivas en el panel izquierdo."""

    def __init__(self, header, lyrics, parent=None):
        super().__init__(parent)
        self.setObjectName("previewCard")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(17, 15, 17, 15)
        layout.setSpacing(8)

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
                    background-color: #1b2b25;
                    border: 2px solid #65c5a6;
                    border-radius: 8px;
                }
            """)
            self.header_label.setStyleSheet("""
                color: #70d7bd;
                font-size: 12px;
                font-weight: bold;
                background: transparent;
            """)
            self.lyrics_label.setStyleSheet("""
                color: #f5f5f4;
                font-size: 15px;
                font-weight: bold;
                background: transparent;
            """)
        else:
            self.setStyleSheet("""
                QFrame#previewCard {
                    background-color: #171d1b;
                    border: 1px solid #29342f;
                    border-radius: 8px;
                }
                QFrame#previewCard:hover {
                    background-color: #1c2521;
                    border: 1px solid #3b6658;
                }
            """)
            self.header_label.setStyleSheet("""
                color: #9ba8a2;
                font-size: 12px;
                font-weight: bold;
                background: transparent;
            """)
            self.lyrics_label.setStyleSheet("""
                color: #e7ece9;
                font-size: 15px;
                font-weight: 600;
                background: transparent;
            """)


class LibrarySongRowWidget(QFrame):
    """Fila de canto para la biblioteca. Sin botón Proyectar. Clickeable directamente."""

    def __init__(self, title, category, on_add_clicked, on_edit_clicked, on_delete_clicked, parent=None):
        super().__init__(parent)
        self.setObjectName("librarySongRow")

        layout = QHBoxLayout(self)
        layout.setContentsMargins(12, 8, 12, 8)
        layout.setSpacing(12)

        # Icono circular de música
        self.icon_lbl = QLabel("♫")
        self.icon_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.icon_lbl.setFixedSize(36, 36)
        self.icon_lbl.setStyleSheet("""
            background-color: #202923;
            color: #a9bbb2;
            border: 1px solid #3a4b40;
            border-radius: 6px;
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
                border-color: #52525b;
            }
        """)
        self.edit_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.edit_btn.clicked.connect(on_edit_clicked)
        layout.addWidget(self.edit_btn)

        # Botón Eliminar
        self.del_btn = QPushButton("×")
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
                border-color: #991b1b;
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

        self.set_in_guion(False)

    def set_in_guion(self, in_guion):
        self.is_in_guion = in_guion
        if in_guion:
            self.setStyleSheet("""
                QFrame#librarySongRow {
                    background-color: rgba(99, 102, 241, 0.12);
                    border: 1.5px solid #6366f1;
                    border-radius: 8px;
                }
                QFrame#librarySongRow:hover {
                    background-color: rgba(99, 102, 241, 0.20);
                    border: 1.5px solid #818cf8;
                }
            """)
            self.add_btn.setText("✓ En Guión")
            self.add_btn.setStyleSheet("""
                QPushButton {
                    background-color: #1e1b4b;
                    color: #a5b4fc;
                    border: 1px solid #6366f1;
                    border-radius: 14px;
                    padding: 6px 14px;
                    font-size: 11px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #312e81;
                    color: #ffffff;
                    border: 1px solid #818cf8;
                }
            """)
        else:
            self.setStyleSheet("""
                QFrame#librarySongRow {
                    background-color: transparent;
                    border: 1px solid transparent;
                    border-radius: 8px;
                }
                QFrame#librarySongRow:hover {
                    background-color: #27272a;
                    border: 1px solid #3f3f46;
                }
            """)
            self.add_btn.setText("+ Guión")
            self.add_btn.setStyleSheet("""
                QPushButton {
                    background-color: #2563eb;
                    color: #ffffff;
                    border: 1px solid #3b82f6;
                    border-radius: 14px;
                    padding: 6px 14px;
                    font-size: 11px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #1d4ed8;
                    border: 1px solid #60a5fa;
                }
            """)


class CleanProjectionWidget(QFrame):
    """Vista previa nativa de la salida que renderiza el contenido proyectado."""

    def __init__(self, is_external=False, parent=None):
        super().__init__(parent)
        self.setObjectName("projectionOutputFrame")
        # Establecer fondo transparente para que nuestro paintEvent controle el dibujado completo
        self.setStyleSheet(
            "background-color: transparent; border: none; border-radius: 0px;")

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 15, 20, 15)

        # Encabezado de la salida local; las pantallas se administran en el panel lateral.
        top_layout = QHBoxLayout()
        self.projection_title = QLabel("PROYECCIÓN")
        self.projection_title.setObjectName("projectionLabelTitle")
        top_layout.addWidget(self.projection_title)
        top_layout.addStretch()

        self.btn_screens = QPushButton("▣ Pantallas ▾")
        self.btn_screens.setObjectName("subtleCtrlBtn")
        self.btn_screens.setStyleSheet("padding: 2px 8px; font-size: 11px;")
        self.btn_screens.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_screens.setToolTip("Configurar proyección en múltiples pantallas")
        self.btn_screens.clicked.connect(self._on_screens_menu_clicked)
        top_layout.addWidget(self.btn_screens)

        main_layout.addLayout(top_layout)

        self.header_label = QLabel("")
        self.header_label.setObjectName("projectionHeader")
        self.header_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.header_label.hide()
        main_layout.addWidget(self.header_label)

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
        main_layout.addWidget(self.text_label, stretch=1)

        # Línea inferior (Pie de página y ayuda)
        bottom_layout = QHBoxLayout()
        bottom_layout.addStretch()

        self.footer_label = QLabel("")
        self.footer_label.setObjectName("projectionFooterTitle")
        self.footer_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        bottom_layout.addWidget(self.footer_label)

        bottom_layout.addStretch()

        self.version_label = QLabel("")
        self.version_label.setObjectName("projectionVersion")
        self.version_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.version_label.hide()
        bottom_layout.addWidget(self.version_label)

        self.help_btn = QPushButton("?")
        self.help_btn.setObjectName("helpRoundBtn")
        self.help_btn.setToolTip("Opciones de proyección multipantalla")
        self.help_btn.clicked.connect(self._on_help_clicked)
        bottom_layout.addWidget(self.help_btn)

        main_layout.addLayout(bottom_layout)

        self.is_external = is_external
        self.current_display_mode = "standard"
        self.current_verse_number = None
        self.current_chapter_title = ""
        self.current_plain_text = ""

        # La miniatura ayuda a identificar imágenes en la vista local sin alterar la salida externa.
        self.image_thumbnail_label = QLabel(self)
        self.image_thumbnail_label.setObjectName("outputImageThumbnail")
        self.image_thumbnail_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.image_thumbnail_label.setAttribute(
            Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        self.image_thumbnail_label.hide()

        # Ocultar controles administrativos en ventana externa
        if is_external:
            self.projection_title.hide()
            self.btn_screens.hide()
            self.help_btn.hide()
        else:
            self.setCursor(Qt.CursorShape.PointingHandCursor)
            self.setToolTip("Haz clic derecho o en 'Pantallas ▾' para proyectar en una o varias pantallas")
            self.text_label.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
            self.projection_title.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
            self.footer_label.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)

        self.raw_pixmap = None
        self.general_bg_pixmap = None
        self.is_black_screen = True
        self.current_song_title = ""
        self.current_slide_name = ""
        self.font_size_offset = 0

        # Inicializar fuentes
        self.adjust_font_size()

    def adjust_font_size(self):
        w = self.width()
        h = self.height()
        if w <= 0 or h <= 0:
            return

        is_bible = getattr(self, "current_display_mode", "standard") == "bible"

        if is_bible:
            base_size = int(min(w * 0.042, h * 0.075))
            length = len(self.current_plain_text)

            if length > 260:
                scale = 0.72
            elif length > 180:
                scale = 0.82
            elif length > 110:
                scale = 0.9
            else:
                scale = 1.0

            offset = getattr(self, "font_size_offset", 0)
            min_font_size = max(18, int(min(w * 0.025, h * 0.04)))
            max_font_size = max(min_font_size, int(min(w * 0.075, h * 0.12)))
            font_size = max(min_font_size, min(max_font_size, int(base_size * scale) + offset))
            self.text_label.setStyleSheet(f"""
                color: #f5efe6;
                font-family: 'Times New Roman', Georgia, serif;
                font-size: {font_size}px;
                font-weight: 600;
                line-height: 1.35;
                background: transparent;
            """)

            header_font_size = max(20, int(min(w * 0.034, h * 0.06)))
            self.header_label.setStyleSheet(f"""
                color: #70d7bd;
                font-family: 'Times New Roman', Georgia, serif;
                font-size: {header_font_size}px;
                font-weight: 700;
                letter-spacing: 2px;
                background: transparent;
            """)

            footer_font_size = max(14, int(h * 0.026))
            self.footer_label.setStyleSheet(f"""
                color: #d7b76d;
                font-family: 'Times New Roman', Georgia, serif;
                font-size: {footer_font_size}px;
                font-style: italic;
                font-weight: 500;
                background: transparent;
            """)
            self.version_label.setStyleSheet(f"""
                color: #cbd5d1;
                font-family: 'Segoe UI', sans-serif;
                font-size: {max(12, int(h * 0.022))}px;
                font-weight: 600;
                background: transparent;
            """)

            margin_w = max(40, int(w * 0.08))
            margin_h = max(30, int(h * 0.06))
            self.layout().setContentsMargins(margin_w, margin_h, margin_w, margin_h)
            return

        # Base font size: roughly 8.5% of height or 5.5% of width, whichever is smaller
        base_size = int(min(w * 0.055, h * 0.085))

        # Adaptive factor based on text length to prevent overflow
        length = len(self.current_plain_text)

        if length > 200:
            scale = 0.65
        elif length > 120:
            scale = 0.75
        elif length > 60:
            scale = 0.85
        else:
            scale = 1.0

        offset = getattr(self, "font_size_offset", 0)
        max_font_size = max(18, int(min(w * 0.085, h * 0.13)))
        font_size = max(16, min(max_font_size, int(base_size * scale) + offset))

        self.text_label.setStyleSheet(f"""
            color: #ffffff; 
            font-size: {font_size}px; 
            font-weight: bold; 
            background: transparent;
        """)

        footer_font_size = max(11, int(h * 0.022))
        self.footer_label.setStyleSheet(f"""
            color: #71717a;
            font-size: {footer_font_size}px;
            font-weight: 500;
            background: transparent;
        """)

        margin_w = max(20, int(w * 0.05))
        margin_h = max(15, int(h * 0.04))
        self.layout().setContentsMargins(margin_w, margin_h, margin_w, margin_h)

    def display_text(self, text, slide_name="", song_title="", display_mode="standard", verse_number=None, chapter_title=""):
        self.is_black_screen = False
        self.raw_pixmap = None
        self.image_thumbnail_label.clear()
        self.image_thumbnail_label.hide()
        self.current_display_mode = display_mode
        self.current_verse_number = verse_number
        self.current_chapter_title = chapter_title
        self.current_plain_text = text or ""
        self.current_slide_name = slide_name
        self.current_song_title = song_title

        if display_mode == "bible":
            self.header_label.setText(slide_name or "")
            self.header_label.show()

            verse_text = self.current_plain_text.strip()
            if verse_number is not None and verse_text:
                # Escapar el texto evita que signos del versículo se interpreten como HTML.
                verse_text = (
                    f"<sup><span style='color:#70d7bd;font-size:65%;font-weight:700'>"
                    f"{escape_html(str(verse_number))}</span></sup>&nbsp;"
                    f"{escape_html(verse_text)}"
                )
            else:
                verse_text = escape_html(verse_text)
            self.text_label.setText(f"<div>{verse_text}</div>")
            self.text_label.show()
            self.footer_label.setText(chapter_title or "")
            self.footer_label.setVisible(bool(chapter_title))
            self.version_label.setText(song_title or "")
            self.version_label.setVisible(bool(song_title))
        else:
            self.header_label.hide()
            self.text_label.setText(text)
            self.text_label.show()
            self.version_label.clear()
            self.version_label.hide()

            if self.current_slide_name and self.current_song_title:
                self.footer_label.setText(f"{self.current_slide_name} · {self.current_song_title}")
            else:
                self.footer_label.setText(self.current_song_title or self.current_slide_name)
            self.footer_label.show()

        self.adjust_font_size()
        self.update()

    def display_image(self, pixmap_path, slide_name="", song_title=""):
        self.is_black_screen = False
        self.current_display_mode = "standard"
        self.current_verse_number = None
        self.current_chapter_title = ""
        self.current_plain_text = ""
        self.header_label.hide()
        self.version_label.clear()
        self.version_label.hide()
        self.text_label.hide()
        # Las imágenes ocupan la salida completa; no superponer títulos ni nombres.
        self.footer_label.clear()
        self.footer_label.hide()
        self.raw_pixmap = QPixmap(pixmap_path)
        if not self.is_external and not self.raw_pixmap.isNull():
            self._update_image_thumbnail()
            self.image_thumbnail_label.show()
            self.image_thumbnail_label.raise_()
        else:
            self.image_thumbnail_label.clear()
            self.image_thumbnail_label.hide()
        self.update()

        if song_title:
            self.current_song_title = song_title
        if slide_name:
            self.current_slide_name = slide_name


    def set_black_screen(self):
        self.is_black_screen = True
        self.current_display_mode = "standard"
        self.current_verse_number = None
        self.current_chapter_title = ""
        self.current_plain_text = ""
        self.header_label.hide()
        self.image_thumbnail_label.clear()
        self.image_thumbnail_label.hide()
        self.text_label.hide()
        self.raw_pixmap = None
        self.footer_label.setText("")
        self.version_label.clear()
        self.version_label.hide()
        self.update()

    def clear_text_only(self):
        self.header_label.hide()
        self.current_plain_text = ""
        self.image_thumbnail_label.clear()
        self.image_thumbnail_label.hide()
        self.text_label.setText("")
        self.footer_label.setText("")
        self.version_label.clear()
        self.version_label.hide()
        self.update()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.adjust_font_size()
        if self.image_thumbnail_label.isVisible() and self.raw_pixmap:
            self._update_image_thumbnail()
        self.update()

    def _update_image_thumbnail(self):
        """Escala y posiciona una miniatura dentro de la vista local."""
        thumbnail_width = min(144, max(88, int(self.width() * 0.2)))
        thumbnail_height = int(thumbnail_width * 9 / 16) + 8
        self.image_thumbnail_label.setFixedSize(
            thumbnail_width, thumbnail_height)
        self.image_thumbnail_label.setPixmap(self.raw_pixmap.scaled(
            thumbnail_width - 8,
            thumbnail_height - 8,
            Qt.AspectRatioMode.KeepAspectRatio,
            Qt.TransformationMode.SmoothTransformation,
        ))
        margin = 14
        top_offset = 42
        self.image_thumbnail_label.move(
            max(margin, self.width() - thumbnail_width - margin), top_offset)

    def paintEvent(self, event):
        from PyQt6.QtGui import QPainter
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform, True)

        rect = self.contentsRect()
        # Tamaño físico real (considera DPR de pantallas HiDPI)
        dpr = self.devicePixelRatioF()
        phys_w = int(rect.width() * dpr)
        phys_h = int(rect.height() * dpr)
        from PyQt6.QtCore import QSize
        phys_size = QSize(phys_w, phys_h)

        # Fondo negro siempre
        painter.fillRect(rect, Qt.GlobalColor.black)

        # Fondo de letras (solo cuando NO hay imagen/diapositiva activa)
        if (not getattr(self, "is_black_screen", False)
                and self.general_bg_pixmap
                and not self.general_bg_pixmap.isNull()
                and self.raw_pixmap is None):
            scaled = self.general_bg_pixmap.scaled(
                phys_size,
                Qt.AspectRatioMode.KeepAspectRatioByExpanding,
                Qt.TransformationMode.SmoothTransformation
            )
            # Dibujar centrado en coordenadas lógicas
            lw = scaled.width() / dpr
            lh = scaled.height() / dpr
            x = rect.x() + (rect.width() - lw) / 2
            y = rect.y() + (rect.height() - lh) / 2
            painter.drawPixmap(int(x), int(y), int(lw), int(lh), scaled)

        # Imagen o diapositiva activa — cubre toda la pantalla
        if self.raw_pixmap and not self.raw_pixmap.isNull():
            scaled = self.raw_pixmap.scaled(
                phys_size,
                Qt.AspectRatioMode.KeepAspectRatioByExpanding,
                Qt.TransformationMode.SmoothTransformation
            )
            lw = scaled.width() / dpr
            lh = scaled.height() / dpr
            x = rect.x() + (rect.width() - lw) / 2
            y = rect.y() + (rect.height() - lh) / 2
            painter.drawPixmap(int(x), int(y), int(lw), int(lh), scaled)

        painter.end()
        super().paintEvent(event)

    def _on_screens_menu_clicked(self):
        main_win = self.window()
        if hasattr(main_win, "show_screen_menu"):
            main_win.show_screen_menu(self.btn_screens)

    def _on_help_clicked(self):
        main_win = self.window()
        if hasattr(main_win, "show_screen_menu"):
            main_win.show_screen_menu(self.help_btn)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.RightButton and not getattr(self, "is_external", False):
            main_win = self.window()
            if hasattr(main_win, "show_screen_menu"):
                main_win.show_screen_menu(event.globalPosition().toPoint())
                return
        super().mousePressEvent(event)


class ProjectionWindow(QWidget):
    """Ventana externa independiente (Proyector/Pantalla Secundaria)."""

    def __init__(self, close_callback=None):
        super().__init__()
        self.close_callback = close_callback
        self.setWindowTitle("Station Selah · Proyección")
        if os.path.exists(LOGO_PATH):
            self.setWindowIcon(QIcon(LOGO_PATH))
        self.setStyleSheet("background-color: #000000;")
        self.resize(800, 600)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self.projection_widget = CleanProjectionWidget(is_external=True)
        layout.addWidget(self.projection_widget)

    def closeEvent(self, event):
        cb = self.close_callback
        self.close_callback = None
        if cb:
            try:
                cb()
            except Exception as e:
                print(f"Error in close_callback: {e}")
        super().closeEvent(event)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Escape:
            # ESC → pantalla negra desde la ventana de proyección
            main_window = self.parent() if self.parent() else None
            # Buscar la MainWindow a través del callback
            if hasattr(self, "_main_window_ref") and self._main_window_ref:
                self._main_window_ref.project_black()
            elif self.isFullScreen():
                self.showNormal()
        elif event.key() == Qt.Key.Key_F11:
            if self.isFullScreen():
                self.showNormal()
            else:
                self.showFullScreen()
        else:
            super().keyPressEvent(event)

    def mouseDoubleClickEvent(self, event):
        if self.isFullScreen():
            self.showNormal()
        else:
            self.showFullScreen()
        super().mouseDoubleClickEvent(event)


class SectionBlockWidget(QFrame):
    """Bloque individual de sección (Título, Estrofa, Coro) dentro del editor."""
    COLORS = {
        "Título":   {"border": "#355346", "header_bg": "#20352b", "header_fg": "#92d8b2", "label": "#92d8b2"},
        "Estrofa":  {"border": "#355346", "header_bg": "#20352b", "header_fg": "#92d8b2", "label": "#92d8b2"},
        "Coro":     {"border": "#355346", "header_bg": "#20352b", "header_fg": "#92d8b2", "label": "#92d8b2"},
        "Puente":   {"border": "#355346", "header_bg": "#20352b", "header_fg": "#92d8b2", "label": "#92d8b2"},
    }

    def __init__(self, section_type="Estrofa", number=1, content="", on_delete=None, parent=None):
        super().__init__(parent)
        self.on_delete = on_delete
        self.section_type = section_type
        self.number = number

        colors = self.COLORS.get(section_type, self.COLORS["Estrofa"])

        self.setStyleSheet(f"""
            QFrame {{
                background-color: #151c19;
                border: 1px solid {colors['border']};
                border-radius: 6px;
            }}
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(8)

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
                border: 1px solid #52525b;
                border-radius: 12px;
                font-size: 11px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #7f1d1d;
                color: #fca5a5;
                border-color: #991b1b;
            }
        """)
        del_btn.clicked.connect(self._delete_self)
        header_row.addWidget(del_btn)

        layout.addLayout(header_row)

        # Text area
        self.text_edit = QTextEdit()
        self.text_edit.setPlaceholderText("Escribe la letra aquí...")
        self.text_edit.setPlainText(content)
        self.text_edit.setMinimumHeight(100)
        self.text_edit.setMaximumHeight(190)
        self.text_edit.setStyleSheet("""
            QTextEdit {
                background-color: #101513;
                color: #f4f4f5;
                border: 1px solid #29352f;
                border-radius: 4px;
                padding: 10px;
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
        self.setObjectName("songEditorDialog")
        self.setMinimumSize(900, 620)
        self.resize(1080, 720)
        self.setStyleSheet("""
            QDialog#songEditorDialog { background-color: #101614; }
            QLabel { color: #9aa8a0; background: transparent; }
            QLineEdit, QComboBox {
                background-color: #171f1b;
                border: 1px solid #304038;
                border-radius: 5px;
                color: #edf3ef;
                padding: 9px 11px;
                font-size: 12px;
            }
            QLineEdit:focus, QComboBox:focus { border-color: #5da98c; }
            QComboBox::drop-down { border: none; width: 22px; }
            QComboBox QAbstractItemView {
                background-color: #171f1b;
                color: #edf3ef;
                selection-background-color: #244235;
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
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # Encabezado del editor.
        header = QFrame()
        header.setObjectName("songEditorHeader")
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(20, 14, 20, 12)
        header_layout.setSpacing(12)

        icon = QLabel("✎" if song_id else "+")
        icon.setObjectName("songEditorIcon")
        icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon.setFixedSize(34, 34)
        header_layout.addWidget(icon)

        header_text = QVBoxLayout()
        header_text.setSpacing(2)
        eyebrow = QLabel("BIBLIOTECA DE CANTOS")
        eyebrow.setObjectName("songEditorEyebrow")
        header_text.addWidget(eyebrow)
        dialog_title = QLabel("Editar canto" if song_id else "Agregar canto")
        dialog_title.setObjectName("songEditorTitle")
        header_text.addWidget(dialog_title)
        header_layout.addLayout(header_text)
        header_layout.addStretch()

        close_btn = QPushButton("×")
        close_btn.setObjectName("songEditorCloseBtn")
        close_btn.setFixedSize(32, 32)
        close_btn.setToolTip("Cerrar")
        close_btn.clicked.connect(self.reject)
        header_layout.addWidget(close_btn)
        main_layout.addWidget(header)

        # Campos principales en una sola fila, como en la referencia.
        meta_panel = QFrame()
        meta_panel.setObjectName("songEditorMeta")
        meta_row = QHBoxLayout(meta_panel)
        meta_row.setContentsMargins(20, 12, 20, 14)
        meta_row.setSpacing(10)

        self.title_input = QLineEdit()
        self.title_input.setPlaceholderText("Ej. A Dios sea la Gloria")
        self.author_input = QLineEdit()
        self.author_input.setPlaceholderText("Nombre del autor")

        self.category_input = QComboBox()
        self.category_input.setEditable(True)
        self.category_input.addItems(["Alabanza", "Adoración", "Himno"])
        self.category_input.setCurrentText("Adoración")

        self.tone_input = QLineEdit()
        self.tone_input.setPlaceholderText("Ej. Sol")

        for label, control, stretch in (
            ("TÍTULO DEL CANTO", self.title_input, 3),
            ("AUTOR", self.author_input, 2),
            ("CATEGORÍA", self.category_input, 1),
            ("TONO", self.tone_input, 1),
        ):
            field = QVBoxLayout()
            field.setSpacing(5)
            field_label = QLabel(label)
            field_label.setObjectName("songEditorFieldLabel")
            field.addWidget(field_label)
            field.addWidget(control)
            meta_row.addLayout(field, stretch=stretch)
        main_layout.addWidget(meta_panel)

        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("background-color: #29352f; max-height: 1px;")
        main_layout.addWidget(sep)

        # Editor de secciones y preview en columnas independientes.
        body = QWidget()
        body_layout = QHBoxLayout(body)
        body_layout.setContentsMargins(20, 14, 20, 14)
        body_layout.setSpacing(16)

        editor_column = QWidget()
        editor_layout = QVBoxLayout(editor_column)
        editor_layout.setContentsMargins(0, 0, 0, 0)
        editor_layout.setSpacing(10)

        sections_header = QHBoxLayout()
        section_heading = QVBoxLayout()
        section_heading.setSpacing(2)
        section_kicker = QLabel("ESTRUCTURA")
        section_kicker.setObjectName("songEditorEyebrow")
        section_heading.addWidget(section_kicker)
        section_title = QLabel("Secciones del canto")
        section_title.setObjectName("songEditorSectionTitle")
        section_heading.addWidget(section_title)
        sections_header.addLayout(section_heading)
        sections_header.addStretch()
        self.section_count_label = QLabel("0 bloques")
        self.section_count_label.setObjectName("songEditorCount")
        sections_header.addWidget(self.section_count_label)
        editor_layout.addLayout(sections_header)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.scroll_area.setObjectName("songSectionsScroll")

        self.sections_container = QWidget()
        self.sections_container.setStyleSheet("background: transparent;")
        self.sections_layout = QVBoxLayout(self.sections_container)
        self.sections_layout.setContentsMargins(0, 2, 8, 2)
        self.sections_layout.setSpacing(10)
        self.sections_layout.addStretch()

        self.scroll_area.setWidget(self.sections_container)
        editor_layout.addWidget(self.scroll_area, stretch=1)

        # Atajos para insertar tipos comunes de sección.
        add_row = QHBoxLayout()
        add_row.setSpacing(7)
        add_label = QLabel("AGREGAR")
        add_label.setObjectName("songEditorFieldLabel")
        add_row.addWidget(add_label)

        for label, sec_type in [("+ Título", "Título"), ("+ Estrofa", "Estrofa"), ("+ Coro", "Coro"), ("+ Puente", "Puente")]:
            btn = QPushButton(label)
            btn.setObjectName("songEditorAddSectionBtn")
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.clicked.connect(
                lambda checked, st=sec_type: self.add_section(st))
            add_row.addWidget(btn)

        add_row.addStretch()
        editor_layout.addLayout(add_row)
        body_layout.addWidget(editor_column, stretch=7)

        preview_panel = QFrame()
        preview_panel.setObjectName("songEditorPreviewPanel")
        preview_layout = QVBoxLayout(preview_panel)
        preview_layout.setContentsMargins(14, 14, 14, 14)
        preview_layout.setSpacing(10)

        preview_heading = QVBoxLayout()
        preview_heading.setSpacing(2)
        preview_kicker = QLabel("VISTA PREVIA")
        preview_kicker.setObjectName("songEditorEyebrow")
        preview_heading.addWidget(preview_kicker)
        preview_subtitle = QLabel("Así se proyectará")
        preview_subtitle.setObjectName("songEditorSectionTitle")
        preview_heading.addWidget(preview_subtitle)
        preview_layout.addLayout(preview_heading)

        self.preview_canvas = QFrame()
        self.preview_canvas.setObjectName("songPreviewCanvas")
        self.preview_canvas.setMinimumHeight(220)
        self.preview_canvas.setMaximumHeight(340)
        canvas_layout = QVBoxLayout(self.preview_canvas)
        canvas_layout.setContentsMargins(20, 20, 20, 20)
        canvas_layout.addStretch()

        self.preview_section_label = QLabel("ESTROFA 1")
        self.preview_section_label.setObjectName("songPreviewSection")
        self.preview_section_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        canvas_layout.addWidget(self.preview_section_label)

        self.preview_lyrics_label = QLabel("La letra aparecerá aquí mientras escribes.")
        self.preview_lyrics_label.setObjectName("songPreviewLyrics")
        self.preview_lyrics_label.setWordWrap(True)
        self.preview_lyrics_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        canvas_layout.addWidget(self.preview_lyrics_label)
        canvas_layout.addStretch()

        self.preview_song_title = QLabel("TÍTULO DEL CANTO")
        self.preview_song_title.setObjectName("songPreviewFooter")
        self.preview_song_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        canvas_layout.addWidget(self.preview_song_title)
        preview_layout.addWidget(self.preview_canvas)
        preview_layout.addStretch()
        body_layout.addWidget(preview_panel, stretch=4)
        main_layout.addWidget(body, stretch=1)

        # Pie fijo para las acciones principales.
        footer = QFrame()
        footer.setObjectName("songEditorFooter")
        bottom_row = QHBoxLayout()
        bottom_row.setContentsMargins(20, 12, 20, 12)
        bottom_row.setSpacing(10)
        footer_note = QLabel("Los cambios se guardarán en tu biblioteca local.")
        footer_note.setObjectName("songEditorCount")
        bottom_row.addWidget(footer_note)
        bottom_row.addStretch()

        cancel_btn = QPushButton("Cancelar")
        cancel_btn.setObjectName("songEditorCancelBtn")
        cancel_btn.clicked.connect(self.reject)
        bottom_row.addWidget(cancel_btn)

        self.save_btn = QPushButton("✓  Guardar cambios" if song_id else "✓  Agregar canto")
        self.save_btn.setObjectName("songEditorSaveBtn")
        self.save_btn.clicked.connect(self.save_song)
        bottom_row.addWidget(self.save_btn)
        footer.setLayout(bottom_row)
        main_layout.addWidget(footer)

        self.title_input.textChanged.connect(
            lambda text: self.save_btn.setEnabled(bool(text.strip())))
        self.title_input.textChanged.connect(self.update_song_preview)
        self.category_input.currentTextChanged.connect(
            self.update_song_preview)

        if self.song_id:
            self._load_existing_song()
        else:
            self.add_section("Título")
            self.add_section("Estrofa")
            self.add_section("Coro")
        self.save_btn.setEnabled(bool(self.title_input.text().strip()))
        self.update_song_preview()

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
        block.text_edit.textChanged.connect(
            lambda section=block: self.update_song_preview(section))
        block.type_combo.currentTextChanged.connect(
            lambda _, section=block: self.update_song_preview(section))
        block.num_combo.currentTextChanged.connect(
            lambda _, section=block: self.update_song_preview(section))
        self.section_count_label.setText(
            f"{len(self.section_widgets)} bloques")
        self.update_song_preview(block)

        QTimer.singleShot(50, lambda: self.scroll_area.verticalScrollBar().setValue(
            self.scroll_area.verticalScrollBar().maximum()
        ))

    def remove_section(self, block):
        self.sections_layout.removeWidget(block)
        self.section_widgets.remove(block)
        block.deleteLater()
        self.section_count_label.setText(
            f"{len(self.section_widgets)} bloques")
        self.update_song_preview()

    def update_song_preview(self, preferred_section=None):
        """Actualiza el canvas de proyección con la sección editada."""
        selected_section = preferred_section
        if selected_section not in self.section_widgets:
            selected_section = next(
                (section for section in self.section_widgets
                 if section.text_edit.toPlainText().strip()),
                self.section_widgets[0] if self.section_widgets else None,
            )

        if selected_section:
            section_label, section_text = selected_section.get_data()
            self.preview_section_label.setText(section_label.upper())
            self.preview_lyrics_label.setText(
                section_text or "La letra aparecerá aquí mientras escribes.")
        else:
            self.preview_section_label.setText("SECCIÓN")
            self.preview_lyrics_label.setText(
                "La letra aparecerá aquí mientras escribes.")

        self.preview_song_title.setText(
            self.title_input.text().strip().upper() or "TÍTULO DEL CANTO")

    def _load_existing_song(self):
        conn = sqlite3.connect(database.DB_PATH)
        cursor = conn.cursor()
        cursor.execute(
            "SELECT title, category, tone, lyrics, author FROM songs WHERE id = ?",
            (self.song_id,))
        row = cursor.fetchone()
        conn.close()
        if not row:
            return
        title, category, tone, lyrics, author = row
        self.title_input.setText(title or "")
        self.author_input.setText(author or "")
        self.category_input.setCurrentText(category or "")
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

        category = self.category_input.currentText().strip()
        author = self.author_input.text().strip()
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
                "UPDATE songs SET title=?, category=?, tone=?, lyrics=?, author=? WHERE id=?",
                (title, category, tone, lyrics, author, self.song_id)
            )
        else:
            cursor.execute(
                "INSERT INTO songs (title, category, tone, lyrics, author) VALUES (?, ?, ?, ?, ?)",
                (title, category, tone, lyrics, author)
            )
        conn.commit()
        conn.close()
        self.accept()


# ---------------------------------------------------------------------------
# Main window and UI bootstrap
# ---------------------------------------------------------------------------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Station Selah")
        if os.path.exists(LOGO_PATH):
            self.setWindowIcon(QIcon(LOGO_PATH))
        self.resize(1200, 750)

        self.load_styles()

        # Session state
        self.guion = {
            "songs": [],
            "bible": [],
            "images": [],
            "pptx": []
        }

        # Projection state
        self.active_item_data = None
        self.projection_windows = {}
        self._is_external_obs_mode = False

        # Screen monitoring timer
        self._screen_monitor_timer = QTimer()
        self._screen_monitor_timer.setInterval(2000)
        self._screen_monitor_timer.timeout.connect(self._check_screens_connected)
        self._screen_monitor_timer.start()
        self.external_font_size_offset = 0

        self.current_projection_mode = "black"
        self.last_projected_mode = None
        self.last_projected_text = ""
        self.last_projected_header = ""
        self.last_projected_song_title = ""
        self.last_projected_image_path = ""
        self.last_projected_display_mode = "standard"
        self.last_projected_verse_number = None
        self.last_projected_chapter_title = ""

        self.setup_ui()
        self.load_songs_library()
        self.load_images_library()
        self.load_pptx_library()
        self.load_persisted_lyrics_bg()

        esc_shortcut = QShortcut(QKeySequence(Qt.Key.Key_Escape), self)
        esc_shortcut.activated.connect(self.project_black)

    def load_styles(self):
        """Loads the app stylesheet from the packaged or local project directory."""
        candidates = []
        if is_frozen_app() and hasattr(sys, "_MEIPASS"):
            candidates.append(os.path.join(sys._MEIPASS, "styles.qss"))
        candidates.append(os.path.join(PROJECT_DIR, "styles.qss"))
        candidates.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), "styles.qss"))

        qss_content = ""
        for qss_path in candidates:
            if not os.path.exists(qss_path):
                continue
            try:
                with open(qss_path, "r", encoding="utf-8") as file:
                    qss_content = file.read()
                if qss_content:
                    break
            except Exception as exc:
                print(f"Error cargando stylesheet desde {qss_path}: {exc}")

        if qss_content:
            self.setStyleSheet(qss_content)
            app_instance = QApplication.instance()
            if app_instance:
                app_instance.setStyleSheet(qss_content)

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

        # Identidad de la aplicación junto al símbolo de marca.
        brand = QWidget()
        brand.setObjectName("appBrand")
        brand_layout = QHBoxLayout(brand)
        brand_layout.setContentsMargins(0, 0, 12, 0)
        brand_layout.setSpacing(8)

        brand_icon = QLabel()
        brand_icon.setObjectName("appBrandIcon")
        brand_icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        brand_icon.setFixedSize(34, 34)
        logo_pixmap = QPixmap(LOGO_PATH) if os.path.exists(LOGO_PATH) else QPixmap()
        if not logo_pixmap.isNull():
            brand_icon.setPixmap(logo_pixmap.scaled(
                28, 28,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation,
            ))
        else:
            brand_icon.setText("S")
        brand_layout.addWidget(brand_icon)

        brand_text = QVBoxLayout()
        brand_text.setContentsMargins(0, 0, 0, 0)
        brand_text.setSpacing(0)
        brand_station = QLabel("STATION")
        brand_station.setObjectName("appBrandStation")
        brand_name = QLabel("SELAH")
        brand_name.setObjectName("appBrandName")
        brand_text.addWidget(brand_station)
        brand_text.addWidget(brand_name)
        brand_layout.addLayout(brand_text)
        top_bar_layout.addWidget(brand)

        self.nav_buttons = {}
        nav_items = [
            ("Cantos", "♫"),
            ("Biblias", "▤"),
            ("Imágenes", "▧"),
            ("PPTX", "▣"),
            ("Guión", "☷")
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
            "Buscar cita (ej. Juan 3 16) y Enter...")
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

        splitter.addWidget(preview_container)

        # ---- Panel Derecho: Biblioteca y salida de proyección ----
        self.right_splitter = QSplitter(Qt.Orientation.Vertical)
        self.right_splitter.setObjectName("rightSplitter")
        self.right_splitter.setChildrenCollapsible(False)

        # Stacked Widget para la biblioteca (Cerebro)
        self.library_stack = QStackedWidget()

        self.setup_songs_page()
        self.setup_bible_page()
        self.setup_images_page()
        self.setup_pptx_page()
        self.setup_guion_page()

        self.right_splitter.addWidget(self.library_stack)

        # Panel Inferior: vista de salida y controles nativos
        self.obs_container = QFrame()
        self.obs_container.setObjectName("projectionWorkspaceFrame")
        self.obs_container.setMinimumHeight(340)

        obs_layout = QHBoxLayout(self.obs_container)
        obs_layout.setContentsMargins(8, 8, 8, 8)
        obs_layout.setSpacing(8)

        self.local_projection_widget = CleanProjectionWidget(is_external=False)
        self.local_projection_widget.btn_screens.hide()
        obs_layout.addWidget(self.local_projection_widget)

        # Los controles viven junto a la salida para dejar más espacio a los bloques de Preview.
        self.projection_controls_panel = QFrame()
        self.projection_controls_panel.setObjectName("projectionControlPanel")
        self.projection_controls_panel.setFixedWidth(205)
        controls_panel_layout = QVBoxLayout(self.projection_controls_panel)
        controls_panel_layout.setContentsMargins(12, 12, 12, 12)
        controls_panel_layout.setSpacing(6)

        outputs_title = QLabel("SALIDAS")
        outputs_title.setObjectName("projectionSectionLabel")
        controls_panel_layout.addWidget(outputs_title)

        self.projection_output_rows = {}
        projection_outputs_container = QWidget()
        projection_outputs_container.setObjectName("projectionOutputsContainer")
        self.projection_outputs_layout = QVBoxLayout(projection_outputs_container)
        self.projection_outputs_layout.setContentsMargins(0, 0, 0, 0)
        self.projection_outputs_layout.setSpacing(6)
        controls_panel_layout.addWidget(projection_outputs_container)

        controls_divider = QFrame()
        controls_divider.setFrameShape(QFrame.Shape.HLine)
        controls_divider.setObjectName("projectionControlsDivider")
        controls_panel_layout.addWidget(controls_divider)

        controls_title = QLabel("CONTROLES")
        controls_title.setObjectName("projectionSectionLabel")
        controls_panel_layout.addWidget(controls_title)

        self.btn_clear_black = QPushButton("Pantalla negra")
        self.btn_clear_black.setObjectName("projectionDangerBtn")
        self.btn_clear_black.setToolTip("Poner todas las salidas en negro")
        self.btn_clear_black.clicked.connect(self.project_black)
        controls_panel_layout.addWidget(self.btn_clear_black)

        # Las tres acciones secundarias se mantienen juntas como botones tipo tarjeta.
        action_row = QHBoxLayout()
        action_row.setSpacing(6)
        self.btn_clear_text = QToolButton()
        self.btn_clear_text.setText("Limpiar")
        self.btn_clear_text.setIcon(
            self.style().standardIcon(QStyle.StandardPixmap.SP_TrashIcon))
        self.btn_clear_text.setIconSize(QSize(18, 18))
        self.btn_clear_text.setToolButtonStyle(
            Qt.ToolButtonStyle.ToolButtonTextUnderIcon)
        self.btn_clear_text.setObjectName("projectionTileBtn")
        self.btn_clear_text.setToolTip("Ocultar el texto proyectado")
        self.btn_clear_text.clicked.connect(self.project_clear_text)
        action_row.addWidget(self.btn_clear_text)

        self.btn_restore = QToolButton()
        self.btn_restore.setText("Restaurar")
        self.btn_restore.setIcon(
            self.style().standardIcon(QStyle.StandardPixmap.SP_BrowserReload))
        self.btn_restore.setIconSize(QSize(18, 18))
        self.btn_restore.setToolButtonStyle(
            Qt.ToolButtonStyle.ToolButtonTextUnderIcon)
        self.btn_restore.setObjectName("projectionTileBtn")
        self.btn_restore.setToolTip("Restaurar el último contenido proyectado")
        self.btn_restore.clicked.connect(self.project_restore)
        action_row.addWidget(self.btn_restore)

        self.btn_lyrics_bg = QToolButton()
        self.btn_lyrics_bg.setText("Fondo")
        self.btn_lyrics_bg.setIcon(
            self.style().standardIcon(QStyle.StandardPixmap.SP_DesktopIcon))
        self.btn_lyrics_bg.setIconSize(QSize(18, 18))
        self.btn_lyrics_bg.setToolButtonStyle(
            Qt.ToolButtonStyle.ToolButtonTextUnderIcon)
        self.btn_lyrics_bg.setObjectName("projectionTileBtn")
        self.btn_lyrics_bg.setToolTip("Cambiar el fondo de la proyección")
        self.btn_lyrics_bg.clicked.connect(self.change_lyrics_background)
        action_row.addWidget(self.btn_lyrics_bg)
        controls_panel_layout.addLayout(action_row)

        font_controls_row = QHBoxLayout()
        font_controls_row.setSpacing(5)
        font_title = QLabel("T")
        font_title.setObjectName("projectionSectionLabel")
        font_title.setToolTip("Tamaño de letra")
        font_controls_row.addWidget(font_title)

        self.btn_font_down = QPushButton("−")
        self.btn_font_down.setObjectName("fontStepBtn")
        self.btn_font_down.setToolTip("Reducir tamaño de letra")
        self.btn_font_down.clicked.connect(self.decrease_font_size)
        font_controls_row.addWidget(self.btn_font_down)

        self.btn_font_reset = QPushButton("M")
        self.btn_font_reset.setObjectName("fontSizeIndicatorBtn")
        self.btn_font_reset.setToolTip("Restablecer tamaño de letra")
        self.btn_font_reset.clicked.connect(self.reset_font_size)
        font_controls_row.addWidget(self.btn_font_reset)

        self.btn_font_up = QPushButton("+")
        self.btn_font_up.setObjectName("fontStepBtn")
        self.btn_font_up.setToolTip("Aumentar tamaño de letra")
        self.btn_font_up.clicked.connect(self.increase_font_size)
        font_controls_row.addWidget(self.btn_font_up)
        controls_panel_layout.addLayout(font_controls_row)

        quick_actions_row = QHBoxLayout()
        quick_actions_row.setSpacing(6)
        self.btn_quick_bienvenida = QPushButton("Bienvenida")
        self.btn_quick_bienvenida.setObjectName("projectionActionBtn")
        self.btn_quick_bienvenida.setToolTip("Clic: proyectar bienvenida; clic derecho: configurar imagen")
        self.btn_quick_bienvenida.clicked.connect(
            lambda: self.project_quick_image("bienvenida", "Bienvenida"))
        self.btn_quick_bienvenida.setContextMenuPolicy(
            Qt.ContextMenuPolicy.CustomContextMenu)
        self.btn_quick_bienvenida.customContextMenuRequested.connect(
            lambda pos: self.show_quick_image_menu(
                "bienvenida", "Bienvenida", self.btn_quick_bienvenida.mapToGlobal(pos)))
        quick_actions_row.addWidget(self.btn_quick_bienvenida)

        self.btn_quick_despedida = QPushButton("Despedida")
        self.btn_quick_despedida.setObjectName("projectionActionBtn")
        self.btn_quick_despedida.setToolTip("Clic: proyectar despedida; clic derecho: configurar imagen")
        self.btn_quick_despedida.clicked.connect(
            lambda: self.project_quick_image("despedida", "Despedida"))
        self.btn_quick_despedida.setContextMenuPolicy(
            Qt.ContextMenuPolicy.CustomContextMenu)
        self.btn_quick_despedida.customContextMenuRequested.connect(
            lambda pos: self.show_quick_image_menu(
                "despedida", "Despedida", self.btn_quick_despedida.mapToGlobal(pos)))
        quick_actions_row.addWidget(self.btn_quick_despedida)
        controls_panel_layout.addLayout(quick_actions_row)
        controls_panel_layout.addStretch()

        obs_layout.addWidget(self.projection_controls_panel)

        self.right_splitter.addWidget(self.obs_container)

        # Set initial stretch/sizes (equivalent to 2:4 stretch ratio)
        self.right_splitter.setSizes([2000, 4000])

        # Disable vertical splitter dragging initially
        handle = self.right_splitter.handle(1)
        if handle:
            handle.setEnabled(False)
            handle.setCursor(Qt.CursorShape.ArrowCursor)

        splitter.addWidget(self.right_splitter)

        splitter.setSizes([420, 780])

        self.set_active_nav_button("Cantos")
        self._update_projection_status()

    # ---------------------------------------------------------------------------
    # Library pages
    # ---------------------------------------------------------------------------

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
        self.song_search.setPlaceholderText("Buscar cantos...")
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

        lbl = QLabel(f"▤  Sección de {name} (Próximamente)")
        lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl.setStyleSheet(
            "color: #71717a; font-size: 14px; font-weight: bold;")
        layout.addWidget(lbl)

        self.library_stack.addWidget(container)

    # ---------------------------------------------------------------------------
    # Bible module
    # ---------------------------------------------------------------------------

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
        self.bible_current_chapter = None   # capítulo activo en la preview
        self.bible_current_verse = None     # versículo activo (para mantener foco al cambiar versión)
        self.bible_search_mode = False

        # Los selectores siguen el flujo: versión, libro y capítulo.
        selector_row = QHBoxLayout()
        selector_row.setSpacing(10)

        self.bible_version_combo = QComboBox()
        self.bible_version_combo.setObjectName("bibleSelectorCombo")
        self.bible_version_combo.addItems(["RVR60", "NVI", "RV1569", "DHH"])
        self.bible_version_combo.setCurrentText(self.bible_current_version)
        self.bible_version_combo.currentTextChanged.connect(
            self.on_bible_version_changed)

        self.bible_book_combo = QComboBox()
        self.bible_book_combo.setObjectName("bibleSelectorCombo")
        self._populate_book_combo()
        self.bible_book_combo.currentTextChanged.connect(
            self.on_bible_book_changed)

        self.bible_chapter_combo = QComboBox()
        self.bible_chapter_combo.setObjectName("bibleSelectorCombo")
        self.bible_chapter_combo.addItem("Selecciona capítulo", None)
        self.bible_chapter_combo.activated.connect(
            lambda index: self.on_bible_chapter_changed(
                self.bible_chapter_combo.itemData(index)))

        for label, control in (
            ("VERSIÓN", self.bible_version_combo),
            ("LIBRO", self.bible_book_combo),
            ("CAPÍTULO", self.bible_chapter_combo),
        ):
            field_layout = QVBoxLayout()
            field_layout.setContentsMargins(0, 0, 0, 0)
            field_layout.setSpacing(4)
            field_label = QLabel(label)
            field_label.setObjectName("bibleSelectorLabel")
            field_layout.addWidget(field_label)
            field_layout.addWidget(control)
            selector_row.addLayout(field_layout, stretch=1)

        layout.addLayout(selector_row)

        # La búsqueda temporal alterna con la vista de versículos del capítulo.
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

        # La lista bajo los selectores funciona como vista seleccionable del capítulo.
        verses_page = QWidget()
        verses_page.setStyleSheet("background: transparent;")
        verses_layout = QVBoxLayout(verses_page)
        verses_layout.setContentsMargins(0, 4, 0, 0)
        verses_layout.setSpacing(6)

        self.bible_book_section_label = QLabel("")
        self.bible_book_section_label.setObjectName("bibleBookSection")
        verses_layout.addWidget(self.bible_book_section_label)

        verse_heading_row = QHBoxLayout()
        self.bible_chapter_label = QLabel("Selecciona un capítulo")
        self.bible_chapter_label.setObjectName("bibleChapterHeading")
        verse_heading_row.addWidget(self.bible_chapter_label)
        verse_heading_row.addStretch()

        self.bible_verse_count_label = QLabel("")
        self.bible_verse_count_label.setObjectName("bibleVerseCount")
        verse_heading_row.addWidget(self.bible_verse_count_label)
        verses_layout.addLayout(verse_heading_row)

        # Permite encontrar un versículo por texto o saltar a él desde un selector.
        verse_navigation_row = QHBoxLayout()
        verse_navigation_row.setSpacing(8)

        self.bible_verse_search = QLineEdit()
        self.bible_verse_search.setObjectName("bibleVerseSearch")
        self.bible_verse_search.setPlaceholderText(
            "Buscar versículo por número o texto...")
        self.bible_verse_search.textChanged.connect(
            self.filter_bible_verses)
        verse_navigation_row.addWidget(self.bible_verse_search, stretch=1)

        self.bible_verse_combo = QComboBox()
        self.bible_verse_combo.setObjectName("bibleVerseCombo")
        self.bible_verse_combo.addItem("Ir a versículo", None)
        self.bible_verse_combo.activated.connect(
            lambda index: self._select_bible_verse(
                self.bible_verse_combo.itemData(index)))
        verse_navigation_row.addWidget(self.bible_verse_combo)
        verses_layout.addLayout(verse_navigation_row)

        self.bible_verses_list = QListWidget()
        self.bible_verses_list.setObjectName("bibleVerseList")
        self.bible_verses_list.setSpacing(3)
        self.bible_verses_list.itemClicked.connect(
            self._on_bible_verse_clicked)
        verses_layout.addWidget(self.bible_verses_list, stretch=1)

        self.bible_inner_stack.addWidget(verses_page)  # index 1
        self.bible_inner_stack.setCurrentIndex(1)

        self.library_stack.addWidget(container)

        # Mostrar el primer capítulo como consulta, sin proyectarlo al iniciar.
        chapters = self.load_bible_chapters(self.bible_current_book)
        if chapters:
            self.bible_chapter_combo.setCurrentIndex(
                self.bible_chapter_combo.findData(chapters[0]))
            self._load_bible_library_verses(
                self.bible_current_book, chapters[0])

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
        if self.bible_current_chapter:
            self.load_bible_chapter_to_preview(
                self.bible_current_book,
                self.bible_current_chapter,
                focus_verse=self.bible_current_verse
            )
        else:
            chapters = self.load_bible_chapters(self.bible_current_book)
            if chapters:
                chapter_num = self.bible_chapter_combo.currentData()
                if chapter_num not in chapters:
                    chapter_num = chapters[0]
                    self.bible_chapter_combo.setCurrentIndex(
                        self.bible_chapter_combo.findData(chapter_num))
                self._load_bible_library_verses(
                    self.bible_current_book, chapter_num)

    def on_bible_book_changed(self, book_name):
        if book_name.startswith("──"):
            return
        self.bible_current_book = book_name
        self.bible_current_chapter = None
        self.bible_current_verse = None
        chapters = self.load_bible_chapters(book_name)
        if chapters:
            first_chapter = chapters[0]
            self.bible_chapter_combo.setCurrentIndex(
                self.bible_chapter_combo.findData(first_chapter))
            self._load_bible_library_verses(book_name, first_chapter)

    def on_bible_chapter_changed(self, chapter_num):
        if chapter_num is None:
            return
        self.load_bible_chapter_to_preview(
            self.bible_current_book, int(chapter_num))

    def load_bible_chapters(self, book_name, highlight_chapter=None):
        """Actualiza el selector con los capítulos disponibles del libro."""
        conn = sqlite3.connect(database.DB_PATH)
        cursor = conn.cursor()
        cursor.execute(
            "SELECT DISTINCT chapter FROM bible WHERE book = ? AND version = ? ORDER BY chapter",
            (book_name, self.bible_current_version)
        )
        chapters = [row[0] for row in cursor.fetchall()]
        conn.close()

        self.bible_chapter_combo.blockSignals(True)
        self.bible_chapter_combo.clear()
        self.bible_chapter_combo.addItem("Selecciona capítulo", None)
        for chapter_num in chapters:
            self.bible_chapter_combo.addItem(str(chapter_num), chapter_num)

        if highlight_chapter is not None:
            selected_index = self.bible_chapter_combo.findData(highlight_chapter)
            if selected_index >= 0:
                self.bible_chapter_combo.setCurrentIndex(selected_index)
        self.bible_chapter_combo.blockSignals(False)

        if not chapters:
            self.bible_chapter_label.setText("No hay capítulos disponibles")
            self.bible_book_section_label.setText(
                f"{book_name.upper()} · {self.bible_current_version}")
            self.bible_verse_count_label.clear()
            self.bible_verses_list.clear()
            self.bible_verse_search.clear()
            self.bible_verse_combo.clear()
            self.bible_verse_combo.addItem("Ir a versículo", None)
            return chapters
        if highlight_chapter is None:
            self.bible_chapter_label.setText("Selecciona un capítulo")
            self.bible_book_section_label.setText(
                f"{book_name.upper()} · {self.bible_current_version}")
            self.bible_verse_count_label.clear()
            self.bible_verses_list.clear()
        return chapters

    def _load_bible_library_verses(self, book_name, chapter_num):
        """Carga versículos para consulta sin modificar la proyección activa."""
        conn = sqlite3.connect(database.DB_PATH)
        cursor = conn.cursor()
        cursor.execute(
            "SELECT verse, text FROM bible "
            "WHERE book = ? AND chapter = ? AND version = ? ORDER BY verse",
            (book_name, chapter_num, self.bible_current_version)
        )
        verses = cursor.fetchall()
        conn.close()

        if verses:
            self._populate_bible_verse_list(
                verses, book_name, chapter_num, verses[0][0])
        else:
            self.bible_verses_list.clear()
            self.bible_verse_count_label.clear()
            self.bible_chapter_label.setText("Sin versículos disponibles")

    def _populate_bible_verse_list(self, verses, book_name, chapter_num, selected_verse):
        """Muestra el capítulo en la biblioteca con filas seleccionables."""
        self.bible_verses_list.clear()
        self.bible_verse_search.blockSignals(True)
        self.bible_verse_search.clear()
        self.bible_verse_search.blockSignals(False)
        self.bible_verse_combo.blockSignals(True)
        self.bible_verse_combo.clear()
        self.bible_verse_combo.addItem("Ir a versículo", None)
        gospels = {"Mateo", "Marcos", "Lucas", "Juan"}
        if book_name in gospels:
            section_title = f"EVANGELIO SEGÚN {book_name.upper()}"
        elif book_name in self.BIBLE_BOOKS_OT:
            section_title = "ANTIGUO TESTAMENTO"
        else:
            section_title = "NUEVO TESTAMENTO"

        self.bible_book_section_label.setText(section_title)
        self.bible_chapter_label.setText(f"Capítulo {chapter_num}")
        self.bible_verse_count_label.setText(
            f"{len(verses)} versículos")

        selected_item = None
        for verse_num, verse_text in verses:
            self.bible_verse_combo.addItem(
                f"Versículo {verse_num}", verse_num)
            item = QListWidgetItem(self.bible_verses_list)
            item.setData(Qt.ItemDataRole.UserRole, verse_num)

            row = QWidget()
            row.setObjectName("bibleVersePreviewRow")
            row.setProperty("selected", verse_num == selected_verse)
            row_layout = QHBoxLayout(row)
            row_layout.setContentsMargins(8, 6, 8, 6)
            row_layout.setSpacing(10)

            number_label = QLabel(str(verse_num))
            number_label.setObjectName("bibleVerseNumber")
            number_label.setFixedWidth(24)
            number_label.setAlignment(
                Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignHCenter)
            row_layout.addWidget(number_label)

            text_label = QLabel(verse_text)
            text_label.setObjectName("bibleVerseText")
            text_label.setWordWrap(True)
            text_label.setAttribute(
                Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
            number_label.setAttribute(
                Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
            row_layout.addWidget(text_label, stretch=1)

            item.setSizeHint(row.sizeHint())
            self.bible_verses_list.setItemWidget(item, row)
            if verse_num == selected_verse:
                selected_item = item

        selected_combo_index = self.bible_verse_combo.findData(selected_verse)
        if selected_combo_index >= 0:
            self.bible_verse_combo.setCurrentIndex(selected_combo_index)
        self.bible_verse_combo.blockSignals(False)

        if selected_item:
            self.bible_verses_list.setCurrentItem(selected_item)
            self.bible_verses_list.scrollToItem(selected_item)

    def filter_bible_verses(self, query):
        """Filtra las tarjetas del capítulo por número o contenido."""
        normalized_query = remove_accents(query.strip())
        for index in range(self.bible_verses_list.count()):
            item = self.bible_verses_list.item(index)
            verse_number = str(item.data(Qt.ItemDataRole.UserRole) or "")
            row = self.bible_verses_list.itemWidget(item)
            text_label = row.findChild(QLabel, "bibleVerseText") if row else None
            verse_text = text_label.text() if text_label else ""
            searchable_text = remove_accents(
                f"{verse_number} {verse_text}")
            item.setHidden(bool(normalized_query) and
                           normalized_query not in searchable_text)

    def _select_bible_verse(self, verse_num):
        """Salta a un versículo seleccionado desde el combo."""
        if verse_num is None:
            return
        self.bible_verse_search.clear()
        for index in range(self.bible_verses_list.count()):
            item = self.bible_verses_list.item(index)
            if item.data(Qt.ItemDataRole.UserRole) == verse_num:
                self._on_bible_verse_clicked(item)
                return

    def _on_bible_verse_clicked(self, item):
        """Selecciona y proyecta el versículo equivalente en la Preview izquierda."""
        verse_num = item.data(Qt.ItemDataRole.UserRole)
        chapter_num = self.bible_chapter_combo.currentData()
        if chapter_num is not None:
            self.load_bible_chapter_to_preview(
                self.bible_current_book,
                int(chapter_num),
                focus_verse=verse_num,
            )

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
        self._update_guion_buttons()

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
        self.bible_current_chapter = chapter_num
        # Si no hay focus_verse, resetear el versículo recordado
        if focus_verse is None:
            self.bible_current_verse = None
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
            reference = f"{book_name} {chapter_num}"

            list_item = QListWidgetItem(self.preview_list)
            list_item.setData(Qt.ItemDataRole.UserRole, {
                "mode": "text",
                "text": verse_text,
                "header": reference,
                "song_title": self.bible_current_version,
                "display_mode": "bible",
                "chapter_title": "",
                "bible_verse": verse_num
            })

            # Contenedor: tarjeta + botón "+ Guión" a la derecha
            wrapper = QFrame()
            wrapper.setObjectName("biblePreviewRow")
            wrapper.setStyleSheet("""
                QFrame#biblePreviewRow {
                    background: transparent;
                    border: 1.5px solid transparent;
                    border-radius: 8px;
                }
            """)
            h = QHBoxLayout(wrapper)
            h.setContentsMargins(4, 2, 4, 2)
            h.setSpacing(6)

            card_widget = PreviewCardWidget(header=header, lyrics=verse_text)
            h.addWidget(card_widget, stretch=1)

            wrapper.set_selected = card_widget.set_selected

            guion_btn = QPushButton("+ Guión")
            guion_btn.setObjectName("addToGuionBtn")
            guion_btn.setFixedWidth(82)
            guion_btn.setCursor(Qt.CursorShape.PointingHandCursor)
            guion_btn.setStyleSheet("""
                QPushButton {
                    background-color: #2563eb;
                    color: #ffffff;
                    border: 1px solid #3b82f6;
                    border-radius: 14px;
                    padding: 6px 12px;
                    font-size: 11px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #1d4ed8;
                    border: 1px solid #60a5fa;
                }
            """)
            guion_btn.clicked.connect(
                lambda _, b=book_name, c=chapter_num, v=verse_num, t=verse_text:
                self.add_bible_to_guion(b, c, v, t)
            )
            h.addWidget(guion_btn, alignment=Qt.AlignmentFlag.AlignVCenter)

            def make_set_in_guion(w=wrapper, g=guion_btn):
                def set_in_guion(in_guion):
                    if in_guion:
                        w.setStyleSheet("""
                            QFrame#biblePreviewRow {
                                background-color: rgba(99, 102, 241, 0.12);
                                border: 1.5px solid #6366f1;
                                border-radius: 8px;
                            }
                        """)
                        g.setText("✓ En Guión")
                        g.setStyleSheet("""
                            QPushButton {
                                background-color: #1e1b4b;
                                color: #a5b4fc;
                                border: 1px solid #6366f1;
                                border-radius: 14px;
                                padding: 6px 12px;
                                font-size: 11px;
                                font-weight: bold;
                            }
                            QPushButton:hover {
                                background-color: #312e81;
                                color: #ffffff;
                                border: 1px solid #818cf8;
                            }
                        """)
                    else:
                        w.setStyleSheet("""
                            QFrame#biblePreviewRow {
                                background: transparent;
                                border: 1.5px solid transparent;
                                border-radius: 8px;
                            }
                        """)
                        g.setText("+ Guión")
                        g.setStyleSheet("""
                            QPushButton {
                                background-color: #2563eb;
                                color: #ffffff;
                                border: 1px solid #3b82f6;
                                border-radius: 14px;
                                padding: 6px 12px;
                                font-size: 11px;
                                font-weight: bold;
                            }
                            QPushButton:hover {
                                background-color: #1d4ed8;
                                border: 1px solid #60a5fa;
                            }
                        """)
                return set_in_guion

            wrapper.set_in_guion = make_set_in_guion()

            list_item.setSizeHint(wrapper.minimumSizeHint())
            self.preview_list.setItemWidget(list_item, wrapper)

            if focus_verse and verse_num == focus_verse:
                focus_item = list_item

        # Actualizar estado de botones/sombreado de versículos en Guión
        self._update_guion_buttons()

        # Seleccionar y proyectar: el versículo buscado, o el primero por defecto
        target_item = focus_item if focus_item else self.preview_list.item(0)
        if target_item:
            target_data = target_item.data(Qt.ItemDataRole.UserRole) or {}
            self._populate_bible_verse_list(
                verses,
                book_name,
                chapter_num,
                target_data.get("bible_verse"),
            )
            self.preview_list.setCurrentItem(target_item)
            self.preview_list.scrollToItem(target_item)
            self.on_preview_card_clicked(target_item)

    # ---------------------------------------------------------------------------
    # Script module
    # ---------------------------------------------------------------------------

    # Paleta visual por tipo de item
    GUION_COLORS = {
        "song":   {"icon": "♫",  "label": "CANTO"},
        "bible":  {"icon": "▤", "label": "BIBLIA"},
        "image":  {"icon": "▧", "label": "IMAGEN"},
        "pptx":   {"icon": "▣", "label": "PPTX"},
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

        clear_btn = QPushButton("Limpiar todo")
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
        self.guion_list.setObjectName("guionList")
        self.guion_list.setSpacing(7)
        self.guion_list.itemClicked.connect(self.on_guion_item_clicked)
        self.guion_list.itemDoubleClicked.connect(
            self.on_guion_item_double_clicked)
        layout.addWidget(self.guion_list)

        # Estado vacío
        self.guion_empty_lbl = QLabel(
            "Agrega cantos, versículos, imágenes o\npresentaciones desde las otras pestañas")
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

            widget = self._make_guion_row(
                item_type, item_data, colors, self.guion_list.count())
            list_item.setSizeHint(widget.minimumSizeHint())
            self.guion_list.setItemWidget(list_item, widget)

    def _make_guion_row(self, item_type, item_data, colors, order):
        """Crea el widget visual de una fila del Guión."""
        row = QFrame()
        row.setObjectName("guionRow")
        layout = QHBoxLayout(row)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(12)

        order_lbl = QLabel(f"{order:02d}")
        order_lbl.setObjectName("guionOrder")
        order_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        order_lbl.setFixedWidth(34)
        order_lbl.setAttribute(
            Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        layout.addWidget(order_lbl)

        # Icono neutro y etiqueta de tipo para escanear la secuencia rápidamente.
        icon_lbl = QLabel(colors["icon"])
        icon_lbl.setObjectName("guionTypeIcon")
        icon_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon_lbl.setFixedSize(28, 28)
        icon_lbl.setAttribute(
            Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        layout.addWidget(icon_lbl)

        type_lbl = QLabel(colors["label"])
        type_lbl.setObjectName("guionTypeLabel")
        type_lbl.setFixedWidth(74)
        type_lbl.setAttribute(
            Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        layout.addWidget(type_lbl)

        title_lbl = QLabel(item_data.get(
            "title", item_data.get("name", "Sin título")))
        title_lbl.setObjectName("guionItemTitle")
        title_lbl.setMinimumWidth(120)
        title_lbl.setAttribute(
            Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        layout.addWidget(title_lbl, stretch=1)

        detail_text = ""
        if item_type == "bible":
            detail_text = f"{item_data.get('book', '')} {item_data.get('chapter', '')}"
        elif item_type == "image":
            detail_text = "Archivo local"
        elif item_type == "pptx":
            detail_text = "Presentación"

        detail_lbl = QLabel(detail_text)
        detail_lbl.setObjectName("guionItemDetail")
        detail_lbl.setMinimumWidth(90)
        detail_lbl.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        detail_lbl.setAttribute(
            Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        layout.addWidget(detail_lbl)

        del_btn = QPushButton("×")
        del_btn.setObjectName("guionRemoveBtn")
        del_btn.setFixedSize(26, 26)
        del_btn.setToolTip("Quitar del Guion")
        del_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        del_btn.clicked.connect(
            lambda: self.guion_remove_item(item_type, item_data))
        layout.addWidget(del_btn)

        arrow_lbl = QLabel("›")
        arrow_lbl.setObjectName("guionRowArrow")
        arrow_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        arrow_lbl.setFixedWidth(18)
        arrow_lbl.setAttribute(
            Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        layout.addWidget(arrow_lbl)

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
        """Actualiza el estado visual y bordes/sombreado de los ítems en Guión en todas las pestañas."""
        # 1. Cantos
        if hasattr(self, "songs_list"):
            for i in range(self.songs_list.count()):
                item = self.songs_list.item(i)
                widget = self.songs_list.itemWidget(item)
                if widget and hasattr(widget, "set_in_guion"):
                    data = item.data(Qt.ItemDataRole.UserRole)
                    if data:
                        in_guion = any(s.get("id") == data.get("id") for s in self.guion.get("songs", []))
                        widget.set_in_guion(in_guion)

        # 2. Imágenes
        if hasattr(self, "images_list"):
            for i in range(self.images_list.count()):
                item = self.images_list.item(i)
                widget = self.images_list.itemWidget(item)
                if widget and hasattr(widget, "set_in_guion"):
                    data = item.data(Qt.ItemDataRole.UserRole)
                    if data:
                        in_guion = any(
                            (img.get("id") is not None and img.get("id") == data.get("id")) or
                            (img.get("name") and img.get("name") == data.get("name"))
                            for img in self.guion.get("images", [])
                        )
                        widget.set_in_guion(in_guion)

        # 3. PPTX / PDF
        if hasattr(self, "pptx_list"):
            for i in range(self.pptx_list.count()):
                item = self.pptx_list.item(i)
                widget = self.pptx_list.itemWidget(item)
                if widget and hasattr(widget, "set_in_guion"):
                    data = item.data(Qt.ItemDataRole.UserRole)
                    if data:
                        in_guion = any(
                            (p.get("id") is not None and p.get("id") == data.get("id")) or
                            (p.get("name") and p.get("name") == data.get("name"))
                            for p in self.guion.get("pptx", [])
                        )
                        widget.set_in_guion(in_guion)

        # 4. Versículos de la Biblia en la lista de previsualización
        if hasattr(self, "preview_list") and hasattr(self, "bible_current_book") and self.bible_current_book:
            for i in range(self.preview_list.count()):
                item = self.preview_list.item(i)
                widget = self.preview_list.itemWidget(item)
                if widget and hasattr(widget, "set_in_guion"):
                    data = item.data(Qt.ItemDataRole.UserRole)
                    if data and "bible_verse" in data:
                        v = data["bible_verse"]
                        b = self.bible_current_book
                        c = self.bible_current_chapter
                        in_guion = any(
                            bv.get("book") == b and bv.get("chapter") == c and bv.get("verse") == v
                            for bv in self.guion.get("bible", [])
                        )
                        widget.set_in_guion(in_guion)

    def on_guion_item_clicked(self, list_item):
        """Click simple → carga en preview."""
        data = list_item.data(Qt.ItemDataRole.UserRole)
        if not data:
            return
        self._set_active_guion_row(list_item)
        self._load_guion_item_to_preview(data["type"], data["data"])

    def on_guion_item_double_clicked(self, list_item):
        """Doble click → carga en preview y proyecta directo."""
        data = list_item.data(Qt.ItemDataRole.UserRole)
        if not data:
            return
        self._set_active_guion_row(list_item)
        self._load_guion_item_to_preview(
            data["type"], data["data"], auto_project=True)

    def _set_active_guion_row(self, active_item):
        """Resalta solo la fila activa del guion."""
        for index in range(self.guion_list.count()):
            item = self.guion_list.item(index)
            widget = self.guion_list.itemWidget(item)
            if not widget:
                continue
            widget.setProperty("active", item is active_item)
            widget.style().unpolish(widget)
            widget.style().polish(widget)

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
        self.image_search.setPlaceholderText("Buscar imágenes...")
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
            elif not os.path.exists(abs_path):
                base_name = os.path.basename(file_path)
                candidate1 = os.path.join(PROJECT_DIR, "imported_images", base_name)
                candidate2 = os.path.join(PROJECT_DIR, base_name)
                if os.path.exists(candidate1):
                    abs_path = candidate1
                elif os.path.exists(candidate2):
                    abs_path = candidate2

            item = QListWidgetItem(self.images_list)
            item.setData(Qt.ItemDataRole.UserRole, {
                         "id": image_id, "name": name, "file_path": abs_path})

            def make_add_callback(iid=image_id, iname=name, ipath=abs_path):
                return lambda: self.add_image_to_guion(iid, iname, ipath)

            def make_delete_callback(iid=image_id, iname=name, ipath=abs_path):
                return lambda: self.delete_image(iid, iname, ipath)

            widget = LibraryImageRowWidget(
                title=name,
                file_path=abs_path,
                on_add_clicked=make_add_callback(),
                on_delete_clicked=make_delete_callback()
            )
            item.setSizeHint(widget.minimumSizeHint())
            self.images_list.setItemWidget(item, widget)

        conn.close()
        self._update_guion_buttons()

    def delete_image(self, image_id, name, file_path):
        reply = QMessageBox.question(
            self,
            "Eliminar imagen",
            f"¿Estás seguro que quieres eliminar «{name}»?\nEsta acción no se puede deshacer.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Cancel
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        conn = sqlite3.connect(database.DB_PATH)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM images WHERE id = ?", (image_id,))
        conn.commit()
        conn.close()

        # Borrar el archivo del disco
        if file_path and os.path.exists(file_path):
            try:
                os.remove(file_path)
            except Exception:
                pass

        self.preview_list.clear()
        self.load_images_library(self.image_search.text())

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
        self._update_guion_buttons()

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
        self.pptx_search.setPlaceholderText("Buscar presentaciones...")
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
        self._update_guion_buttons()

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
        self._update_guion_buttons()

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

                        # Cargar el pixmap limpio (el paintEvent maneja el DPR)
                        pixmap = QPixmap(abs_dest_path)
                        self.local_projection_widget.general_bg_pixmap = pixmap
                        for win in list(self.projection_windows.values()):
                            if win and win.isVisible():
                                win.projection_widget.general_bg_pixmap = pixmap
                                win.projection_widget.update()

                        self.local_projection_widget.update()

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
            for win in list(self.projection_windows.values()):
                if win and win.isVisible():
                    win.projection_widget.general_bg_pixmap = None
                    win.projection_widget.update()

            self.local_projection_widget.update()

            QMessageBox.information(
                self, "Fondo Quitado", "Se ha quitado el fondo general para las letras (ahora es negro sólido).")

    def load_persisted_lyrics_bg(self):
        bg_config_path = os.path.join(PROJECT_DIR, "lyrics_bg_path.txt")
        if not os.path.exists(bg_config_path):
            if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
                meipass_bg = os.path.join(sys._MEIPASS, "lyrics_bg_path.txt")
                if os.path.exists(meipass_bg):
                    bg_config_path = meipass_bg

        if os.path.exists(bg_config_path):
            try:
                with open(bg_config_path, "r", encoding="utf-8") as f:
                    path = f.read().strip()
                if path:
                    if not os.path.exists(path):
                        base = os.path.basename(path)
                        candidate1 = os.path.join(PROJECT_DIR, "imported_images", base)
                        candidate2 = os.path.join(PROJECT_DIR, base)
                        if os.path.exists(candidate1):
                            path = candidate1
                        elif os.path.exists(candidate2):
                            path = candidate2

                    if os.path.exists(path):
                        pixmap = QPixmap(path)
                        if not pixmap.isNull():
                            self.local_projection_widget.general_bg_pixmap = pixmap
                            for win in list(self.projection_windows.values()):
                                if win and win.isVisible():
                                    win.projection_widget.general_bg_pixmap = pixmap
                                    win.projection_widget.update()
                            self.local_projection_widget.update()
            except Exception as e:
                print(f"Error cargando fondo de letras persistido: {e}")

    # ---------------------------------------------------------------------------
    # Quick image shortcuts
    # ---------------------------------------------------------------------------

    def get_quick_image_path(self, image_type, title_name):
        """Obtiene la ruta asignada a la imagen rápida (Bienvenida o Despedida)."""
        config_file = os.path.join(PROJECT_DIR, f"quick_{image_type}_path.txt")
        if os.path.exists(config_file):
            try:
                with open(config_file, "r", encoding="utf-8") as f:
                    saved_path = f.read().strip()
                if saved_path == "NONE":
                    return None
                if saved_path:
                    if os.path.isabs(saved_path) and os.path.exists(saved_path):
                        return saved_path
                    cand = os.path.join(PROJECT_DIR, saved_path)
                    if os.path.exists(cand):
                        return cand
                    base = os.path.basename(saved_path)
                    cand2 = os.path.join(PROJECT_DIR, "imported_images", base)
                    if os.path.exists(cand2):
                        return cand2
            except Exception as e:
                print(f"Error leyendo config de acceso rápido {image_type}: {e}")

        # Buscar por defecto en la carpeta imported_images
        imported_dir = os.path.join(PROJECT_DIR, "imported_images")
        if os.path.exists(imported_dir):
            for filename in os.listdir(imported_dir):
                name_without_ext, _ = os.path.splitext(filename)
                if name_without_ext.strip().lower() == title_name.strip().lower():
                    cand = os.path.join(imported_dir, filename)
                    if os.path.exists(cand):
                        return cand

        # Buscar en la base de datos
        try:
            conn = sqlite3.connect(database.DB_PATH)
            conn.create_function("remove_accents", 1, remove_accents)
            cursor = conn.cursor()
            norm_title = remove_accents(title_name)
            cursor.execute(
                "SELECT file_path FROM images WHERE remove_accents(name) LIKE ? ORDER BY id DESC LIMIT 1",
                (f"%{norm_title}%",)
            )
            row = cursor.fetchone()
            conn.close()
            if row:
                db_path = row[0]
                abs_path = db_path if os.path.isabs(db_path) else os.path.join(PROJECT_DIR, db_path)
                if os.path.exists(abs_path):
                    return abs_path
                base = os.path.basename(db_path)
                cand = os.path.join(PROJECT_DIR, "imported_images", base)
                if os.path.exists(cand):
                    return cand
        except Exception as e:
            print(f"Error buscando {title_name} en DB: {e}")

        return None

    def project_quick_image(self, image_type, title_name):
        """Proyecta directamente la imagen de Bienvenida o Despedida."""
        path = self.get_quick_image_path(image_type, title_name)
        if not path or not os.path.exists(path):
            reply = QMessageBox.question(
                self,
                f"Configurar {title_name}",
                f"No hay una imagen asignada para «{title_name}».\n¿Deseas seleccionar una imagen ahora?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes
            )
            if reply == QMessageBox.StandardButton.Yes:
                self.change_quick_image(image_type, title_name)
            return

        self.load_image_to_preview(title_name, path)
        if self.preview_list.count() > 0:
            first_item = self.preview_list.item(0)
            self.preview_list.setCurrentItem(first_item)
            self.on_preview_card_clicked(first_item)

    def change_quick_image(self, image_type, title_name):
        """Permite seleccionar una nueva imagen desde el PC para Bienvenida o Despedida."""
        file_dialog = QFileDialog(self)
        file_dialog.setWindowTitle(f"Seleccionar Imagen para {title_name}")
        file_dialog.setNameFilter("Imágenes (*.png *.jpg *.jpeg *.bmp *.webp)")
        if file_dialog.exec():
            selected_files = file_dialog.selectedFiles()
            if not selected_files:
                return

            file_path = selected_files[0]
            if not os.path.exists(file_path):
                return

            imported_dir = os.path.join(PROJECT_DIR, "imported_images")
            os.makedirs(imported_dir, exist_ok=True)

            _, ext = os.path.splitext(file_path)
            dest_filename = f"{title_name}{ext}"
            dest_path = os.path.join(imported_dir, dest_filename)

            try:
                shutil.copy2(file_path, dest_path)
                abs_dest_path = os.path.abspath(dest_path)

                config_file = os.path.join(PROJECT_DIR, f"quick_{image_type}_path.txt")
                with open(config_file, "w", encoding="utf-8") as f:
                    f.write(abs_dest_path)

                conn = sqlite3.connect(database.DB_PATH)
                cursor = conn.cursor()
                rel_path = os.path.join("imported_images", dest_filename)
                cursor.execute("SELECT id FROM images WHERE name = ?", (dest_filename,))
                row = cursor.fetchone()
                if not row:
                    cursor.execute("INSERT INTO images (name, file_path) VALUES (?, ?)", (dest_filename, rel_path))
                else:
                    cursor.execute("UPDATE images SET file_path = ? WHERE id = ?", (rel_path, row[0]))
                conn.commit()
                conn.close()

                if hasattr(self, "images_list") and hasattr(self, "image_search"):
                    self.load_images_library(self.image_search.text())

                # Proyectar inmediatamente la nueva imagen
                self.project_quick_image(image_type, title_name)

                QMessageBox.information(
                    self,
                    "Imagen Guardada",
                    f"Se ha asignado y proyectado correctamente la imagen de «{title_name}»."
                )
            except Exception as e:
                QMessageBox.warning(
                    self,
                    "Error",
                    f"No se pudo guardar la imagen: {str(e)}"
                )

    def remove_quick_image(self, image_type, title_name):
        """Quita la asignación actual de la imagen rápida."""
        reply = QMessageBox.question(
            self,
            "Quitar Imagen Rápida",
            f"¿Estás seguro de que deseas desasignar la imagen de «{title_name}»?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Cancel
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        config_file = os.path.join(PROJECT_DIR, f"quick_{image_type}_path.txt")
        try:
            with open(config_file, "w", encoding="utf-8") as f:
                f.write("NONE")
        except Exception:
            pass

        QMessageBox.information(
            self,
            "Imagen Removida",
            f"Se ha quitado la asignación rápida de «{title_name}»."
        )

    def show_quick_image_menu(self, image_type, title_name, global_pos):
        """Muestra el menú contextual con opciones para la imagen rápida."""
        menu = QMenu(self)
        menu.setStyleSheet("""
            QMenu {
                background-color: #18181b;
                color: #f4f4f5;
                border: 1px solid #3f3f46;
                border-radius: 8px;
                padding: 4px;
            }
            QMenu::item {
                padding: 8px 20px;
                border-radius: 4px;
            }
            QMenu::item:selected {
                background-color: #3f3f46;
            }
            QMenu::separator {
                height: 1px;
                background: #3f3f46;
                margin: 4px 8px;
            }
        """)

        proj_action = menu.addAction(f"▧  Proyectar {title_name}")
        proj_action.triggered.connect(
            lambda: self.project_quick_image(image_type, title_name))

        change_action = menu.addAction("📁  Cambiar Imagen...")
        change_action.triggered.connect(
            lambda: self.change_quick_image(image_type, title_name))

        path = self.get_quick_image_path(image_type, title_name)
        if path and os.path.exists(path):
            menu.addSeparator()
            remove_action = menu.addAction("✕  Quitar Imagen")
            remove_action.triggered.connect(
                lambda: self.remove_quick_image(image_type, title_name))

        menu.exec(global_pos)

    # ---------------------------------------------------------------------------
    # Song library load and filtering
    # ---------------------------------------------------------------------------

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
        self._update_guion_buttons()

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

    # ---------------------------------------------------------------------------
    # Navigation
    # ---------------------------------------------------------------------------

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
        icons = {"book": "▤", "chapter": "▣", "verse": "•"}
        t = sug["type"]

        icon_lbl = QLabel(icons.get(t, "▤"))
        icon_lbl.setFixedSize(30, 30)
        icon_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon_lbl.setStyleSheet("""
            background-color: #202923;
            color: #a9bbb2;
            border: 1px solid #3a4b40;
            border-radius: 5px;
            font-size: 14px;
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
        self.current_nav = active_name
        for name, btn in self.nav_buttons.items():
            is_active = (name == active_name)
            btn.setProperty("active", "true" if is_active else "false")
            btn.style().polish(btn)
        self._resize_library_for_active_page()
        self._update_guion_buttons()

    def _resize_library_for_active_page(self):
        """Amplía el Guion sin quitarle demasiado espacio a la salida."""
        if not hasattr(self, "right_splitter") or self.right_splitter.height() <= 0:
            return

        total_height = self.right_splitter.height()
        guion_active = self.current_nav == "Guión"
        self.obs_container.setMinimumHeight(280 if guion_active else 320)

        if guion_active:
            library_height = int(total_height * 0.54)
            output_height = total_height - library_height
            self.right_splitter.setSizes([library_height, output_height])
        elif self._is_external_obs_mode:
            library_height = max(int(total_height * 0.60), 300)
            self.right_splitter.setSizes(
                [library_height, total_height - library_height])
        else:
            self.right_splitter.setSizes([2000, 4000])

    # ---------------------------------------------------------------------------
    # Library interaction
    # ---------------------------------------------------------------------------

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

    # ---------------------------------------------------------------------------
    # Preview panel loader
    # ---------------------------------------------------------------------------

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

    # ---------------------------------------------------------------------------
    # Projection actions
    # ---------------------------------------------------------------------------

    @property
    def projection_window(self):
        """Propiedad de compatibilidad: retorna la primera ventana de proyección activa."""
        for win in self.projection_windows.values():
            if win and win.isVisible():
                return win
        return None

    @property
    def projection_window_2(self):
        """Propiedad de compatibilidad: retorna la segunda ventana de proyección activa."""
        wins = [w for w in self.projection_windows.values() if w and w.isVisible()]
        return wins[1] if len(wins) > 1 else None

    # ---------------------------------------------------------------------------
    # Stable multi-screen projection actions
    # ---------------------------------------------------------------------------

    def project_selected_card(self, item):
        data = item.data(Qt.ItemDataRole.UserRole)
        if data:
            self.project_active_item(data)

    def project_active_item(self, data):
        if not data:
            return

        mode = data.get("mode")
        if mode == "text":
            text = data["text"]
            header = data.get("header", "")
            song_title = data.get("song_title", "")
            display_mode = data.get("display_mode", "standard")
            verse_number = data.get("bible_verse")
            chapter_title = data.get("chapter_title", "")

            if "bible_verse" in data:
                self.bible_current_verse = data["bible_verse"]
                self.bible_verse_search.clear()
                combo_index = self.bible_verse_combo.findData(verse_number)
                if combo_index >= 0:
                    self.bible_verse_combo.setCurrentIndex(combo_index)
                for index in range(self.bible_verses_list.count()):
                    verse_item = self.bible_verses_list.item(index)
                    verse_row = self.bible_verses_list.itemWidget(verse_item)
                    if verse_row:
                        is_selected = (
                            verse_item.data(Qt.ItemDataRole.UserRole)
                            == verse_number
                        )
                        verse_row.setProperty("selected", is_selected)
                        verse_row.style().unpolish(verse_row)
                        verse_row.style().polish(verse_row)
                    if verse_item.data(Qt.ItemDataRole.UserRole) == verse_number:
                        self.bible_verses_list.setCurrentItem(verse_item)
                        self.bible_verses_list.scrollToItem(verse_item)
                        break

            self.current_projection_mode = "text"
            self.last_projected_mode = "text"
            self.last_projected_text = text
            self.last_projected_header = header
            self.last_projected_song_title = song_title
            self.last_projected_display_mode = display_mode
            self.last_projected_verse_number = verse_number
            self.last_projected_chapter_title = chapter_title

            self.local_projection_widget.display_text(
                text,
                header,
                song_title,
                display_mode=display_mode,
                verse_number=verse_number,
                chapter_title=chapter_title,
            )

            for win in list(self.projection_windows.values()):
                if win and win.isVisible():
                    win.projection_widget.display_text(
                        text,
                        header,
                        song_title,
                        display_mode=display_mode,
                        verse_number=verse_number,
                        chapter_title=chapter_title,
                    )

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

            # Todas las pantallas externas activas
            for win in list(self.projection_windows.values()):
                if win and win.isVisible():
                    win.projection_widget.display_image(file_path, header, name)

    def project_black(self):
        self.current_projection_mode = "black"
        self.local_projection_widget.set_black_screen()
        for win in list(self.projection_windows.values()):
            if win and win.isVisible():
                win.projection_widget.set_black_screen()

    def project_clear_text(self):
        if self.current_projection_mode == "text":
            self.local_projection_widget.clear_text_only()
            for win in list(self.projection_windows.values()):
                if win and win.isVisible():
                    win.projection_widget.clear_text_only()

    def project_restore(self):
        if self.last_projected_mode == "text":
            self.current_projection_mode = "text"
            self.local_projection_widget.display_text(
                self.last_projected_text,
                self.last_projected_header,
                self.last_projected_song_title,
                display_mode=self.last_projected_display_mode,
                verse_number=self.last_projected_verse_number,
                chapter_title=self.last_projected_chapter_title,
            )
            for win in list(self.projection_windows.values()):
                if win and win.isVisible():
                    win.projection_widget.display_text(
                        self.last_projected_text,
                        self.last_projected_header,
                        self.last_projected_song_title,
                        display_mode=self.last_projected_display_mode,
                        verse_number=self.last_projected_verse_number,
                        chapter_title=self.last_projected_chapter_title,
                    )
        elif self.last_projected_mode == "image":
            self.current_projection_mode = "image"
            self.local_projection_widget.display_image(
                self.last_projected_image_path, self.last_projected_header, self.last_projected_song_title
            )
            for win in list(self.projection_windows.values()):
                if win and win.isVisible():
                    win.projection_widget.display_image(
                        self.last_projected_image_path, self.last_projected_header, self.last_projected_song_title
                    )

    # ---------------------------------------------------------------------------
    # Multi-screen projection manager
    # ---------------------------------------------------------------------------

    def is_screen_projecting(self, screen):
        """Verifica si una pantalla específica tiene una ventana de proyección abierta y visible."""
        win = self.projection_windows.get(screen)
        return bool(win and win.isVisible())

    def toggle_screen_projection(self, screen):
        """Alterna (activa/desactiva) la proyección en una pantalla específica sin afectar las demás."""
        if self.is_screen_projecting(screen):
            self.stop_projection_on_screen(screen)
        else:
            self.start_projection_on_screen(screen)

    def start_projection_on_screen(self, screen):
        """Abre la proyección en la pantalla indicada a pantalla completa."""
        # Si ya existe una ventana activa en esta pantalla, no hacer nada
        if screen in self.projection_windows:
            old_win = self.projection_windows[screen]
            if old_win and old_win.isVisible():
                return
            elif old_win:
                try:
                    old_win.close()
                except Exception:
                    pass

        win = ProjectionWindow(close_callback=lambda s=screen: self.on_projection_screen_closed(s))
        win._main_window_ref = self

        screens = QApplication.screens()
        screen_idx = screens.index(screen) + 1 if screen in screens else ""
        win.setWindowTitle(f"Salida de Proyección - Pantalla {screen_idx}")

        # Sincronizar fondo y offset de fuente
        win.projection_widget.general_bg_pixmap = self.local_projection_widget.general_bg_pixmap
        win.projection_widget.font_size_offset = self.external_font_size_offset

        # Posicionar ventana en la pantalla correspondiente en modo pantalla completa
        geo = screen.geometry()
        win.setGeometry(geo)
        win.show()
        if win.windowHandle():
            win.windowHandle().setScreen(screen)
        win.setGeometry(geo)
        win.showFullScreen()

        self.projection_windows[screen] = win

        # Sincronizar contenido proyectado actualmente
        self._sync_projection_window(win)

        self.adjust_obs_container_size(is_external_visible=True)
        self._update_projection_status()

    def stop_projection_on_screen(self, screen):
        """Detiene y cierra la proyección en una pantalla específica."""
        if screen in self.projection_windows:
            win = self.projection_windows.pop(screen, None)
            if win:
                win.close_callback = None
                win.close()

        has_active = any(w and w.isVisible() for w in self.projection_windows.values())
        self.adjust_obs_container_size(is_external_visible=has_active)
        self._update_projection_status()

    def on_projection_screen_closed(self, screen):
        """Callback cuando una ventana de proyección externa es cerrada por el usuario o SO."""
        if screen in self.projection_windows:
            self.projection_windows.pop(screen, None)
        has_active = any(w and w.isVisible() for w in self.projection_windows.values())
        self.adjust_obs_container_size(is_external_visible=has_active)
        self._update_projection_status()

    def project_all_secondary_screens(self):
        """Activa la proyección en todas las pantallas secundarias detectadas."""
        screens = QApplication.screens()
        primary = QApplication.primaryScreen()
        secondary = [s for s in screens if s != primary]
        targets = secondary if secondary else screens
        for s in targets:
            self.start_projection_on_screen(s)

    def project_all_screens(self):
        """Activa la proyección en absolutamente todas las pantallas conectadas."""
        for s in QApplication.screens():
            self.start_projection_on_screen(s)

    def close_all_projections(self):
        """Cierra todas las pantallas de proyección activas."""
        for screen, win in list(self.projection_windows.items()):
            if win:
                win.close_callback = None
                win.close()
        self.projection_windows.clear()
        self.adjust_obs_container_size(is_external_visible=False)
        self._update_projection_status()

    def show_screen_menu(self, target=None, pos=None, *args, **kwargs):
        """Muestra el menú interactivo para gestionar la proyección individual o múltiple."""
        screens = QApplication.screens()
        primary_screen = QApplication.primaryScreen()
        menu = QMenu(self)
        menu.setStyleSheet("""
            QMenu {
                background-color: #18181b;
                color: #f4f4f5;
                border: 1px solid #3f3f46;
                border-radius: 8px;
                padding: 6px;
            }
            QMenu::item {
                padding: 8px 24px;
                border-radius: 4px;
                font-size: 12px;
            }
            QMenu::item:selected {
                background-color: #27272a;
                color: #ffffff;
            }
            QMenu::separator {
                height: 1px;
                background: #3f3f46;
                margin: 5px 8px;
            }
        """)

        active_screens = [s for s in screens if self.is_screen_projecting(s)]
        secondary_screens = [s for s in screens if s != primary_screen]

        # Acciones masivas
        if len(screens) > 1:
            all_sec_active = bool(secondary_screens and all(self.is_screen_projecting(s) for s in secondary_screens))
            if not all_sec_active:
                act_all_sec = menu.addAction("▣  Proyectar en TODAS las pantallas secundarias")
                act_all_sec.triggered.connect(self.project_all_secondary_screens)

            if len(active_screens) < len(screens) and len(secondary_screens) > 1:
                act_all = menu.addAction("▣  Proyectar en TODAS las pantallas (incluida principal)")
                act_all.triggered.connect(self.project_all_screens)

            menu.addSeparator()

        if active_screens:
            count_str = f" ({len(active_screens)} activa{'s' if len(active_screens) > 1 else ''})"
            close_all_action = menu.addAction(f"✕  Cerrar todas las proyecciones{count_str}")
            close_all_action.triggered.connect(self.close_all_projections)
            menu.addSeparator()

        # Lista de pantallas individuales
        for i, screen in enumerate(screens):
            geo = screen.geometry()
            is_active = self.is_screen_projecting(screen)
            is_primary = (screen == primary_screen)

            extra_info = []
            if is_primary:
                extra_info.append("Principal / Este monitor")
            if is_active:
                extra_info.append("● PROYECTANDO - Clic para apagar")
            else:
                extra_info.append("Clic para proyectar")

            info_str = " (" + " · ".join(extra_info) + ")" if extra_info else ""
            icon_status = "●" if is_active else "▣"
            label = f"{icon_status}  Pantalla {i + 1}  —  {geo.width()}×{geo.height()}{info_str}"

            action = menu.addAction(label)
            action.triggered.connect(lambda checked, s=screen: self.toggle_screen_projection(s))

        # Posicionamiento del menú
        if isinstance(pos, QPoint):
            menu.exec(pos)
        elif isinstance(target, QPoint):
            menu.exec(target)
        elif isinstance(target, QWidget):
            menu.exec(target.mapToGlobal(target.rect().center()))
        else:
            from PyQt6.QtGui import QCursor
            menu.exec(QCursor.pos())

    def _sync_projection_window(self, win):
        """Sincroniza el contenido actual de proyección a una ventana externa."""
        win.projection_widget.is_black_screen = \
            self.local_projection_widget.is_black_screen
        if self.current_projection_mode == "text":
            if not self.local_projection_widget.text_label.text():
                win.projection_widget.display_text(
                    "",
                    self.last_projected_header,
                    self.last_projected_song_title,
                    display_mode=self.last_projected_display_mode,
                    verse_number=self.last_projected_verse_number,
                    chapter_title=self.last_projected_chapter_title,
                )
            else:
                win.projection_widget.display_text(
                    self.last_projected_text,
                    self.last_projected_header,
                    self.last_projected_song_title,
                    display_mode=self.last_projected_display_mode,
                    verse_number=self.last_projected_verse_number,
                    chapter_title=self.last_projected_chapter_title,
                )
        elif self.current_projection_mode == "image":
            win.projection_widget.display_image(
                self.last_projected_image_path,
                self.last_projected_header,
                self.last_projected_song_title
            )
        elif self.current_projection_mode == "black":
            win.projection_widget.set_black_screen()

    def _set_projection_output(self, screen, enabled):
        """Abre o cierra la ventana nativa de proyección en una pantalla."""
        if enabled != self.is_screen_projecting(screen):
            self.toggle_screen_projection(screen)

    def _refresh_projection_outputs(self):
        """Actualiza la lista de pantallas y el estado de sus interruptores."""
        if not hasattr(self, "projection_outputs_layout"):
            return

        screens = QApplication.screens()
        for screen, row_data in list(self.projection_output_rows.items()):
            if screen not in screens:
                self.projection_outputs_layout.removeWidget(row_data["frame"])
                row_data["frame"].deleteLater()
                del self.projection_output_rows[screen]

        primary_screen = QApplication.primaryScreen()
        for index, screen in enumerate(screens, start=1):
            row_data = self.projection_output_rows.get(screen)
            if row_data is None:
                # Cada interruptor controla una salida de pantalla real.
                row = QFrame()
                row.setObjectName("projectionScreenRow")
                row_layout = QHBoxLayout(row)
                row_layout.setContentsMargins(7, 5, 7, 5)
                row_layout.setSpacing(7)

                icon = QLabel("▣")
                icon.setObjectName("projectionScreenIcon")
                icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
                icon.setFixedSize(30, 30)
                row_layout.addWidget(icon)

                text_layout = QVBoxLayout()
                text_layout.setContentsMargins(0, 0, 0, 0)
                text_layout.setSpacing(2)
                title = QLabel()
                title.setObjectName("projectionScreenName")
                detail = QLabel()
                detail.setObjectName("projectionScreenDetails")
                text_layout.addWidget(title)
                text_layout.addWidget(detail)
                row_layout.addLayout(text_layout, stretch=1)

                toggle = QCheckBox()
                toggle.setObjectName("projectionScreenToggle")
                toggle.setToolTip("Activar o detener la proyección en esta pantalla")
                toggle.toggled.connect(
                    lambda enabled, target=screen:
                    self._set_projection_output(target, enabled)
                )
                row_layout.addWidget(toggle)

                row_data = {
                    "frame": row,
                    "title": title,
                    "detail": detail,
                    "toggle": toggle,
                }
                self.projection_output_rows[screen] = row_data
                self.projection_outputs_layout.addWidget(row)

            geometry = screen.geometry()
            is_primary = screen == primary_screen
            screen_name = "Pantalla principal" if is_primary else (screen.name() or f"Pantalla {index}")
            is_active = self.is_screen_projecting(screen)
            row_data["title"].setText(screen_name)
            status = "Proyectando" if is_active else "Disponible"
            row_data["detail"].setText(
                f"{geometry.width()} × {geometry.height()} · {status}"
            )
            row_data["toggle"].setAccessibleName(screen_name)
            row_data["toggle"].blockSignals(True)
            row_data["toggle"].setChecked(is_active)
            row_data["toggle"].blockSignals(False)
            row = row_data["frame"]
            row.setProperty("active", is_active)
            row.style().unpolish(row)
            row.style().polish(row)

    def _update_projection_status(self):
        """Sincroniza el encabezado y las filas de salida nativa."""
        self.local_projection_widget.projection_title.setText("PROYECCIÓN")
        self._refresh_projection_outputs()

    def _check_screens_connected(self):
        """Verifica cada 2s si las pantallas asignadas siguen conectadas físicamente."""
        available = QApplication.screens()
        disconnected = [s for s in list(self.projection_windows.keys()) if s not in available]
        for s in disconnected:
            win = self.projection_windows.pop(s, None)
            if win:
                try:
                    win.close_callback = None
                    win.close()
                except Exception:
                    pass

        if disconnected:
            has_active = any(w and w.isVisible() for w in self.projection_windows.values())
            self.adjust_obs_container_size(is_external_visible=has_active)
        self._update_projection_status()

    def toggle_projection_window(self):
        """Alterna la proyección en la primera pantalla secundaria disponible."""
        screens = QApplication.screens()
        primary = QApplication.primaryScreen()
        secondary = [s for s in screens if s != primary]
        target = secondary[0] if secondary else primary
        self.toggle_screen_projection(target)

    def toggle_projection_window_2(self):
        """Alterna la proyección en la segunda pantalla secundaria disponible."""
        screens = QApplication.screens()
        primary = QApplication.primaryScreen()
        secondary = [s for s in screens if s != primary]
        if len(secondary) > 1:
            self.toggle_screen_projection(secondary[1])
        elif screens:
            self.toggle_screen_projection(screens[-1])

    def close_projection_window(self, window_num=1):
        """Cierra proyecciones externas."""
        self.close_all_projections()

    def on_projection_window_closed(self):
        has_active = any(w and w.isVisible() for w in self.projection_windows.values())
        self.adjust_obs_container_size(is_external_visible=has_active)
        self._update_projection_status()

    def on_projection_window_2_closed(self):
        has_active = any(w and w.isVisible() for w in self.projection_windows.values())
        self.adjust_obs_container_size(is_external_visible=has_active)
        self._update_projection_status()

    def adjust_obs_container_size(self, is_external_visible):
        has_active = any(w and w.isVisible() for w in self.projection_windows.values()) if is_external_visible else False
        if getattr(self, "_is_external_obs_mode", None) == has_active:
            return
        self._is_external_obs_mode = has_active

        if has_active:
            handle = self.right_splitter.handle(1)
            if handle:
                handle.setEnabled(True)
                handle.setCursor(Qt.CursorShape.SplitVCursor)

            self.obs_container.setMinimumHeight(80)
            self.obs_container.setMaximumHeight(16777215)

            total_height = self.right_splitter.height()
            # Dar al menos 60% al panel de biblioteca, 40% al OBS
            lib_height = max(int(total_height * 0.60), 300)
            obs_height = total_height - lib_height
            self.right_splitter.setSizes([lib_height, obs_height])
        else:
            handle = self.right_splitter.handle(1)
            if handle:
                handle.setEnabled(False)
                handle.setCursor(Qt.CursorShape.ArrowCursor)

            self.obs_container.setMinimumHeight(320)
            self.obs_container.setMaximumHeight(16777215)
            self.right_splitter.setSizes([2000, 4000])

    def _update_font_size_indicator(self):
        """Refleja el nivel de tamaño actual en el botón de restablecimiento."""
        offset = getattr(self, "external_font_size_offset", 0)
        label = "S" if offset < 0 else "L" if offset > 0 else "M"
        self.btn_font_reset.setText(label)

    def increase_font_size(self):
        self.external_font_size_offset = getattr(self, "external_font_size_offset", 0) + 4
        self._update_font_size_indicator()
        self.local_projection_widget.font_size_offset = self.external_font_size_offset
        self.local_projection_widget.adjust_font_size()
        for win in list(self.projection_windows.values()):
            if win and win.isVisible():
                win.projection_widget.font_size_offset = self.external_font_size_offset
                win.projection_widget.adjust_font_size()

    def decrease_font_size(self):
        self.external_font_size_offset = getattr(self, "external_font_size_offset", 0) - 4
        self._update_font_size_indicator()
        self.local_projection_widget.font_size_offset = self.external_font_size_offset
        self.local_projection_widget.adjust_font_size()
        for win in list(self.projection_windows.values()):
            if win and win.isVisible():
                win.projection_widget.font_size_offset = self.external_font_size_offset
                win.projection_widget.adjust_font_size()

    def reset_font_size(self):
        self.external_font_size_offset = 0
        self._update_font_size_indicator()
        self.local_projection_widget.font_size_offset = 0
        self.local_projection_widget.adjust_font_size()
        for win in list(self.projection_windows.values()):
            if win and win.isVisible():
                win.projection_widget.font_size_offset = 0
                win.projection_widget.adjust_font_size()

    # ---------------------------------------------------------------------------
    # Script management
    # ---------------------------------------------------------------------------

    def add_song_to_guion(self, song_id, title):
        if any(s["id"] == song_id for s in self.guion["songs"]):
            QMessageBox.information(
                self, "Guión", f"«{title}» ya está en el Guión.")
            return
        self.guion["songs"].append({"id": song_id, "title": title})
        self.refresh_guion_list()
        self._update_guion_buttons()

    def closeEvent(self, event):
        for win in list(self.projection_windows.values()):
            if win:
                win.close_callback = None
                win.close()
        self.projection_windows.clear()
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
    window.setMinimumSize(1024, 600)
    window.showMaximized()
    sys.exit(app.exec())
    