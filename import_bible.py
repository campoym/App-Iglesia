"""
import_bible.py
Importa un archivo XML de biblia (formato Beblia) a la tabla `bible` de iglesia.db

Uso:
    python import_bible.py SpanishRevisedRVR1960Bible.xml RVR60
"""

import sys
import sqlite3
import xml.etree.ElementTree as ET

DB_PATH = "iglesia.db"

# Mapeo número de libro -> nombre en español (orden estándar protestante, 66 libros)
BOOK_NAMES = {
    1: "Génesis", 2: "Éxodo", 3: "Levítico", 4: "Números", 5: "Deuteronomio",
    6: "Josué", 7: "Jueces", 8: "Rut", 9: "1 Samuel", 10: "2 Samuel",
    11: "1 Reyes", 12: "2 Reyes", 13: "1 Crónicas", 14: "2 Crónicas",
    15: "Esdras", 16: "Nehemías", 17: "Ester", 18: "Job", 19: "Salmos",
    20: "Proverbios", 21: "Eclesiastés", 22: "Cantares", 23: "Isaías",
    24: "Jeremías", 25: "Lamentaciones", 26: "Ezequiel", 27: "Daniel",
    28: "Oseas", 29: "Joel", 30: "Amós", 31: "Abdías", 32: "Jonás",
    33: "Miqueas", 34: "Nahúm", 35: "Habacuc", 36: "Sofonías", 37: "Hageo",
    38: "Zacarías", 39: "Malaquías",
    # Nuevo Testamento (continúa la numeración 40-66)
    40: "Mateo", 41: "Marcos", 42: "Lucas", 43: "Juan", 44: "Hechos",
    45: "Romanos", 46: "1 Corintios", 47: "2 Corintios", 48: "Gálatas",
    49: "Efesios", 50: "Filipenses", 51: "Colosenses", 52: "1 Tesalonicenses",
    53: "2 Tesalonicenses", 54: "1 Timoteo", 55: "2 Timoteo", 56: "Tito",
    57: "Filemón", 58: "Hebreos", 59: "Santiago", 60: "1 Pedro", 61: "2 Pedro",
    62: "1 Juan", 63: "2 Juan", 64: "3 Juan", 65: "Judas", 66: "Apocalipsis",
}


def import_bible(xml_path, version_code):
    print(f"Leyendo {xml_path} ...")
    tree = ET.parse(xml_path)
    root = tree.getroot()

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    # Asegurar columna testament por si no existe (migración segura)
    try:
        cursor.execute("ALTER TABLE bible ADD COLUMN testament TEXT")
        conn.commit()
    except Exception:
        pass

    # Borrar datos previos de esta versión para evitar duplicados al re-importar
    cursor.execute("DELETE FROM bible WHERE version = ?", (version_code,))
    conn.commit()

    rows_to_insert = []
    book_global_counter = 0  # contador absoluto de libros (1..66) sin importar el testamento

    for testament in root.findall("testament"):
        testament_name = testament.get("name", "")  # "Old" o "New"
        testament_es = "AT" if testament_name == "Old" else "NT"

        for book in testament.findall("book"):
            book_global_counter += 1
            book_name = BOOK_NAMES.get(book_global_counter, f"Libro {book_global_counter}")

            for chapter in book.findall("chapter"):
                chapter_num = int(chapter.get("number"))

                for verse in chapter.findall("verse"):
                    verse_num = int(verse.get("number"))
                    verse_text = (verse.text or "").strip()

                    rows_to_insert.append((
                        book_name,
                        chapter_num,
                        verse_num,
                        verse_text,
                        version_code,
                        testament_es
                    ))

    print(f"Insertando {len(rows_to_insert)} versículos...")
    cursor.executemany(
        "INSERT INTO bible (book, chapter, verse, text, version, testament) VALUES (?, ?, ?, ?, ?, ?)",
        rows_to_insert
    )
    conn.commit()
    conn.close()
    print(f"¡Listo! Biblia '{version_code}' importada con {len(rows_to_insert)} versículos.")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Uso: python import_bible.py archivo.xml CODIGO_VERSION")
        print("Ejemplo: python import_bible.py SpanishRevisedRVR1960Bible.xml RVR60")
        sys.exit(1)

    xml_file = sys.argv[1]
    version = sys.argv[2]
    import_bible(xml_file, version)