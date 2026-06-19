import os
import sqlite3

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "iglesia.db")

def init_db():
    db_exists = os.path.exists(DB_PATH)

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    # Tabla de Cantos con columna de categoría (para el subtítulo de "Alabanza" o "Adoración")
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS songs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT NOT NULL,
        lyrics TEXT NOT NULL,
        category TEXT DEFAULT 'Alabanza',
        author TEXT,
        key TEXT
    )
    """)

    # Tabla de la Biblia
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS bible (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        book TEXT NOT NULL,
        chapter INTEGER NOT NULL,
        verse INTEGER NOT NULL,
        text TEXT NOT NULL,
        version TEXT DEFAULT 'RV1960'
    )
    """)

    # Tabla de imágenes
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS images (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        file_path TEXT NOT NULL
    )
    """)

    # Tabla de PPTX
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS pptx_files (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        file_path TEXT NOT NULL
    )
    """)

    conn.commit()
    
    if not db_exists:
        insert_sample_songs(conn)
        insert_sample_bible(conn)
        
    conn.close()

def insert_sample_songs(conn):
    cursor = conn.cursor()
    
    songs_data = [
        (
            "Sublime Gracia",
            "Sublime gracia del Señor,\nQue a un infeliz salvó;\nFui ciego mas hoy veo yo,\nPerdido y Él me halló.\n\n[Coro]\nSublime gracia\nCuán dulce es\nQue salva al pecador\nFui ciego mas hoy veo\nPerdido y Él me halló\n\n[Estrofa 2]\nEn los peligros o aflicción\nQue yo he tenido aquí\nSu gracia siempre me libró\nY me guiará feliz.",
            "Alabanza",
            "John Newton",
            "Mi"
        ),
        (
            "Cuán Grande Es Él",
            "Señor, mi Dios, al contemplar los cielos,\nEl firmamento y las estrellas mil;\nAl oír tu voz en los potentes truenos\nY ver brillar al sol en su cenit.\n\n[Coro]\nMi corazón entona la canción:\n¡Cuán grande es Él! ¡Cuán grande es Él!\nMi corazón entona la canción:\n¡Cuán grande es Él! ¡Cuán grande es Él!\n\n[Estrofa 2]\nAl recorrer los montes y los valles\nY ver las bellas flores al pasar;\nAl escuchar el canto de las aves\nY el murmurar del claro manantial.",
            "Adoración",
            "Stuart K. Hine",
            "La"
        ),
        (
            "Digno es el Cordero",
            "Gracias por la cruz, oh Dios\nEl precio que pagaste por mí\nLlevando mi pecado allí\nSublime gracia\n\n[Coro]\nDigno es el Cordero celestial\nCoronado en majestad\nReinas con poder\n¡Digno eres, Señor!\n\n[Estrofa 2]\nGracias por tu amor, Señor\nTus manos heridas vi\nMe lavaste en tu sangre, Señor\nY hoy sé que perdonado fui",
            "Adoración",
            "Darlene Zschech",
            "Sol"
        ),
        (
            "A Dios sea la Gloria",
            "A Dios sea la gloria,\nPor su gran amor,\nQue dio a su Hijo,\nEl Salvador.\n\n[Coro]\nExaltad al Señor,\nEscuchad su voz;\nExaltad al Señor,\nAlabad a Dios;\nVenid al Padre\nPor Cristo el Hijo,\nY dadle la gloria\nPor su gran amor.\n\n[Estrofa 2]\nQuien abrió las puertas\nDe la salvación,\nY limpio de manchas\nAl pecador.",
            "Alabanza",
            "Fanny Crosby",
            "Sol"
        )
    ]
    
    cursor.executemany("INSERT INTO songs (title, lyrics, category, author, key) VALUES (?, ?, ?, ?, ?)", songs_data)
    conn.commit()

def insert_sample_bible(conn):
    cursor = conn.cursor()
    
    bible_data = [
        ("Génesis", 1, 1, "En el principio creó Dios los cielos y la tierra.", "RV1960"),
        ("Génesis", 1, 2, "Y la tierra estaba desordenada y vacía, y las tinieblas estaban sobre la faz del abismo, y el Espíritu de Dios se movía sobre la faz de las aguas.", "RV1960"),
        ("Génesis", 1, 3, "Y dijo Dios: Sea la luz; y fue la luz.", "RV1960"),
        ("Génesis", 1, 4, "Y vio Dios que la luz era buena; y separó Dios la luz de las tinieblas.", "RV1960"),
        ("Génesis", 1, 5, "Y llamó Dios a la luz Día, y a las tinieblas llamó Noche. Y fue la tarde y la mañana un día.", "RV1960"),
        
        ("Salmos", 23, 1, "Jehová es mi pastor; nada me faltará.", "RV1960"),
        ("Salmos", 23, 2, "En lugares de delicados pastos me hará descansar; Junto a aguas de reposo me pastoreará.", "RV1960"),
        ("Salmos", 23, 3, "Confortará mi alma; Me guiará por sendas de justicia por amor de su nombre.", "RV1960"),
        
        ("Juan", 3, 16, "Porque de tal manera amó Dios al mundo, que ha dado a su Hijo unigénito, para que todo aquel que en él cree, no se pierda, mas tenga vida eterna.", "RV1960")
    ]
    
    cursor.executemany("INSERT INTO bible (book, chapter, verse, text, version) VALUES (?, ?, ?, ?, ?)", bible_data)
    conn.commit()

if __name__ == "__main__":
    init_db()
