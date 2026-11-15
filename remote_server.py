"""Local-only authenticated HTTP control server for Station Selah."""

import hashlib
import http.server
import ipaddress
import json
import secrets
import sqlite3
import threading
import time
import unicodedata
from http.cookies import SimpleCookie
from pathlib import Path
from urllib.parse import parse_qs, quote, urlsplit


COOKIE_NAME = "station_selah_session"
PAIRING_WINDOW_SECONDS = 600
SESSION_TTL_SECONDS = 12 * 60 * 60
MAX_PAIR_ATTEMPTS = 5
PAIR_ATTEMPT_WINDOW_SECONDS = 60
MAX_REQUESTS_PER_SESSION = 60
SESSION_RATE_WINDOW_SECONDS = 10
MAX_ACTIVE_SESSIONS = 12
MAX_BODY_BYTES = 2048
SUPPORTED_BIBLE_VERSIONS = {"RVR60", "NVI", "RV1569", "DHH"}
ALLOWED_COMMANDS = {
    "select_song", "select_verse", "select_script_item",
    "project_welcome", "project_farewell", "font_increase", "font_decrease",
    "previous", "next", "black", "restore", "clear",
}


def normalize_search(value):
    normalized = unicodedata.normalize("NFD", value or "")
    return "".join(
        char for char in normalized
        if unicodedata.category(char) != "Mn"
    ).lower()


class LANRemoteServer:
    """HTTP controller restricted to one private IPv4 interface and subnet."""

    def __init__(self, host, network, database_path, command_callback,
                 script_items_callback=None):
        address = ipaddress.IPv4Address(host)
        subnet = ipaddress.IPv4Network(network, strict=False)
        if (not address.is_private or address.is_loopback or address.is_link_local
                or address not in subnet):
            raise ValueError("La dirección debe pertenecer a una interfaz IPv4 privada.")

        self.host = str(address)
        self.network = subnet
        self.database_path = str(database_path)
        self.command_callback = command_callback
        self.script_items_callback = script_items_callback
        self.pairing_code = ""
        self.pairing_expires_at = 0.0
        self._sessions = {}
        self._pair_attempts = {}
        self._lock = threading.Lock()
        self._httpd = None
        self._thread = None
        self.port = None
        self.origin = None

    @property
    def is_running(self):
        return self._httpd is not None and self._thread is not None and self._thread.is_alive()

    @property
    def pairing_url(self):
        if not self.is_running:
            return ""
        return f"{self.origin}/#pair={quote(self.pairing_code)}"

    @property
    def pairing_seconds_remaining(self):
        return max(0, int(self.pairing_expires_at - time.monotonic()))

    def start(self):
        if self.is_running:
            return self.pairing_url

        self.rotate_pairing_code()
        handler = self._make_handler()
        httpd = http.server.ThreadingHTTPServer((self.host, 0), handler)
        httpd.daemon_threads = True
        httpd.block_on_close = False
        self._httpd = httpd
        self.port = httpd.server_address[1]
        self.origin = f"http://{self.host}:{self.port}"
        self._thread = threading.Thread(
            target=httpd.serve_forever,
            kwargs={"poll_interval": 0.2},
            name="StationSelahLANRemote",
            daemon=True,
        )
        self._thread.start()
        return self.pairing_url

    def rotate_pairing_code(self):
        with self._lock:
            self.pairing_code = secrets.token_urlsafe(24)
            self.pairing_expires_at = time.monotonic() + PAIRING_WINDOW_SECONDS
            self._pair_attempts.clear()
        return self.pairing_url

    def stop(self):
        httpd = self._httpd
        thread = self._thread
        self._httpd = None
        self._thread = None
        self.port = None
        self.origin = None
        if httpd is not None:
            httpd.shutdown()
            httpd.server_close()
        if thread is not None and thread.is_alive():
            thread.join(timeout=2)
        with self._lock:
            self._sessions.clear()
            self._pair_attempts.clear()
            self.pairing_code = ""
            self.pairing_expires_at = 0.0

    def _make_handler(self):
        controller = self

        class RequestHandler(http.server.BaseHTTPRequestHandler):
            server_version = "StationSelah"
            sys_version = ""
            protocol_version = "HTTP/1.1"

            def log_message(self, _format, *_args):
                return

            def send_error(self, code, message=None, explain=None):
                self._send_json(code, {"ok": False, "error": "Solicitud no válida."})

            def _send_json(self, status, payload, extra_headers=None):
                body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
                self.close_connection = True
                self.send_response(status)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Content-Length", str(len(body)))
                self.send_header("Connection", "close")
                self.send_header("Cache-Control", "no-store")
                self.send_header("X-Content-Type-Options", "nosniff")
                self.send_header("Referrer-Policy", "no-referrer")
                if extra_headers:
                    for key, value in extra_headers.items():
                        self.send_header(key, value)
                self.end_headers()
                self.wfile.write(body)

            def _send_page(self):
                page_path = Path(__file__).with_name("remote_control.html")
                try:
                    page = page_path.read_text(encoding="utf-8")
                except OSError:
                    self._send_json(503, {"ok": False, "error": "No se encontró la interfaz remota."})
                    return

                nonce = secrets.token_urlsafe(18)
                page = page.replace("__CSP_NONCE__", nonce).encode("utf-8")
                self.close_connection = True
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(page)))
                self.send_header("Connection", "close")
                self.send_header("Cache-Control", "no-store")
                self.send_header("X-Content-Type-Options", "nosniff")
                self.send_header("Referrer-Policy", "no-referrer")
                self.send_header(
                    "Content-Security-Policy",
                    "default-src 'none'; "
                    f"script-src 'nonce-{nonce}'; "
                    f"style-src 'nonce-{nonce}'; "
                    "connect-src 'self'; img-src 'self' data:; "
                    "base-uri 'none'; form-action 'self'; frame-ancestors 'none'",
                )
                self.end_headers()
                self.wfile.write(page)

            def _client_allowed(self):
                try:
                    client_ip = ipaddress.ip_address(self.client_address[0])
                except ValueError:
                    return False
                return client_ip.version == 4 and client_ip in controller.network

            def _host_allowed(self):
                expected = f"{controller.host}:{controller.port}".lower()
                return self.headers.get("Host", "").strip().lower() == expected

            def _origin_allowed(self):
                return self.headers.get("Origin", "").rstrip("/") == controller.origin

            def _read_json(self):
                if not self.headers.get("Content-Type", "").lower().startswith("application/json"):
                    return None
                try:
                    content_length = int(self.headers.get("Content-Length", "0"))
                except ValueError:
                    return None
                if content_length <= 0 or content_length > MAX_BODY_BYTES:
                    return None
                try:
                    payload = json.loads(self.rfile.read(content_length).decode("utf-8"))
                except (UnicodeDecodeError, json.JSONDecodeError):
                    return None
                return payload if isinstance(payload, dict) else None

            def _record_pair_attempt(self):
                now = time.monotonic()
                address = self.client_address[0]
                with controller._lock:
                    attempts = [
                        started for started in controller._pair_attempts.get(address, [])
                        if now - started < PAIR_ATTEMPT_WINDOW_SECONDS
                    ]
                    if len(attempts) >= MAX_PAIR_ATTEMPTS:
                        controller._pair_attempts[address] = attempts
                        return False
                    attempts.append(now)
                    controller._pair_attempts[address] = attempts
                    return True

            def _session_key(self):
                cookie = SimpleCookie()
                try:
                    cookie.load(self.headers.get("Cookie", ""))
                except Exception:
                    return None
                morsel = cookie.get(COOKIE_NAME)
                if morsel is None or not morsel.value:
                    return None
                digest = hashlib.sha256(morsel.value.encode("utf-8")).hexdigest()
                now = time.monotonic()
                with controller._lock:
                    session = controller._sessions.get(digest)
                    if not session or session["expires"] <= now:
                        controller._sessions.pop(digest, None)
                        return None
                    return digest

            def _consume_session_request(self, session_key):
                now = time.monotonic()
                with controller._lock:
                    session = controller._sessions.get(session_key)
                    if not session or session["expires"] <= now:
                        controller._sessions.pop(session_key, None)
                        return False
                    recent = [
                        timestamp for timestamp in session["requests"]
                        if now - timestamp < SESSION_RATE_WINDOW_SECONDS
                    ]
                    if len(recent) >= MAX_REQUESTS_PER_SESSION:
                        session["requests"] = recent
                        return False
                    recent.append(now)
                    session["requests"] = recent
                    return True

            def _authorized_session(self):
                session_key = self._session_key()
                if session_key is None:
                    self._send_json(401, {"ok": False, "error": "Escanea nuevamente el QR para conectar."})
                    return None
                if not self._consume_session_request(session_key):
                    self._send_json(429, {"ok": False, "error": "Demasiadas solicitudes. Espera un momento."})
                    return None
                return session_key

            def do_GET(self):
                if not self._client_allowed() or not self._host_allowed():
                    self._send_json(403, {"ok": False, "error": "Acceso solo disponible en esta red local."})
                    return

                parsed = urlsplit(self.path)
                if parsed.path == "/":
                    self._send_page()
                    return
                if parsed.path == "/api/state":
                    if self._authorized_session() is None:
                        return
                    self._send_json(200, {"ok": True, "connected": True})
                    return
                if parsed.path == "/api/songs":
                    if self._authorized_session() is None:
                        return
                    query = parse_qs(parsed.query).get("q", [""])[0].strip()
                    if len(query) > 80:
                        self._send_json(400, {"ok": False, "error": "La búsqueda es demasiado larga."})
                        return
                    normalized = normalize_search(query)
                    if len(normalized) < 2:
                        self._send_json(200, {"ok": True, "songs": []})
                        return
                    pattern = f"%{normalized}%"
                    try:
                        connection = sqlite3.connect(controller.database_path, timeout=2)
                        connection.create_function("normalize_search", 1, normalize_search)
                        rows = connection.execute(
                            "SELECT id, title, category FROM songs "
                            "WHERE normalize_search(title) LIKE ? "
                            "OR normalize_search(lyrics) LIKE ? "
                            "ORDER BY title LIMIT 30",
                            (pattern, pattern),
                        ).fetchall()
                        connection.close()
                    except sqlite3.Error:
                        self._send_json(503, {"ok": False, "error": "No se pudo consultar la biblioteca."})
                        return
                    songs = [
                        {"id": row[0], "title": row[1], "category": row[2] or ""}
                        for row in rows
                    ]
                    self._send_json(200, {"ok": True, "songs": songs})
                    return
                if parsed.path == "/api/guion":
                    if self._authorized_session() is None:
                        return
                    if not controller.script_items_callback:
                        self._send_json(503, {"ok": False, "error": "El Guion remoto no está disponible."})
                        return
                    try:
                        items = controller.script_items_callback()
                    except Exception:
                        self._send_json(503, {"ok": False, "error": "No se pudo consultar el Guion."})
                        return
                    self._send_json(200, {"ok": True, "items": items[:200]})
                    return
                if parsed.path == "/api/bible/books":
                    if self._authorized_session() is None:
                        return
                    version = parse_qs(parsed.query).get("version", ["RVR60"])[0]
                    if version not in SUPPORTED_BIBLE_VERSIONS:
                        self._send_json(400, {"ok": False, "error": "Versión bíblica no permitida."})
                        return
                    try:
                        connection = sqlite3.connect(controller.database_path, timeout=2)
                        rows = connection.execute(
                            "SELECT book FROM bible WHERE version = ? "
                            "GROUP BY book ORDER BY MIN(id)", (version,)
                        ).fetchall()
                        connection.close()
                    except sqlite3.Error:
                        self._send_json(503, {"ok": False, "error": "No se pudo consultar la lista de libros."})
                        return
                    self._send_json(200, {"ok": True, "books": [row[0] for row in rows]})
                    return
                if parsed.path == "/api/bible/chapters":
                    if self._authorized_session() is None:
                        return
                    params = parse_qs(parsed.query)
                    version = params.get("version", ["RVR60"])[0]
                    book = params.get("book", [""])[0]
                    if version not in SUPPORTED_BIBLE_VERSIONS or not book or len(book) > 80:
                        self._send_json(400, {"ok": False, "error": "Libro o versión no válidos."})
                        return
                    try:
                        connection = sqlite3.connect(controller.database_path, timeout=2)
                        rows = connection.execute(
                            "SELECT DISTINCT chapter FROM bible "
                            "WHERE version = ? AND book = ? ORDER BY chapter",
                            (version, book),
                        ).fetchall()
                        connection.close()
                    except sqlite3.Error:
                        self._send_json(503, {"ok": False, "error": "No se pudieron consultar los capítulos."})
                        return
                    self._send_json(200, {"ok": True, "chapters": [row[0] for row in rows]})
                    return
                if parsed.path == "/api/bible/verses":
                    if self._authorized_session() is None:
                        return
                    params = parse_qs(parsed.query)
                    version = params.get("version", ["RVR60"])[0]
                    book = params.get("book", [""])[0]
                    raw_chapter = params.get("chapter", [""])[0]
                    if (version not in SUPPORTED_BIBLE_VERSIONS or not book
                            or len(book) > 80 or not raw_chapter.isdigit()):
                        self._send_json(400, {"ok": False, "error": "Capítulo, libro o versión no válidos."})
                        return
                    chapter = int(raw_chapter)
                    if chapter < 1:
                        self._send_json(400, {"ok": False, "error": "Capítulo no válido."})
                        return
                    try:
                        connection = sqlite3.connect(controller.database_path, timeout=2)
                        rows = connection.execute(
                            "SELECT verse, text FROM bible "
                            "WHERE version = ? AND book = ? AND chapter = ? "
                            "ORDER BY verse LIMIT 200",
                            (version, book, chapter),
                        ).fetchall()
                        connection.close()
                    except sqlite3.Error:
                        self._send_json(503, {"ok": False, "error": "No se pudieron consultar los versículos."})
                        return
                    verses = [{"verse": row[0], "text": row[1]} for row in rows]
                    self._send_json(200, {"ok": True, "verses": verses})
                    return
                if parsed.path == "/api/bible":
                    if self._authorized_session() is None:
                        return
                    params = parse_qs(parsed.query)
                    query = params.get("q", [""])[0].strip()
                    version = params.get("version", ["RVR60"])[0]
                    if len(query) > 80:
                        self._send_json(400, {"ok": False, "error": "La búsqueda es demasiado larga."})
                        return
                    if version not in SUPPORTED_BIBLE_VERSIONS:
                        self._send_json(400, {"ok": False, "error": "Versión bíblica no permitida."})
                        return
                    normalized = normalize_search(query)
                    if len(normalized) < 2:
                        self._send_json(200, {"ok": True, "verses": []})
                        return
                    pattern = f"%{normalized}%"
                    try:
                        connection = sqlite3.connect(controller.database_path, timeout=2)
                        connection.create_function("normalize_search", 1, normalize_search)
                        rows = connection.execute(
                            "SELECT book, chapter, verse, text, version FROM bible "
                            "WHERE version = ? AND ("
                            "normalize_search(book || ' ' || chapter || ':' || verse) LIKE ? "
                            "OR normalize_search(text) LIKE ?) "
                            "ORDER BY book, chapter, verse LIMIT 40",
                            (version, pattern, pattern),
                        ).fetchall()
                        connection.close()
                    except sqlite3.Error:
                        self._send_json(503, {"ok": False, "error": "No se pudo consultar la Biblia."})
                        return
                    verses = [
                        {
                            "book": row[0],
                            "chapter": row[1],
                            "verse": row[2],
                            "text": row[3],
                            "version": row[4],
                        }
                        for row in rows
                    ]
                    self._send_json(200, {"ok": True, "verses": verses})
                    return
                self._send_json(404, {"ok": False, "error": "Recurso no disponible."})

            def do_POST(self):
                if not self._client_allowed() or not self._host_allowed():
                    self._send_json(403, {"ok": False, "error": "Acceso solo disponible en esta red local."})
                    return
                if not self._origin_allowed():
                    self._send_json(403, {"ok": False, "error": "Origen no autorizado."})
                    return

                parsed = urlsplit(self.path)
                payload = self._read_json()
                if payload is None:
                    self._send_json(400, {"ok": False, "error": "Solicitud JSON no válida."})
                    return

                if parsed.path == "/api/pair":
                    if not self._record_pair_attempt():
                        self._send_json(429, {"ok": False, "error": "Demasiados intentos de emparejamiento."})
                        return
                    supplied_code = payload.get("code")
                    if (not isinstance(supplied_code, str)
                            or not supplied_code.isascii()
                            or len(supplied_code) > 128
                            or not controller.pairing_code):
                        self._send_json(401, {"ok": False, "error": "Código QR no válido o expirado."})
                        return
                    if (controller.pairing_seconds_remaining <= 0
                            or not secrets.compare_digest(supplied_code, controller.pairing_code)):
                        self._send_json(401, {"ok": False, "error": "Código QR no válido o expirado."})
                        return

                    with controller._lock:
                        active_sessions = sum(
                            session["expires"] > time.monotonic()
                            for session in controller._sessions.values()
                        )
                        if active_sessions >= MAX_ACTIVE_SESSIONS:
                            self._send_json(429, {"ok": False, "error": "Se alcanzó el límite de dispositivos emparejados."})
                            return
                        session_token = secrets.token_urlsafe(32)
                        session_key = hashlib.sha256(session_token.encode("utf-8")).hexdigest()
                        controller._sessions[session_key] = {
                            "expires": time.monotonic() + SESSION_TTL_SECONDS,
                            "requests": [],
                        }

                    cookie = (
                        f"{COOKIE_NAME}={session_token}; Path=/; HttpOnly; "
                        f"SameSite=Strict; Max-Age={SESSION_TTL_SECONDS}"
                    )
                    self._send_json(
                        200,
                        {"ok": True, "message": "Dispositivo conectado."},
                        {"Set-Cookie": cookie},
                    )
                    return

                if parsed.path == "/api/command":
                    if self._authorized_session() is None:
                        return
                    action = payload.get("action")
                    if not isinstance(action, str) or action not in ALLOWED_COMMANDS:
                        self._send_json(400, {"ok": False, "error": "Comando no permitido."})
                        return
                    allowed_fields = {
                        "select_song": {"action", "song_id"},
                        "select_verse": {
                            "action", "book", "chapter", "verse", "version"
                        },
                        "select_script_item": {"action", "index"},
                    }.get(action, {"action"})
                    if set(payload) - allowed_fields:
                        self._send_json(400, {"ok": False, "error": "Parámetros no permitidos."})
                        return
                    command = {"action": action}
                    if action == "select_song":
                        song_id = payload.get("song_id")
                        if type(song_id) is not int or song_id <= 0:
                            self._send_json(400, {"ok": False, "error": "Canto no válido."})
                            return
                        connection = sqlite3.connect(controller.database_path, timeout=2)
                        exists = connection.execute(
                            "SELECT 1 FROM songs WHERE id = ?", (song_id,)
                        ).fetchone()
                        connection.close()
                        if not exists:
                            self._send_json(404, {"ok": False, "error": "Canto no encontrado."})
                            return
                        command["song_id"] = song_id
                    elif action == "select_verse":
                        book = payload.get("book")
                        chapter = payload.get("chapter")
                        verse = payload.get("verse")
                        version = payload.get("version")
                        if (not isinstance(book, str) or not book.strip()
                                or len(book) > 80
                                or type(chapter) is not int or chapter < 1
                                or type(verse) is not int or verse < 1
                                or version not in SUPPORTED_BIBLE_VERSIONS):
                            self._send_json(400, {"ok": False, "error": "Referencia bíblica no válida."})
                            return
                        connection = sqlite3.connect(controller.database_path, timeout=2)
                        exists = connection.execute(
                            "SELECT 1 FROM bible WHERE book = ? AND chapter = ? "
                            "AND verse = ? AND version = ?",
                            (book, chapter, verse, version),
                        ).fetchone()
                        connection.close()
                        if not exists:
                            self._send_json(404, {"ok": False, "error": "Versículo no encontrado."})
                            return
                        command.update({
                            "book": book,
                            "chapter": chapter,
                            "verse": verse,
                            "version": version,
                        })
                    elif action == "select_script_item":
                        item_index = payload.get("index")
                        if type(item_index) is not int or item_index < 0:
                            self._send_json(400, {"ok": False, "error": "Elemento de Guion no válido."})
                            return
                        if not controller.script_items_callback:
                            self._send_json(503, {"ok": False, "error": "El Guion remoto no está disponible."})
                            return
                        try:
                            items = controller.script_items_callback()
                        except Exception:
                            self._send_json(503, {"ok": False, "error": "No se pudo consultar el Guion."})
                            return
                        if item_index >= len(items):
                            self._send_json(404, {"ok": False, "error": "Elemento de Guion no encontrado."})
                            return
                        command["index"] = item_index
                    try:
                        controller.command_callback(command)
                    except Exception:
                        self._send_json(503, {"ok": False, "error": "No se pudo enviar el comando a la aplicación."})
                        return
                    self._send_json(202, {"ok": True, "message": "Comando recibido."})
                    return

                self._send_json(404, {"ok": False, "error": "Recurso no disponible."})

        return RequestHandler
