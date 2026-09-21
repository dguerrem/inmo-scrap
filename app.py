"""
Scraper táctico de inmobiliarias en Google Maps (Valencia + área metropolitana).

Stack:
    - Backend:  FastAPI (este único archivo).
    - Frontend: templates/index.html  (Tailwind CDN + Vanilla JS, sin frameworks).
    - Scraping: Playwright (Chromium headless) con URLs de búsqueda directas
                por población/zona. NO se arrastra el mapa ni se usan coordenadas.
    - Datos:    data_temp.json (SIN base de datos) + Excel generado en memoria.

Puesta en marcha:
    python -m venv .venv && source .venv/bin/activate
    pip install -r requirements.txt
    python -m playwright install chromium
    uvicorn app:app --host 127.0.0.1 --port 8000
"""
from __future__ import annotations

import io
import itertools
import json
import logging
import math
import os
import random
import re
import threading
import time
import unicodedata
from collections import deque
from datetime import datetime
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Any
from urllib.parse import quote_plus

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse, StreamingResponse
from openpyxl import Workbook
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from pydantic import BaseModel, Field
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

# --------------------------------------------------------------------------- #
# Configuración
# --------------------------------------------------------------------------- #
BASE_DIR = Path(__file__).resolve().parent
TEMPLATES_DIR = BASE_DIR / "templates"
DATA_FILE = BASE_DIR / "data_temp.json"
DESCARTES_FILE = BASE_DIR / "descartados_temp.json"
ZONAS_FILE = BASE_DIR / "zonas_valencia.json"

PLANTILLA_URL = "https://www.google.com/maps/search/{termino}?hl=es"
USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
)

# Selectores (Google Maps cambia clases a menudo; estos son los estables por rol/atributo)
SEL_FEED = 'div[role="feed"]'                    # panel lateral de resultados
SEL_ENLACE = 'div[role="feed"] a.hfpxzc'         # tarjeta -> enlace a la ficha
SEL_FIN_LISTA = "span.HlvSq"                     # "Has llegado al final de la lista"
SEL_DIRECCION = 'button[data-item-id="address"]'
SEL_TELEFONO = 'button[data-item-id^="phone:tel:"]'
SEL_WEB = 'a[data-item-id="authority"]'

SEL_CONSENTIMIENTO = (
    'button:has-text("Aceptar todo")',
    'button:has-text("Rechazar todo")',
    'button:has-text("Accept all")',
    'button:has-text("Reject all")',
)

# Columnas del Excel / del JSON (clave interna -> cabecera visible)
COLUMNAS: tuple[tuple[str, str], ...] = (
    ("nombre", "Nombre"),
    ("direccion", "Dirección"),
    ("barrio", "Barrio/Población"),
    ("distancia_km", "Distancia (km)"),
    ("telefono", "Teléfono"),
    ("web", "Web"),
)

MAX_CICLOS_SCROLL = 40       # ciclos máximos de scroll dentro del feed, por zona
ESPERA_SCROLL_MS = 1200      # pausa tras cada scroll (carga dinámica)
CICLOS_SIN_NOVEDAD = 3       # se corta cuando N ciclos no aportan resultados nuevos
RECICLAR_PAGINA_CADA = 25    # recrea la pestaña cada N zonas (higiene de memoria)
ESPERA_DIRECCION_MS = 4000   # espera máxima a que la ficha muestre la dirección

# Extrae de una vez todos los enlaces del feed (más rápido que N llamadas al DOM)
JS_ENLACES = (
    "() => Array.from(document.querySelectorAll('div[role=\"feed\"] a.hfpxzc'))"
    ".map(a => ({ href: a.href, nombre: (a.getAttribute('aria-label') || '').trim() }))"
)

# --------------------------------------------------------------------------- #
# Filtrado geográfico
#
# Google Maps devuelve, para la misma consulta, conjuntos de resultados muy
# distintos: buscando "inmobiliarias en Carpesa" se han llegado a recibir 117
# resultados con negocios a 709 km. No respeta la zona del texto ni el viewport
# de la URL. Lo único fiable es que cada enlace del feed lleva las coordenadas
# exactas del negocio, así que filtramos por distancia al centro real de la zona.
# --------------------------------------------------------------------------- #
COORD_ENLACE = re.compile(r"!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)")
_descartes: list[dict[str, Any]] = []      # resultados fuera de radio, para la hoja "Descartados"
_lock_descartes = threading.Lock()
_zonas_de_esta_ejecucion: set[str] = set()


def _cargar_zonas() -> dict[str, dict]:
    """Centros y radios por zona. Si falta el fichero, se trabaja sin filtro."""
    if not ZONAS_FILE.exists():
        return {}
    try:
        datos = json.loads(ZONAS_FILE.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        # Ojo: aquí todavía no existe el logger, así que no se puede usar `log`.
        print(f"[!] {ZONAS_FILE.name} está corrupto: se trabaja sin filtro geográfico")
        return {}
    return datos if isinstance(datos, dict) else {}


ZONAS_GEO = _cargar_zonas()


def _coords_de_enlace(href: str) -> tuple[float, float] | None:
    """Las coordenadas del negocio vienen dentro del propio href del feed."""
    coincidencia = COORD_ENLACE.search(href or "")
    return (float(coincidencia.group(1)), float(coincidencia.group(2))) if coincidencia else None


def km_entre(a: tuple[float, float], b: tuple[float, float]) -> float:
    """Distancia en kilómetros entre dos puntos (Haversine)."""
    radio = 6371.0
    lat1, lat2 = math.radians(a[0]), math.radians(b[0])
    dlat = lat2 - lat1
    dlng = math.radians(b[1] - a[1])
    h = math.sin(dlat / 2) ** 2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlng / 2) ** 2
    return 2 * radio * math.asin(math.sqrt(h))


def _distancia_a_zona(coords: tuple[float, float], zona: str) -> float | None:
    """Distancia al centro de la zona, o None si esa zona no tiene centro conocido."""
    centro = ZONAS_GEO.get(zona)
    if not centro:
        return None
    return km_entre(coords, (centro["lat"], centro["lng"]))


def _dentro_de_zona(coords: tuple[float, float] | None, zona: str) -> bool:
    """¿Está el resultado dentro del radio de la zona? Sin datos, se acepta."""
    centro = ZONAS_GEO.get(zona)
    if not centro:
        return True          # zona sin geolocalizar: no filtramos
    if coords is None:
        return True          # enlace sin coordenadas: no podemos juzgar, se acepta
    return km_entre(coords, (centro["lat"], centro["lng"])) <= centro["radio_km"]


def _registrar_descarte(nombre: str, zona: str, distancia: float | None, motivo: str) -> None:
    with _lock_descartes:
        _descartes.append(
            {
                "nombre": nombre,
                "barrio": zona,
                "distancia_km": round(distancia, 1) if distancia is not None else "",
                "motivo": motivo,
            }
        )


def _zona_mas_cercana(
    coords: tuple[float, float] | None, barrio_original: str
) -> str:
    """
    Reasigna cada inmobiliaria a la zona cuyo centro tiene más cerca.

    Sin esto, al rastrear "Valencia" con sus 8 km de radio se quedaría con casi
    todo, y hojas como Ruzafa o El Carmen saldrían casi vacías.
    """
    if coords is None:
        return barrio_original
    candidatas = [z for z in ZONAS_GEO if z in _zonas_de_esta_ejecucion] or [barrio_original]
    mejor, mejor_distancia = barrio_original, None
    for zona in candidatas:
        distancia = _distancia_a_zona(coords, zona)
        if distancia is not None and (mejor_distancia is None or distancia < mejor_distancia):
            mejor, mejor_distancia = zona, distancia
    return mejor


_zonas_de_esta_ejecucion: set[str] = set()

# --------------------------------------------------------------------------- #
# Log en directo: stdout + archivo rotativo + buffer en memoria para el frontend
# --------------------------------------------------------------------------- #
LOG_FILE = BASE_DIR / "scraping.log"
MAX_LINEAS_LOG = 20_000      # líneas retenidas en memoria para la consola del frontend

_registro: deque[dict[str, Any]] = deque(maxlen=MAX_LINEAS_LOG)
_secuencia = itertools.count(1)   # numeración monótona: nunca se reinicia al limpiar
_lock_log = threading.Lock()


class _HandlerMemoria(logging.Handler):
    """Guarda cada línea en un buffer circular que el frontend consume por HTTP."""

    def emit(self, record: logging.LogRecord) -> None:
        try:
            with _lock_log:
                _registro.append(
                    {
                        "n": next(_secuencia),
                        "hora": datetime.fromtimestamp(record.created).strftime("%H:%M:%S"),
                        "nivel": record.levelname.lower(),
                        "texto": record.getMessage(),
                    }
                )
        except Exception:  # noqa: BLE001 - el log jamás debe tumbar el scraping
            pass


def _configurar_log() -> logging.Logger:
    """Logger único con tres destinos: terminal, archivo rotativo y memoria."""
    logger = logging.getLogger("inmo")
    logger.setLevel(logging.INFO)
    logger.propagate = False

    if not logger.handlers:
        consola = logging.StreamHandler()
        consola.setFormatter(logging.Formatter("%(asctime)s │ %(message)s", "%H:%M:%S"))

        archivo = RotatingFileHandler(
            LOG_FILE, maxBytes=5_000_000, backupCount=2, encoding="utf-8"
        )
        archivo.setFormatter(
            logging.Formatter(
                "%(asctime)s │ %(levelname)-7s │ %(message)s", "%Y-%m-%d %H:%M:%S"
            )
        )

        for handler in (consola, archivo, _HandlerMemoria()):
            handler.setLevel(logging.INFO)
            logger.addHandler(handler)

    return logger


log = _configurar_log()

# --------------------------------------------------------------------------- #
# Progreso (barra del frontend). Se actualiza mientras el scraping avanza.
# --------------------------------------------------------------------------- #
_progreso: dict[str, Any] = {
    "activo": False,
    "total_zonas": 0,
    "zona_actual": 0,
    "zona_nombre": "",
    "fichas_hechas": 0,
    "fichas_totales": 0,
    "acumuladas": 0,
    "segundos": 0.0,
}
_lock_progreso = threading.Lock()
_inicio_ejecucion = 0.0


def _reiniciar_progreso(total_zonas: int) -> None:
    global _inicio_ejecucion
    with _lock_progreso:
        _progreso.update(
            activo=True,
            total_zonas=total_zonas,
            zona_actual=0,
            zona_nombre="",
            fichas_hechas=0,
            fichas_totales=0,
            acumuladas=0,
            segundos=0.0,
        )
        _inicio_ejecucion = time.perf_counter()


def _actualizar_progreso(**campos: Any) -> None:
    with _lock_progreso:
        _progreso.update(campos)


def _cerrar_progreso() -> None:
    """Marca la ejecución como terminada incluso si algo falló a mitad."""
    with _lock_progreso:
        _progreso["activo"] = False
        _progreso["zona_nombre"] = ""
        _progreso["segundos"] = round(time.perf_counter() - _inicio_ejecucion, 1)


def _instantanea_progreso() -> dict[str, Any]:
    """Progreso + porcentaje y estimación de tiempo restante."""
    with _lock_progreso:
        datos = dict(_progreso)

    if datos["activo"]:
        datos["segundos"] = round(time.perf_counter() - _inicio_ejecucion, 1)

    total = datos["total_zonas"]
    if not total:
        datos.update(porcentaje=0.0, zonas_restantes=0, eta_segundos=None)
        return datos

    # Media zona de crédito por cada zona ya terminada, más la fracción de la actual.
    completadas = max(datos["zona_actual"] - 1, 0)
    fraccion = 0.0
    if datos["fichas_totales"]:
        fraccion = min(datos["fichas_hechas"] / datos["fichas_totales"], 1.0)

    datos["porcentaje"] = round(min((completadas + fraccion) / total * 100, 100), 1)
    # Al terminar, no queda nada pendiente aunque zona_actual apunte a la última.
    datos["zonas_restantes"] = 0 if not datos["activo"] else max(total - completadas, 0)

    # ETA extrapolando el ritmo real por zona (más honesto que usar el porcentaje,
    # porque las zonas tienen tamaños muy dispares).
    if completadas and datos["activo"]:
        datos["eta_segundos"] = round(datos["segundos"] / completadas * datos["zonas_restantes"])
    else:
        datos["eta_segundos"] = None
    return datos


app = FastAPI(title="Scraper Inmobiliarias · Google Maps")
_candado = threading.Lock()   # evita que dos scrapings se pisen


class ScrapeRequest(BaseModel):
    zonas: list[str] = Field(default_factory=list)


# --------------------------------------------------------------------------- #
# Endpoints
# --------------------------------------------------------------------------- #
@app.get("/")
def index() -> FileResponse:
    plantilla = TEMPLATES_DIR / "index.html"
    if not plantilla.exists():
        raise HTTPException(status_code=500, detail="Falta templates/index.html")
    return FileResponse(plantilla, media_type="text/html")


@app.post("/api/scrape")
def scrape(peticion: ScrapeRequest) -> dict[str, Any]:
    """Pasos 2-4: recorre las zonas, extrae, deduplica y guarda data_temp.json."""
    zonas = _limpiar_zonas(peticion.zonas)
    if not zonas:
        raise HTTPException(status_code=400, detail="No se ha recibido ninguna zona.")

    if not _candado.acquire(blocking=False):
        raise HTTPException(status_code=409, detail="Ya hay un scraping en curso.")

    try:
        inicio = time.time()
        with _lock_log:
            _registro.clear()
        with _lock_descartes:
            _descartes.clear()
        _zonas_de_esta_ejecucion.clear()
        _zonas_de_esta_ejecucion.update(zonas)
        _reiniciar_progreso(len(zonas))
        log.info("═══ Nueva recopilación · %d zonas ═══", len(zonas))

        _geolocalizar_desconocidas(zonas)
        sin_geo = [z for z in zonas if z not in ZONAS_GEO]
        if sin_geo:
            log.warning(
                "Sin datos geográficos para %d zona(s): se aceptarán todos sus resultados. "
                "Añádelas con: .venv/bin/python generar_zonas.py %s",
                len(sin_geo),
                " ".join(f'"{z}"' for z in sin_geo[:6]),
            )
        else:
            radios = [ZONAS_GEO[z]["radio_km"] for z in zonas if z in ZONAS_GEO]
            log.info(
                "Filtro geográfico activo · radio %s km según el tamaño real de cada zona",
                (
                    f"{min(radios)}-{max(radios)}"
                    if radios and min(radios) != max(radios)
                    else (f"{radios[0]}" if radios else "?")
                ),
            )

        registros, errores = _scrape_zonas(zonas)
        with _lock_descartes:
            descartados = list(_descartes)
        unicos = reasignar_zonas(deduplicar(registros))
        _guardar_datos(_sin_coords_auxiliares(unicos))
        _guardar_descartes(descartados)

        log.info(
            "═══ Terminado · %d únicas de %d brutos · %d duplicadas · %d descartadas "
            "por distancia · %.1f min ═══",
            len(unicos),
            len(registros),
            len(registros) - len(unicos),
            len(descartados),
            (time.time() - inicio) / 60,
        )
        return {
            "ok": True,
            "zonas": len(zonas),
            "resultados": len(registros),
            "unicos": len(unicos),
            "duplicados": len(registros) - len(unicos),
            "descartados": len(descartados),
            "sin_geo": sin_geo,
            "errores": errores,
            "segundos": round(time.time() - inicio, 1),
        }
    except HTTPException:
        raise
    except Exception as exc:  # noqa: BLE001 - se devuelve el motivo en crudo al frontend
        mensaje = f"{type(exc).__name__}: {exc}"
        if "Executable doesn't exist" in mensaje or "playwright install" in mensaje:
            mensaje = "Chromium no está instalado. Ejecuta: python -m playwright install chromium"
        log.error("Recopilación abortada · %s", mensaje)
        raise HTTPException(status_code=500, detail=mensaje) from exc
    finally:
        _cerrar_progreso()
        _candado.release()


@app.get("/api/log")
def api_log(desde: int = 0) -> dict[str, Any]:
    """Log en directo: solo las líneas nuevas desde el índice `desde`."""
    with _lock_log:
        entradas = [entrada for entrada in _registro if entrada["n"] > desde]
        ultimo = _registro[-1]["n"] if _registro else 0
    return {"entradas": entradas, "ultimo": ultimo}


@app.get("/api/log/descargar")
def descargar_log() -> FileResponse:
    """Log completo en disco, incluidas ejecuciones anteriores."""
    if not LOG_FILE.exists():
        raise HTTPException(status_code=404, detail="Todavía no hay log.")
    return FileResponse(LOG_FILE, media_type="text/plain", filename="scraping.log")


@app.get("/api/estado")
def api_estado() -> dict[str, Any]:
    """Si hay una recopilación en curso, el frontend se engancha a su log."""
    return {"en_curso": _candado.locked()}


@app.get("/api/progreso")
def api_progreso() -> dict[str, Any]:
    """Progreso de la recopilación en curso, para la barra del frontend."""
    return _instantanea_progreso()


@app.get("/api/data")
def data() -> dict[str, Any]:
    """Extra opcional: datos guardados para la vista previa del frontend."""
    items = _leer_datos()
    return {"total": len(items), "items": items}


@app.get("/api/download")
def download() -> StreamingResponse:
    """Paso 6: lee el JSON y devuelve un .xlsx generado en memoria."""
    items = _leer_datos()
    if not items:
        raise HTTPException(
            status_code=404,
            detail="Todavía no hay datos. Ejecuta primero «Recopilar datos».",
        )
    buffer = _excel_en_memoria(items, _leer_descartes())
    nombre = f"inmobiliarias_valencia_{datetime.now():%Y%m%d_%H%M}.xlsx"
    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{nombre}"'},
    )


# --------------------------------------------------------------------------- #
# Scraping (Playwright, sincrónico -> FastAPI lo ejecuta en un hilo aparte)
# --------------------------------------------------------------------------- #
def _guardar_zonas_geo() -> None:
    try:
        ZONAS_FILE.write_text(
            json.dumps(ZONAS_GEO, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except OSError as exc:
        log.warning("No se pudo guardar %s: %s", ZONAS_FILE.name, exc)


def _geolocalizar_desconocidas(zonas: list[str]) -> None:
    """
    Geolocaliza al vuelo las zonas que no estén en `zonas_valencia.json`.

    Sin esto, escribir una zona nueva (por ejemplo "El Palmar") dejaría la
    búsqueda sin filtro y volverían a colarse resultados de toda España. Se pide
    una sola vez por zona y queda cacheada para las siguientes ejecuciones.
    """
    desconocidas = [z for z in zonas if z not in ZONAS_GEO]
    if not desconocidas:
        return

    log.info("Geolocalizando %d zona(s) nueva(s): %s", len(desconocidas), ", ".join(desconocidas))
    try:
        import generar_zonas  # import local: solo se necesita si hay zonas nuevas
    except Exception as exc:  # noqa: BLE001
        log.warning(
            "No se pudo cargar generar_zonas (%s): esas zonas irán sin filtro geográfico",
            exc,
        )
        return

    nuevas = 0
    for zona in desconocidas:
        try:
            datos, _ = generar_zonas.geocodificar(zona)
        except Exception as exc:  # noqa: BLE001 - sin red o Nominatim caído
            log.warning("  «%s»: no se pudo geolocalizar (%s)", zona, exc)
            continue
        if not datos:
            log.warning("  «%s»: sin coincidencia fiable, irá sin filtro", zona)
            continue
        ZONAS_GEO[zona] = datos
        nuevas += 1
        log.info(
            "  «%s»: %.5f,%.5f · radio %.1f km", zona, datos["lat"], datos["lng"], datos["radio_km"]
        )
    if nuevas:
        _guardar_zonas_geo()
        log.info("%d zona(s) añadidas a %s", nuevas, ZONAS_FILE.name)


def _limpiar_zonas(zonas: list[str]) -> list[str]:
    """Normaliza la lista y elimina zonas repetidas respetando el orden."""
    limpias: list[str] = []
    vistas: set[str] = set()
    for bruto in zonas:
        for trozo in re.split(r"[\n,;]+", str(bruto)):
            zona = trozo.strip()
            if zona and zona.casefold() not in vistas:
                vistas.add(zona.casefold())
                limpias.append(zona)
    return limpias


def _url_zona(zona: str) -> str:
    """https://www.google.com/maps/search/inmobiliarias+en+<ZONA>+valencia"""
    termino = quote_plus(f"inmobiliarias en {zona} valencia")
    return PLANTILLA_URL.format(termino=termino)


def _scrape_zonas(zonas: list[str]) -> tuple[list[dict[str, str]], list[dict[str, str]]]:
    registros: list[dict[str, str]] = []
    errores: list[dict[str, str]] = []
    total = len(zonas)
    arranque = time.perf_counter()

    with sync_playwright() as pw:
        navegador = pw.chromium.launch(
            headless=True,
            args=["--disable-blink-features=AutomationControlled", "--no-sandbox"],
        )
        contexto = navegador.new_context(
            locale="es-ES",
            viewport={"width": 1440, "height": 900},
            user_agent=USER_AGENT,
        )
        pagina = contexto.new_page()
        pagina.set_default_timeout(30_000)
        log.info("Chromium listo · %d zonas en cola", total)

        try:
            for indice, zona in enumerate(zonas, start=1):
                if indice > 1 and (indice - 1) % RECICLAR_PAGINA_CADA == 0:
                    log.info("Reciclando pestaña tras %d zonas", indice - 1)
                    pagina.close()
                    pagina = contexto.new_page()
                    pagina.set_default_timeout(30_000)

                log.info("[%d/%d] «%s» · buscando…", indice, total, zona)
                _actualizar_progreso(
                    zona_actual=indice,
                    zona_nombre=zona,
                    fichas_hechas=0,
                    fichas_totales=0,
                )
                t_zona = time.perf_counter()
                try:
                    items = _scrape_zona(pagina, zona)
                    registros.extend(items)
                    # Salvaguarda: si el proceso muere en la zona 60, las anteriores ya están a salvo
                    unicos = reasignar_zonas(deduplicar(registros))
                    _guardar_datos(_sin_coords_auxiliares(unicos))
                    _actualizar_progreso(acumuladas=len(unicos))
                    log.info(
                        "[%d/%d] «%s» · %d inmobiliarias en %.1f s · %d únicas acumuladas",
                        indice,
                        total,
                        zona,
                        len(items),
                        time.perf_counter() - t_zona,
                        len(unicos),
                    )
                except Exception as exc:  # noqa: BLE001 - una zona rota no tumba el resto
                    errores.append({"zona": zona, "error": f"{type(exc).__name__}: {exc}"})
                    log.error("[%d/%d] «%s» · FALLO: %s", indice, total, zona, exc)
                time.sleep(random.uniform(1.0, 2.0))
        finally:
            contexto.close()
            navegador.close()

    log.info(
        "Recorrido terminado · %d zonas · %d registros · %.1f min",
        total,
        len(registros),
        (time.perf_counter() - arranque) / 60,
    )
    return registros, errores


def _scrape_zona(pagina, zona: str) -> list[dict[str, str]]:
    """Scrapea una zona completa: busca, hace scroll en el panel y abre cada ficha."""
    t = time.perf_counter()
    pagina.goto(_url_zona(zona), wait_until="domcontentloaded", timeout=60_000)
    _aceptar_consentimiento(pagina)
    pagina.wait_for_timeout(1500)
    log.info("      búsqueda cargada en %.1f s", time.perf_counter() - t)

    if "/sorry/" in pagina.url:
        raise RuntimeError(
            "Google ha cortado las peticiones (página /sorry/). Espera unos minutos "
            "y reintenta con menos zonas."
        )

    # Si la búsqueda tiene un único resultado, Google redirige a la ficha directamente.
    if "/maps/place/" in pagina.url:
        item = _extraer_ficha(pagina)
        item["barrio"] = zona
        coords = _coords_de_enlace(pagina.url)
        _anotar_coords(item, coords)
        _actualizar_progreso(fichas_hechas=1, fichas_totales=1)
        log.info("      un único resultado: redirigido directamente a la ficha")
        return [item] if item["nombre"] else []

    t = time.perf_counter()
    try:
        pagina.locator(SEL_FEED).wait_for(state="visible", timeout=20_000)
    except PlaywrightTimeoutError:
        log.warning("      sin panel de resultados (la zona no devolvió nada)")
        return []
    log.info("      panel de resultados listo en %.1f s", time.perf_counter() - t)

    t = time.perf_counter()
    enlaces = _recolectar_enlaces(pagina, zona)
    total_fichas = len(enlaces)
    _actualizar_progreso(fichas_hechas=0, fichas_totales=total_fichas)
    log.info(
        "      %d resultados tras %.1f s de scroll · extrayendo fichas…",
        total_fichas,
        time.perf_counter() - t,
    )

    items: list[dict[str, str]] = []
    for posicion, (href, nombre) in enumerate(enlaces.items(), start=1):
        t = time.perf_counter()
        item: dict[str, str] = {
            "nombre": nombre,
            "direccion": "",
            "barrio": zona,
            "telefono": "",
            "web": "",
        }
        _anotar_coords(item, _coords_de_enlace(href))
        problema = ""
        try:
            pagina.goto(href, wait_until="domcontentloaded", timeout=45_000)
            pagina.wait_for_selector("h1", timeout=15_000)
            try:
                pagina.wait_for_selector(SEL_DIRECCION, timeout=ESPERA_DIRECCION_MS)
            except PlaywrightTimeoutError:
                pass  # ficha sin dirección visible: seguimos con lo que haya
            pagina.wait_for_timeout(500)

            ficha = _extraer_ficha(pagina)
            for campo in ("nombre", "direccion", "telefono", "web"):
                if ficha.get(campo):
                    item[campo] = ficha[campo]
        except Exception as exc:  # noqa: BLE001 - se conserva lo que ya teníamos
            problema = f"{type(exc).__name__}: {exc}"

        items.append(item)
        _actualizar_progreso(fichas_hechas=posicion)
        log.info(
            "      %3d/%d · %5.1f s · %s · tel %s · web %s%s",
            posicion,
            total_fichas,
            time.perf_counter() - t,
            (nombre or "(sin nombre)")[:58],
            "sí" if item["telefono"] else "—",
            "sí" if item["web"] else "—",
            f" · FALLO: {problema}" if problema else "",
        )
        time.sleep(random.uniform(0.25, 0.75))

    return items


def _recolectar_enlaces(pagina, zona: str) -> dict[str, str]:
    """
    Hace scroll DENTRO del panel de resultados y acumula {href: nombre}.

    Aquí se aplica el filtro geográfico: Google devuelve a menudo negocios de
    cualquier parte del país, y como cada enlace lleva sus coordenadas podemos
    descartarlos ANTES de abrir la ficha, que es la parte lenta. Además de
    limpiar los datos, esto puede ahorrar la mayor parte del tiempo de la zona.
    """
    encontrados: dict[str, str] = {}
    anterior = -1
    sin_novedad = 0
    fuera_de_radio = 0

    for ciclo in range(MAX_CICLOS_SCROLL):
        try:
            for fila in pagina.evaluate(JS_ENLACES):
                href = (fila.get("href") or "").strip()
                if not href or href in encontrados:
                    continue
                nombre = (fila.get("nombre") or "").strip()
                coords = _coords_de_enlace(href)
                if _dentro_de_zona(coords, zona):
                    encontrados[href] = nombre
                else:
                    fuera_de_radio += 1
                    _registrar_descarte(
                        nombre,
                        zona,
                        _distancia_a_zona(coords, zona),
                        "fuera del radio de la zona",
                    )
        except Exception:  # noqa: BLE001
            pass

        if len(encontrados) == anterior:
            sin_novedad += 1
        else:
            sin_novedad = 0
            log.info(
                "      scroll %d · %d dentro de la zona%s",
                ciclo + 1,
                len(encontrados),
                f" (+{fuera_de_radio} descartados por distancia)" if fuera_de_radio else "",
            )
        anterior = len(encontrados)

        if sin_novedad >= CICLOS_SIN_NOVEDAD or _fin_de_lista(pagina):
            break

        try:
            pagina.locator(SEL_FEED).evaluate("(el) => el.scrollTo(0, el.scrollHeight)")
        except Exception:  # noqa: BLE001
            pass
        pagina.wait_for_timeout(ESPERA_SCROLL_MS)

    if fuera_de_radio:
        log.info(
            "      filtro geográfico: %d descartados, %d se quedan",
            fuera_de_radio,
            len(encontrados),
        )
    return encontrados


def _anotar_coords(item: dict[str, Any], coords: tuple[float, float] | None) -> None:
    """Guarda las coordenadas del negocio y su distancia al centro de la zona."""
    if coords is None:
        return
    item["_lat"], item["_lng"] = coords
    distancia = _distancia_a_zona(coords, item.get("barrio", ""))
    if distancia is not None:
        item["distancia_km"] = round(distancia, 2)


def _fin_de_lista(pagina) -> bool:
    try:
        return pagina.locator(SEL_FIN_LISTA).count() > 0
    except Exception:  # noqa: BLE001
        return False


def _extraer_ficha(pagina) -> dict[str, str]:
    """Lee nombre, dirección, teléfono y web de una ficha de Google Maps.

    Comprueba count() antes de leer cada campo: si no existe, devuelve "" al
    instante en lugar de agotar el timeout de 4 s de Playwright. Sin esta guarda,
    cada ficha sin teléfono ni web pierde 8 s de espera.
    """
    ficha = {"nombre": "", "direccion": "", "telefono": "", "web": ""}

    if _contar(pagina.locator("h1")):
        ficha["nombre"] = _texto(pagina.locator("h1").first)

    loc_direccion = pagina.locator(SEL_DIRECCION)
    if _contar(loc_direccion):
        ficha["direccion"] = _limpiar_etiqueta(
            _atributo(loc_direccion.first, "aria-label")
            or _texto(loc_direccion.first)
        )

    loc_telefono = pagina.locator(SEL_TELEFONO)
    if _contar(loc_telefono):
        crudo = _atributo(loc_telefono.first, "data-item-id")
        if "phone:tel:" in crudo:
            ficha["telefono"] = crudo.split("phone:tel:", 1)[1].strip()
        else:
            ficha["telefono"] = _limpiar_etiqueta(
                _atributo(loc_telefono.first, "aria-label")
                or _texto(loc_telefono.first)
            )

    loc_web = pagina.locator(SEL_WEB)
    if _contar(loc_web):
        ficha["web"] = _atributo(loc_web.first, "href")

    return ficha


def _aceptar_consentimiento(pagina) -> bool:
    """Cierra el banner de cookies (ES/EN) si aparece. Best-effort: nunca rompe el flujo."""
    for marco in (None, 'iframe[src*="consent"]'):
        for selector in SEL_CONSENTIMIENTO:
            try:
                objetivo = (
                    pagina.locator(selector)
                    if marco is None
                    else pagina.frame_locator(marco).locator(selector)
                )
                if objetivo.count() > 0:
                    objetivo.first.click(timeout=2500)
                    pagina.wait_for_timeout(800)
                    return True
            except Exception:  # noqa: BLE001
                continue
    return False


def _contar(locator) -> int:
    """Número de coincidencias. count() es inmediato: NO espera como get_attribute."""
    try:
        return locator.count()
    except Exception:  # noqa: BLE001
        return 0


def _texto(locator) -> str:
    try:
        return (locator.inner_text(timeout=4000) or "").strip()
    except Exception:  # noqa: BLE001
        return ""


def _atributo(locator, nombre: str) -> str:
    try:
        return (locator.get_attribute(nombre, timeout=4000) or "").strip()
    except Exception:  # noqa: BLE001
        return ""


def _limpiar_etiqueta(valor: str) -> str:
    """Quita prefijos tipo 'Dirección: ' / 'Website: ' de los aria-label."""
    if not valor:
        return ""
    valor = valor.strip()
    for etiqueta in (
        "Dirección:", "Address:", "Teléfono:", "Phone:", "Tel:",
        "Sitio web:", "Website:", "Web:",
    ):
        if valor.lower().startswith(etiqueta.lower()):
            return valor[len(etiqueta):].strip()
    return valor


# --------------------------------------------------------------------------- #
# Deduplicación y persistencia
# --------------------------------------------------------------------------- #
def deduplicar(registros: list[dict[str, str]]) -> list[dict[str, str]]:
    """
    Elimina duplicados por teléfono O por nombre normalizado (zonas colindantes
    devuelven las mismas inmobiliarias). Al fusionar, rellena campos vacíos.
    Los registros sin nombre ni teléfono (extracción fallida) se descartan.
    """
    unicos: list[dict[str, str]] = []
    por_nombre: dict[str, int] = {}
    por_telefono: dict[str, int] = {}

    for registro in registros:
        clave_nombre = _normalizar_nombre(registro.get("nombre", ""))
        clave_telefono = _normalizar_telefono(registro.get("telefono", ""))
        if not clave_nombre and not clave_telefono:
            continue  # fila inservible: ni nombre ni teléfono

        indice: int | None = None
        if clave_telefono and clave_telefono in por_telefono:
            indice = por_telefono[clave_telefono]
        elif clave_nombre and clave_nombre in por_nombre:
            indice = por_nombre[clave_nombre]

        if indice is None:
            unicos.append(dict(registro))
            indice = len(unicos) - 1
        else:
            for campo in ("nombre", "direccion", "telefono", "web"):
                if not unicos[indice].get(campo) and registro.get(campo):
                    unicos[indice][campo] = registro[campo]

        if clave_nombre:
            por_nombre.setdefault(clave_nombre, indice)
        if clave_telefono:
            por_telefono.setdefault(clave_telefono, indice)

    return unicos


def _normalizar_nombre(nombre: str) -> str:
    base = unicodedata.normalize("NFKD", (nombre or "").lower())
    base = "".join(c for c in base if not unicodedata.combining(c))
    base = re.sub(
        r"\b(s\.?\s?l\.?u?|s\.?\s?a\.?|sociedad limitada|sociedad anonima)\s*$", "", base
    )
    base = re.sub(r"[^a-z0-9]+", " ", base).strip()
    return re.sub(r"\s+", " ", base)


def _normalizar_telefono(telefono: str) -> str:
    return re.sub(r"\D", "", telefono or "")


def _guardar_datos(items: list[dict[str, str]]) -> None:
    """Escritura atómica: una lectura concurrente nunca ve un JSON a medias."""
    temporal = DATA_FILE.with_suffix(".json.tmp")
    temporal.write_text(
        json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    os.replace(temporal, DATA_FILE)


def _guardar_descartes(items: list[dict[str, Any]]) -> None:
    temporal = DESCARTES_FILE.with_suffix(".json.tmp")
    temporal.write_text(
        json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    os.replace(temporal, DESCARTES_FILE)


def _leer_descartes() -> list[dict[str, Any]]:
    if not DESCARTES_FILE.exists():
        return []
    try:
        datos = json.loads(DESCARTES_FILE.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return []
    return datos if isinstance(datos, list) else []


def reasignar_zonas(items: list[dict[str, str]]) -> list[dict[str, str]]:
    """
    Reparte cada inmobiliaria entre las zonas de esta ejecución según cuál tenga
    el centro más próximo, no según la búsqueda que la encontró.

    Es necesario porque las búsquedas se solapan: al rastrear "Valencia" (radio de
    8 km) entraría casi todo, y hojas como Ruzafa o El Carmen saldrían vacías.
    Si una zona no tiene centro conocido, se respeta el barrio original.

    Conserva las coordenadas auxiliares: se llama también en el guardado
    incremental, así que limpiarlas aquí rompería las pasadas siguientes.
    """
    for item in items:
        coords = None
        if item.get("_lat") is not None and item.get("_lng") is not None:
            coords = (item["_lat"], item["_lng"])
        original = item.get("barrio", "")
        asignada = _zona_mas_cercana(coords, original)
        if asignada != original:
            item["barrio_buscado"] = original
        item["barrio"] = asignada
        if coords is not None:
            distancia = _distancia_a_zona(coords, asignada)
            if distancia is not None:
                item["distancia_km"] = round(distancia, 2)
    return items


def _sin_coords_auxiliares(items: list[dict[str, str]]) -> list[dict[str, str]]:
    """Quita las coordenadas internas: no pintan nada en el JSON ni en el Excel."""
    for item in items:
        item.pop("_lat", None)
        item.pop("_lng", None)
    return items


def _leer_datos() -> list[dict[str, str]]:
    if not DATA_FILE.exists():
        return []
    try:
        datos = json.loads(DATA_FILE.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return []
    return datos if isinstance(datos, list) else []


# --------------------------------------------------------------------------- #
# Excel: formato cuidado, hoja de resumen + una hoja por barrio/población
# --------------------------------------------------------------------------- #
AZUL = "1F3B63"           # cabeceras y título
AZUL_BANDA = "F2F7FC"     # filas alternas
AZUL_CLARO = "E4EDF7"     # subtotales
GRIS_BORDE = "C9D6E4"
GRIS_TEXTO = "5F7284"
AZUL_ENLACE = "0B5FA5"
AZUL_DATO = "1F3B63"

BORDE_FINO = Side(style="thin", color=GRIS_BORDE)
BORDE_CELDA = Border(left=BORDE_FINO, right=BORDE_FINO, top=BORDE_FINO, bottom=BORDE_FINO)

ANCHOS_COLUMNA = (42, 56, 26, 14, 17, 48)   # una por columna de COLUMNAS
CARACTERES_INVALIDOS_HOJA = re.compile(r"[:\\/?*\[\]]")


def _excel_en_memoria(
    items: list[dict[str, str]], descartados: list[dict[str, Any]] | None = None
) -> io.BytesIO:
    """Libro con una hoja resumen + una hoja por barrio/población, todo formateado."""
    libro = Workbook()
    libro.remove(libro.active)  # sin la hoja "Sheet" vacía por defecto

    reservados = {"resumen"}
    _hoja_resumen(libro, items, descartados or [])
    for zona, filas in _agrupar_por_zona(items):
        _hoja_zona(libro, _nombre_hoja(zona, reservados), filas)
    if descartados:
        _hoja_descartados(libro, _nombre_hoja("Descartados", reservados), descartados)

    buffer = io.BytesIO()
    libro.save(buffer)
    buffer.seek(0)
    return buffer


def _hoja_descartados(
    libro: Workbook, nombre: str, descartados: list[dict[str, Any]]
) -> None:
    """
    Resultados que Google devolvió para una zona pero estaban fuera de su radio.

    Existe para poder auditar el filtro: si aquí aparece un negocio que debería
    estar en la hoja de su zona, es que ese radio se queda corto y hay que subirlo
    en `zonas_valencia.json`.
    """
    hoja = libro.create_sheet(nombre)

    hoja.merge_cells("A1:D1")
    titulo = hoja["A1"]
    titulo.value = (
        f"{len(descartados)} resultados descartados por estar fuera del radio de su zona"
    )
    titulo.fill = PatternFill("solid", fgColor=AZUL)
    titulo.font = Font(bold=True, size=12, color="FFFFFF")
    titulo.alignment = Alignment(vertical="center", indent=1)
    hoja.row_dimensions[1].height = 30

    hoja.merge_cells("A2:D2")
    hoja["A2"].value = (
        "Google devuelve resultados de toda la geografía para una misma búsqueda. "
        "Se conservan aquí para poder ajustar los radios si hiciera falta."
    )
    hoja["A2"].font = Font(size=9, italic=True, color=GRIS_TEXTO)
    hoja["A2"].alignment = Alignment(vertical="center", indent=1)
    hoja.row_dimensions[2].height = 20

    _pintar_cabecera(hoja, 4, ("Nombre", "Zona buscada", "Distancia al centro (km)", "Motivo"))

    ordenados = sorted(
        descartados,
        key=lambda d: (d.get("barrio", ""), -(float(d.get("distancia_km") or 0))),
    )
    for indice, descartado in enumerate(ordenados):
        fila = 5 + indice
        valores = (
            descartado.get("nombre", ""),
            descartado.get("barrio", ""),
            descartado.get("distancia_km", ""),
            descartado.get("motivo", ""),
        )
        for columna, valor in enumerate(valores, start=1):
            celda = hoja.cell(row=fila, column=columna, value=valor)
            celda.border = BORDE_CELDA
            celda.font = Font(size=10)
            celda.alignment = Alignment(vertical="center", indent=1)
            if indice % 2:
                celda.fill = PatternFill("solid", fgColor=AZUL_BANDA)
        hoja.cell(row=fila, column=1).font = Font(size=10, bold=True, color=AZUL_DATO)
        hoja.cell(row=fila, column=3).alignment = Alignment(
            horizontal="center", vertical="center"
        )
        hoja.row_dimensions[fila].height = 18

    for columna, ancho in enumerate((44, 22, 24, 34), start=1):
        hoja.column_dimensions[get_column_letter(columna)].width = ancho
    hoja.freeze_panes = "A5"
    hoja.auto_filter.ref = f"A4:D{4 + len(ordenados)}"
    hoja.sheet_view.showGridLines = False


def _agrupar_por_zona(items: list[dict[str, str]]) -> list[tuple[str, list[dict[str, str]]]]:
    """Agrupa por barrio/población y ordena alfabéticamente (sin acentos)."""
    grupos: dict[str, list[dict[str, str]]] = {}
    for item in items:
        grupos.setdefault((item.get("barrio") or "Sin zona").strip(), []).append(item)
    return sorted(grupos.items(), key=lambda par: _normalizar_nombre(par[0]))


def _nombre_hoja(zona: str, reservados: set[str]) -> str:
    """
    Excel limita los nombres de hoja a 31 caracteres, sin : \\ / ? * [ ] y únicos.
    'Riba-roja de Túria' pasa tal cual; otros se sanean.
    """
    nombre = CARACTERES_INVALIDOS_HOJA.sub("-", (zona or "").strip()).strip("'")
    nombre = (nombre or "Sin zona")[:31]
    base = nombre
    sufijo = 2
    while nombre.casefold() in reservados:
        marca = f" ({sufijo})"
        nombre = base[: 31 - len(marca)] + marca
        sufijo += 1
    reservados.add(nombre.casefold())
    return nombre


def _hoja_resumen(
    libro: Workbook, items: list[dict[str, str]], descartados: list[dict[str, Any]]
) -> None:
    """Portada con totales y desglose por zona, listo para segmentar de un vistazo."""
    hoja = libro.create_sheet("Resumen", 0)
    grupos = _agrupar_por_zona(items)
    total = len(items)
    con_tel = sum(1 for item in items if item.get("telefono"))
    con_web = sum(1 for item in items if item.get("web"))
    con_contacto = sum(1 for item in items if item.get("telefono") or item.get("web"))
    ultima_columna = get_column_letter(len(COLUMNAS))

    # --- Título ---
    hoja.merge_cells(f"A1:{ultima_columna}1")
    titulo = hoja["A1"]
    titulo.value = "Inmobiliarias · Google Maps · Valencia y área metropolitana"
    titulo.fill = PatternFill("solid", fgColor=AZUL)
    titulo.font = Font(bold=True, size=14, color="FFFFFF")
    titulo.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    hoja.row_dimensions[1].height = 38

    hoja.merge_cells(f"A2:{ultima_columna}2")
    subtitulo = hoja["A2"]
    subtitulo.value = (
        f"Generado el {datetime.now():%d/%m/%Y a las %H:%M} · "
        f"{total} inmobiliarias en {len(grupos)} zonas"
        + (f" · {len(descartados)} descartadas por distancia" if descartados else "")
    )
    subtitulo.font = Font(size=10, italic=True, color=GRIS_TEXTO)
    subtitulo.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    hoja.row_dimensions[2].height = 22

    # --- Indicadores generales ---
    fila = 4
    for columna, titulo in enumerate(("Métrica", "Total", "% del total"), start=1):
        encabezado = hoja.cell(row=fila, column=columna, value=titulo)
        encabezado.font = Font(bold=True, size=10, color=AZUL_DATO)
        encabezado.fill = PatternFill("solid", fgColor=AZUL_CLARO)
        encabezado.border = BORDE_CELDA
        encabezado.alignment = Alignment(
            horizontal="left" if columna == 1 else "center",
            vertical="center",
            indent=1 if columna == 1 else 0,
        )
    hoja.row_dimensions[fila].height = 20

    indicadores = (
        ("Inmobiliarias únicas", total, "0"),
        ("Con teléfono", con_tel, "0"),
        ("Con web", con_web, "0"),
        ("Con teléfono o web", con_contacto, "0"),
    )
    fila = 5
    for etiqueta, valor, formato in indicadores:
        celda_etiqueta = hoja.cell(row=fila, column=1, value=etiqueta)
        celda_etiqueta.font = Font(bold=True, size=10, color=AZUL_DATO)
        celda_etiqueta.fill = PatternFill("solid", fgColor=AZUL_CLARO)
        celda_etiqueta.border = BORDE_CELDA
        celda_etiqueta.alignment = Alignment(vertical="center", indent=1)

        celda_valor = hoja.cell(row=fila, column=2, value=valor)
        celda_valor.number_format = formato
        celda_valor.font = Font(bold=True, size=11, color=AZUL_DATO)
        celda_valor.alignment = Alignment(horizontal="center", vertical="center")
        celda_valor.border = BORDE_CELDA

        porcentaje = (valor / total) if total else 0
        celda_pct = hoja.cell(row=fila, column=3, value=porcentaje)
        celda_pct.number_format = "0.0%"
        celda_pct.font = Font(size=10, color=GRIS_TEXTO)
        celda_pct.alignment = Alignment(horizontal="center", vertical="center")
        celda_pct.border = BORDE_CELDA
        hoja.row_dimensions[fila].height = 19
        fila += 1

    # --- Desglose por zona ---
    fila += 1
    hoja.cell(row=fila, column=1, value="Desglose por barrio / población").font = Font(
        bold=True, size=11, color=AZUL_DATO
    )
    fila += 1

    cabeceras = ("Barrio / Población", "Inmobiliarias", "Con teléfono", "Con web", "% con contacto")
    _pintar_cabecera(hoja, fila, cabeceras)
    primera_fila_datos = fila + 1

    for indice, (zona, filas_zona) in enumerate(grupos):
        fila += 1
        tamano = len(filas_zona)
        tel_zona = sum(1 for item in filas_zona if item.get("telefono"))
        web_zona = sum(1 for item in filas_zona if item.get("web"))
        contacto = sum(
            1 for item in filas_zona if item.get("telefono") or item.get("web")
        )

        valores = (
            zona,
            tamano,
            tel_zona,
            web_zona,
            (contacto / tamano) if tamano else 0,
        )
        for columna, valor in enumerate(valores, start=1):
            celda = hoja.cell(row=fila, column=columna, value=valor)
            celda.border = BORDE_CELDA
            celda.font = Font(size=10)
            if indice % 2:
                celda.fill = PatternFill("solid", fgColor=AZUL_BANDA)
            if columna == 1:
                celda.alignment = Alignment(vertical="center", indent=1)
                celda.font = Font(size=10, bold=True, color=AZUL_DATO)
            else:
                celda.alignment = Alignment(horizontal="center", vertical="center")
            celda.number_format = "0.0%" if columna == 5 else "0"
        hoja.row_dimensions[fila].height = 18

    # Fila de totales
    fila += 1
    totales = ("TOTAL", total, con_tel, con_web, (con_contacto / total) if total else 0)
    for columna, valor in enumerate(totales, start=1):
        celda = hoja.cell(row=fila, column=columna, value=valor)
        celda.fill = PatternFill("solid", fgColor=AZUL_CLARO)
        celda.font = Font(bold=True, size=10, color=AZUL_DATO)
        celda.border = Border(
            left=BORDE_FINO, right=BORDE_FINO, bottom=BORDE_FINO, top=Side(style="medium", color=AZUL)
        )
        celda.number_format = "0.0%" if columna == 5 else "0"
        celda.alignment = (
            Alignment(vertical="center", indent=1)
            if columna == 1
            else Alignment(horizontal="center", vertical="center")
        )
    hoja.row_dimensions[fila].height = 20

    # Barra de datos en la columna de volumen: se ve el peso de cada zona de un vistazo
    hoja.conditional_formatting.add(
        f"B{primera_fila_datos}:B{fila - 1}",
        DataBarRule(start_type="num", start_value=0, end_type="max", color="5B9BD5"),
    )

    for columna, ancho in enumerate((34, 15, 14, 12, 15), start=1):
        hoja.column_dimensions[get_column_letter(columna)].width = ancho
    hoja.sheet_view.showGridLines = False


def _pintar_cabecera(hoja, fila: int, titulos: tuple[str, ...]) -> None:
    for columna, titulo in enumerate(titulos, start=1):
        celda = hoja.cell(row=fila, column=columna, value=titulo)
        celda.fill = PatternFill("solid", fgColor=AZUL)
        celda.font = Font(bold=True, size=10, color="FFFFFF")
        celda.alignment = Alignment(
            horizontal="left" if columna == 1 else "center",
            vertical="center",
            indent=1 if columna == 1 else 0,
        )
        celda.border = BORDE_CELDA
    hoja.row_dimensions[fila].height = 24


def _hoja_zona(libro: Workbook, nombre: str, filas: list[dict[str, str]]) -> None:
    """Una hoja por barrio, con cabecera coloreada, filas alternas y web clicable."""
    hoja = libro.create_sheet(nombre)

    _pintar_cabecera(hoja, 1, tuple(titulo for _, titulo in COLUMNAS))

    for indice, item in enumerate(filas):
        fila = indice + 2
        banda = PatternFill("solid", fgColor=AZUL_BANDA) if indice % 2 else None

        for columna, (clave, _) in enumerate(COLUMNAS, start=1):
            celda = hoja.cell(row=fila, column=columna, value=str(item.get(clave) or ""))
            celda.border = BORDE_CELDA
            celda.font = Font(size=10)
            celda.alignment = Alignment(vertical="center", indent=1)
            if banda:
                celda.fill = banda

        # Nombre destacado y teléfono centrado
        hoja.cell(row=fila, column=1).font = Font(size=10, bold=True, color=AZUL_DATO)
        hoja.cell(row=fila, column=4).alignment = Alignment(
            horizontal="center", vertical="center"
        )

        # Web como enlace real
        celda_web = hoja.cell(row=fila, column=5)
        if celda_web.value:
            celda_web.hyperlink = celda_web.value
            celda_web.font = Font(size=10, color=AZUL_ENLACE, underline="single")

        hoja.row_dimensions[fila].height = 18

    for columna, ancho in enumerate(ANCHOS_COLUMNA, start=1):
        hoja.column_dimensions[get_column_letter(columna)].width = ancho

    hoja.freeze_panes = "A2"
    hoja.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNAS))}{len(filas) + 1}"
    hoja.sheet_view.showGridLines = False


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("app:app", host="127.0.0.1", port=8000)
