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

import pandas as pd
from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse, StreamingResponse
from pydantic import BaseModel, Field
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

# --------------------------------------------------------------------------- #
# Configuración
# --------------------------------------------------------------------------- #
BASE_DIR = Path(__file__).resolve().parent
TEMPLATES_DIR = BASE_DIR / "templates"
DATA_FILE = BASE_DIR / "data_temp.json"

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
        log.info("═══ Nueva recopilación · %d zonas ═══", len(zonas))

        registros, errores = _scrape_zonas(zonas)
        unicos = deduplicar(registros)
        _guardar_datos(unicos)

        log.info(
            "═══ Terminado · %d únicas de %d brutos · %d duplicadas · %.1f min ═══",
            len(unicos),
            len(registros),
            len(registros) - len(unicos),
            (time.time() - inicio) / 60,
        )
        return {
            "ok": True,
            "zonas": len(zonas),
            "resultados": len(registros),
            "unicos": len(unicos),
            "duplicados": len(registros) - len(unicos),
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
    buffer = _excel_en_memoria(items)
    nombre = f"inmobiliarias_valencia_{datetime.now():%Y%m%d_%H%M}.xlsx"
    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{nombre}"'},
    )


# --------------------------------------------------------------------------- #
# Scraping (Playwright, sincrónico -> FastAPI lo ejecuta en un hilo aparte)
# --------------------------------------------------------------------------- #
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
                t_zona = time.perf_counter()
                try:
                    items = _scrape_zona(pagina, zona)
                    registros.extend(items)
                    # Salvaguarda: si el proceso muere en la zona 60, las anteriores ya están a salvo
                    unicos = deduplicar(registros)
                    _guardar_datos(unicos)
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
    enlaces = _recolectar_enlaces(pagina)
    total_fichas = len(enlaces)
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


def _recolectar_enlaces(pagina) -> dict[str, str]:
    """
    Hace scroll DENTRO del panel de resultados y va acumulando {href: nombre}.
    Se acumula sobre la marcha para no perder tarjetas si Google virtualiza el feed.
    """
    encontrados: dict[str, str] = {}
    anterior = -1
    sin_novedad = 0

    for ciclo in range(MAX_CICLOS_SCROLL):
        try:
            for fila in pagina.evaluate(JS_ENLACES):
                href = (fila.get("href") or "").strip()
                if href and href not in encontrados:
                    encontrados[href] = (fila.get("nombre") or "").strip()
        except Exception:  # noqa: BLE001
            pass

        if len(encontrados) == anterior:
            sin_novedad += 1
        else:
            sin_novedad = 0
            log.info(
                "      scroll %d · %d resultados (+%d)",
                ciclo + 1,
                len(encontrados),
                len(encontrados) - max(anterior, 0),
            )
        anterior = len(encontrados)

        if sin_novedad >= CICLOS_SIN_NOVEDAD or _fin_de_lista(pagina):
            break

        try:
            pagina.locator(SEL_FEED).evaluate("(el) => el.scrollTo(0, el.scrollHeight)")
        except Exception:  # noqa: BLE001
            pass
        pagina.wait_for_timeout(ESPERA_SCROLL_MS)

    return encontrados


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


def _leer_datos() -> list[dict[str, str]]:
    if not DATA_FILE.exists():
        return []
    try:
        datos = json.loads(DATA_FILE.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return []
    return datos if isinstance(datos, list) else []


# --------------------------------------------------------------------------- #
# Excel
# --------------------------------------------------------------------------- #
def _excel_en_memoria(items: list[dict[str, str]]) -> io.BytesIO:
    filas = [
        {titulo: str(item.get(clave) or "") for clave, titulo in COLUMNAS}
        for item in items
    ]
    # dtype=str: evita que pandas convierta los teléfonos en números (644130714.0).
    dataframe = pd.DataFrame(
        filas, columns=[titulo for _, titulo in COLUMNAS], dtype=str
    )
    dataframe = dataframe.fillna("")

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name="Inmobiliarias")
        hoja = writer.sheets["Inmobiliarias"]
        for columna, ancho in {"A": 42, "B": 52, "C": 22, "D": 18, "E": 48}.items():
            hoja.column_dimensions[columna].width = ancho
        hoja.freeze_panes = "A2"

    buffer.seek(0)
    return buffer


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("app:app", host="127.0.0.1", port=8000)
