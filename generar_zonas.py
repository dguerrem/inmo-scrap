"""
Genera (y mantiene) `zonas_valencia.json`: centro geográfico, radio de filtrado y
código postal de cada zona.

Para qué sirve: Google Maps no respeta la zona buscada y devuelve resultados de
cualquier parte (se han medido casos a 700 km). Este fichero permite filtrar cada
resultado por distancia real, que es la única forma fiable de que la hoja de
"Carpesa" contenga solo lo que hay alrededor de Carpesa.

Se ejecuta una sola vez y el resultado queda cacheado; el scraper funciona después
sin conexión. Si añades zonas nuevas, vuelve a ejecutar:

    .venv/bin/python generar_zonas.py

Fuente de datos: Nominatim (OpenStreetMap). Sin clave de API.

Tres salvaguardas evitan centros absurdos:
  1. Guardián geográfico: todo el listado es Valencia + área metropolitana, así que
     cualquier candidato fuera de esa caja se descarta sin más.
  2. El nombre del candidato debe parecerse al de la zona (si no, se descarta).
  3. A igualdad de condiciones gana el candidato más próximo a Valencia capital,
     porque un falso positivo casi siempre acaba lejos.
"""
from __future__ import annotations

import json
import math
import re
import sys
import time
import unicodedata
import urllib.parse
import urllib.request
from difflib import SequenceMatcher
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
ZONAS_FILE = BASE_DIR / "zonas_valencia.json"

UA = {"User-Agent": "inmo-scrap/1.0 (herramienta local de uso interno)"}
PAUSA_S = 1.1            # Nominatim pide como máximo 1 petición por segundo

# Caja que engloba Valencia capital y su área metropolitana: todo el listado de
# zonas cae dentro. Es la salvaguarda más eficaz contra geocodificaciones basura.
CAJA = {"lat_min": 39.15, "lat_max": 39.75, "lng_min": -0.80, "lng_max": -0.20}
CENTRO_VALENCIA = (39.46971, -0.37633)

RADIO_MIN_KM = 1.2       # suelo: barrios muy pequeños siguen teniendo margen
RADIO_MAX_KM = 8.0       # tope para términos municipales grandes

# El radio se deduce del tamaño real de la zona (mitad de su extensión mayor), no
# de una cifra fija. Un radio fijo de 3,5 km para todas era un error: alrededor de
# Carpesa (1,7 x 2,8 km) caben Rascanya, Benicalap, Poble Nou y Benimàmet, así que
# la hoja de Carpesa acababa con ~70 inmobiliarias de barrios vecinos.
FACTOR_RADIO = 0.5

CARRETERAS = {"highway", "railway", "waterway", "natural", "landuse", "barrier"}

# Nominatim devuelve a menudo el topónimo en castellano aunque se pregunte en
# valenciano ("Ciudad Vieja" por Ciutat Vella, "Ensanche" por L'Eixample), así que
# cada zona admite varias formas. La comparación es además difusa, para tolerar
# variantes ortográficas (Burjassot/Burjasot, Puçol/Puzol).
SINONIMOS = {
    "ciutat vella": ["ciudad vieja"],
    "l'eixample": ["ensanche", "eixample"],
    "extramurs": ["extramuros"],
    "la saidia": ["zaidia", "saidia", "la zaidia"],
    "el pla del real": ["pla del real", "llano del real"],
    "l'olivereta": ["olivereta"],
    "poblats maritims": ["maritim", "maritimo", "poblats maritims", "poblados maritimos"],
    "castellar-oliveral": ["castellar", "oliveral", "castellar oliveral"],
    "forn d'alcedo": ["horno de alcedo", "forn d'alcedo", "forn dalcedo"],
    "benimaclet": ["benimaclet"],
    "rascanya": ["rascana", "raxcanya"],
    "benicalap": ["benicalap"],
    "benimamet": ["benimamet"],
    "el carmen": ["carme", "carmen"],
    "el cabanyal": ["cabanyal", "cabañal"],
    "algirós": ["algirós", "algirós"],
    "quatre carreres": ["cuatro carreras"],
    "camins al grau": ["caminos al grau", "camins al grau"],
    "poble nou": ["pueblo nuevo", "poble nou"],
    "pobla de farnals": ["puebla de farnals", "pobla de farnals"],
    "l'eliana": ["eliana"],
    "silla": ["silla"],
    "albal": ["albal"],
    "catarroja": ["catarroja"],
    "torrent": ["torrente"],
    "el saler": ["saler"],
    "la torre": ["torre"],
    "pinedo": ["pinedo"],
    "carpesa": ["carpesa"],
}


def _parecido(a: str, b: str) -> float:
    return SequenceMatcher(None, a, b).ratio()


def _nombre_valido(zona: str, cand: dict) -> bool:
    """¿Se reconoce el nombre de la zona en el candidato? Comparación difusa."""
    zn = _sin_acentos(zona)
    nucleo = re.sub(r"^(l|la|el|els|les|los|las)\s+", "", zn)
    variantes = {zn, nucleo} | {_sin_acentos(s) for s in SINONIMOS.get(zn, [])}

    partes = [_sin_acentos(p).strip() for p in (cand.get("display_name") or "").split(",")]
    if cand.get("name"):
        partes.append(_sin_acentos(cand["name"]).strip())

    for variante in variantes:
        if len(variante) < 4:
            continue
        for parte in partes:
            if not parte:
                continue
            if variante == parte or variante in parte or parte in variante:
                return True
            if _parecido(variante, parte) >= 0.78:
                return True
    return False

# Nombres que Nominatim no resuelve bien tal cual
ALIASES = {
    "poblats maritims": "Marítim, València",
    "avenida del puerto": "Avinguda del Port, València",
    "forn d'alcedo": "Forn d'Alcedo, València",
    "pobla de farnals": "la Pobla de Farnals",
    "l'eliana": "l'Eliana, València",
    "l'eixample": "l'Eixample, València",
    "l'olivereta": "l'Olivereta, València",
    "la saidia": "la Saïdia, València",
    "el pla del real": "el Pla del Real, València",
    "castellar-oliveral": "Castellar-Oliveral, València",
    "el saler": "el Saler, València",
    "el carmen": "el Carme, València",
    "ruzafa": "Russafa, València",
    "benimamet": "Benimàmet, València",
    "camins al grau": "Camins al Grau, València",
    "poble nou": "Poble Nou, València",
    "rascanya": "Rascanya, València",
    "benicalap": "Benicalap, València",
    "beniferri": "Beniferri, València",
    "la torre": "la Torre, València",
    "carpesa": "Carpesa, València",
    "pinedo": "Pinedo, València",
    "el palmar": "el Palmar, València",
    "catarroja": "Catarroja, València",
    "silla": "Silla, València",
    "albal": "Albal, València",
    "alfafar": "Alfafar, València",
}

ZONAS_POR_DEFECTO = [
    "Valencia", "Ciutat Vella", "El Carmen", "Ruzafa", "L'Eixample", "Extramurs",
    "Campanar", "La Saïdia", "El Pla del Real", "L'Olivereta", "Patraix", "Jesús",
    "Quatre Carreres", "Poblats Marítims", "El Cabanyal", "Camins al Grau",
    "Avenida del Puerto", "Algirós", "Benimaclet", "Rascanya", "Benicalap",
    "Benimàmet", "Beniferri", "Poble Nou", "Carpesa", "Castellar-Oliveral",
    "La Torre", "Forn d'Alcedo", "Pinedo", "El Saler", "El Palmar", "Mislata",
    "Quart de Poblet", "Manises", "Aldaia", "Alaquàs", "Torrent", "Xirivella",
    "Paiporta", "Picanya", "Sedaví", "Benetússer", "Alfafar", "Massanassa",
    "Catarroja", "Albal", "Silla", "Picassent", "Burjassot", "Paterna", "Godella",
    "Rocafort", "Moncada", "Tavernes Blanques", "Almàssera", "Meliana", "Foios",
    "Albalat dels Sorells", "Museros", "Massamagrell", "Pobla de Farnals",
    "Rafelbunyol", "El Puig", "Puçol", "Llíria", "Riba-roja de Túria", "L'Eliana",
    "Loriguilla",
]

# Puntos de referencia para autoverificar el resultado (tolerancia en km).
# Si alguna zona se aleja más de lo permitido, el script avisa: así un cambio en
# Nominatim no pasa desapercibido.
VERIFICACION = {
    "Valencia": (39.4697, -0.3763, 3.0),
    "Carpesa": (39.5170, -0.3780, 3.0),
    "Ruzafa": (39.4623, -0.3721, 3.0),
    "Torrent": (39.4367, -0.4656, 3.0),
    "Llíria": (39.6251, -0.5953, 3.0),
    "Mislata": (39.4751, -0.4179, 3.0),
    "El Cabanyal": (39.4720, -0.3287, 3.0),
    "Benimàmet": (39.5002, -0.4189, 3.0),
    "Picassent": (39.3635, -0.4612, 3.0),
    "Puçol": (39.6172, -0.3028, 3.0),
}


def _sin_acentos(texto: str) -> str:
    base = unicodedata.normalize("NFKD", (texto or "").lower())
    return "".join(c for c in base if not unicodedata.combining(c))


def km_entre(a: tuple[float, float], b: tuple[float, float]) -> float:
    r = 6371.0
    lat1, lat2 = math.radians(a[0]), math.radians(b[0])
    dlat = lat2 - lat1
    dlng = math.radians(b[1] - a[1])
    h = math.sin(dlat / 2) ** 2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlng / 2) ** 2
    return 2 * r * math.asin(math.sqrt(h))


def _get(url: str):
    try:
        with urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=30) as r:
            return json.load(r)
    except Exception:
        return None


def _buscar(consulta: str) -> list[dict]:
    params = urllib.parse.urlencode(
        {
            "q": consulta,
            "format": "json",
            "limit": 8,
            "countrycodes": "es",
            "addressdetails": 1,
            "accept-language": "es",
        }
    )
    return _get(f"https://nominatim.openstreetmap.org/search?{params}") or []


def _reverse(lat: float, lng: float) -> dict:
    params = urllib.parse.urlencode(
        {"lat": lat, "lon": lng, "format": "json", "addressdetails": 1,
         "accept-language": "es", "zoom": 18}
    )
    return _get(f"https://nominatim.openstreetmap.org/reverse?{params}") or {}


def _extension_km(bbox: list[str]) -> float:
    sur, norte, oeste, este = (float(x) for x in bbox)
    alto = (norte - sur) * 111.0
    ancho = (este - oeste) * 111.0 * math.cos(math.radians((norte + sur) / 2))
    return max(ancho, alto)


def _evaluar(cand: dict, zona: str) -> tuple[float, str] | None:
    """Devuelve (puntuación, motivo_descarte) o None si el candidato sirve."""
    direccion = cand.get("address") or {}
    if direccion.get("country_code", "es") != "es":
        return None, "no es España"

    try:
        lat, lng = float(cand["lat"]), float(cand["lon"])
    except (KeyError, TypeError, ValueError):
        return None, "sin coordenadas"

    if not (CAJA["lat_min"] <= lat <= CAJA["lat_max"] and CAJA["lng_min"] <= lng <= CAJA["lng_max"]):
        return None, "fuera del área de Valencia y su entorno"

    # El nombre debe reconocerse en alguna de sus formas: sin esto, cualquier
    # homónimo dentro de la caja geográfica valdría.
    if not _nombre_valido(zona, cand):
        return None, f"el nombre «{cand.get('name')}» no coincide con «{zona}»"

    clase, tipo = cand.get("class"), cand.get("type")
    if clase in CARRETERAS:
        # Una vía no es una zona, pero su punto medio sirve como centroide cuando
        # no hay nada mejor: para "Avenida del Puerto" o "Forn d'Alcedo" (que en
        # OSM solo existe como apeadero) es la única referencia disponible. La
        # penalización garantiza que cualquier límite administrativo gane.
        return 1.0, ""

    # Se prefiere el límite administrativo (el barrio o municipio "de verdad") a un
    # place/locality, que muchas veces es una urbanización homónima dentro de otro
    # municipio. Es el caso de "El Palmar": el barrio de València es
    # boundary/administrative, mientras que una urbanización de Moncada con el
    # mismo nombre es place/locality. NO se penaliza por distancia a València:
    # hizo que ganara la urbanización de Moncada solo por estar más cerca.
    if clase == "boundary" and tipo == "administrative":
        puntos = 10.0
    elif clase == "place":
        puntos = 8.0
    else:
        puntos = 5.0

    if tipo in {"city", "town", "village", "municipality", "suburb", "neighbourhood", "quarter"}:
        puntos += 1.5
    if nucleo := re.sub(r"^(l|la|el|els|les|los|las)\s+", "", _sin_acentos(zona)):
        nombre_cand = _sin_acentos(cand.get("name") or "")
        if nucleo == nombre_cand or _parecido(nucleo, nombre_cand) >= 0.9:
            puntos += 1.0
    return puntos, ""


def _candidatos(zona: str) -> tuple[list[tuple[float, dict]], list[str]]:
    """Candidatos válidos ordenados por puntuación, y los motivos de descarte."""
    nombre = ALIASES.get(_sin_acentos(zona), zona)
    consultas = [
        f"{nombre}, València, Espanya",
        f"{nombre}, Valencia, España",
        nombre,
    ]
    vistos: dict[str, dict] = {}
    motivos: list[str] = []
    for consulta in consultas:
        for cand in _buscar(consulta):
            clave = f"{cand.get('osm_type')}{cand.get('osm_id')}"
            if clave in vistos:
                continue
            vistos[clave] = cand
            _, motivo = _evaluar(cand, zona)
            if motivo:
                motivos.append(f"{cand.get('name')!r}: {motivo}")
        time.sleep(PAUSA_S)
        validos = [c for c in vistos.values() if _evaluar(c, zona)[0] is not None]
        if len(validos) >= 3:
            break

    puntuados: list[tuple[float, dict]] = []
    for cand in vistos.values():
        puntos, _ = _evaluar(cand, zona)
        if puntos is not None:
            puntuados.append((puntos, cand))
    # A igual puntuación gana el más próximo a València: es solo un desempate entre
    # homónimos, nunca debe pesar más que el tipo de elemento.
    puntuados.sort(
        key=lambda par: (
            -par[0],
            km_entre((float(par[1]["lat"]), float(par[1]["lon"])), CENTRO_VALENCIA),
        )
    )
    # Deduplicar por coordenadas: Nominatim repite el mismo sitio varias veces.
    unicos: list[tuple[float, dict]] = []
    vistas: set[tuple[float, float]] = set()
    for puntos, cand in puntuados:
        clave = (round(float(cand["lat"]), 4), round(float(cand["lon"]), 4))
        if clave in vistas:
            continue
        vistas.add(clave)
        unicos.append((puntos, cand))
    return unicos, motivos


def geocodificar(zona: str) -> tuple[dict | None, list[str]]:
    """Devuelve (datos de la zona, motivos de descarte) — los motivos solo si falla."""
    candidatos, motivos = _candidatos(zona)
    if not candidatos:
        return None, motivos

    mejor_puntos, mejor = candidatos[0]
    segundo = candidatos[1][0] if len(candidatos) > 1 else None
    lat, lng = float(mejor["lat"]), float(mejor["lon"])

    try:
        extension = _extension_km(mejor["boundingbox"])
    except Exception:
        extension = 0.0

    radio = min(max(extension * FACTOR_RADIO, RADIO_MIN_KM), RADIO_MAX_KM)

    # El CP se guarda solo como información: se ha comprobado que usarlo como
    # filtro estricto descarta resultados legítimos (Cercasa está a 2,5 km del
    # centro de Carpesa pero su CP, 46019, es el de Rascaña).
    cp = (mejor.get("address") or {}).get("postcode")
    if not cp:
        inversa = _reverse(lat, lng).get("address") or {}
        cp = inversa.get("postcode")
        time.sleep(PAUSA_S)

    return {
        "lat": round(lat, 6),
        "lng": round(lng, 6),
        "radio_km": round(radio, 1),
        "cp": cp,
        "extension_km": round(extension, 1),
        "puntos": round(mejor_puntos, 1),
        "margen": round(mejor_puntos - segundo, 1) if segundo is not None else None,
        "nombre_osm": mejor.get("name") or "",
    }, []


def cargar() -> dict:
    if not ZONAS_FILE.exists():
        return {}
    try:
        return json.loads(ZONAS_FILE.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}


def verificar(datos: dict) -> list[str]:
    """Comprueba las zonas de referencia. Devuelve la lista de fallos."""
    fallos = []
    for zona, (lat, lng, tolerancia) in VERIFICACION.items():
        d = datos.get(zona)
        if not d:
            fallos.append(f"{zona}: no se generó")
            continue
        distancia = km_entre((d["lat"], d["lng"]), (lat, lng))
        if distancia > tolerancia:
            fallos.append(
                f"{zona}: {d['lat']:.5f},{d['lng']:.5f} está a {distancia:.1f} km "
                f"del punto esperado (máx {tolerancia} km)"
            )
    return fallos


def recalcular_radios() -> int:
    """
    Recalcula los radios ya guardados a partir de la extensión almacenada.

    Sirve para ajustar RADIO_MIN_KM o FACTOR_RADIO sin volver a consultar
    Nominatim: la extensión de cada zona ya está en el fichero.
    """
    datos = cargar()
    if not datos:
        return 0
    for zona, info in datos.items():
        if not info.get("extension_km"):
            continue
        info["radio_km"] = round(
            min(max(info["extension_km"] * FACTOR_RADIO, RADIO_MIN_KM), RADIO_MAX_KM), 1
        )
    ZONAS_FILE.write_text(json.dumps(datos, ensure_ascii=False, indent=2), encoding="utf-8")
    return len(datos)


def main() -> None:
    if "--recalcular-radios" in sys.argv:
        total = recalcular_radios()
        print(f"Radios recalculados en {ZONAS_FILE.name} para {total} zonas "
              f"(suelo {RADIO_MIN_KM} km, factor {FACTOR_RADIO}, tope {RADIO_MAX_KM} km)")
        return

    zonas = [a for a in sys.argv[1:] if not a.startswith("--")] or ZONAS_POR_DEFECTO
    existente = cargar()
    # Se parte de lo ya calculado y se conservan las zonas que no se piden ahora:
    # generar una zona suelta no debe borrar el resto del fichero.
    salida: dict[str, dict] = dict(existente)

    print(f"Geolocalizando {len(zonas)} zonas (Nominatim, ~1 petición/s)...")
    for i, zona in enumerate(zonas, 1):
        if existente.get(zona):
            print(f"  [{i}/{len(zonas)}] {zona}: ya estaba, se reutiliza")
            continue

        datos, motivos = geocodificar(zona)
        if not datos:
            print(f"  [{i}/{len(zonas)}] {zona}: SIN GEOLOCALIZAR (sin filtro para esta zona)")
            for motivo in dict.fromkeys(motivos):
                print(f"          descartado: {motivo}")
            continue
        salida[zona] = datos
        margen = f"margen {datos['margen']:+.1f}" if datos["margen"] is not None else "único"
        print(
            f"  [{i}/{len(zonas)}] {zona:22} {datos['lat']:9.5f},{datos['lng']:10.5f} "
            f"radio {datos['radio_km']:>3} km  CP {datos['cp'] or '-':>6}  {margen}"
        )

    ZONAS_FILE.write_text(json.dumps(salida, ensure_ascii=False, indent=2), encoding="utf-8")

    fallos = verificar(salida)
    print()
    print(f"Escrito {ZONAS_FILE.name}: {len(salida)} zonas · "
          f"{sum(1 for v in salida.values() if v['cp'])} con CP")
    if fallos:
        print()
        print("AVISO: la autoverificación ha detectado desviaciones:")
        for f in fallos:
            print("  -", f)
    else:
        print(f"Autoverificación correcta: las {len(VERIFICACION)} zonas de referencia "
              "están donde se espera.")


if __name__ == "__main__":
    main()
