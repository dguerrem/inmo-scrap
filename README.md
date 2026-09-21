# inmo-scrap

Extrae inmobiliarias de Google Maps (Valencia + área metropolitana) y las descarga en
Excel. Sin base de datos: los resultados viven en `data_temp.json` mientras el proceso
está en marcha.

---

# 🚀 Guía rápida (para mí, dentro de un mes)

## Paso 0 · Levantar la web

Abre una terminal en esta carpeta y ejecuta:

```bash
cd ~/Desktop/Everything/Projects/Other/inmo-scrap
.venv/bin/python -m uvicorn app:app --host 127.0.0.1 --port 8000
```

Luego abre en el navegador: **<http://127.0.0.1:8000>**

> Si el comando falla porque `.venv` no existe o está roto, salta a
> [Reinstalar desde cero](#-reinstalar-desde-cero).
>
> Si sale un error tipo `[Errno 48] Address already in use`, **ya hay un servidor
> arrancado** en ese puerto (quizá en otra terminal o pestaña olvidada). Para
> liberarlo: `pkill -f "uvicorn app:app"` y vuelve a lanzarlo.

## Paso 1 · Elegir zonas

Arriba verás un recuadro con **68 zonas ya escritas**, una por línea
(barrios de Valencia, municipios del área metropolitana, Llíria, Riba-roja…).

- Borra las que no quieras, añade las que falten.
- **Una zona por línea. Nada de comas.**
- Para una prueba rápida, deja solo 2 o 3 zonas.

## Paso 2 · Recopilar datos

Pulsa **「Recopilar datos」**.

- El botón se bloquea y pasa a "Scrapeando…".
- Se abre la **consola negra** con la traza en directo.

## Paso 3 · Mirar la consola (lo interesante)

Es un terminal en vivo. Cada línea es un paso del bot:

| Lo que lees                                        | Qué significa                                             |
| -------------------------------------------------- | --------------------------------------------------------- |
| `═══ Nueva recopilación · 68 zonas ═══`            | Empieza una ejecución.                                    |
| `[12/68] «Patraix» · buscando…`                    | Va por la zona 12 de 68.                                  |
| `búsqueda cargada en 2.9 s`                        | Tardó 2,9 s en abrir la búsqueda de esa zona.             |
| `scroll 3 · 48 resultados (+9)`                    | El scroll ya ha cargado 48 tarjetas (9 nuevas).           |
| `48 resultados tras 12.4 s de scroll`              | Fin del scroll: hay 48 inmobiliarias que visitar.         |
| `  7/48 · 2.1 s · Casa Nova · tel sí · web —`      | Ficha 7 de 48, tardó 2,1 s; tiene teléfono y no web.      |
| `[12/68] «Patraix» · 48 inmobiliarias en 341.2 s`  | Zona terminada: 48 fichas en 5 min 41 s.                  |
| `filtro geográfico: 224 descartados, 70 se quedan` | Cuántos resultados de esa búsqueda caían fuera del radio. |
| `═══ Terminado · 812 únicas de 903 brutos ═══`     | Fin: 812 tras eliminar 91 duplicados.                     |

- **Color**: gris = información normal, ámbar = aviso, rojo = fallo.
- **`FALLO: ...`** al final de una línea = esa ficha no se pudo leer. No es grave:
  el bot continúa y guarda lo que ya tenía. `TimeoutError` es lo más habitual.
- Casilla **autoscroll**: desmárcala si quieres subir a leer sin que te arrastre.
- Enlace **descargar**: baja el log completo a un `.txt`, incluido el de ejecuciones
  anteriores. Útil para analizar tiempos con calma.

> **Puedes cerrar o recargar la pestaña.** El bot sigue trabajando en el servidor.
> Al volver a entrar, la web detecta la ejecución en marcha y se reengancha sola a
> la consola.

## Paso 4 · Seguir la barra de progreso

Bajo el botón aparece una **barra de progreso** con:

| Elemento                     | Qué te dice                                                  |
| ---------------------------- | ------------------------------------------------------------ |
| `Zona 12/68 · «Patraix»`     | Por dónde va y cuántas zonas tiene en total.                 |
| `38 %`                       | Porcentaje global (zonas terminadas + la actual).            |
| `Fichas de esta zona: 37/48` | Inmobiliarias dentro del radio ya procesadas en esta zona.   |
| `24 únicas acumuladas`       | Lo que lleva deduplicado (se actualiza al cerrar cada zona). |
| `Quedan ~12 min 30 s`        | Estimación según el ritmo real que lleva.                    |

> El tiempo restante es una **estimación**: se calcula con la media de las zonas ya
> terminadas. Al principio puede bailar bastante (las zonas tienen tamaños muy
> dispares) y se va afinando según avanza. Si aún no ha terminado ninguna zona
> muestra solo el tiempo transcurrido.

## Paso 5 · Descargar el Excel

Cuando termina aparece el botón verde **「⬇️ Descargar Excel」** y el mensaje
`✅ La recopilación ha terminado`.

- El botón solo aparece al terminar **todas** las zonas.
- Debajo tienes una **Vista previa** con las filas recogidas, para comprobar que
  todo tiene sentido antes de descargar.
- El archivo se llama `inmobiliarias_valencia_AAAAMMDD_HHMM.xlsx`.

### Cómo viene el Excel

Trae **una pestaña de resumen + una pestaña por cada barrio/población**:

- **`Resumen`** (primera pestaña): totales generales (cuántas tienen teléfono, web,
  porcentajes), y una tabla con el desglose zona a zona. La columna de volumen lleva
  una **barra de datos azul** para ver de un vistazo qué zonas concentran más
  inmobiliarias. La última fila es el `TOTAL`.
- **Una pestaña por zona** (ordenadas alfabéticamente), con:
  - Cabecera azul marino con texto blanco y **filas alternas** en azul muy claro.
  - Fila de cabecera **congelada** y **autofiltro** activado: puedes filtrar y
    ordenar sin perder de vista los títulos.
  - **Web como enlace clicable** (azul subrayado).
  - Teléfono centrado y en **formato texto** (evita que Excel se coma el `0`
    inicial o lo convierta en notación científica).
  - **`Distancia (km)`**: lo lejos que está de verdad del centro de esa zona. Si ves
    un valor alto, es que el radio de esa zona se queda corto.
  - Anchos de columna ya ajustados y sin líneas de cuadrícula.
- **`Descartados`** (solo si hay algo que descartar): los resultados que Google
  devolvió para una zona pero estaban fuera de su radio, con la distancia y el
  motivo. Sirve para auditar el filtro.

Nombres de zona como `Riba-roja de Túria` o `L'Eixample` dan pestañas válidas; si
una zona tuviera caracteres prohibidos por Excel (`: \ / ? * [ ]`) o superase los
31 caracteres, se sanea automáticamente.

---

## 🎯 Por qué cada zona tiene solo sus inmobiliarias

Esto es **lo más importante de entender** de la herramienta.

Google Maps **no respeta la zona que buscas**. Para la misma consulta
`inmobiliarias+en+carpesa+valencia` devuelve cosas muy distintas según su estado
interno. Medido en este proyecto:

| Consulta (idéntica)                | Resultados | Más lejano |
| ---------------------------------- | ---------- | ---------- |
| Primera vez                        | 4          | 2,8 km     |
| Justo después de buscar "valencia" | 117        | **709 km** |

Ese resultado a 709 km es un negocio de Murcia. Es decir: sin filtro, la pestaña de
"Carpesa" acababa con **102 inmobiliarias**, la mayoría de otros barrios.

Se probó y **no funciona** intentar que Google se porte bien: forzar el viewport en
la URL (`@lat,lng,zoom`) lo ignora, y los metadatos de la página dan coordenadas de
Galicia. Lo único fiable es que **cada resultado trae sus coordenadas reales**, así
que la herramienta filtra por distancia.

### Cómo se decide qué es "de Carpesa"

1. **Centro y radio por zona** en `zonas_valencia.json`. El centro sale de
   OpenStreetMap (una sola vez, queda cacheado) y el radio se deduce del **tamaño
   real** de la zona: mitad de su extensión mayor, entre 1,2 y 8 km.
   - Carpesa (1,7 × 2,8 km) → radio **1,4 km**
   - Ruzafa (1,6 × 1,4 km) → radio **1,2 km**
   - Torrent (15,6 km) → radio **7,8 km**
   - Llíria (22 km) → radio **8 km** (tope)

   Un radio fijo no vale: **3,5 km alrededor de Carpesa** incluye Rascanya,
   Benicalap, Poble Nou y Benimàmet, y la hoja acababa con 70 filas de otros barrios.

2. **Filtro al vuelo**: los resultados fuera del radio se descartan **antes** de
   abrir su ficha, así que además de limpiar los datos ahorran casi todo el tiempo.
   Carpesa pasó de 112 fichas y 6 minutos a 2 fichas y 13 segundos.

3. **Reparto por cercanía**: cada inmobiliaria se asigna a la zona cuyo centro tiene
   más cerca, no a la búsqueda que la encontró. Así, si buscas Carpesa y Rascanya a
   la vez, Cercasa (que está en Rascanya, a 2,8 km de Carpesa) aparece en la hoja de
   **Rascanya**, que es donde le corresponde.

4. **Zonas nuevas**: si escribes una zona que no está en el fichero, se geolocaliza
   sola en ese momento y queda guardada. Si no hubiera cobertura de red, avisa en el
   log y esa zona va sin filtro.

> **Consecuencia práctica:** la hoja de Carpesa tiene **2 filas**, no las 5 que
> devolvía Google. Las otras 3 están en Rascanya y Torrefiel, que **también son
> zonas de la lista**, así que no se pierden: salen en su pestaña. Si prefieres
> radios más generosos, súbelo en `zonas_valencia.json` o ajusta `RADIO_MIN_KM` /
> `FACTOR_RADIO` en `generar_zonas.py` y ejecuta `--recalcular-radios`.

### Si quieres retocar los radios

```bash
# Ver y ajustar un radio a mano (edita zonas_valencia.json directamente)
# o recalcular todos con otras reglas, sin volver a consultar la red:
.venv/bin/python generar_zonas.py --recalcular-radios

# Geolocalizar zonas nuevas (o todas, si borras el fichero)
.venv/bin/python generar_zonas.py "Mi Zona Nueva" "Otra Zona"
```

## Paso 6 · Parar el servidor

En la terminal donde lo lanzaste: `Ctrl + C`.

---

## ⏱️ ¿Cuánto va a tardar?

Cada inmobiliaria es **una visita a su ficha**, y eso cuesta ~2-4 s. Tiempos medidos
con el filtro geográfico ya activo:

| Zona               | Dentro del radio | Descartados | Tiempo |
| ------------------ | ---------------- | ----------- | ------ |
| Carpesa            | 2                | 3           | 13 s   |
| El Palmar          | 3                | 7           | 13 s   |
| Carpesa + Rascanya | 38 únicas        | 471         | 2 min  |

**El filtro es lo que hace esto viable:** descarta los resultados lejanos **antes** de
abrir sus fichas, así que una búsqueda que devuelve 500 resultados se resuelve en
minutos en lugar de horas. Carpesa, que antes tardaba 6 minutos en procesar 112
fichas ajenas, ahora tarda 13 segundos.

**Cálculo rápido:** ~**3 s por inmobiliaria que sí entra**. Antes de lanzar las 68
zonas de golpe, ten en cuenta que los barrios céntricos (Ruzafa, El Carmen,
Benimaclet) tienen muchas dentro de su radio y pueden tardar **5-15 minutos cada uno**.

> ⚠️ **Estrategia recomendada:** ir por bloques (primero los barrios de Valencia,
> luego el área metropolitana), **descargando el Excel al final de cada bloque**.
>
> **Cada ejecución sobrescribe `data_temp.json`.** Si descargas solo al final de todo,
> el Excel tendrá todo. Si vas por bloques, descarga cada bloque antes de lanzar el
> siguiente o perderás los anteriores.
>
> 💡 **Cuantas más zonas juntas en una misma ejecución, mejor reparto.** Al buscar
> zonas vecinas a la vez, las inmobiliarias limítrofes se colocan en la hoja que les
> corresponde por cercanía en lugar de quedarse en la primera que las encontró.

> ℹ️ **Sobre los resultados "raros":** Google no respeta la coletilla "valencia" de
> la búsqueda, y de ahí venían los resultados de Murcia o Almería (hasta 709 km). Con
> el filtro ya no se cuelan; los descartados quedan en la pestaña `Descartados` por si
> quieres revisarlos.

---

## 🔧 Problemas típicos

| Síntoma                                     | Qué hacer                                                                  |
| ------------------------------------------- | -------------------------------------------------------------------------- |
| `❌ Ya hay un scraping en curso.`           | Hay una ejecución viva. La consola se engancha sola a su traza: espérala.  |
| Una hoja tiene menos filas de las esperadas | Es correcto: solo entran las que están dentro del radio de esa zona.       |
| Una hoja sale vacía                         | Su radio es pequeño y no hay nada dentro. Súbelo en `zonas_valencia.json`. |
| Una inmobiliaria está en la hoja equivocada | Ajusta el radio; se asigna por cercanía al centro de cada zona.            |
| Quiero ver qué se descartó y por qué        | Pestaña `Descartados` del Excel.                                           |
| El log avisa de una zona sin filtro         | No se pudo geolocalizar. Añádela con `generar_zonas.py` (ver arriba).      |
| La zona sale en ámbar con `FALLO`           | Google cortó las peticiones. Espera unos minutos y prueba con menos zonas. |
| `Sin datos para descargar`                  | Todavía no has ejecutado ninguna recopilación.                             |
| Muchas líneas `FALLO: TimeoutError`         | Google va lento o te está limitando. Reduce las zonas por bloque.          |
| Quiero menos ruido en la consola            | Sube `CICLOS_SIN_NOVEDAD` en `app.py`, o simplemente ignóralo.             |
| El tiempo restante baila mucho              | Normal en las primeras zonas. Se afina según van cerrando zonas.           |
| `[Errno 48] Address already in use`         | Ya hay un servidor en el puerto. `pkill -f "uvicorn app:app"` y reintenta. |

---

## 🧱 Qué hace por dentro (resumen)

1. El textarea se envía a `POST /api/scrape` como lista de zonas.
2. Playwright abre `https://www.google.com/maps/search/inmobiliarias+en+<ZONA>+valencia`,
   espera al panel de resultados, hace scroll dentro de él y abre cada ficha para leer
   **nombre, dirección, teléfono y web**.
3. Deduplica por **teléfono o nombre normalizado** (las zonas colindantes devuelven las
   mismas inmobiliarias).
4. Guarda en `data_temp.json` **al terminar cada zona**: si algo falla a mitad, no
   pierdes lo anterior.
5. `GET /api/download` agrupa por barrio y genera el `.xlsx` en memoria (una pestaña
   por zona + `Resumen`) y lo descarga.

- **Backend:** FastAPI, todo en `app.py`.
- **Frontend:** `templates/index.html` (Tailwind por CDN + Vanilla JS, cero frameworks).
- **Scraping:** Playwright (Chromium headless).
- **Exportación:** OpenPyXL (generación y formato directo, sin pandas).

### Endpoints

| Método | Ruta                 | Descripción                                              |
| ------ | -------------------- | -------------------------------------------------------- |
| GET    | `/`                  | Interfaz web.                                            |
| POST   | `/api/scrape`        | Recorre las zonas, deduplica y guarda el JSON.           |
| GET    | `/api/download`      | Devuelve el Excel (Resumen + una hoja por barrio).       |
| GET    | `/api/data`          | Datos guardados, para la vista previa.                   |
| GET    | `/api/log?desde=N`   | Líneas de log nuevas desde el índice N.                  |
| GET    | `/api/log/descargar` | Log completo en `.txt` (incluye ejecuciones anteriores). |
| GET    | `/api/estado`        | Si hay una recopilación en curso.                        |
| GET    | `/api/progreso`      | Progreso, porcentaje y ETA para la barra del frontend.   |

---

## ⚙️ Ajustes rápidos

Constantes al inicio de `app.py`:

| Constante              | Por defecto | Para qué sirve                                       |
| ---------------------- | ----------- | ---------------------------------------------------- |
| `MAX_CICLOS_SCROLL`    | `40`        | Tope de scroll dentro del panel por zona.            |
| `ESPERA_SCROLL_MS`     | `1200`      | Pausa tras cada scroll (carga dinámica).             |
| `CICLOS_SIN_NOVEDAD`   | `3`         | Corta cuando N ciclos seguidos no traen resultados.  |
| `ESPERA_DIRECCION_MS`  | `4000`      | Espera máxima a que la ficha muestre la dirección.   |
| `RECICLAR_PAGINA_CADA` | `25`        | Recrea la pestaña cada N zonas (higiene de memoria). |
| `MAX_LINEAS_LOG`       | `20000`     | Líneas retenidas en memoria para la consola.         |
| `SEL_*`                | —           | Selectores del DOM de Google Maps.                   |

Aspecto del Excel (también al inicio de `app.py`): `AZUL` (cabeceras), `AZUL_BANDA`
(filas alternas), `ANCHOS_COLUMNA` y `CARACTERES_INVALIDOS_HOJA`.

Archivos generados (ignorados por git, se pueden borrar sin miedo):

- `data_temp.json` — los datos de la última ejecución.
- `scraping.log` — log en texto; rota cada 5 MB con 2 backups.

---

## 🔁 Reinstalar desde cero

Si `.venv` se rompe o te llevas el proyecto a otro Mac:

```bash
cd ~/Desktop/Everything/Projects/Other/inmo-scrap
rm -rf .venv
/opt/homebrew/bin/python3.14 -m venv .venv      # ⚠️ usa python3.14, NO python3
.venv/bin/python -m pip install --upgrade pip
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python -m playwright install chromium
```

> **Importante:** en este Mac, `/opt/homebrew/bin/python3` crea entornos virtuales
> corruptos (mezcla Python 3.9 y 3.14). Usa siempre el binario **`python3.14`**.

---

## ⚠️ Advertencias

- **Sobrescritura:** cada ejecución reemplaza `data_temp.json`. Descarga el Excel antes
  de lanzar la siguiente si te interesa conservarlo.
- **Deduplicación:** la clave es el teléfono o el nombre normalizado, así que dos
  franquicias con nombre genérico en municipios distintos podrían colapsar en una. Si
  ocurre, cambia la clave a `nombre + barrio` en `deduplicar()`.
- **Legal:** el scraping de Google Maps va contra sus Términos de Servicio. Úsalo como
  herramienta interna puntual, no como extracción masiva recurrente.
