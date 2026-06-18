# Good Fit Test — Solver

## Instalación (una sola vez)

```bash
pip install openpyxl pandas openai python-dotenv
```

## Configuración

Editá el archivo `.env` en la misma carpeta del script y ponés tu API key:

```
OPENAI_API_KEY=sk-...tu_key_aqui...
```

Sin la key el script igual corre usando extracción determinística (regex/keywords).

OPENAI_LOG=1 python3 solve_cinesis_test.py cinesis_good_fit_test_clean output/cinesis_good_fit_test_output.xlsx
OPENAI_LOG=1 python3 solve_cinesis_test.py Tulsa_Reefer_cinesis_good_fit_test_clean output/Tulsa_Reefer_cinesis_good_fit_test_output.xlsx
OPENAI_LOG=1 python3 solve_cinesis_test.py VAN_cinesis_good_fit_test_clean output/VAN_cinesis_good_fit_test_output.xlsx
OPENAI_LOG=1 python3 solve_cinesis_test.py Hotshot_cinesis_good_fit_test_clean output/Hotshot_cinesis_good_fit_test_output.xlsx
OPENAI_LOG=1 python3 solve_cinesis_test.py Flatbed_cinesis_good_fit_test_clean output/Flatbed_cinesis_good_fit_test_output.xlsx
---

## Uso

```bash
python3 solve_cinesis_test.py <archivo_entrada> <archivo_salida>
```

- La extensión `.xlsx` en el input es opcional — el script la agrega si falta.
- La carpeta de output se crea automáticamente si no existe.
- Si omitís los argumentos usa `cinesis_good_fit_test_clean.xlsx` y `cinesis_good_fit_test_completed.xlsx` por defecto.

### Con logs de OpenAI visibles en consola

```bash
OPENAI_LOG=1 python3 solve_cinesis_test.py <entrada> <salida>
```

Activa la impresión del prompt completo enviado a OpenAI y la respuesta JSON cruda (con tokens usados y finish_reason) tanto en consola como en el archivo de log.

---

## Ejemplos reales

```bash
# Conversación original del test
OPENAI_LOG=1 python3 solve_cinesis_test.py \
  cinesis_good_fit_test_clean \
  output/cinesis_good_fit_test_output.xlsx

# Driver con equipo Reefer saliendo de Tulsa
OPENAI_LOG=1 python3 solve_cinesis_test.py \
  Tulsa_Reefer_cinesis_good_fit_test_clean \
  output/Tulsa_Reefer_cinesis_good_fit_test_output.xlsx

# Driver con Van
OPENAI_LOG=1 python3 solve_cinesis_test.py \
  VAN_cinesis_good_fit_test_clean \
  output/VAN_cinesis_good_fit_test_output.xlsx

# Driver Hotshot
OPENAI_LOG=1 python3 solve_cinesis_test.py \
  Hotshot_cinesis_good_fit_test_clean \
  output/Hotshot_cinesis_good_fit_test_output.xlsx

# Driver Flatbed
OPENAI_LOG=1 python3 solve_cinesis_test.py \
  Flatbed_cinesis_good_fit_test_clean \
  output/Flatbed_cinesis_good_fit_test_output.xlsx
```

---

## Logs

Cada ejecución genera automáticamente un archivo en la carpeta `logs/`:

```
logs/log_20260618_181600.log
```

El nombre incluye fecha y hora (`YYYYMMDD_HHMMSS`) para que cada run quede registrado por separado sin sobreescribirse.

**Qué contiene el log:**
- Todo lo que se ve en consola (timestamps, perfil extraído, skips, rechazos, ranking final)
- Si `OPENAI_LOG=1`: el system prompt, el user message completo (transcripción + tabla de ciudades), y la respuesta JSON de OpenAI con tokens usados

El archivo de log guarda **siempre todo en nivel DEBUG**, independientemente de si `OPENAI_LOG` está activo o no en consola.

---

## Cómo funciona el script

### Estructura del workbook de entrada

El `.xlsx` debe tener estos tabs:

| Tab | Contenido |
|---|---|
| `Sample Conversation` | Transcripción del diálogo driver–dispatch (columnas Speaker / Dialogue) |
| `Loads` | Tabla de cargas disponibles con origen, destino, coordenadas, trailer, peso, precio |
| `Part A (Fill In)` | Template donde el script escribe el perfil extraído |
| `Part B (Fill In)` | Template donde el script escribe el top 3 de cargas |

### Part A — Extracción del perfil

El script lee la transcripción del tab `Sample Conversation` y extrae un objeto JSON estructurado con evidencia y confianza para cada campo:

| Campo | Cómo se extrae |
|---|---|
| `current_location` | Frase del driver tipo "I'm in Dallas" — solo líneas del driver |
| `home_base` | Dispatch menciona ciudad → driver confirma ("Yes, that's correct / usually in that area") |
| `minimum_rate_per_mile` | Frase del driver con "per mile" — solo líneas del driver para no confundir con las cotizaciones del dispatcher |
| `equipment_types` | Solo líneas del driver que contienen marcadores de posesión ("I run", "I drive", "I have") — excluye preguntas retóricas como "¿Do y'all deal with hotshots?" |
| `weight_capacity_lb` | Si el driver lo dice explícitamente se usa ese valor. Si no, se infiere del equipo (`inferred: true`) con una nota de la suposición |
| `constraints` / `notes` | Preferencias de lanes, factoring, días de trabajo, etc. |

**Extracción con OpenAI (si hay API key):** usa GPT-4o con structured outputs y un JSON Schema estricto. El modelo recibe la transcripción más la tabla de ciudades con coordenadas.

**Fallback determinístico (sin API key):** regex y keywords sobre las líneas del driver. Produce el mismo objeto JSON con los mismos campos.

Después de cualquier extracción, el script rellena coordenadas nulas usando la tabla de ciudades local (sin APIs externas).

### Part B — Filtrado y ranking

La fórmula usada es la que indica el tab `Part B (Fill In)` del workbook:

```
effective_rate_per_mile = price ÷ (deadhead_to_origin + loaded_miles + deadhead_home)
```

Todas las distancias se calculan con la fórmula **Haversine** (distancia en línea recta en millas, *great-circle*).

#### Por qué se excluyen las filas con MISSING

Las filas `L06` y `L07` tienen campos marcados como `MISSING` en el workbook:

- **L06:** el precio está `MISSING` → sin precio no se puede calcular el effective rate, la fórmula divide por algo desconocido. Se registra el motivo y se omite.
- **L07:** el destino y sus coordenadas están `MISSING` → sin destino no se puede calcular `loaded_miles` ni `deadhead_home`. Se registra el motivo y se omite.

El script también descarta filas donde esos valores son vacíos, `NaN`, o no numéricos — no sólo el string literal `"MISSING"` — para ser robusto ante cualquier variante de dato ausente.

#### Filtros aplicados antes del ranking (en orden)

**1. Tipo de trailer (equipment)**

El tipo de trailer de la carga se normaliza a forma canónica y se compara contra el equipo extraído del perfil:

| Input en el workbook | Forma canónica |
|---|---|
| `"hot shot"`, `"hot-shot"` | `hotshot` |
| `"goose neck"`, `"goose-neck"` | `gooseneck` |
| `"flat bed"`, `"flat-bed"` | `flatbed` |
| `"dry van"`, `"dry-van"` | `van` |

Se usa **interpretación conservadora**: Flatbed no se asume compatible con Hotshot/Gooseneck a menos que el driver lo diga explícitamente. Van nunca matchea Hotshot.

Las cargas rechazadas por este filtro muestran la forma normalizada para facilitar el debug:
```
REJECT L05: Trailer 'Flatbed' (→ 'flatbed') not in driver equipment ['Hotshot', 'Gooseneck'] (→ ['hotshot', 'gooseneck'])
```

**2. Peso**

`load.weight ≤ weight_capacity_lb` extraído del perfil.  
Si el weight capacity es `null` (no se pudo extraer ni inferir), la carga se rechaza con una nota explicando que no se puede verificar el peso.

**3. Effective rate mínimo**

`effective_rate ≥ minimum_rate_per_mile` extraído del perfil.  
Si el mínimo es `null` (no mencionado en la transcripción), este filtro no se aplica y se deja constancia en el log.

#### Ranking

Las cargas que pasan los tres filtros se ordenan de mayor a menor effective rate. Las top 3 se escriben en el tab `Part B (Fill In)` con el rate formateado como número flotante con 3 decimales y formato Excel `$#,##0.000` (para evitar ambigüedad entre separador decimal y de miles según la configuración regional).

---

## Resultados — conversación original

### Part A

| Campo | Valor |
|---|---|
| Current Location | Dallas, TX (32.7767, −96.797) |
| Home Base | San Antonio, TX (29.4241, −98.4936) |
| Min Rate/Mile | $2.00 |
| Equipment | Hotshot, Gooseneck |
| Weight Capacity | 15,000 lb *(inferred)* |

### Part B — Top 3

| Rank | Load | Ruta | Effective Rate | Detalle |
|---|---|---|---|---|
| 1 | L03 | Austin → Corpus Christi | **$3.098/mi** | DH 182 + 172 loaded + 130 DH-home = 484 mi |
| 2 | L08 | Dallas → McAllen | **$2.480/mi** | DH 0 + 462 loaded + 223 DH-home = 685 mi |
| 3 | L02 | Houston → Laredo | **$2.418/mi** | DH 225 + 293 loaded + 144 DH-home = 662 mi |

### Skipped (datos incompletos)

| Carga | Motivo |
|---|---|
| L06 | `Price ($)` = MISSING — sin precio no se puede calcular effective rate |
| L07 | Destination, Dest Lat, Dest Lon = MISSING — sin destino no se puede calcular la ruta |

### Rechazadas (inelegibles)

| Carga | Motivo |
|---|---|
| L01 | Trailer `Van` — no matchea equipo del driver |
| L04 | Trailer `Van` — no matchea equipo del driver (aunque el precio es $1,500, igual al #1) |
| L05 | Trailer `Flatbed` — excluida por interpretación conservadora |
