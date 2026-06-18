# Good Fit Test — Solver

## Installation (once)

```bash
pip install openpyxl pandas openai python-dotenv
```

## Configuration

Edit the `.env` file in the same folder as the script and add your API key:

```
OPENAI_API_KEY=sk-...your_key_here...
```

Without the key the script still runs using deterministic extraction (regex/keywords).

---

## Usage

```bash
python3 solve_cinesis_test.py <input_file> <output_file>
```

- The `.xlsx` extension on the input is optional — the script appends it if missing.
- The output folder is created automatically if it does not exist.
- If arguments are omitted, defaults to `good_fit_test_clean.xlsx` → `good_fit_test_completed.xlsx`.

### With full OpenAI logs printed to console

```bash
OPENAI_LOG=1 python3 solve_cinesis_test.py <input> <output>
```

Prints the full prompt sent to OpenAI and the raw JSON response (including token usage and finish_reason) to both console and log file.

---

## Examples

```bash
# Original test conversation
OPENAI_LOG=1 python3 solve_cinesis_test.py \
  good_fit_test_clean \
  output/good_fit_test_output.xlsx

# Reefer driver out of Tulsa
OPENAI_LOG=1 python3 solve_cinesis_test.py \
  Tulsa_Reefer_good_fit_test_clean \
  output/Tulsa_Reefer_good_fit_test_output.xlsx

# Van driver
OPENAI_LOG=1 python3 solve_cinesis_test.py \
  VAN_good_fit_test_clean \
  output/VAN_good_fit_test_output.xlsx

# Hotshot driver
OPENAI_LOG=1 python3 solve_cinesis_test.py \
  Hotshot_good_fit_test_clean \
  output/Hotshot_good_fit_test_output.xlsx

# Flatbed driver
OPENAI_LOG=1 python3 solve_cinesis_test.py \
  Flatbed_good_fit_test_clean \
  output/Flatbed_good_fit_test_output.xlsx
```

---

## Logs

Each run automatically creates a file in the `logs/` folder:

```
logs/log_20260618_181600.log
```

The filename includes date and time (`YYYYMMDD_HHMMSS`) so each run is saved separately without overwriting.

**What the log contains:**
- Everything printed to console (timestamps, extracted profile, skips, rejections, final ranking)
- If `OPENAI_LOG=1`: the full system prompt, the complete user message (transcript + city table), and the OpenAI JSON response with token usage

The log file always saves everything at DEBUG level regardless of whether `OPENAI_LOG` is active in the console.

---

## How it works

### Input workbook structure

The `.xlsx` must have these tabs:

| Tab | Content |
|---|---|
| `Sample Conversation` | Driver–dispatch dialogue transcript (Speaker / Dialogue columns) |
| `Loads` | Available loads with origin, destination, coordinates, trailer type, weight, price |
| `Part A (Fill In)` | Template where the script writes the extracted driver profile |
| `Part B (Fill In)` | Template where the script writes the top 3 loads |

### Part A — Profile extraction

The script reads the transcript from the `Sample Conversation` tab and extracts a structured JSON object with evidence and confidence for every field:

| Field | How it is extracted |
|---|---|
| `current_location` | Driver phrase like "I'm in Dallas" — driver lines only |
| `home_base` | Dispatch mentions a city → driver confirms ("Yes, that's correct / usually in that area") |
| `minimum_rate_per_mile` | Driver phrase containing "per mile" — driver lines only to avoid matching dispatcher rate quotes |
| `equipment_types` | Driver lines containing ownership markers ("I run", "I drive", "I have") — excludes rhetorical questions like "Do y'all deal with hotshots?" |
| `weight_capacity_lb` | Used directly if the driver states it explicitly. Otherwise inferred from equipment (`inferred: true`) with a documented assumption |
| `constraints` / `notes` | Lane preferences, factoring requirements, working schedule, etc. |

**With OpenAI (API key set):** uses GPT-4o with structured outputs and a strict JSON Schema. The model receives the transcript and returns coordinates directly — GPT-4o knows the coordinates of any city in the world. No external mapping APIs are called.

**Deterministic fallback (no API key):** regex and keyword matching over driver lines. Produces the same JSON object with the same fields. A local city coordinate table is used as backup to fill in any null lat/lon values.

### Part B — Filtering and ranking

The formula used is the one stated in the `Part B (Fill In)` tab of the workbook:

```
effective_rate_per_mile = price ÷ (deadhead_to_origin + loaded_miles + deadhead_home)
```

All distances are calculated using the **Haversine** formula (straight-line great-circle miles).

#### Why MISSING rows are excluded

Some rows in the loads table have fields marked as `MISSING`:

- **No price:** the effective rate cannot be calculated (the formula needs the numerator). The reason is recorded and the row is skipped.
- **No destination or destination coordinates:** `loaded_miles` and `deadhead_home` cannot be calculated. The reason is recorded and the row is skipped.

The script also discards rows where those values are empty, `NaN`, or non-numeric — not just the literal string `"MISSING"` — to handle any variant of missing data.

#### Filters applied before ranking (in order)

**1. Trailer type (equipment)**

The load's trailer type is normalized to canonical form and compared against the equipment extracted from the profile:

| Value in workbook | Canonical form |
|---|---|
| `"hot shot"`, `"hot-shot"` | `hotshot` |
| `"goose neck"`, `"goose-neck"` | `gooseneck` |
| `"flat bed"`, `"flat-bed"` | `flatbed` |
| `"dry van"`, `"dry-van"` | `van` |

**Conservative interpretation:** Flatbed is not assumed compatible with Hotshot/Gooseneck unless the driver explicitly says so. Van never matches Hotshot.

Rejected loads show the normalized form to make debugging easier:
```
REJECT L05: Trailer 'Flatbed' (→ 'flatbed') not in driver equipment ['Hotshot', 'Gooseneck'] (→ ['hotshot', 'gooseneck'])
```

**2. Weight**

`load.weight ≤ weight_capacity_lb` extracted from the profile.
If weight capacity is `null` (could not be extracted or inferred), the load is rejected with a note explaining that weight cannot be verified.

**3. Minimum effective rate**

`effective_rate ≥ minimum_rate_per_mile` extracted from the profile.
If the minimum is `null` (not mentioned in the transcript), this filter is not applied and the assumption is recorded in the log.

#### Ranking

Loads that pass all three filters are sorted from highest to lowest effective rate. The top 3 are written to the `Part B (Fill In)` tab as floating-point numbers with 3 decimal places and Excel number format `$#,##0.000` (to avoid ambiguity between decimal and thousands separators across regional settings).

---

## Sample results — original conversation

### Part A

| Field | Value |
|---|---|
| Current Location | Dallas, TX (32.7767, −96.797) |
| Home Base | San Antonio, TX (29.4241, −98.4936) |
| Min Rate/Mile | $2.00 |
| Equipment | Hotshot, Gooseneck |
| Weight Capacity | 15,000 lb *(inferred)* |

### Part B — Top 3

| Rank | Load | Route | Effective Rate | Breakdown |
|---|---|---|---|---|
| 1 | L03 | Austin → Corpus Christi | **$3.098/mi** | DH 182 + 172 loaded + 130 DH-home = 484 mi |
| 2 | L08 | Dallas → McAllen | **$2.480/mi** | DH 0 + 462 loaded + 223 DH-home = 685 mi |
| 3 | L02 | Houston → Laredo | **$2.418/mi** | DH 225 + 293 loaded + 144 DH-home = 662 mi |

### Skipped (incomplete data)

| Load | Reason |
|---|---|
| L06 | `Price ($)` = MISSING — cannot calculate effective rate without price |
| L07 | Destination, Dest Lat, Dest Lon = MISSING — cannot calculate route without destination |

### Rejected (ineligible)

| Load | Reason |
|---|---|
| L01 | Trailer `Van` — does not match driver equipment |
| L04 | Trailer `Van` — does not match driver equipment (price is $1,500, same as #1, but equipment filter eliminates it) |
| L05 | Trailer `Flatbed` — excluded by conservative interpretation |
