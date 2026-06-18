# Cinesis Good Fit Test — Solution

## Usage

```bash
pip install openpyxl pandas openai   # openai is optional
python solve_cinesis_test.py [input.xlsx]
```

Output: `cinesis_good_fit_test_completed.xlsx`

Set `OPENAI_API_KEY` in your environment to use GPT-4o structured outputs for
profile extraction. Without it the script uses deterministic keyword extraction.
The script produces the same output structure either way.

---

## Design

### Generic and reusable

No hardcoded answers anywhere. Every value used in Part B — current location
coordinates, home base coordinates, equipment list, weight capacity, minimum
rate — flows from the extracted profile dict. Replace the `Sample Conversation`
tab with any driver transcript and the script re-derives everything.

### Profile JSON structure

Extraction produces a structured JSON object with evidence and confidence for
every field:

```json
{
  "current_location": { "city_state", "lat", "lon", "evidence", "confidence" },
  "home_base":        { "city_state", "lat", "lon", "evidence", "confidence" },
  "minimum_rate_per_mile": { "value", "evidence", "confidence" },
  "equipment_types":  { "value": ["Hotshot", "Gooseneck"], "evidence", "confidence" },
  "weight_capacity_lb": { "value", "evidence", "confidence", "inferred", "assumption" },
  "constraints": [ { "type", "value", "evidence" }, ... ],
  "notes": [ ... ]
}
```

### City coordinate lookup

A seed table of ~40 cities is built into the script. At startup it is
supplemented with every city found in the workbook's Loads tab (with its
coordinates from that tab). No external APIs are called at runtime. If a
transcript city is not in the table, lat/lon are set to null and Part B
records the reason it cannot calculate those loads.

### Equipment normalization

A canonical-form alias map handles spacing/hyphenation variants:

| Input | Canonical |
|---|---|
| `"hot shot"`, `"hot-shot"` | `hotshot` |
| `"goose neck"`, `"goose-neck"` | `gooseneck` |
| `"flat bed"`, `"flat-bed"` | `flatbed` |
| `"dry van"`, `"dry-van"` | `van` |

Matching is done on canonical forms so `"Goose Neck"` in a future load board
will still match a driver who says `"gooseneck"`.

### Extraction heuristics (deterministic fallback)

| Field | Source | Key decision |
|---|---|---|
| Current location | `"I'm in Dallas"` — driver lines only | Regex `i'm in <city>` |
| Home base | Dispatch asks about San Antonio; driver confirms | Pattern match + "usually"/"that's correct" confirmation |
| Min rate | `"above $2 per mile"` — driver lines only | Regex requires `per mile` suffix to avoid matching earnings like `$2,300` |
| Equipment | `"I run a hotshot gooseneck trailer"` | Only ownership lines (`"I run"`, `"I drive"`, etc.) — rhetorical questions excluded |
| Weight capacity | Not stated | Inferred as 15,000 lb for hotshot/gooseneck; `inferred=true`, `assumption` documented |

### Part B — load ranking formula

```
effective_rate = price / (deadhead_to_origin + loaded_miles + deadhead_home)
```

All distances use **Haversine great-circle miles**.

Filters applied in order before ranking:
1. Trailer type must match driver equipment (after normalization)
2. Load weight ≤ driver weight capacity (null capacity → reject with reason)
3. Effective rate ≥ driver minimum rate (null minimum → no rate filter, noted)

---

## Results (sample transcript)

### Part A — Driver profile

| Field | Value |
|---|---|
| Current Location | Dallas, TX (32.7767, −96.797) |
| Home Base | San Antonio, TX (29.4241, −98.4936) |
| Min Rate/Mile | $2.00 |
| Equipment | Hotshot, Gooseneck |
| Weight Capacity | 15,000 lb *(inferred)* |

### Part B — Top 3 loads

| Rank | Load | Route | Effective Rate | Breakdown |
|---|---|---|---|---|
| 1 | L03 | Austin → Corpus Christi | **$3.098/mi** | DH 182 + 172 loaded + 130 DH-home = 484 mi |
| 2 | L08 | Dallas → McAllen | **$2.480/mi** | DH 0 + 462 loaded + 223 DH-home = 685 mi |
| 3 | L02 | Houston → Laredo | **$2.418/mi** | DH 225 + 293 loaded + 144 DH-home = 662 mi |

### Skipped (incomplete data)

| Load | Reason |
|---|---|
| L06 | Price = MISSING |
| L07 | Destination and coordinates = MISSING |

### Notable ineligible load rejected

**L04** (Plano → Memphis, Van, $1,500) — `Van` trailer does not match driver
equipment `['Hotshot', 'Gooseneck']`. The price is identical to the #1 load
but the equipment filter eliminates it before the rate is even evaluated.

**L05** (Waco → San Antonio, Flatbed, $640) — `Flatbed` is excluded under the
conservative interpretation: the driver said *"I run a hotshot gooseneck
trailer"* and the script does not assume Flatbed compatibility unless the
driver explicitly claims it or the transcript clearly implies it.
