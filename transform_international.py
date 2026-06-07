import json
import re
from pathlib import Path
from openpyxl import Workbook

INPUT_FILE = Path("processing/fedex international 4.pdf_extracted.json")
OUTPUT_DIR = Path("output")
OUTPUT_DIR.mkdir(exist_ok=True)

# --- Configuration ---

CARRIER_MAP = {
    "Germany": ("FEDEX DE", "DE"),
    "Spain": ("FEDEX ES", "ES"),
    "Hungary": ("FEDEX HU", "HU"),
    "Poland": ("FEDEX PL", "PL"),
    "Romania": ("FEDEX RO", "RO"),
    "Austria": ("TNT AT", "AT"),
    "Portugal": ("TNT EXP PT", "PT"),
    "France": ("TNT FR (P01)", "FR"),
    "Slovenia": ("TNT SE", "SE"),
    "Slovakia": ("TNT SK", "SK"),
}

IN_SCOPE_COUNTRIES = set(CARRIER_MAP.keys())

CURRENCY = "EUR"

# PartnerCountry# -> Country name (derived from self-zone=0 analysis)
PARTNER_COUNTRY_MAP = {
    1: "Andorra", 2: "Austria", 3: "Belgium", 4: "Bulgaria",
    5: "Canary Islands", 6: "Channel Islands", 7: "Croatia",
    8: "Cyprus", 9: "Czech Republic", 10: "Denmark",
    11: "Estonia", 12: "Finland", 13: "France", 14: "Germany",
    15: "Greece", 16: "Hungary", 17: "Ireland", 18: "Israel",
    19: "Italy", 20: "Latvia", 21: "Liechtenstein", 22: "Lithuania",
    23: "Luxembourg", 24: "Malta", 25: "Monaco", 26: "Netherlands",
    27: "Norway", 28: "Poland", 29: "Portugal", 30: "Rest of Europe",
    31: "Romania", 32: "Russian Federation", 33: "San Marino",
    34: "Slovakia", 35: "Slovenia", 36: "Spain", 37: "Sweden",
    38: "Switzerland", 39: "Turkey", 40: "Ukraine",
    41: "United Kingdom", 42: "Holy See", 43: "North America",
    44: "Latin America", 45: "Asia 1", 46: "Asia 2",
    47: "Pacific", 48: "Pacific 2", 49: "Middle East",
    50: "India", 51: "Africa",
}

COUNTRY_TO_PC = {v: k for k, v in PARTNER_COUNTRY_MAP.items()}

# Country name -> ISO 3166-1 alpha-2 (for CountryRegions output)
COUNTRY_NAME_TO_ISO = {
    "Andorra": "AD", "Austria": "AT", "Belgium": "BE", "Bulgaria": "BG",
    "Canary Islands": "IC", "Channel Islands": "GG", "Croatia": "HR",
    "Cyprus": "CY", "Czech Republic": "CZ", "Denmark": "DK", "Estonia": "EE",
    "Finland": "FI", "France": "FR", "Germany": "DE", "Greece": "GR",
    "Hungary": "HU", "Ireland": "IE", "Israel": "IL", "Italy": "IT",
    "Latvia": "LV", "Liechtenstein": "LI", "Lithuania": "LT", "Luxembourg": "LU",
    "Malta": "MT", "Monaco": "MC", "Morocco": "MA", "Netherlands": "NL",
    "Norway": "NO", "Poland": "PL", "Portugal": "PT", "Romania": "RO",
    "Russian Federation": "RU", "San Marino": "SM", "Slovakia": "SK",
    "Slovenia": "SI", "Spain": "ES", "Sweden": "SE", "Switzerland": "CH",
    "Turkey": "TR", "Ukraine": "UA", "United Kingdom": "GB", "Holy See": "VA",
    "United States": "US", "U.S.A.": "US", "U.S.A": "US", "USA": "US",
    "Canada": "CA", "Mexico": "MX", "Brazil": "BR", "Argentina": "AR",
    "China": "CN", "Japan": "JP", "India": "IN", "Australia": "AU",
    "New Zealand": "NZ", "Singapore": "SG", "Hong Kong SAR (China)": "HK",
    "Taiwan (China)": "TW", "South Africa": "ZA", "Egypt": "EG",
    "Saudi Arabia": "SA", "United Arab Emirates": "AE", "Albania": "AL",
    "Algeria": "DZ", "Afghanistan": "AF", "Belarus": "BY",
    "Bosnia and Herzegovina": "BA", "Faroe Islands": "FO", "Gibraltar": "GI",
    "Kosovo": "XK", "North Macedonia": "MK", "Moldova Republic of": "MD",
    "Montenegro": "ME", "Serbia": "RS", "Chile": "CL", "Indonesia": "ID",
    "Malaysia": "MY", "Philippines": "PH", "Thailand": "TH",
    "Korea Republic of": "KR", "Viet Nam": "VN", "Fiji": "FJ",
    "Bahrain": "BH", "Kuwait": "KW", "Oman": "OM", "Qatar": "QA",
    "American Samoa": "AS", "Cook Islands": "CK", "French Polynesia": "PF",
    "Guam": "GU", "Kiribati": "KI", "Marshall Islands": "MH",
    "Micronesia": "FM", "Nauru": "NR", "New Caledonia": "NC", "Niue": "NU",
    "Northern Mariana Islands": "MP", "Palau": "PW", "Papua New Guinea": "PG",
    "Samoa": "WS", "Solomon Islands": "SB", "Tonga": "TO", "Tuvalu": "TV",
    "Vanuatu": "VU", "Wallis and Futuna": "WF",
}
# Regional group labels used in zoning (not ISO countries)
REGION_NAME_TO_CODE = {
    "Rest of Europe": "RoE",
    "North America": "NA",
    "Latin America": "LA",
    "Asia 1": "AS1",
    "Asia 2": "AS2",
    "Pacific": "PAC",
    "Pacific 2": "PAC2",
    "Middle East": "ME",
    "Africa": "AFR",
    "Rest of World": "RoW",
    "India": "IN",
}

# Expand regional zoning labels into constituent countries (CountryRegions output)
REGION_EXPANSION = {
    "Rest of Europe": [
        "Albania", "Belarus", "Bosnia and Herzegovina", "Faroe Islands",
        "Gibraltar", "Kosovo", "North Macedonia", "Moldova Republic of",
        "Montenegro", "Serbia",
    ],
    "North America": ["Canada", "United States"],
    "Latin America": ["Argentina", "Brazil", "Chile", "Mexico"],
    "Asia 1": [
        "China", "Hong Kong SAR (China)", "Indonesia", "Malaysia",
        "Philippines", "Singapore", "Thailand",
    ],
    "Asia 2": ["Japan", "Korea Republic of", "Taiwan (China)", "Viet Nam"],
    "Pacific": ["Australia", "New Zealand", "Fiji"],
    "Pacific 2": [
        "American Samoa", "Cook Islands", "Fiji", "French Polynesia", "Guam",
        "Kiribati", "Marshall Islands", "Micronesia", "Nauru", "New Caledonia",
        "Niue", "Northern Mariana Islands", "Palau", "Papua New Guinea", "Samoa",
        "Solomon Islands", "Tonga", "Tuvalu", "Vanuatu", "Wallis and Futuna",
    ],
    "India": ["India"],
    "Middle East": [
        "Bahrain", "Kuwait", "Oman", "Qatar", "Saudi Arabia",
        "United Arab Emirates",
    ],
    "Africa": ["South Africa", "Egypt", "Morocco"],
}

REGION_CODE_TO_NAME = {code: name for name, code in REGION_NAME_TO_CODE.items()}


def country_to_iso(name: str) -> str:
    """Resolve a country or region label to a short code for CountryRegions output."""
    if not name:
        return name
    name = name.strip()
    if len(name) == 2 and name.isalpha():
        return name.upper()
    if name in COUNTRY_NAME_TO_ISO:
        return COUNTRY_NAME_TO_ISO[name]
    if name in REGION_NAME_TO_CODE:
        return REGION_NAME_TO_CODE[name]
    for key, code in COUNTRY_NAME_TO_ISO.items():
        if key.lower() == name.lower():
            return code
    return name


def expand_to_country_codes(name_or_code: str) -> list[str]:
    """Return ISO country codes, expanding regional labels when defined."""
    if not name_or_code:
        return []

    region_name = REGION_CODE_TO_NAME.get(name_or_code, name_or_code)
    for key in REGION_EXPANSION:
        if key.lower() == name_or_code.lower():
            region_name = key
            break

    if region_name in REGION_EXPANSION:
        return sorted({country_to_iso(country) for country in REGION_EXPANSION[region_name]})

    return [country_to_iso(name_or_code)]


# --- Morocco (MA) international ---

MA_COUNTRY = "MA"
MA_CARRIER = "FEDEX MA"
MA_CURRENCY = "MAD"
MA_SERVICE_ORDER = ["ENV", "PAK", "PACKAGE"]
MA_SERVICE_INDEX = {name: idx for idx, name in enumerate(MA_SERVICE_ORDER)}


def is_ma_international(fields: dict) -> bool:
    """True when the document uses MainCostsMA / RegionsMA instead of standard fields."""
    return bool(fields.get("MainCostsMA", {}).get("value"))


def extract_zone_number(raw_value: str) -> int | None:
    """Extract zone number from a potentially messy OCR value.
    Takes the last numeric token found.
    """
    if not raw_value:
        return None
    tokens = re.split(r'[\s\n]+', raw_value.strip())
    numbers = []
    for t in tokens:
        try:
            numbers.append(int(t))
        except ValueError:
            pass
    if not numbers:
        return None
    return numbers[-1]


def parse_zoning_matrix(zm_data: list) -> tuple[dict, dict]:
    """Parse ZoningMatrix.
    Returns:
      - zone_map: {origin_country: {dest_country: zone_number}} (in-scope pairs only)
      - all_zones_per_country: {origin_country: sorted list of all used zones}
    """
    zone_map = {}
    full_zone_map = {}
    all_zones_per_country = {}

    for entry in zm_data:
        ic = entry.get("InvoicingCountry", "")
        if ic not in IN_SCOPE_COUNTRIES:
            continue

        in_scope_zones = {}
        used_zones = set()

        all_dest_zones = {}

        for pc_num in range(1, 52):
            key = f"PartnerCountry{pc_num}"
            raw = entry.get(key, "")
            zone = extract_zone_number(raw)
            if zone is None or zone == 0:
                continue

            used_zones.add(zone)
            dest_country = PARTNER_COUNTRY_MAP.get(pc_num)
            if dest_country and dest_country != ic:
                all_dest_zones[dest_country] = zone
                if dest_country in IN_SCOPE_COUNTRIES:
                    in_scope_zones[dest_country] = zone

        zone_map[ic] = in_scope_zones
        full_zone_map[ic] = all_dest_zones
        all_zones_per_country[ic] = sorted(used_zones)

    return zone_map, full_zone_map, all_zones_per_country


def build_country_regions(full_zone_map: dict) -> dict:
    """Build {origin_country: {zone_num: [dest_countries]}} with ALL countries."""
    regions = {}
    for origin, dest_zones in full_zone_map.items():
        by_zone = {}
        for dest, zone_num in dest_zones.items():
            for code in expand_to_country_codes(dest):
                by_zone.setdefault(zone_num, []).append(code)
        for z in by_zone:
            by_zone[z] = sorted(set(by_zone[z]))
        regions[origin] = dict(sorted(by_zone.items()))
    return regions


def write_country_regions_txt(regions: dict, output_path: Path):
    """Write CountryRegions.txt."""
    lines = []
    lines.append("=" * 70)
    lines.append("COUNTRY REGIONS - International Zoning Matrix")
    lines.append("=" * 70)

    for origin in sorted(regions.keys()):
        _, code = CARRIER_MAP[origin]
        lines.append("")
        lines.append(f"--- {origin} ({code}) ---")
        for zone_num, countries in regions[origin].items():
            lines.append(
                f"  {code} Zone {zone_num:02d}: {', '.join(countries)}"
            )

    lines.append("")
    lines.append("=" * 70)
    output_path.write_text("\n".join(lines), encoding="utf-8")
    print(f"  Written: {output_path}")


def parse_main_costs(mc_data: list) -> dict:
    """Parse MainCosts into rate blocks.
    Returns {rate_name: {'weights': [...], 'rates': {zone_num: [...]}}}
    """
    blocks = []
    current_weights = []
    current_rates = {}
    current_name = ""

    for entry in mc_data:
        rate_name = entry.get("RateName", "")
        weight = entry.get("Weight", "")

        if weight == "Weight (kg)":
            if current_weights:
                blocks.append((current_name, current_weights, current_rates))
            current_name = rate_name
            current_weights = []
            current_rates = {}
            continue

        if rate_name:
            current_name = rate_name

        current_weights.append(weight)
        for i in range(1, 16):
            zone_key = f"Zone{i}"
            val = entry.get(zone_key, "")
            if val:
                current_rates.setdefault(i, []).append(val)

    if current_weights:
        blocks.append((current_name, current_weights, current_rates))

    # Assign names
    result = {}
    if len(blocks) == 1:
        name = blocks[0][0] if blocks[0][0] else "Express"
        result[name] = {"weights": blocks[0][1], "rates": blocks[0][2]}
    elif len(blocks) == 2:
        name1 = blocks[0][0] if blocks[0][0] else "Express"
        name2 = blocks[1][0] if blocks[1][0] else "Economy Express"
        result[name1] = {"weights": blocks[0][1], "rates": blocks[0][2]}
        result[name2] = {"weights": blocks[1][1], "rates": blocks[1][2]}
    else:
        for i, (name, weights, rates) in enumerate(blocks):
            key = name if name else f"Block {i+1}"
            result[key] = {"weights": weights, "rates": rates}

    return result


def filter_weights_and_rates(weights: list, rates: dict) -> tuple[list, dict]:
    """Remove regular weight entries between Add-on rows.
    Keep: initial regular weights (before first Add-on) + all Add-on rows.
    For Add-on labels: use the last regular weight before the NEXT Add-on
    (i.e., the end of the segment that follows this Add-on).
    Last Add-on gets "> last_weight".
    Returns (filtered_brackets, filtered_rates).
    """
    first_addon_idx = None
    for idx, w in enumerate(weights):
        if w.lower().startswith("add"):
            first_addon_idx = idx
            break

    if first_addon_idx is None:
        brackets = [(f"<={w}", "Flat") for w in weights]
        return brackets, rates

    # Build brackets for initial regular weights
    brackets = []
    keep_indices = []

    for idx in range(first_addon_idx):
        brackets.append((f"<={weights[idx]}", "Flat"))
        keep_indices.append(idx)

    # Collect all add-on indices
    addon_indices = [i for i, w in enumerate(weights) if w.lower().startswith("add")]

    for pos, addon_idx in enumerate(addon_indices):
        is_last = (pos == len(addon_indices) - 1)

        if is_last:
            # Last add-on: find the last regular weight before it
            last_weight = None
            for j in range(addon_idx - 1, -1, -1):
                if not weights[j].lower().startswith("add"):
                    last_weight = weights[j]
                    break
            brackets.append((f">{last_weight}", "p/unit"))
        else:
            # Find the last regular weight before the NEXT add-on
            next_addon_idx = addon_indices[pos + 1]
            boundary_weight = None
            for j in range(next_addon_idx - 1, -1, -1):
                if not weights[j].lower().startswith("add"):
                    boundary_weight = weights[j]
                    break
            brackets.append((f"<={boundary_weight}", "p/unit"))

        keep_indices.append(addon_idx)

    # Filter rates to only keep the selected indices
    filtered_rates = {}
    for zone_num, values in rates.items():
        filtered_rates[zone_num] = [values[i] for i in keep_indices if i < len(values)]

    return brackets, filtered_rates


def build_lanes(zone_map: dict, all_zones: dict) -> list:
    """Build lane rows sorted by carrier: forward lanes first, then duplicates.
    Forward: Origin Country + Dest Region filled.
    Duplicate: switched — Origin Country value → Destination,
               Dest Region value → Origin Country Region,
               Origin Country and Dest Region become empty.
    """
    lanes = []

    # Group by carrier: forward lanes then duplicates for each
    for origin in sorted(all_zones.keys()):
        carrier, origin_code = CARRIER_MAP[origin]
        zones = all_zones[origin]

        # Forward lanes for this carrier
        for zone_num in zones:
            dest_region = f"{origin_code} Zone {zone_num:02d}"
            lanes.append({
                "carrier": carrier,
                "origin_country": origin_code,
                "origin_region": "",
                "destination": "",
                "dest_region": dest_region,
                "zone_num": zone_num,
                "currency": CURRENCY,
            })

        # Duplicate (switched) lanes for this carrier
        for zone_num in zones:
            dest_region = f"{origin_code} Zone {zone_num:02d}"
            lanes.append({
                "carrier": carrier,
                "origin_country": "",
                "origin_region": dest_region,
                "destination": origin_code,
                "dest_region": "",
                "zone_num": zone_num,
                "currency": CURRENCY,
            })

    return lanes


def write_output(lanes: list, rate_blocks: dict, output_dir: Path):
    """Write all cost blocks into a single xlsx and a single txt file.
    Each cost block appears as consecutive column groups after the shipment details.
    """
    header_cols = [
        "Lane #", "Carrier Name", "Origin Country",
        "Origin Country Region", "Destination",
        "Destination Country Region",
    ]
    num_ship_cols = len(header_cols)

    # Prepare bracket info for each rate block (filter out intermediate weights)
    block_info = []
    for rate_name, block_data in rate_blocks.items():
        brackets, filtered_rates = filter_weights_and_rates(
            block_data["weights"], block_data["rates"]
        )
        block_info.append({
            "name": rate_name,
            "cost_name": f"Transport cost ({rate_name})",
            "brackets": brackets,
            "labels": [b[0] for b in brackets],
            "types": [b[1] for b in brackets],
            "rates": filtered_rates,
        })

    # --- XLSX ---
    wb = Workbook()
    ws = wb.active
    ws.title = "International Rates"

    # Each block occupies: 1 (Currency) + len(labels) columns
    # Row 1: cost names (placed at the Currency column of each block)
    row1 = [""] * num_ship_cols
    for bi in block_info:
        row1.append(bi["cost_name"])
        row1.extend([""] * len(bi["labels"]))
    ws.append(row1)

    # Row 2: Applies if:
    row2 = [""] * num_ship_cols
    for bi in block_info:
        row2.append("Applies if:")
        row2.extend([""] * len(bi["labels"]))
    ws.append(row2)

    # Row 3: Rate by
    row3 = [""] * num_ship_cols
    for bi in block_info:
        row3.append("Rate by: Weight/chargeable kg")
        row3.extend([""] * len(bi["labels"]))
    ws.append(row3)

    # Row 4: Regular + rule
    row4 = [""] * num_ship_cols
    for bi in block_info:
        row4.append("Regular + rule")
        row4.extend([""] * len(bi["labels"]))
    ws.append(row4)

    # Row 5: Rounding rule
    row5 = [""] * num_ship_cols
    for bi in block_info:
        row5.append("Rounding rule:")
        row5.extend([""] * len(bi["labels"]))
    ws.append(row5)

    # Row 6: column headers (shipment details + weight brackets for each block)
    row6 = list(header_cols)
    for bi in block_info:
        row6.append("Currency")
        row6.extend(bi["labels"])
    ws.append(row6)

    # Row 7: value types (Flat / p/unit)
    row7 = [""] * num_ship_cols
    for bi in block_info:
        row7.append("Currency")
        row7.extend(bi["types"])
    ws.append(row7)

    # Data rows
    for lane_idx, lane in enumerate(lanes, 1):
        zone_num = lane["zone_num"]
        row = [
            lane_idx,
            lane["carrier"],
            lane["origin_country"],
            lane["origin_region"],
            lane["destination"],
            lane["dest_region"],
        ]
        for bi in block_info:
            row.append(lane["currency"])
            zone_rates = bi["rates"].get(zone_num, [])
            row.extend(zone_rates)
            expected = len(bi["labels"])
            actual = len(zone_rates)
            if actual < expected:
                row.extend([""] * (expected - actual))
        ws.append(row)

    xlsx_path = output_dir / "international_rates.xlsx"
    wb.save(xlsx_path)
    print(f"  Written: {xlsx_path}")


def _looks_like_zone_letter(value: str) -> bool:
    if not value:
        return False
    token = value.strip().split("\n")[0].strip()
    return len(token) == 1 and token.isalpha()


def _is_half_kg(weight: str) -> bool:
    if not weight:
        return False
    for part in weight.replace(",", ".").split("\n"):
        part = part.strip()
        if not part:
            continue
        try:
            if abs(float(part) - 0.5) < 0.01:
                return True
        except ValueError:
            continue
    return False


def _is_ma_block_header(rate_name: str) -> bool:
    if not rate_name:
        return False
    rn = rate_name.upper()
    if "GRILLE" in rn:
        return True
    if "EXPORT" in rn:
        return True
    if "IMPORT" in rn:
        return True
    return False


def _block_direction(block_rows: list, block_index: int, block_count: int) -> str:
    for entry in block_rows[:5]:
        rn = (entry.get("RateName") or "").upper()
        if "IMPORT" in rn:
            return "import"
        if "EXPORT" in rn:
            return "export"
    if block_count == 2:
        return "export" if block_index == 0 else "import"
    return "export" if block_index == 0 else "import"


def _identify_ma_service(rate_name: str, half_kg_index: int) -> str:
    if rate_name:
        rn = rate_name.upper()
        if "ENV" in rn or rn.strip() == "DOC":
            return "ENV"
        if "PACKAGE" in rn:
            return "PACKAGE"
        if "PAK" in rn or rn.strip() == "PACK":
            return "PAK"
    if half_kg_index < len(MA_SERVICE_ORDER):
        return MA_SERVICE_ORDER[half_kg_index]
    return "PACKAGE"


def _ma_cost_name(service: str, direction: str) -> str:
    if service == "ENV":
        return "Transport cost (ENV/DOC)"
    if service == "PACKAGE":
        tag = "exp" if direction == "export" else "imp"
        return f"Transport cost (PACKAGE ({tag}))"
    return f"Transport cost ({service})"


def _merge_env_pak_ma_blocks(blocks: list[dict]) -> list[dict]:
    """When ENV and PAK share identical brackets, emit one cost block and merge rates."""
    env_block = next((b for b in blocks if b["service"] == "ENV"), None)
    pak_block = next((b for b in blocks if b["service"] == "PAK"), None)
    others = [b for b in blocks if b["service"] not in ("ENV", "PAK")]

    if not env_block or not pak_block:
        return blocks

    if (env_block["labels"] != pak_block["labels"]
            or env_block["types"] != pak_block["types"]):
        return blocks

    merged_rates: dict[int, list[str]] = {}
    for zone_num in range(1, 12):
        env_vals = env_block["rates"].get(zone_num, [])
        pak_vals = pak_block["rates"].get(zone_num, [])
        combined = []
        for idx in range(len(env_block["labels"])):
            env_val = env_vals[idx] if idx < len(env_vals) else ""
            pak_val = pak_vals[idx] if idx < len(pak_vals) else ""
            combined.append(env_val or pak_val)
        merged_rates[zone_num] = combined

    return [{
        **env_block,
        "service": "ENV",
        "cost_name": _ma_cost_name("ENV", env_block["direction"]),
        "rates": merged_rates,
    }, *others]


def _merge_duplicate_ma_cost_blocks(blocks: list[dict]) -> list[dict]:
    """When export/import blocks share cost_name and identical brackets, emit once."""
    groups: dict[tuple, list[dict]] = {}
    order: list[tuple] = []
    for block in blocks:
        key = (block["cost_name"], tuple(block["labels"]), tuple(block["types"]))
        if key not in groups:
            order.append(key)
        groups.setdefault(key, []).append(block)

    merged: list[dict] = []
    for key in order:
        group = groups[key]
        if len(group) == 1:
            merged.append(group[0])
            continue

        directions = {b["direction"] for b in group}
        if len(directions) < 2:
            merged.extend(group)
            continue

        base = group[0]
        merged.append({
            **base,
            "direction": "both",
            "rates_by_direction": {b["direction"]: b["rates"] for b in group},
        })
    return merged


def _zone_rate_value(entry: dict, zone_num: int, part_index: int = 0) -> str:
    raw = entry.get(f"Zone{zone_num}", "")
    if not raw:
        return ""
    parts = [p.strip() for p in raw.split("\n") if p.strip()]
    if not parts:
        return ""
    if part_index < len(parts):
        return parts[part_index]
    return parts[0]


def _append_ma_weight_row(services: dict, service: str, weight_label: str,
                          entry: dict, zone_part_index: int = 0):
    bucket = services[service]
    bucket["weights"].append(weight_label)
    for zone_num in range(1, 12):
        bucket["rates"][zone_num].append(
            _zone_rate_value(entry, zone_num, zone_part_index)
        )


def parse_zone_letters_ma(block_rows: list) -> dict[int, str]:
    """Build {zone_num: letter} from zone header rows inside a MainCostsMA block."""
    zone_letters: dict[int, str] = {}
    for entry in block_rows:
        weight = (entry.get("Weight") or "").strip().upper()
        if weight and weight not in ("WEIGHT KG", "KG") and not _is_half_kg(entry.get("Weight", "")):
            zone_values = [entry.get(f"Zone{i}", "") for i in range(1, 12) if entry.get(f"Zone{i}")]
            if zone_values and not all(_looks_like_zone_letter(v) for v in zone_values):
                continue
        for zone_num in range(1, 12):
            raw = entry.get(f"Zone{zone_num}", "")
            if not raw or zone_num in zone_letters:
                continue
            if _looks_like_zone_letter(raw):
                zone_letters[zone_num] = raw.strip().split("\n")[0].strip().upper()
            elif "\n" in raw:
                parts = [p.strip() for p in raw.split("\n") if p.strip()]
                if parts and _looks_like_zone_letter(parts[0]):
                    zone_letters[zone_num] = parts[0].upper()
                if len(parts) > 1 and _looks_like_zone_letter(parts[1]):
                    next_zone = zone_num + 1
                    if next_zone <= 11 and next_zone not in zone_letters:
                        zone_letters[next_zone] = parts[1].upper()
    return zone_letters


def split_main_costs_ma_blocks(mc_data: list) -> list[tuple[str, list]]:
    """Split MainCostsMA into (direction, rows) blocks."""
    raw_blocks: list[list] = []
    current: list = []

    for entry in mc_data:
        rate_name = entry.get("RateName", "")
        if rate_name and _is_ma_block_header(rate_name) and current:
            raw_blocks.append(current)
            current = [entry]
        else:
            current.append(entry)
    if current:
        raw_blocks.append(current)

    blocks = []
    for idx, block_rows in enumerate(raw_blocks):
        direction = _block_direction(block_rows, idx, len(raw_blocks))
        blocks.append((direction, block_rows))
    return blocks


def parse_main_costs_ma_block(block_rows: list) -> tuple[dict[int, str], dict]:
    """Parse one export/import block into zone letters and per-service weight/rate tables."""
    zone_letters = parse_zone_letters_ma(block_rows)
    services = {
        name: {"weights": [], "rates": {zone: [] for zone in range(1, 12)}}
        for name in MA_SERVICE_ORDER
    }

    current_service = None
    half_kg_counter = 0

    for entry in block_rows:
        weight = entry.get("Weight", "")
        rate_name = entry.get("RateName", "")

        if not weight:
            continue
        weight_upper = weight.upper().strip()
        if weight_upper in ("WEIGHT KG", "KG"):
            continue

        if "\n" in weight:
            parts = [p.strip() for p in weight.split("\n") if p.strip()]
            if len(parts) >= 2 and _is_half_kg(parts[-1]):
                if current_service and not _is_half_kg(parts[0]):
                    _append_ma_weight_row(services, current_service, parts[0], entry, 0)
                half_kg_counter += 1
                current_service = _identify_ma_service("", half_kg_counter - 1)
                _append_ma_weight_row(services, current_service, "0.5", entry, 1)
                continue

        if _is_half_kg(weight):
            if rate_name:
                service = _identify_ma_service(rate_name, half_kg_counter)
            else:
                service = _identify_ma_service("", half_kg_counter)
            half_kg_counter += 1
            current_service = service
            _append_ma_weight_row(services, current_service, "0.5", entry, 0)
            continue

        if current_service is None:
            continue

        weight_label = weight.split("\n")[0].strip()
        _append_ma_weight_row(services, current_service, weight_label, entry, 0)

    return zone_letters, services


def build_ma_brackets(weights: list) -> list[tuple[str, str]]:
    """Convert raw weight labels to bracket labels and Flat / p/unit types."""
    brackets = []
    for idx, weight in enumerate(weights):
        label = weight.strip()
        is_last = idx == len(weights) - 1

        if label.endswith("+") and is_last:
            base = label[:-1].strip()
            brackets.append((f">{base}", "p/unit"))
            continue

        if "-" in label and not label.endswith("+"):
            upper = label.split("-")[-1].strip()
            brackets.append((f"<={upper}", "Flat"))
            continue

        numeric = label.replace(",", ".")
        brackets.append((f"<={numeric}", "Flat"))
    return brackets


def build_ma_lanes(zone_letters: dict[int, str], direction: str) -> list[dict]:
    """Build lane rows for one export or import direction."""
    lanes = []
    for zone_num in sorted(zone_letters):
        letter = zone_letters[zone_num]
        region = f"Zone {letter}"
        if direction == "export":
            lanes.append({
                "carrier": MA_CARRIER,
                "origin_country": MA_COUNTRY,
                "origin_region": "",
                "destination": "",
                "dest_region": region,
                "zone_num": zone_num,
                "direction": direction,
                "currency": MA_CURRENCY,
            })
        else:
            lanes.append({
                "carrier": MA_CARRIER,
                "origin_country": "",
                "origin_region": region,
                "destination": MA_COUNTRY,
                "dest_region": "",
                "zone_num": zone_num,
                "direction": direction,
                "currency": MA_CURRENCY,
            })
    return lanes


def build_regions_ma(regions_data: list) -> dict[str, list[str]]:
    """Group RegionsMA country codes by zone letter."""
    regions: dict[str, list[str]] = {}
    for entry in regions_data:
        zone = (entry.get("Zone") or "").strip().upper()
        country = (entry.get("Country") or "").strip()
        code = (entry.get("CountryCode") or "").strip().upper()
        if not zone:
            continue
        if code:
            regions.setdefault(zone, []).append(code)
        elif country:
            for iso_code in expand_to_country_codes(country):
                regions.setdefault(zone, []).append(iso_code.upper())
    for zone in regions:
        regions[zone] = sorted(set(regions[zone]))
    return dict(sorted(regions.items()))


def write_regions_ma_txt(regions: dict[str, list[str]], output_path: Path):
    """Write Morocco zone-to-country mapping."""
    lines = [
        "=" * 70,
        "COUNTRY REGIONS - Morocco (RegionsMA)",
        "=" * 70,
        "",
    ]
    for zone, countries in regions.items():
        lines.append(f"--- Zone {zone} ---")
        lines.append(f"  {', '.join(countries)}")
        lines.append("")
    lines.append("=" * 70)
    output_path.write_text("\n".join(lines), encoding="utf-8")
    print(f"  Written: {output_path}")


def parse_main_costs_ma(mc_data: list) -> tuple[list[dict], list[dict]]:
    """Parse MainCostsMA into lane rows and cost block definitions."""
    blocks = split_main_costs_ma_blocks(mc_data)
    parsed_blocks = []
    canonical_zones: dict[int, str] = {}

    for direction, block_rows in blocks:
        zone_letters, services = parse_main_costs_ma_block(block_rows)
        if len(zone_letters) > len(canonical_zones):
            canonical_zones = zone_letters
        parsed_blocks.append((direction, zone_letters, services))

    all_lanes: list[dict] = []
    cost_blocks: list[dict] = []

    for direction, zone_letters, services in parsed_blocks:
        merged_zones = {**canonical_zones, **zone_letters}
        all_lanes.extend(build_ma_lanes(merged_zones, direction))

        direction_blocks = []
        for service in MA_SERVICE_ORDER:
            data = services[service]
            if not data["weights"]:
                continue
            brackets = build_ma_brackets(data["weights"])
            direction_blocks.append({
                "direction": direction,
                "service": service,
                "cost_name": _ma_cost_name(service, direction),
                "brackets": brackets,
                "labels": [b[0] for b in brackets],
                "types": [b[1] for b in brackets],
                "rates": data["rates"],
            })
        cost_blocks.extend(_merge_env_pak_ma_blocks(direction_blocks))

    cost_blocks = _merge_duplicate_ma_cost_blocks(cost_blocks)
    return all_lanes, cost_blocks


def write_ma_output(lanes: list, cost_blocks: list, output_dir: Path):
    """Write Morocco international rates xlsx."""
    header_cols = [
        "Lane #", "Carrier Name", "Origin Country",
        "Origin Country Region", "Destination",
        "Destination Country Region",
    ]
    num_ship_cols = len(header_cols)

    wb = Workbook()
    ws = wb.active
    ws.title = "International Rates"

    row1 = [""] * num_ship_cols
    for block in cost_blocks:
        row1.append(block["cost_name"])
        row1.extend([""] * len(block["labels"]))
    ws.append(row1)

    row2 = [""] * num_ship_cols
    for block in cost_blocks:
        row2.append("")
        row2.extend([""] * len(block["labels"]))
    ws.append(row2)

    row3 = [""] * num_ship_cols
    for block in cost_blocks:
        row3.append("Rate by: Weight/chargeable kg")
        row3.extend([""] * len(block["labels"]))
    ws.append(row3)

    row4 = [""] * num_ship_cols
    for block in cost_blocks:
        row4.append("Regular + rule")
        row4.extend([""] * len(block["labels"]))
    ws.append(row4)

    row5 = [""] * num_ship_cols
    for block in cost_blocks:
        row5.append("Rounding rule:")
        row5.extend([""] * len(block["labels"]))
    ws.append(row5)

    row6 = list(header_cols)
    for block in cost_blocks:
        row6.append("Currency")
        row6.extend(block["labels"])
    ws.append(row6)

    row7 = [""] * num_ship_cols
    for block in cost_blocks:
        row7.append("Currency")
        row7.extend(block["types"])
    ws.append(row7)

    for lane_idx, lane in enumerate(lanes, 1):
        row = [
            lane_idx,
            lane["carrier"],
            lane["origin_country"],
            lane["origin_region"],
            lane["destination"],
            lane["dest_region"],
        ]
        for block in cost_blocks:
            rates_by_direction = block.get("rates_by_direction")
            if rates_by_direction:
                row.append(lane["currency"])
                zone_rates = rates_by_direction.get(lane["direction"], {}).get(
                    lane["zone_num"], []
                )
            elif block["direction"] != lane["direction"]:
                row.append("")
                row.extend([""] * len(block["labels"]))
                continue
            else:
                row.append(lane["currency"])
                zone_rates = block["rates"].get(lane["zone_num"], [])
            row.extend(zone_rates)
            if len(zone_rates) < len(block["labels"]):
                row.extend([""] * (len(block["labels"]) - len(zone_rates)))
        ws.append(row)

    xlsx_path = output_dir / "international_rates.xlsx"
    wb.save(xlsx_path)
    print(f"  Written: {xlsx_path}")


def main():
    with open(INPUT_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)

    fields = data["documents"][0]["fields"]

    # 1. Parse ZoningMatrix
    print("1. Parsing ZoningMatrix...")
    zm_data = fields["ZoningMatrix"]["value"]
    zone_map, full_zone_map, all_zones = parse_zoning_matrix(zm_data)

    for country in sorted(all_zones.keys()):
        _, code = CARRIER_MAP[country]
        print(f"  {code}: zones {all_zones[country]}")

    # 2. Build Country Regions (all countries in each zone)
    print("2. Building Country Regions...")
    regions = build_country_regions(full_zone_map)
    write_country_regions_txt(regions, OUTPUT_DIR / "CountryRegions.txt")

    # 3. Parse MainCosts
    print("3. Parsing MainCosts...")
    mc_data = fields["MainCosts"]["value"]
    rate_blocks = parse_main_costs(mc_data)
    for name, block in rate_blocks.items():
        print(f"  {name}: {len(block['weights'])} weight entries, "
              f"{len(block['rates'])} zones with data")

    # 4. Build lanes
    print("4. Building lanes...")
    lanes = build_lanes(zone_map, all_zones)
    forward_count = sum(len(z) for z in all_zones.values())
    print(f"  Forward lanes: {forward_count}")
    print(f"  Duplicate (switched) lanes: {forward_count}")
    print(f"  Total lanes: {len(lanes)}")

    # 5. Write output
    print("5. Writing output...")
    write_output(lanes, rate_blocks, OUTPUT_DIR)

    print("\nDone!")


if __name__ == "__main__":
    main()
