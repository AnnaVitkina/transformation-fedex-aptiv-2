import json
from pathlib import Path
from openpyxl import Workbook

INPUT_FILE = Path("processing/fedex domestic 6(with pln).pdf_extracted.json")
OUTPUT_DIR = Path("output")
OUTPUT_DIR.mkdir(exist_ok=True)

CURRENCY = "EUR"


def check_zones_identical(entries: list) -> bool:
    """Check if all zone values are identical for every weight bracket (skip headers)."""
    for entry in entries:
        zones = {}
        for key, val in entry.items():
            if key.startswith("Zone") and val:
                zones[key] = val
        values = list(zones.values())
        # Skip header rows (contain text like "Zone 1")
        if values and any(v.startswith("Zone") for v in values):
            continue
        if len(set(values)) > 1:
            return False
    return True


def parse_domestic_at(data: list) -> dict:
    """Parse DomesticMainCostsAT into rate structure."""
    weights = []
    rates = []

    for entry in data:
        weight = entry.get("Weight", "")
        rate_name = entry.get("RateName", "")

        # Skip header/label rows
        if not weight or weight.startswith("Gewicht") or weight.startswith("Zone"):
            continue
        if rate_name and not weight:
            continue

        # Get rate value from Zone1 (all zones are identical)
        zone1 = entry.get("Zone1", "")
        if not zone1:
            continue

        weights.append(weight)
        rates.append(zone1)

    return {"weights": weights, "rates": rates}


def build_domestic_brackets_at(weights: list) -> list:
    """Build bracket labels for AT domestic.
    Last entry 'jede weiteren 25 kg' becomes >100 p/25 units.
    """
    brackets = []
    for w in weights:
        if "weiteren" in w.lower() or "add" in w.lower() or "jede" in w.lower():
            # Last regular weight before this
            prev = weights[weights.index(w) - 1]
            # Clean the weight value
            prev_clean = prev.rstrip("0").rstrip(".") if "." in prev else prev
            brackets.append((f">{prev_clean}", "p/25 units"))
        else:
            w_clean = w.rstrip("0").rstrip(".") if "." in w else w
            brackets.append((f"<={w_clean}", "Flat"))
    return brackets


def parse_domestic_ro(data: list) -> dict:
    """Parse DomesticMainCostsRO into Express (4 zones) and Economy (1 zone) blocks."""
    SERVICE_MAP = {
        "Zone1": "EXP",
        "Zone2": "EXP900",
        "Zone3": "EXP1000",
        "Zone4": "EXP1200",
    }

    express_weights = []
    express_rates = {"Zone1": [], "Zone2": [], "Zone3": [], "Zone4": []}
    economy_weights = []
    economy_rates = []

    section = None  # "express" or "economy"
    seen_weights_express = set()

    for entry in data:
        weight = entry.get("Weight", "")
        rate_name = entry.get("RateName", "")
        zone1 = entry.get("Zone1", "")

        # Detect section headers
        if "DOMESTIC EXPRESS" in rate_name or "DOMESTIC 09" in rate_name:
            section = "express"
            if weight == "Weight (kg)":
                continue
        elif "DOMESTIC ECONOMY" in rate_name:
            section = "economy"
            if weight == "Weight (kg)":
                continue

        if not weight or weight == "Weight (kg)":
            continue

        if section == "express":
            if weight in seen_weights_express:
                continue
            seen_weights_express.add(weight)
            express_weights.append(weight)
            express_rates["Zone1"].append(entry.get("Zone1", ""))
            express_rates["Zone2"].append(entry.get("Zone2", ""))
            express_rates["Zone3"].append(entry.get("Zone3", ""))
            express_rates["Zone4"].append(entry.get("Zone4", ""))

        elif section == "economy":
            economy_weights.append(weight)
            economy_rates.append(zone1)

    return {
        "express_weights": express_weights,
        "express_rates": express_rates,
        "economy_weights": economy_weights,
        "economy_rates": economy_rates,
        "service_map": SERVICE_MAP,
    }


def build_domestic_brackets_ro_express(weights: list) -> list:
    """Build bracket labels for RO Express domestic.
    'add per kg' becomes >last_weight, p/unit.
    """
    brackets = []
    for i, w in enumerate(weights):
        if "add" in w.lower() and "per" in w.lower() and "kg" in w.lower():
            prev = weights[i - 1]
            brackets.append((f">{prev}", "p/unit"))
        else:
            brackets.append((f"<={w}", "Flat"))
    return brackets


def build_domestic_brackets_ro_economy(weights: list) -> list:
    """Build bracket labels for RO Economy domestic.
    'add per 100 kg additional' becomes >last_weight, p/100 units.
    """
    brackets = []
    for i, w in enumerate(weights):
        if "add" in w.lower() and "100" in w:
            prev = weights[i - 1]
            brackets.append((f">{prev}", "p/100 units"))
        else:
            brackets.append((f"<={w}", "Flat"))
    return brackets


def parse_domestic_hu(data: list) -> dict:
    """Parse DomesticMainCostsHU into weights and zone rates."""
    weights = []
    rates_z1 = []
    rates_z2 = []

    for entry in data:
        weight = entry.get("Weight", "")
        rate_name = entry.get("RateName", "")

        if not weight or weight == "Weight (kg)":
            continue
        if rate_name and rate_name in ("HUNGARY", "TNT DOMESTIC EXPRESS"):
            continue

        zone1 = entry.get("Zone1", "")
        zone2 = entry.get("Zone2", "")
        if not zone1:
            continue

        # Normalize "Min 3" to "3"
        w_label = weight
        if weight.lower().startswith("min"):
            w_label = weight.split()[-1]

        weights.append(w_label)
        rates_z1.append(zone1)
        rates_z2.append(zone2)

    return {"weights": weights, "rates_z1": rates_z1, "rates_z2": rates_z2}


def filter_weights_hu(weights: list, rates_z1: list, rates_z2: list):
    """Filter HU weights using the same add-on logic as international.
    Remove intermediate weights between add-ons, label add-ons correctly.
    """
    is_addon = [("add" in w.lower() and "per" in w.lower()) for w in weights]

    # Find positions of add-ons
    addon_positions = [i for i, v in enumerate(is_addon) if v]

    # Build filtered output
    filtered_brackets = []
    filtered_z1 = []
    filtered_z2 = []

    for i, w in enumerate(weights):
        if is_addon[i]:
            # Find the last regular weight before the NEXT add-on (or end)
            next_addon_idx = None
            for j in addon_positions:
                if j > i:
                    next_addon_idx = j
                    break

            if next_addon_idx is not None:
                # There's another add-on after this one
                # Find last regular weight before next add-on
                last_regular = None
                for k in range(i + 1, next_addon_idx):
                    if not is_addon[k]:
                        last_regular = weights[k]
                label = f"<={last_regular}" if last_regular else f">={weights[i-1]}"
            else:
                # Last add-on
                prev_regular = weights[i - 1] if i > 0 else "0"
                # Find last regular weight before this add-on
                for k in range(i - 1, -1, -1):
                    if not is_addon[k]:
                        prev_regular = weights[k]
                        break
                label = f">{prev_regular}"

            filtered_brackets.append((label, "p/unit"))
            filtered_z1.append(rates_z1[i])
            filtered_z2.append(rates_z2[i])

        else:
            # Regular weight - check if it's between two add-ons (skip it)
            between_addons = False
            if addon_positions:
                prev_addon = None
                next_addon = None
                for ap in addon_positions:
                    if ap < i:
                        prev_addon = ap
                    elif ap > i and next_addon is None:
                        next_addon = ap
                if prev_addon is not None and next_addon is not None:
                    between_addons = True

            if not between_addons:
                filtered_brackets.append((f"<={w}", "Flat"))
                filtered_z1.append(rates_z1[i])
                filtered_z2.append(rates_z2[i])

    return filtered_brackets, filtered_z1, filtered_z2


def parse_domestic_de(data: list) -> dict:
    """Parse DomesticMainCostsDE into Express (5 zones) and BusinessPak (4 zones)."""
    EXPRESS_SERVICE_MAP = {
        "Zone1": "EXP",
        "Zone2": "EXP1200",
        "Zone3": "EXP1000",
        "Zone4": "EXP900",
        "Zone5": "EXP800",
    }
    BPAK_SERVICE_MAP = {
        "Zone1": "EXP1200_BusinessPak",
        "Zone2": "EXP1000_BusinessPak",
        "Zone3": "EXP900_BusinessPak",
        "Zone4": "EXP800_BusinessPak",
    }

    express_brackets = []
    express_rates = {f"Zone{i}": [] for i in range(1, 6)}
    bpak_brackets = []
    bpak_rates = {f"Zone{i}": [] for i in range(1, 5)}

    section = None  # "express" or "bpak"

    for entry in data:
        rate_name = entry.get("RateName", "")
        weight = entry.get("Weight", "")

        if "EXPRESS DOMESTIC" in rate_name:
            section = "express"
            continue
        elif "BUSINESS PAK" in rate_name:
            section = "bpak"
            if weight == "Base price":
                continue

        if not weight or weight == "weight in kg" or weight == "Base price":
            continue

        if section == "express":
            # Parse weight range: "0.01\n-\n2.00" → <=2
            if "add" in weight.lower() or "each" in weight.lower():
                # Last bracket → >last_upper_bound, p/unit
                last_upper = express_brackets[-1][0].replace("<=", "")
                express_brackets.append((f">{last_upper}", "p/unit"))
                for z in range(1, 6):
                    val = entry.get(f"Zone{z}", "").replace("€", "").strip()
                    express_rates[f"Zone{z}"].append(val)
            elif "\n-\n" in weight:
                parts = weight.split("\n-\n")
                upper = parts[1].rstrip("0").rstrip(".")
                express_brackets.append((f"<={upper}", "Flat"))
                for z in range(1, 6):
                    val = entry.get(f"Zone{z}", "").replace("€", "").strip()
                    express_rates[f"Zone{z}"].append(val)

        elif section == "bpak":
            # Single weight: "Price per padded envelope up to max. 5 kg..."
            if "max" in weight.lower() or "kg" in weight.lower():
                import re
                m = re.search(r"(\d+)\s*kg", weight)
                upper = m.group(1) if m else "5"
                bpak_brackets.append((f"<={upper}", "Flat"))
                for z in range(1, 5):
                    val = entry.get(f"Zone{z}", "").replace("€", "").strip()
                    bpak_rates[f"Zone{z}"].append(val)

    return {
        "express_brackets": express_brackets,
        "express_rates": express_rates,
        "express_service_map": EXPRESS_SERVICE_MAP,
        "bpak_brackets": bpak_brackets,
        "bpak_rates": bpak_rates,
        "bpak_service_map": BPAK_SERVICE_MAP,
    }


def parse_domestic_pl(data: list) -> dict:
    """Parse DomesticMainCostsPL into services with weight brackets."""
    SERVICE_MAP = {
        "15N": "EXP",
        "12N": "EXP1200",
        "10N": "EXP1000",
        "9N": "EXP900",
    }

    services = {}  # service_code -> list of (weight, rate)
    current_service = None
    skip_service = False

    for entry in data:
        rate_name = entry.get("RateName", "").strip().rstrip(",")
        weight = entry.get("Weight", "")
        zone1 = entry.get("Zone1", "")

        # Detect section header
        if rate_name and rate_name in SERVICE_MAP:
            current_service = SERVICE_MAP[rate_name]
            skip_service = False
            if current_service not in services:
                services[current_service] = []
            if weight and zone1:
                services[current_service].append((weight, zone1))
            continue

        # Skip D-services (10D, 9D, etc.)
        if rate_name and ("D" in rate_name or "d" in rate_name):
            skip_service = True
            continue

        # Skip header/date rows
        if rate_name and rate_name not in SERVICE_MAP:
            if any(c.isdigit() and "-" in rate_name for c in rate_name):
                continue
            skip_service = True
            continue

        if skip_service or current_service is None:
            continue

        if not weight:
            continue

        # Only take first set of weights per service (avoid duplicates)
        existing_weights = [w for w, _ in services.get(current_service, [])]
        if weight in existing_weights:
            continue

        services[current_service].append((weight, zone1))

    return {"services": services}


def build_domestic_brackets_pl(weight_entries: list) -> tuple:
    """Build brackets for PL domestic.
    '1' -> <=1 Flat, '3' -> <=3 Flat, '3.000+' -> >3 p/kg
    """
    brackets = []
    rates = []
    for weight, rate in weight_entries:
        if "+" in weight:
            # Add-on: extract base weight (3.000+ -> 3)
            base = weight.replace("+", "").strip()
            # Remove trailing zeros after decimal
            if "." in base:
                base = base.rstrip("0").rstrip(".")
            brackets.append((f">{base}", "p/kg"))
        else:
            brackets.append((f"<={weight}", "Flat"))
        rates.append(rate)
    return brackets, rates


PT_DEST_ZONES = [
    ("Zone1", "Zone 1 exclude Braga"),
    ("Zone2", "Zone 2"),
    ("Zone3", "Madeira"),
    ("Zone4", "Acores"),
]
PT_ORIGIN = "Braga"

PREMIUM_SERVICE_MAP = {
    "9:00 express": ("Transport cost (Exp 900)", "Service equals EXP900"),
    "10:00 express": ("Transport cost (Exp 1000)", "Service equals EXP1000"),
    "12:00 express": ("Transport cost (Exp 1200)", "Service equals EXP1200"),
}


def _clean_pt_rate(val: str) -> str:
    """Strip currency symbols and take first line if multiline."""
    if not val:
        return ""
    val = val.split("\n")[0].strip()
    return val.replace("€", "").strip()


def _is_pt_up_to_5kg(weight: str) -> bool:
    w = weight.lower()
    return "5kg" in w and ("até" in w or "ate" in w or "up to" in w)


def _is_pt_add_per_kg(weight: str) -> bool:
    w = weight.lower()
    return ("adc" in w or "add" in w) and "kg" in w


def parse_domestic_pt(data: list) -> dict:
    """Parse DomesticMainCostsPT into express and premium service blocks."""
    express_brackets = [("<=5", "Flat"), (">5", "p/unit")]
    express_rates = {z: [] for z, _ in PT_DEST_ZONES}

    premium_brackets = [("<=70", "Flat"), (">70", "Flat")]
    premium_services = {}  # service_key -> [rate_<70, rate_>70]

    section = None

    for entry in data:
        rate_name = entry.get("RateName", "").strip()
        weight = entry.get("Weight", "").strip()

        if "TNT DOMESTIC EXPRESS" in rate_name:
            section = "express"
            continue
        if "SERVIÇOS PREMIUM" in rate_name or "SERVICOS PREMIUM" in rate_name:
            section = "premium_header"
            continue

        if rate_name in PREMIUM_SERVICE_MAP:
            section = rate_name
            if weight.lower() == "envio":
                premium_services[rate_name] = [
                    _clean_pt_rate(entry.get("Zone1", "")),
                    _clean_pt_rate(entry.get("Zone2", "")),
                ]
            continue

        if not weight or weight in ("Peso (kg)", "Taxa por"):
            continue

        if section == "express":
            if _is_pt_up_to_5kg(weight):
                for zone_key, _ in PT_DEST_ZONES:
                    express_rates[zone_key].append(_clean_pt_rate(entry.get(zone_key, "")))
            elif _is_pt_add_per_kg(weight):
                for zone_key, _ in PT_DEST_ZONES:
                    express_rates[zone_key].append(_clean_pt_rate(entry.get(zone_key, "")))

    return {
        "express_brackets": express_brackets,
        "express_rates": express_rates,
        "premium_brackets": premium_brackets,
        "premium_services": premium_services,
    }


def build_postal_code_zones_pt_lines() -> list:
    """Build PostalCodeZones PT txt lines (tab-separated)."""
    return [
        "Name of Zone\tCountry\tPostal Code\tExcluded",
        "Zone 2\tPT\t5000-5999, 6050, 7000-8999\t",
        "Madeira\tPT\t90,91,92,93\t",
        "Acores\tPT\t99\t",
        "Braga\tPT\t47\t",
        "Zone 1\tPT\t1,2,3,4,6,94,95,96,97,98\t",
        "Zone 1 exclude Braga\tPT\t1,2,3,4,6,94,95,96,97,98\t47",
    ]


def write_sheet_separate(ws, cost_blocks, txt_lines, sheet_name):
    """Write cost blocks stacked vertically (one after another)."""
    txt_lines.append(f"=== {sheet_name} ===")
    txt_lines.append("")

    for block in cost_blocks:
        cost_name = block["cost_name"]
        brackets = block["brackets"]
        lanes = block["lanes"]
        header_cols = block["header_cols"]
        currency = block.get("currency", CURRENCY)

        num_ship_cols = len(header_cols)
        bracket_labels = [b[0] for b in brackets]
        value_types = [b[1] for b in brackets]

        ws.append([""] * num_ship_cols + [cost_name])
        ws.append([""] * num_ship_cols + ["Applies if:"])
        ws.append([""] * num_ship_cols + ["Rate by: Weight/chargeable kg"])
        ws.append([""] * num_ship_cols + ["Regular + rule"])
        ws.append([""] * num_ship_cols + ["Rounding rule:"])
        ws.append(header_cols + ["Currency"] + bracket_labels)
        ws.append([""] * num_ship_cols + ["Currency"] + value_types)

        for lane in lanes:
            row = list(lane["label_values"])
            row.append(currency)
            row.extend(lane["rates"])
            ws.append(row)

        ws.append([])

        # TXT
        txt_lines.append(cost_name)
        txt_lines.append("Applies if:")
        txt_lines.append("Rate by: Weight/chargeable kg")
        txt_lines.append("Regular + rule")
        txt_lines.append("Rounding rule:")
        txt_lines.append("")
        txt_lines.append("\t".join(header_cols + ["Currency"] + bracket_labels))
        txt_lines.append("\t".join([""] * len(header_cols) + ["Currency"] + value_types))

        for lane in lanes:
            row = [str(v) for v in lane["label_values"]]
            row.append(currency)
            row.extend(lane["rates"])
            txt_lines.append("\t".join(row))

        txt_lines.append("")


def write_sheet_combined(ws, sheet_info, txt_lines, sheet_name):
    """Write cost blocks as consecutive column groups (side by side)."""
    header_cols = sheet_info["header_cols"]
    cost_blocks = sheet_info["cost_blocks"]
    lanes = sheet_info["lanes"]
    currency = sheet_info.get("currency", CURRENCY)
    num_ship_cols = len(header_cols)

    # Build combined header rows
    row1 = [""] * num_ship_cols  # cost names
    row2 = [""] * num_ship_cols  # applies if
    row3 = [""] * num_ship_cols  # rate by
    row4 = [""] * num_ship_cols  # regular + rule
    row5 = [""] * num_ship_cols  # rounding rule
    row6 = list(header_cols)     # column headers + brackets
    row7 = [""] * num_ship_cols  # value types

    for block in cost_blocks:
        brackets = block["brackets"]
        bracket_labels = [b[0] for b in brackets]
        value_types = [b[1] for b in brackets]

        row1.append(block["cost_name"])
        row1.extend([""] * len(bracket_labels))

        applies_if = block.get("applies_if", "")
        row2.append(applies_if if applies_if else "Applies if:")
        row2.extend([""] * len(bracket_labels))

        row3.append("Rate by: Weight/chargeable kg")
        row3.extend([""] * len(bracket_labels))

        row4.append("Regular + rule")
        row4.extend([""] * len(bracket_labels))

        row5.append("Rounding rule:")
        row5.extend([""] * len(bracket_labels))

        row6.append("Currency")
        row6.extend(bracket_labels)

        row7.append("Currency")
        row7.extend(value_types)

    ws.append(row1)
    ws.append(row2)
    ws.append(row3)
    ws.append(row4)
    ws.append(row5)
    ws.append(row6)
    ws.append(row7)

    # Data rows
    for lane in lanes:
        row = list(lane["label_values"])
        for i, block in enumerate(cost_blocks):
            block_rates = lane["block_rates"][i]
            num_brackets = len(block["brackets"])
            if block_rates is not None:
                row.append(currency)
                row.extend(block_rates)
            else:
                row.extend([""] * (num_brackets + 1))
        ws.append(row)

    ws.append([])

    # TXT
    txt_lines.append(f"=== {sheet_name} ===")
    txt_lines.append("")
    txt_lines.append("\t".join(row6))
    txt_lines.append("\t".join(row7))

    for lane in lanes:
        row = [str(v) for v in lane["label_values"]]
        for i, block in enumerate(cost_blocks):
            block_rates = lane["block_rates"][i]
            num_brackets = len(block["brackets"])
            if block_rates is not None:
                row.append(currency)
                row.extend(block_rates)
            else:
                row.extend([""] * (num_brackets + 1))
        txt_lines.append("\t".join(row))

    txt_lines.append("")


def write_domestic_output(sheets: list, output_dir: Path):
    """Write domestic rates to xlsx with multiple sheets."""
    wb = Workbook()
    wb.remove(wb.active)

    for sheet_info in sheets:
        sheet_name = sheet_info["sheet_name"]
        ws = wb.create_sheet(title=sheet_name)

        if sheet_info.get("combined"):
            write_sheet_combined(ws, sheet_info, [], sheet_name)
        else:
            write_sheet_separate(ws, sheet_info["cost_blocks"], [], sheet_name)

    xlsx_path = output_dir / "domestic_rates.xlsx"
    wb.save(xlsx_path)
    print(f"  Written: {xlsx_path}")


def main():
    with open(INPUT_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)

    fields = data["documents"][0]["fields"]

    sheets = []

    # --- DomesticMainCostsAT ---
    print("Processing DomesticMainCostsAT...")
    at_data = fields["DomesticMainCostsAT"]["value"]

    zones_same = check_zones_identical(at_data)
    print(f"  All zones identical: {zones_same}")

    parsed = parse_domestic_at(at_data)
    print(f"  Weight brackets: {len(parsed['weights'])}")
    print(f"  Weights: {parsed['weights'][:5]}...{parsed['weights'][-3:]}")

    brackets = build_domestic_brackets_at(parsed["weights"])
    print(f"  Bracket labels: {[b[0] for b in brackets][:5]}...{[b[0] for b in brackets][-3:]}")

    if zones_same:
        header_cols = ["Lane #", "Carrier Name", "Origin Country", "Destination"]
        lane_values = [1, "TNT AT", "AT", "AT"]
    else:
        header_cols = [
            "Lane #", "Carrier Name", "Origin Country",
            "Origin Country Region", "Destination",
            "Destination Country Region",
        ]
        lane_values = [1, "TNT AT", "AT", "", "AT", ""]

    sheets.append({
        "sheet_name": "Domestic Express AT",
        "cost_blocks": [{
            "cost_name": "Transport cost (Domestic Express AT)",
            "brackets": brackets,
            "header_cols": header_cols,
            "lanes": [{"label_values": lane_values, "rates": parsed["rates"]}],
        }],
    })

    # --- DomesticMainCostsRO ---
    print("\nProcessing DomesticMainCostsRO...")
    ro_data = fields["DomesticMainCostsRO"]["value"]
    ro_parsed = parse_domestic_ro(ro_data)

    # Express block
    express_brackets = build_domestic_brackets_ro_express(ro_parsed["express_weights"])
    print(f"  Express weights: {len(ro_parsed['express_weights'])}")
    print(f"    First: {ro_parsed['express_weights'][:5]}, Last: {ro_parsed['express_weights'][-3:]}")
    print(f"    Brackets: {[b[0] for b in express_brackets][:3]}...{[b[0] for b in express_brackets][-2:]}")

    # Economy block
    economy_brackets = build_domestic_brackets_ro_economy(ro_parsed["economy_weights"])
    print(f"  Economy weights: {len(ro_parsed['economy_weights'])}")
    print(f"    First: {ro_parsed['economy_weights'][:3]}, Last: {ro_parsed['economy_weights'][-3:]}")
    print(f"    Brackets: {[b[0] for b in economy_brackets][:3]}...{[b[0] for b in economy_brackets][-2:]}")

    # Build combined lanes (all 5 services, Express rates + Economy rates side by side)
    service_map = ro_parsed["service_map"]
    combined_lanes = []
    lane_num = 1
    for zone_key in ["Zone1", "Zone2", "Zone3", "Zone4"]:
        combined_lanes.append({
            "label_values": [lane_num, service_map[zone_key]],
            "block_rates": [ro_parsed["express_rates"][zone_key], None],
        })
        lane_num += 1

    combined_lanes.append({
        "label_values": [lane_num, "EXP_EC"],
        "block_rates": [None, ro_parsed["economy_rates"]],
    })

    sheets.append({
        "sheet_name": "Domestic Express RO",
        "combined": True,
        "header_cols": ["Lane #", "Service"],
        "cost_blocks": [
            {
                "cost_name": "Transport cost (Domestic Express RO)",
                "brackets": express_brackets,
            },
            {
                "cost_name": "Transport cost (Domestic Economy RO)",
                "brackets": economy_brackets,
            },
        ],
        "lanes": combined_lanes,
    })

    # --- DomesticMainCostsHU ---
    print("\nProcessing DomesticMainCostsHU...")
    hu_data = fields["DomesticMainCostsHU"]["value"]
    hu_parsed = parse_domestic_hu(hu_data)

    print(f"  Weights: {len(hu_parsed['weights'])}")
    print(f"    First: {hu_parsed['weights'][:5]}, Last: {hu_parsed['weights'][-5:]}")

    hu_brackets, hu_z1, hu_z2 = filter_weights_hu(
        hu_parsed["weights"], hu_parsed["rates_z1"], hu_parsed["rates_z2"]
    )
    print(f"  Filtered brackets: {len(hu_brackets)}")
    print(f"    First: {[b[0] for b in hu_brackets][:5]}, Last: {[b[0] for b in hu_brackets][-3:]}")

    # Lanes:
    # Zone 1 (Budapest-Budapest): Origin=HU, Origin PCZ=HU Zone 1, Dest PCZ=HU Zone 1
    # Zone 2 (all HU to Countryside): Origin=HU, Origin PCZ=(empty), Dest PCZ=HU Zone 2
    hu_header_cols = ["Lane #", "Origin Country", "Origin Postal Code Zone",
                      "Destination Postal Code Zone"]
    hu_lanes = [
        {
            "label_values": [1, "HU", "HU Zone 1", "HU Zone 1"],
            "block_rates": [hu_z1],
        },
        {
            "label_values": [2, "HU", "", "HU Zone 2"],
            "block_rates": [hu_z2],
        },
    ]

    sheets.append({
        "sheet_name": "Domestic Express HU",
        "combined": True,
        "header_cols": hu_header_cols,
        "cost_blocks": [
            {
                "cost_name": "Transport cost (Domestic Express HU)",
                "brackets": hu_brackets,
            },
        ],
        "lanes": hu_lanes,
    })

    # Write Postal Code Zones TXT for HU
    pcz_lines = [
        "Name of Zone\tCountry\tPostal Code\tExcluded",
        "HU Zone 1\tHU\t109\t",
        "HU Zone 2\tHU\t0-9\t109",
    ]
    pcz_path = OUTPUT_DIR / "PostalCodeZones_HU.txt"
    pcz_path.write_text("\n".join(pcz_lines), encoding="utf-8")
    print(f"  Written: {pcz_path}")

    # --- DomesticMainCostsDE ---
    print("\nProcessing DomesticMainCostsDE...")
    de_data = fields["DomesticMainCostsDE"]["value"]
    de_parsed = parse_domestic_de(de_data)

    print(f"  Express brackets: {len(de_parsed['express_brackets'])}")
    print(f"    {[b[0] for b in de_parsed['express_brackets']]}")
    print(f"  BusinessPak brackets: {len(de_parsed['bpak_brackets'])}")
    print(f"    {[b[0] for b in de_parsed['bpak_brackets']]}")

    # Build combined lanes (9 services total)
    de_header_cols = ["Lane #", "Origin Country", "Destination Country", "Service"]
    de_lanes = []
    lane_num = 1

    # Express services (lanes 1-5)
    for zone_key in ["Zone1", "Zone2", "Zone3", "Zone4", "Zone5"]:
        service = de_parsed["express_service_map"][zone_key]
        de_lanes.append({
            "label_values": [lane_num, "DE", "DE", service],
            "block_rates": [de_parsed["express_rates"][zone_key], None],
        })
        lane_num += 1

    # BusinessPak services (lanes 6-9)
    for zone_key in ["Zone1", "Zone2", "Zone3", "Zone4"]:
        service = de_parsed["bpak_service_map"][zone_key]
        de_lanes.append({
            "label_values": [lane_num, "DE", "DE", service],
            "block_rates": [None, de_parsed["bpak_rates"][zone_key]],
        })
        lane_num += 1

    sheets.append({
        "sheet_name": "Domestic Express DE",
        "combined": True,
        "header_cols": de_header_cols,
        "cost_blocks": [
            {
                "cost_name": "Transport cost (Domestic Express DE)",
                "brackets": de_parsed["express_brackets"],
            },
            {
                "cost_name": "Transport cost (Domestic Business Pak DE)",
                "brackets": de_parsed["bpak_brackets"],
            },
        ],
        "lanes": de_lanes,
    })

    # --- DomesticMainCostsPL ---
    print("\nProcessing DomesticMainCostsPL...")
    pl_data = fields["DomesticMainCostsPL"]["value"]
    pl_parsed = parse_domestic_pl(pl_data)

    pl_header_cols = ["Lane #", "Origin Country", "Origin Postal Code",
                      "Destination Country", "Destination Postal Code", "Service"]

    # Build lanes for each service (single cost block since all share same brackets)
    pl_services_order = ["EXP", "EXP1200", "EXP1000", "EXP900"]
    pl_all_brackets = None
    pl_lanes = []
    lane_num = 1

    for svc in pl_services_order:
        if svc not in pl_parsed["services"]:
            continue
        entries = pl_parsed["services"][svc]
        brackets, rates = build_domestic_brackets_pl(entries)
        print(f"  {svc}: brackets={[b[0] for b in brackets]}, rates={rates}")

        if pl_all_brackets is None:
            pl_all_brackets = brackets

        # Pad rates to match bracket count if incomplete
        while len(rates) < len(pl_all_brackets):
            rates.append("")

        pl_lanes.append({
            "label_values": [lane_num, "PL", "", "PL", "", svc],
            "block_rates": [rates],
        })
        lane_num += 1

    sheets.append({
        "sheet_name": "Domestic Express PL",
        "combined": True,
        "header_cols": pl_header_cols,
        "currency": "PLN",
        "cost_blocks": [
            {
                "cost_name": "Transport cost (Parcels/Documents)",
                "brackets": pl_all_brackets,
            },
        ],
        "lanes": pl_lanes,
    })

    # --- DomesticMainCostsPT ---
    if "DomesticMainCostsPT" in fields and fields["DomesticMainCostsPT"].get("value"):
        print("\nProcessing DomesticMainCostsPT...")
        pt_data = fields["DomesticMainCostsPT"]["value"]
        pt_parsed = parse_domestic_pt(pt_data)

        pt_cost_blocks = [
            {
                "cost_name": "Transport cost",
                "brackets": pt_parsed["express_brackets"],
            },
        ]
        for svc_key, (cost_name, applies_if) in PREMIUM_SERVICE_MAP.items():
            if svc_key in pt_parsed["premium_services"]:
                pt_cost_blocks.append({
                    "cost_name": cost_name,
                    "brackets": pt_parsed["premium_brackets"],
                    "applies_if": applies_if,
                })

        pt_lanes = []
        lane_num = 1
        premium_order = list(PREMIUM_SERVICE_MAP.keys())
        for zone_key, dest_label in PT_DEST_ZONES:
            block_rates = [pt_parsed["express_rates"][zone_key]]
            for svc_key in premium_order:
                if svc_key in pt_parsed["premium_services"]:
                    rates = pt_parsed["premium_services"][svc_key]
                    while len(rates) < 2:
                        rates.append("")
                    block_rates.append(rates)
                else:
                    block_rates.append(None)
            pt_lanes.append({
                "label_values": [lane_num, PT_ORIGIN, dest_label],
                "block_rates": block_rates,
            })
            lane_num += 1

        sheets.append({
            "sheet_name": "Domestic Express PT",
            "combined": True,
            "header_cols": ["Lane #", "Origin Postal Code Zone", "Destination Postal Code Zone"],
            "cost_blocks": pt_cost_blocks,
            "lanes": pt_lanes,
        })

        pcz_path = OUTPUT_DIR / "PostalCodeZones_PT.txt"
        pcz_path.write_text("\n".join(build_postal_code_zones_pt_lines()), encoding="utf-8")
        print(f"  PT: {len(pt_lanes)} lanes, {len(pt_cost_blocks)} cost blocks")
        print(f"  Written: {pcz_path}")

    # Write output
    print("\nWriting output...")
    write_domestic_output(sheets, OUTPUT_DIR)
    print("\nDone!")


if __name__ == "__main__":
    main()
