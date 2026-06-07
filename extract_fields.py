import json
import os
from pathlib import Path

INPUT_DIR = Path("input")
OUTPUT_DIR = Path("processing")

OUTPUT_DIR.mkdir(exist_ok=True)


def extract_value(field: dict):
    """Extract the meaningful value from an Azure Document Intelligence field node."""
    if field is None:
        return None

    field_type = field.get("type")

    if field_type == "string":
        val = field.get("valueString")
        if val is None:
            return None
        return val

    if field_type == "number":
        return field.get("valueNumber")

    if field_type == "integer":
        return field.get("valueInteger")

    if field_type == "date":
        return field.get("valueDate")

    if field_type == "array":
        arr = field.get("valueArray", [])
        return [extract_value(item) for item in arr]

    if field_type == "object":
        obj = field.get("valueObject", {})
        result = {}
        for key, sub_field in obj.items():
            val = extract_value(sub_field)
            if val is not None:
                result[key] = val
        return result if result else None

    return field.get("content")


def extract_field_with_confidence(field: dict):
    """Extract value and confidence from a field."""
    value = extract_value(field)
    confidence = field.get("confidence")
    if value is None and confidence is None:
        return None
    return {"value": value, "confidence": confidence}


def process_file(input_path: Path):
    """Process a single Azure Document Intelligence JSON file."""
    with open(input_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    analyze_result = data.get("analyzeResult", {})
    documents = analyze_result.get("documents", [])

    if not documents:
        print(f"  No documents found in {input_path.name}")
        return None

    results = []

    for doc in documents:
        doc_type = doc.get("docType")
        fields = doc.get("fields", {})

        extracted = {"docType": doc_type, "fields": {}}

        for field_name, field_data in fields.items():
            processed = extract_field_with_confidence(field_data)
            if processed and processed["value"] is not None:
                extracted["fields"][field_name] = processed

        results.append(extracted)

    return {
        "sourceFile": input_path.name,
        "modelId": analyze_result.get("modelId"),
        "apiVersion": analyze_result.get("apiVersion"),
        "documents": results,
    }


def main():
    json_files = list(INPUT_DIR.glob("*.json"))
    if not json_files:
        print("No JSON files found in input/")
        return

    print(f"Found {len(json_files)} file(s) to process.\n")

    for input_path in json_files:
        print(f"Processing: {input_path.name}")
        result = process_file(input_path)

        if result is None:
            continue

        output_name = input_path.stem + "_extracted.json"
        output_path = OUTPUT_DIR / output_name

        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(result, f, indent=2, ensure_ascii=False)

        input_size = input_path.stat().st_size / (1024 * 1024)
        output_size = output_path.stat().st_size / (1024 * 1024)
        print(f"  Input:  {input_size:.2f} MB")
        print(f"  Output: {output_size:.2f} MB")
        print(f"  Saved to: {output_path}\n")

    print("Done.")


if __name__ == "__main__":
    main()
