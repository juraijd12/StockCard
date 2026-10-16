from __future__ import annotations

import argparse
import contextlib
import importlib.util
import io
import json
import math
import os
import re
import sys
from collections import Counter, defaultdict
from datetime import date, datetime, time, timezone
from pathlib import Path
from typing import Any

# Explicit imports keep dynamically loaded modules and their dependencies in packaged builds.
import inspect_excel as _packaged_inspect_excel
import validate_stock_cards_v2_3_3 as _packaged_validator

SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm", ".xltx", ".xltm"}
DEFAULT_OUTPUT_NAME = "_stock_card_batch_output"


def load_module(module_name: str, path: Path):
    spec = importlib.util.spec_from_file_location(module_name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Cannot load module: {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module


def get_quantity_value(quantity: Any) -> Any:
    if isinstance(quantity, dict):
        return quantity.get("value")
    return quantity


def get_quantity_unit(quantity: Any) -> Any:
    if isinstance(quantity, dict):
        return quantity.get("unit")
    return None


def to_json_safe(value: Any) -> Any:
    if isinstance(value, (datetime, date, time)):
        return value.isoformat()
    if isinstance(value, float) and not math.isfinite(value):
        return None
    if isinstance(value, dict):
        return {str(key): to_json_safe(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [to_json_safe(item) for item in value]
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    return str(value)


def normalized_transaction(
    tx: dict[str, Any], *, source_file: str, sheet_name: str,
    block_number: int, side_number: int, item: str | None,
    location: str | None, validator: Any,
) -> dict[str, Any]:
    qin = tx.get("quantity_in") or {}
    qout = tx.get("quantity_out") or {}
    tx_type = (tx.get("type") or "UNKNOWN").upper()

    if tx_type == "RECEIVING":
        quantity, unit = get_quantity_value(qin), get_quantity_unit(qin)
    elif tx_type == "ISSUANCE":
        quantity, unit = get_quantity_value(qout), get_quantity_unit(qout)
    else:
        quantity, unit = None, None

    row = tx.get("source_row")
    key = f"{source_file}::{sheet_name}::B{block_number}S{side_number}R{row}"
    return {
        "transaction_id": key,
        "date": tx.get("date"),
        "item_code": None,
        "item_description": item,
        "location": location,
        "unit": unit,
        "transaction_type": tx_type,
        "reference_number": tx.get("reference"),
        "quantity": quantity,
        "quantity_in": get_quantity_value(qin),
        "quantity_out": get_quantity_value(qout),
        "net_movement": validator.movement_of(tx),
        "project": tx.get("issued_to"),
        "supplier": tx.get("provider"),
        "recorded_balance": get_quantity_value(tx.get("recorded_balance")),
        "source_file": source_file,
        "source_sheet": sheet_name,
        "source_block_number": block_number,
        "source_side_number": side_number,
        "source_row": row,
        "raw": to_json_safe(tx.get("raw", {})),
    }


def scan_files(target: Path, output_dir: Path) -> tuple[list[Path], Path]:
    if target.is_file():
        return ([target] if target.suffix.lower() in SUPPORTED_EXTENSIONS else []), target.parent
    if not target.is_dir():
        raise FileNotFoundError(f"Input path does not exist: {target}")

    excluded = output_dir.resolve()
    files = []
    for path in target.rglob("*"):
        if not path.is_file() or path.name.startswith("~$"):
            continue
        if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            continue
        try:
            path.resolve().relative_to(excluded)
            continue
        except ValueError:
            pass
        files.append(path)
    return sorted(files), target


def duplicate_signature(tx: dict[str, Any]) -> tuple[Any, ...] | None:
    reference = re.sub(r"\s+", "", str(tx.get("reference_number") or "")).upper()
    if not reference:
        return None
    item = re.sub(r"\s+", " ", str(tx.get("item_description") or "")).strip().upper()
    return (
        item,
        (tx.get("transaction_type") or "").upper(),
        reference,
        tx.get("date"),
        tx.get("quantity_in"),
        tx.get("quantity_out"),
    )


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Scan Excel stock cards into one normalized dataset and validation report."
    )
    parser.add_argument("input", nargs="+", help="One or more Excel workbooks or folders to scan recursively")
    parser.add_argument(
        "--output",
        help=f"Output folder (default: <input folder>/{DEFAULT_OUTPUT_NAME})",
    )
    args = parser.parse_args()

    targets = [Path(value).expanduser().resolve() for value in args.input]
    output_dir = Path(args.output).expanduser().resolve() if args.output else (
        targets[0].parent / DEFAULT_OUTPUT_NAME if targets[0].is_file() else targets[0] / DEFAULT_OUTPUT_NAME
    )
    file_sets = [scan_files(target, output_dir)[0] for target in targets]
    files = sorted({path for group in file_sets for path in group})
    source_root = Path(os.path.commonpath([str(path.parent) for path in files])) if files else targets[0].parent
    if not files:
        raise SystemExit(f"No supported Excel workbooks found under: {target}")

    output_dir.mkdir(parents=True, exist_ok=True)
    script_dir = Path(__file__).resolve().parent
    extractor = load_module("stock_card_extractor", script_dir / "inspect_excel.py")
    validator = load_module("stock_card_validator", script_dir / "validate_stock_cards_v2_3_3.py")

    file_results = []
    cards = []
    transactions = []
    errors = []
    extraction_warning_count = 0

    for index, workbook_path in enumerate(files, start=1):
        try:
            rel_path = workbook_path.relative_to(source_root).as_posix()
        except ValueError:
            rel_path = workbook_path.name
        print(f"[{index}/{len(files)}] Reading {rel_path}")

        try:
            with contextlib.redirect_stdout(io.StringIO()):
                inspection = extractor.inspect_workbook(str(workbook_path), write_json=False)

            file_tx_count = 0
            file_card_count = 0
            file_active_count = 0
            file_warning_count = 0
            layout_profiles = []
            sheets_without_cards = []
            missing_item_cards = 0
            missing_location_cards = 0
            unknown_transaction_count = 0
            for sheet in inspection.get("sheets", []):
                sheet_name = sheet.get("sheet_name") or ""
                stock_cards = sheet.get("stock_cards", [])
                if not stock_cards:
                    sheets_without_cards.append(sheet_name)
                    layout_profiles.append({
                        "sheet_name": sheet_name,
                        "max_row": sheet.get("max_row"),
                        "max_column": sheet.get("max_column"),
                        "stock_card_blocks": 0,
                        "side_columns": [],
                        "metadata": {"item_missing": 0, "location_missing": 0},
                        "transactions": 0,
                    })
                    continue
                sheet_profile = {
                    "sheet_name": sheet_name,
                    "max_row": sheet.get("max_row"),
                    "max_column": sheet.get("max_column"),
                    "stock_card_blocks": len(stock_cards),
                    "side_columns": sorted({
                        side.get("columns")
                        for block in stock_cards
                        for side in block.get("sides", [])
                        if side.get("columns")
                    }),
                    "metadata": {"item_missing": 0, "location_missing": 0},
                    "transactions": 0,
                }
                for block in stock_cards:
                    for side in block.get("sides", []):
                        result = validator.validate_side(side)
                        block_number = block.get("block_number")
                        side_number = side.get("side_number")
                        card = {
                            "source_file": rel_path,
                            "source_sheet": sheet_name,
                            "block_number": block_number,
                            "block_start_row": block.get("block_start_row"),
                            "block_end_row": block.get("block_end_row"),
                            "header_row": block.get("header_row"),
                            "side_number": side_number,
                            "columns": side.get("columns"),
                            "item": result.get("item"),
                            "location": result.get("location"),
                            "status": result.get("status"),
                            "opening_balance": result.get("opening_balance"),
                            "issues": list(result.get("issues", [])),
                            "extraction_warnings": list(side.get("warnings", [])),
                            "summary": result.get("summary", {}),
                            "negative_balance_events": result.get("negative_balance_events", []),
                            "transactions": result.get("transactions", []),
                        }
                        if side.get("transactions") and not side.get("item"):
                            card["issues"].append("ITEM_IDENTITY_MISSING")
                            card["status"] = "REVIEW_REQUIRED"
                            missing_item_cards += 1
                            sheet_profile["metadata"]["item_missing"] += 1
                        if side.get("transactions") and not side.get("location"):
                            card["issues"].append("LOCATION_MISSING")
                            card["status"] = "REVIEW_REQUIRED"
                            missing_location_cards += 1
                            sheet_profile["metadata"]["location_missing"] += 1

                        cards.append(card)
                        file_card_count += 1
                        if result.get("status") != "EMPTY_TEMPLATE_CARD":
                            file_active_count += 1
                        extraction_warning_count += len(side.get("warnings", []))
                        file_warning_count += len(side.get("warnings", []))

                        for tx in side.get("transactions", []):
                            item = normalized_transaction(
                                tx,
                                source_file=rel_path,
                                sheet_name=sheet_name,
                                block_number=block_number,
                                side_number=side_number,
                                item=result.get("item"),
                                location=result.get("location"),
                                validator=validator,
                            )
                            transactions.append(item)
                            file_tx_count += 1
                            sheet_profile["transactions"] += 1
                            if item["transaction_type"] == "UNKNOWN":
                                unknown_transaction_count += 1

                layout_profiles.append(sheet_profile)

            file_results.append({
                "source_file": rel_path,
                "status": "PROCESSED",
                "sheet_count": len(inspection.get("sheets", [])),
                "stock_card_blocks": inspection.get("summary", {}).get("stock_card_blocks", 0),
                "stock_card_sides": file_card_count,
                "active_sides": file_active_count,
                "transactions": file_tx_count,
                "extraction_warnings": file_warning_count,
                "layout_profiles": layout_profiles,
                "sheets_without_stock_cards": sheets_without_cards,
                "metadata_missing_item_cards": missing_item_cards,
                "metadata_missing_location_cards": missing_location_cards,
                "unknown_transaction_count": unknown_transaction_count,
            })
        except Exception as exc:
            errors.append({"source_file": rel_path, "error": f"{type(exc).__name__}: {exc}"})
            file_results.append({"source_file": rel_path, "status": "FAILED", "error": f"{type(exc).__name__}: {exc}"})
            print(f"  ERROR: {type(exc).__name__}: {exc}")

    duplicate_map: dict[tuple[Any, ...], list[dict[str, Any]]] = defaultdict(list)
    for tx in transactions:
        signature = duplicate_signature(tx)
        if signature is not None:
            duplicate_map[signature].append(tx)
    duplicate_groups = [
        {"signature": list(signature), "occurrences": [
            {
                "source_file": tx["source_file"],
                "source_sheet": tx["source_sheet"],
                "source_block_number": tx["source_block_number"],
                "source_side_number": tx["source_side_number"],
                "source_row": tx["source_row"],
                "transaction_id": tx["transaction_id"],
            } for tx in occurrences
        ]}
        for signature, occurrences in duplicate_map.items()
        if len(occurrences) > 1
    ]

    status_counts = Counter(card["status"] for card in cards)
    issue_counts = Counter(issue for card in cards for issue in card["issues"])
    transaction_type_counts = Counter(tx["transaction_type"] for tx in transactions)
    summary = {
        "files_found": len(files),
        "files_processed": sum(1 for f in file_results if f["status"] == "PROCESSED"),
        "files_failed": len(errors),
        "stock_card_sides": len(cards),
        "active_sides": sum(1 for c in cards if c["status"] != "EMPTY_TEMPLATE_CARD"),
        "empty_template_sides": status_counts.get("EMPTY_TEMPLATE_CARD", 0),
        "fully_validated_sides": status_counts.get("OK", 0),
        "validation_limited_sides": status_counts.get("VALIDATION_LIMITED", 0),
        "review_required_sides": status_counts.get("REVIEW_REQUIRED", 0),
        "transactions": len(transactions),
        "transaction_type_counts": dict(transaction_type_counts),
        "possible_duplicate_groups": len(duplicate_groups),
        "extraction_warnings": extraction_warning_count,
        "metadata_missing_item_cards": sum(
            f.get("metadata_missing_item_cards", 0) for f in file_results
        ),
        "metadata_missing_location_cards": sum(
            f.get("metadata_missing_location_cards", 0) for f in file_results
        ),
        "unknown_transaction_count": sum(
            f.get("unknown_transaction_count", 0) for f in file_results
        ),
        "issue_counts": dict(issue_counts),
    }

    generated_at = datetime.now(timezone.utc).isoformat(timespec="seconds")
    dataset = {
        "schema_version": "normalized-transactions-1.0",
        "generated_at_utc": generated_at,
        "source_root": str(source_root),
        "source_note": "Source Excel files were opened read-only and were not modified.",
        "summary": summary,
        "possible_duplicate_groups": duplicate_groups,
        "transactions": transactions,
    }
    validation_output = {
        "schema_version": "batch-balance-validation-1.0",
        "validator_version": "2.3.3",
        "generated_at_utc": generated_at,
        "source_root": str(source_root),
        "summary": summary,
        "files": file_results,
        "cards": cards,
        "errors": errors,
        "possible_duplicate_groups": duplicate_groups,
    }
    catalog_output = {
        "schema_version": "workbook-layout-catalog-1.0",
        "generated_at_utc": generated_at,
        "source_root": str(source_root),
        "source_note": "Read-only structural inventory for the future GUI and parser selection.",
        "summary": summary,
        "files": file_results,
        "errors": errors,
    }

    normalized_path = output_dir / "normalized_transactions.json"
    validation_path = output_dir / "balance_validation.json"
    catalog_path = output_dir / "layout_catalog.json"
    report_path = output_dir / "batch_report.txt"
    normalized_path.write_text(json.dumps(dataset, indent=2, ensure_ascii=False), encoding="utf-8")
    validation_path.write_text(json.dumps(validation_output, indent=2, ensure_ascii=False), encoding="utf-8")
    catalog_path.write_text(json.dumps(catalog_output, indent=2, ensure_ascii=False), encoding="utf-8")

    lines = [
        "AIYU STOCK CARD BATCH REPORT",
        "=" * 72,
        f"Generated: {generated_at}",
        f"Source: {source_root}",
        "Source workbooks were opened read-only and were not modified.",
        "",
        "SUMMARY",
        "-" * 72,
    ]
    for label, key in [
        ("Excel files found", "files_found"),
        ("Excel files processed", "files_processed"),
        ("Excel files failed", "files_failed"),
        ("Stock-card sides", "stock_card_sides"),
        ("Active sides", "active_sides"),
        ("Empty template sides", "empty_template_sides"),
        ("Fully validated sides", "fully_validated_sides"),
        ("Validation-limited sides", "validation_limited_sides"),
        ("Sides requiring review", "review_required_sides"),
        ("Normalized transactions", "transactions"),
        ("Possible duplicate groups", "possible_duplicate_groups"),
        ("Extraction warnings", "extraction_warnings"),
        ("Cards missing item identity", "metadata_missing_item_cards"),
        ("Cards missing location", "metadata_missing_location_cards"),
        ("Unknown transactions", "unknown_transaction_count"),
    ]:
        lines.append(f"{label + ':':34}{summary[key]}")
    lines.append("")
    lines.append("TRANSACTION TYPES")
    lines.append("-" * 72)
    for kind, count in sorted(transaction_type_counts.items()):
        lines.append(f"{kind:34}{count}")
    lines.append("")
    lines.append("ISSUES")
    lines.append("-" * 72)
    if issue_counts:
        for issue, count in sorted(issue_counts.items()):
            lines.append(f"{issue:34}{count}")
    else:
        lines.append("None.")
    lines.append("")
    lines.append("CARDS REQUIRING REVIEW OR LIMITED VALIDATION")
    lines.append("-" * 72)
    flagged = [c for c in cards if c["status"] in {"REVIEW_REQUIRED", "VALIDATION_LIMITED"}]
    if not flagged:
        lines.append("None.")
    for card in flagged:
        lines.append(
            f"{card['source_file']} | {card['source_sheet']} | "
            f"Block {card['block_number']} side {card['side_number']} | "
            f"{card['item'] or '[ITEM BLANK]'} | {card['status']}"
        )
        lines.append(f"  Issues: {', '.join(card['issues']) or 'None'}")
    if duplicate_groups:
        lines.extend(["", "POSSIBLE DUPLICATE TRANSACTIONS", "-" * 72])
        for group in duplicate_groups:
            lines.append(f"Reference {group['signature'][2]} | item {group['signature'][0]} | {len(group['occurrences'])} occurrences")
    if errors:
        lines.extend(["", "FILES THAT FAILED", "-" * 72])
        for error in errors:
            lines.append(f"{error['source_file']}: {error['error']}")
    lines.extend([
        "",
        "OUTPUTS",
        "-" * 72,
        str(normalized_path),
        str(validation_path),
        str(catalog_path),
        str(report_path),
        "",
        "No source workbook was modified. No separate stock-card Excel files were generated.",
    ])
    report_path.write_text("\n".join(lines), encoding="utf-8")

    print("\nBATCH COMPLETE")
    print(f"Workbooks processed: {summary['files_processed']} / {summary['files_found']}")
    print(f"Transactions normalized: {summary['transactions']}")
    print(f"Cards requiring review: {summary['review_required_sides']}")
    print(f"Validation-limited cards: {summary['validation_limited_sides']}")
    print(f"Possible duplicate groups: {summary['possible_duplicate_groups']}")
    print(f"Output folder: {output_dir}")
    return 0 if not errors else 1


if __name__ == "__main__":
    raise SystemExit(main())
