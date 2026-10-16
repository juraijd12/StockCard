import json
from pathlib import Path
from datetime import datetime


# ============================================================
# CONFIG
# ============================================================

INPUT_FILE = Path(r"D:\Desktop\CEILING WORKS_inspection.json")
OUTPUT_FILE = Path(r"D:\Desktop\CEILING WORKS_balance_validation.json")

SCHEMA_VERSION = "2.3.1"


# ============================================================
# HELPERS
# ============================================================

def load_json(path):
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def save_json(path, data):
    with path.open("w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def value_from_field(field):
    """
    v2.2 stores quantity/balance fields like:

    {
        "raw": "300",
        "value": 300,
        "unit": null
    }

    Return the numeric value.
    """
    if field is None:
        return None

    if isinstance(field, dict):
        value = field.get("value")
    else:
        value = field

    if value is None:
        return None

    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def parse_date(value):
    if not value:
        return None

    value = str(value)

    # v2.2 normally uses YYYY-MM-DD
    try:
        return datetime.strptime(value[:10], "%Y-%m-%d").date()
    except ValueError:
        return None


def clean_number(value):
    """
    Convert 300.0 -> 300 for cleaner JSON output.
    """
    if value is None:
        return None

    if abs(value - round(value)) < 1e-9:
        return int(round(value))

    return value


# ============================================================
# MOVEMENT
# ============================================================

def calculate_movement(transaction):
    transaction_type = transaction.get("type")

    quantity_in = value_from_field(transaction.get("quantity_in"))
    quantity_out = value_from_field(transaction.get("quantity_out"))

    if transaction_type == "RECEIVING":
        if quantity_in is None:
            return None
        return quantity_in

    if transaction_type == "ISSUANCE":
        if quantity_out is None:
            return None
        return -quantity_out

    if transaction_type == "MIXED":
        in_value = quantity_in if quantity_in is not None else 0
        out_value = quantity_out if quantity_out is not None else 0
        return in_value - out_value

    return None


# ============================================================
# CHRONOLOGY
# ============================================================

def detect_chronology_warnings(transactions):
    """
    Check the ORIGINAL Excel/source order.

    We do not modify source order.

    If a later source row has an earlier date than the previous
    transaction, flag it.
    """

    warnings = []

    previous_date = None
    previous_row = None

    for tx in transactions:
        current_date = parse_date(tx.get("date"))

        if current_date is None:
            continue

        if previous_date is not None and current_date < previous_date:
            warnings.append({
                "source_row": tx.get("source_row"),
                "date": tx.get("date"),
                "previous_source_row": previous_row,
                "previous_date": previous_date.isoformat(),
                "message": (
                    f"Source row {tx.get('source_row')} "
                    f"({tx.get('date')}) appears before the previous "
                    f"transaction date {previous_date.isoformat()} "
                    f"at source row {previous_row}."
                )
            })

        previous_date = current_date
        previous_row = tx.get("source_row")

    return warnings


# ============================================================
# TRANSACTION NORMALIZATION
# ============================================================

def normalize_transaction(transaction):
    quantity_in = value_from_field(transaction.get("quantity_in"))
    quantity_out = value_from_field(transaction.get("quantity_out"))
    recorded_balance = value_from_field(transaction.get("recorded_balance"))

    movement = calculate_movement(transaction)

    tx = {
        "source_row": transaction.get("source_row"),
        "date": transaction.get("date"),
        "type": transaction.get("type"),
        "reference": transaction.get("reference"),
        "receiving_ref": transaction.get("receiving_ref"),
        "provider": transaction.get("provider"),
        "issuance_ref": transaction.get("issuance_ref"),
        "issued_to": transaction.get("issued_to"),

        "quantity_in": clean_number(quantity_in),
        "quantity_out": clean_number(quantity_out),

        "movement": clean_number(movement),

        "recorded_balance": clean_number(recorded_balance),

        "unit_in": (
            transaction.get("quantity_in", {}).get("unit")
            if isinstance(transaction.get("quantity_in"), dict)
            else None
        ),

        "unit_out": (
            transaction.get("quantity_out", {}).get("unit")
            if isinstance(transaction.get("quantity_out"), dict)
            else None
        ),

        "balance_unit": (
            transaction.get("recorded_balance", {}).get("unit")
            if isinstance(transaction.get("recorded_balance"), dict)
            else None
        ),
    }

    return tx


# ============================================================
# ACCOUNTING VALIDATION
# ============================================================

def validate_side(side):
    source_transactions = side.get("transactions", [])

    chronology_warnings = detect_chronology_warnings(source_transactions)

    # Normalize first while preserving original source order.
    normalized_source = [
        normalize_transaction(tx)
        for tx in source_transactions
    ]

    # Accounting order:
    #   1. date
    #   2. source row as deterministic tie-breaker
    #
    # IMPORTANT:
    # This does NOT modify the source data.
    accounting_transactions = sorted(
        normalized_source,
        key=lambda tx: (
            parse_date(tx.get("date")) or datetime.max.date(),
            tx.get("source_row") if tx.get("source_row") is not None else 0
        )
    )

    # --------------------------------------------------------
    # Opening balance inference
    # --------------------------------------------------------

    opening_candidates = []

    cumulative_before = 0

    for tx in accounting_transactions:
        movement = tx.get("movement")
        recorded_balance = tx.get("recorded_balance")

        if movement is not None and recorded_balance is not None:
            candidate = recorded_balance - cumulative_before - movement
            opening_candidates.append({
                "source_row": tx.get("source_row"),
                "date": tx.get("date"),
                "candidate": candidate
            })

        if movement is not None:
            cumulative_before += movement

    opening_status = "UNKNOWN"
    opening_balance = None

    if opening_candidates:
        candidate_values = [
            round(float(x["candidate"]), 9)
            for x in opening_candidates
        ]

        unique_candidates = sorted(set(candidate_values))

        if len(unique_candidates) == 1:
            opening_balance = unique_candidates[0]
            opening_status = "INFERRED"

        else:
            opening_status = "AMBIGUOUS"

    # --------------------------------------------------------
    # Running balance validation
    # --------------------------------------------------------

    expected_balance = opening_balance

    validation_transactions = []

    negative_balance_events = []
    unknown_movement_events = []
    recorded_balance_checks = []

    for tx in accounting_transactions:

        movement = tx.get("movement")
        recorded_balance = tx.get("recorded_balance")

        tx_result = dict(tx)

        tx_result["expected_balance_before"] = (
            clean_number(expected_balance)
            if expected_balance is not None
            else None
        )

        # ----------------------------------------------------
        # Unknown movement
        # ----------------------------------------------------

        if movement is None:
            tx_result["status"] = "UNKNOWN_MOVEMENT"
            tx_result["expected_balance_after"] = None
            tx_result["recorded_balance_discrepancy"] = None

            unknown_movement_events.append({
                "source_row": tx.get("source_row"),
                "date": tx.get("date"),
                "type": tx.get("type")
            })

            validation_transactions.append(tx_result)
            continue

        # ----------------------------------------------------
        # Expected balance
        # ----------------------------------------------------

        if expected_balance is None:
            tx_result["expected_balance_after"] = None

            if recorded_balance is not None:
                tx_result["status"] = "OPENING_BALANCE_UNRESOLVED"
            else:
                tx_result["status"] = "OPENING_BALANCE_UNRESOLVED"

            tx_result["recorded_balance_discrepancy"] = None

        else:
            expected_after = expected_balance + movement

            tx_result["expected_balance_after"] = clean_number(
                expected_after
            )

            # ------------------------------------------------
            # Recorded balance exists
            # ------------------------------------------------

            if recorded_balance is not None:

                discrepancy = recorded_balance - expected_after

                tx_result["recorded_balance_discrepancy"] = clean_number(
                    discrepancy
                )

                if abs(discrepancy) < 1e-9:
                    tx_result["status"] = "MATCH"
                else:
                    tx_result["status"] = "MISMATCH"

                recorded_balance_checks.append({
                    "source_row": tx.get("source_row"),
                    "date": tx.get("date"),
                    "expected_balance": clean_number(expected_after),
                    "recorded_balance": clean_number(recorded_balance),
                    "discrepancy": clean_number(discrepancy)
                })

            # ------------------------------------------------
            # No recorded balance
            # ------------------------------------------------

            else:
                tx_result["recorded_balance_discrepancy"] = None
                tx_result["status"] = "NO_RECORDED_BALANCE"

            # ------------------------------------------------
            # Negative balance
            # ------------------------------------------------

            if expected_after < 0:
                negative_balance_events.append({
                    "source_row": tx.get("source_row"),
                    "date": tx.get("date"),
                    "expected_balance": clean_number(expected_after),
                    "movement": clean_number(movement)
                })

            expected_balance = expected_after

        validation_transactions.append(tx_result)

    # --------------------------------------------------------
    # Summary
    # --------------------------------------------------------

    total_receiving = 0
    total_issuance = 0
    total_movement = 0

    for tx in normalized_source:
        movement = tx.get("movement")

        if tx.get("type") == "RECEIVING":
            qty = tx.get("quantity_in")
            if qty is not None:
                total_receiving += qty

        elif tx.get("type") == "ISSUANCE":
            qty = tx.get("quantity_out")
            if qty is not None:
                total_issuance += qty

        elif tx.get("type") == "MIXED":
            qty_in = tx.get("quantity_in") or 0
            qty_out = tx.get("quantity_out") or 0

            total_receiving += qty_in
            total_issuance += qty_out

        if movement is not None:
            total_movement += movement

    # Last chronological transaction
    last_transaction = (
        accounting_transactions[-1]
        if accounting_transactions
        else None
    )

    ending_expected_balance = expected_balance

    # A "last recorded balance" is informational only.
    # It must NOT automatically be treated as the ending balance.
    last_recorded_balance = None
    last_recorded_balance_source_row = None

    for tx in reversed(accounting_transactions):
        if tx.get("recorded_balance") is not None:
            last_recorded_balance = tx.get("recorded_balance")
            last_recorded_balance_source_row = tx.get("source_row")
            break

    # Unit warnings
    unit_issue_events = []

    for tx in normalized_source:

        units = [
            tx.get("unit_in"),
            tx.get("unit_out"),
            tx.get("balance_unit")
        ]

        units = [u for u in units if u]

        if units:
            normalized_units = [str(u).strip().upper() for u in units]

            if len(set(normalized_units)) > 1:
                unit_issue_events.append({
                    "source_row": tx.get("source_row"),
                    "date": tx.get("date"),
                    "units": units,
                    "message": "Multiple units found within one transaction."
                })

            elif normalized_units[0] not in {
                "PCS",
                "PC",
                "PIECE",
                "PIECES",
                "SET",
                "SETS",
                "BOX",
                "BOXES",
                "MTR",
                "MTRS",
                "METER",
                "METERS",
                "KG",
                "KGS",
                "L",
                "LTR",
                "LITER",
                "LITERS"
            }:
                unit_issue_events.append({
                    "source_row": tx.get("source_row"),
                    "date": tx.get("date"),
                    "units": units,
                    "message": "Unrecognized or uncommon unit."
                })

    # --------------------------------------------------------
    # Side result
    # --------------------------------------------------------

    result = {
        "side_number": side.get("side_number"),
        "columns": side.get("columns"),

        "item": side.get("item"),
        "location": side.get("location"),

        "source_order_transactions": normalized_source,

        "accounting_order_transactions": validation_transactions,

        "chronology_warnings": chronology_warnings,

        "unit_issue_events": unit_issue_events,

        "opening_balance": {
            "status": opening_status,
            "value": clean_number(opening_balance),
            "candidates": [
                {
                    **candidate,
                    "candidate": clean_number(candidate["candidate"])
                }
                for candidate in opening_candidates
            ]
        },

        "summary": {
            "transaction_count": len(normalized_source),

            "total_receiving": clean_number(total_receiving),
            "total_issuance": clean_number(total_issuance),
            "net_movement": clean_number(total_movement),

            "ending_expected_balance": clean_number(
                ending_expected_balance
            ),

            "last_recorded_balance": clean_number(
                last_recorded_balance
            ),

            "last_recorded_balance_source_row":
                last_recorded_balance_source_row,

            "recorded_balance_checks":
                len(recorded_balance_checks),

            "recorded_balance_matches":
                sum(
                    1
                    for x in recorded_balance_checks
                    if abs(x["discrepancy"]) < 1e-9
                ),

            "recorded_balance_mismatches":
                sum(
                    1
                    for x in recorded_balance_checks
                    if abs(x["discrepancy"]) >= 1e-9
                ),

            "negative_balance_events":
                len(negative_balance_events),

            "unknown_movement_events":
                len(unknown_movement_events),

            "chronology_warning_count":
                len(chronology_warnings),

            "unit_issue_count":
                len(unit_issue_events)
        },

        "negative_balance_events": negative_balance_events,

        "unknown_movement_events": unknown_movement_events
    }

    return result


# ============================================================
# MAIN
# ============================================================

def main():

    print("=" * 70)
    print("AIYU STOCK CARD BALANCE VALIDATION")
    print(f"Schema version: {SCHEMA_VERSION}")
    print("=" * 70)
    print()

    if not INPUT_FILE.exists():
        print(f"ERROR: Input file not found:")
        print(INPUT_FILE)
        return

    data = load_json(INPUT_FILE)

    # --------------------------------------------------------
    # Locate stock cards
    # --------------------------------------------------------

    try:
        sheets = data["sheets"]
        sheet = sheets[0]
        stock_cards = sheet["stock_cards"]
    except (KeyError, IndexError, TypeError):
        print("ERROR: Could not find:")
        print("root.sheets[0].stock_cards")
        return

    print(f"Input file: {INPUT_FILE}")
    print(f"Detected blocks: {len(stock_cards)}")
    print()

    # --------------------------------------------------------
    # Process all blocks
    # --------------------------------------------------------

    output_blocks = []

    total_sides = 0
    total_transactions = 0

    total_receiving = 0
    total_issuance = 0
    total_net_movement = 0

    total_recorded_checks = 0
    total_matches = 0
    total_mismatches = 0

    total_negative_events = 0
    total_unknown_events = 0
    total_chronology_warnings = 0
    total_unit_issues = 0

    opening_inferred = 0
    opening_unknown = 0
    opening_ambiguous = 0

    sides_with_transactions = 0
    sides_without_transactions = 0

    for block in stock_cards:

        block_result = {
            "block_number": block.get("block_number"),
            "block_start_row": block.get("block_start_row"),
            "block_end_row": block.get("block_end_row"),
            "header_row": block.get("header_row"),
            "sides": [],
            "warnings": block.get("warnings", [])
        }

        for side in block.get("sides", []):

            total_sides += 1

            side_result = validate_side(side)

            block_result["sides"].append(side_result)

            summary = side_result["summary"]

            tx_count = summary["transaction_count"]

            total_transactions += tx_count

            if tx_count > 0:
                sides_with_transactions += 1
            else:
                sides_without_transactions += 1

            total_receiving += summary["total_receiving"] or 0
            total_issuance += summary["total_issuance"] or 0
            total_net_movement += summary["net_movement"] or 0

            total_recorded_checks += summary["recorded_balance_checks"]
            total_matches += summary["recorded_balance_matches"]
            total_mismatches += summary["recorded_balance_mismatches"]

            total_negative_events += summary["negative_balance_events"]
            total_unknown_events += summary["unknown_movement_events"]

            total_chronology_warnings += summary["chronology_warning_count"]
            total_unit_issues += summary["unit_issue_count"]

            opening_status = side_result["opening_balance"]["status"]

            if opening_status == "INFERRED":
                opening_inferred += 1
            elif opening_status == "AMBIGUOUS":
                opening_ambiguous += 1
            else:
                opening_unknown += 1

        output_blocks.append(block_result)

    # --------------------------------------------------------
    # Overall summary
    # --------------------------------------------------------

    output = {
        "schema_version": SCHEMA_VERSION,

        "source": {
            "input_file": str(INPUT_FILE),
            "sheet": sheet.get("name", "Sheet1")
        },

        "validation": {
            "method": {
                "source_order_preserved": True,
                "accounting_order": "DATE_ASCENDING_THEN_SOURCE_ROW",
                "receiving_movement": "POSITIVE",
                "issuance_movement": "NEGATIVE",
                "mixed_movement": "QUANTITY_IN_MINUS_QUANTITY_OUT"
            },

            "summary": {
                "blocks": len(output_blocks),
                "sides": total_sides,
                "sides_with_transactions": sides_with_transactions,
                "sides_without_transactions": sides_without_transactions,
                "transactions": total_transactions,

                "total_receiving": clean_number(total_receiving),
                "total_issuance": clean_number(total_issuance),
                "net_movement": clean_number(total_net_movement),

                "recorded_balance_checks":
                    total_recorded_checks,

                "recorded_balance_matches":
                    total_matches,

                "recorded_balance_mismatches":
                    total_mismatches,

                "negative_balance_events":
                    total_negative_events,

                "unknown_movement_events":
                    total_unknown_events,

                "chronology_warning_events":
                    total_chronology_warnings,

                "unit_issue_events":
                    total_unit_issues,

                "opening_balances_inferred":
                    opening_inferred,

                "opening_balances_unknown":
                    opening_unknown,

                "opening_balances_ambiguous":
                    opening_ambiguous
            }
        },

        "stock_cards": output_blocks
    }

    save_json(OUTPUT_FILE, output)

    # --------------------------------------------------------
    # Console report
    # --------------------------------------------------------

    print("VALIDATION COMPLETE")
    print("-" * 70)

    print(f"Blocks:                    {len(output_blocks)}")
    print(f"Sides:                     {total_sides}")
    print(f"Sides with transactions:   {sides_with_transactions}")
    print(f"Sides without transactions: {sides_without_transactions}")
    print(f"Transactions:              {total_transactions}")
    print()

    print(f"Total receiving:           {clean_number(total_receiving)}")
    print(f"Total issuance:            {clean_number(total_issuance)}")
    print(f"Net movement:              {clean_number(total_net_movement)}")
    print()

    print(f"Recorded balance checks:   {total_recorded_checks}")
    print(f"Recorded balance matches:  {total_matches}")
    print(f"Recorded balance mismatches:{total_mismatches}")
    print()

    print(f"Negative balance events:   {total_negative_events}")
    print(f"Unknown movement events:   {total_unknown_events}")
    print(f"Chronology warnings:       {total_chronology_warnings}")
    print(f"Unit issues:               {total_unit_issues}")
    print()

    print(f"Opening inferred:          {opening_inferred}")
    print(f"Opening unknown:           {opening_unknown}")
    print(f"Opening ambiguous:         {opening_ambiguous}")
    print()

    print(f"Output:")
    print(OUTPUT_FILE)

    print("=" * 70)


if __name__ == "__main__":
    main()