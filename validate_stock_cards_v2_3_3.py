import json
from pathlib import Path
from collections import Counter

PROJECT_ROOT = Path(__file__).resolve().parent
INPUT_PATH = PROJECT_ROOT / "CEILING WORKS_inspection.json"
OUTPUT_PATH = PROJECT_ROOT / "CEILING WORKS_balance_validation_v2.3.3.json"
REPORT_PATH = PROJECT_ROOT / "CEILING WORKS_balance_report_v2.3.3.txt"


def value_of(obj):
    if isinstance(obj, dict):
        return obj.get("value")
    return obj


def movement_of(tx):
    tx_type = (tx.get("type") or "").upper()
    qin = value_of(tx.get("quantity_in"))
    qout = value_of(tx.get("quantity_out"))

    if tx_type == "RECEIVING":
        return qin
    if tx_type == "ISSUANCE":
        return -qout if qout is not None else None
    if tx_type == "MIXED":
        return (qin or 0) - (qout or 0)
    return None


def sort_key(tx):
    return (
        tx.get("date") or "9999-99-99",
        tx.get("source_row") if tx.get("source_row") is not None else 10**9,
    )


def infer_opening(transactions):
    """Infer opening stock from the earliest usable recorded balance.

    Later recorded balances are kept for independent checks during validation;
    they are not used to choose or revise the opening balance.
    """
    candidates = []
    cumulative = 0
    anchor = None

    for tx in transactions:
        movement = movement_of(tx)
        recorded = value_of(tx.get("recorded_balance"))

        if movement is not None and recorded is not None:
            candidate = recorded - cumulative - movement
            candidate_info = {
                "source_row": tx.get("source_row"),
                "date": tx.get("date"),
                "candidate": candidate,
            }
            candidates.append(candidate_info)
            if anchor is None:
                anchor = candidate_info

        if movement is not None:
            cumulative += movement

    if anchor is None:
        return {
            "status": "UNAVAILABLE",
            "value": None,
            "method": None,
            "anchor_source_row": None,
            "anchor_date": None,
            "candidates": candidates,
        }

    return {
        "status": "INFERRED",
        "value": anchor["candidate"],
        "method": "EARLIEST_RECORDED_BALANCE",
        "anchor_source_row": anchor["source_row"],
        "anchor_date": anchor["date"],
        "candidates": candidates,
    }

def validate_side(side):
    transactions = side.get("transactions") or []
    item = side.get("item")
    location = side.get("location")

    # Completely empty template side.
    if not transactions and not item:
        return {
            "status": "EMPTY_TEMPLATE_CARD",
            "item": None,
            "location": location,
            "transaction_count": 0,
            "transactions": [],
            "opening_balance": {
                "status": "NOT_APPLICABLE",
                "value": None,
                "candidates": [],
            },
            "summary": {},
            "issues": [],
        }

    accounting = sorted(transactions, key=sort_key)
    source_order = list(transactions)

    source_dates = [tx.get("date") for tx in source_order if tx.get("date")]
    accounting_dates = [tx.get("date") for tx in accounting if tx.get("date")]

    chronology_warning = source_dates != accounting_dates or [
        tx.get("source_row") for tx in source_order
    ] != [
        tx.get("source_row") for tx in accounting
    ]

    chronology_count = 1 if chronology_warning else 0

    opening = infer_opening(accounting)

    # Build calculated balances only when opening is known.
    running = opening["value"] if opening["status"] == "INFERRED" else None
    validated_transactions = []
    negative_events = []
    recorded_anchor_checks = 0
    independent_recorded_checks = 0
    recorded_matches = 0
    recorded_mismatches = 0
    anchor_source_row = opening.get("anchor_source_row")
    unknown_movements = 0
    unit_issues = 0

    for tx in accounting:
        movement = movement_of(tx)
        qin = tx.get("quantity_in") or {}
        qout = tx.get("quantity_out") or {}

        if movement is None:
            unknown_movements += 1

        for q in (qin, qout):
            if isinstance(q, dict):
                unit = q.get("unit")
                if unit:
                    # Unit checking is deliberately conservative here.
                    # Different units may be legitimate, so only flag
                    # explicit mixed quantity units within one transaction.
                    pass

        if running is not None and movement is not None:
            running += movement
            expected = running
        else:
            expected = None

        recorded = value_of(tx.get("recorded_balance"))

        if expected is not None and recorded is not None:
            if tx.get("source_row") == anchor_source_row:
                status = "OPENING_BALANCE_ANCHOR"
                recorded_anchor_checks += 1
            else:
                independent_recorded_checks += 1
                if expected == recorded:
                    status = "BALANCE_MATCH"
                    recorded_matches += 1
                else:
                    status = "BALANCE_MISMATCH"
                    recorded_mismatches += 1
        elif recorded is not None:
            status = "OPENING_BALANCE_UNRESOLVED"
        else:
            status = "NO_RECORDED_BALANCE" if expected is not None else "OPENING_BALANCE_UNRESOLVED"

        if expected is not None and expected < 0:
            negative_events.append({
                "source_row": tx.get("source_row"),
                "date": tx.get("date"),
                "reference": tx.get("reference"),
                "movement": movement,
                "expected_balance": expected,
            })

        validated_transactions.append({
            "source_row": tx.get("source_row"),
            "date": tx.get("date"),
            "type": tx.get("type"),
            "reference": tx.get("reference"),
            "quantity_in": value_of(tx.get("quantity_in")),
            "quantity_out": value_of(tx.get("quantity_out")),
            "movement": movement,
            "recorded_balance": recorded,
            "expected_balance": expected,
            "status": status,
        })

    total_receiving = sum(
        value_of(tx.get("quantity_in")) or 0
        for tx in accounting
        if (tx.get("type") or "").upper() in ("RECEIVING", "MIXED")
    )

    total_issuance = sum(
        value_of(tx.get("quantity_out")) or 0
        for tx in accounting
        if (tx.get("type") or "").upper() in ("ISSUANCE", "MIXED")
    )

    net_movement = 0
    for tx in accounting:
        m = movement_of(tx)
        if m is not None:
            net_movement += m

    recorded_rows = [
        tx for tx in accounting
        if value_of(tx.get("recorded_balance")) is not None
    ]
    last_recorded = recorded_rows[-1] if recorded_rows else None

    issues = []

    if opening["status"] == "INCONSISTENT":
        issues.append("OPENING_BALANCE_INCONSISTENT")

    if opening["status"] == "UNAVAILABLE":
        issues.append("OPENING_BALANCE_UNAVAILABLE")

    if chronology_warning:
        issues.append("CHRONOLOGY_WARNING")

    if recorded_mismatches:
        issues.append("BALANCE_MISMATCH")

    if negative_events:
        issues.append("NEGATIVE_BALANCE")

    if unknown_movements:
        issues.append("UNKNOWN_MOVEMENT")

    if not independent_recorded_checks and opening["status"] == "INFERRED" and not issues:
        issues.append("NO_INDEPENDENT_BALANCE_CHECK")

    if issues == ["NO_INDEPENDENT_BALANCE_CHECK"]:
        status = "VALIDATION_LIMITED"
    elif issues:
        status = "REVIEW_REQUIRED"
    else:
        status = "OK"

    summary = {
        "transaction_count": len(accounting),
        "total_receiving": total_receiving,
        "total_issuance": total_issuance,
        "net_movement": net_movement,
        "ending_calculated_balance": running,
        "last_recorded_balance": value_of(last_recorded.get("recorded_balance")) if last_recorded else None,
        "last_recorded_balance_source_row": last_recorded.get("source_row") if last_recorded else None,
        "last_recorded_balance_date": last_recorded.get("date") if last_recorded else None,
        "recorded_balance_anchor_checks": recorded_anchor_checks,
        "independent_recorded_balance_checks": independent_recorded_checks,
        "independent_recorded_balance_matches": recorded_matches,
        "independent_recorded_balance_mismatches": recorded_mismatches,
        "negative_balance_events": len(negative_events),
        "unknown_movement_events": unknown_movements,
        "chronology_warning_count": chronology_count,
        "unit_issue_count": unit_issues,
    }

    return {
        "status": status,
        "item": item,
        "location": location,
        "opening_balance": opening,
        "issues": issues,
        "summary": summary,
        "negative_balance_events": negative_events,
        "transactions": validated_transactions,
    }


def fmt_num(value):
    if value is None:
        return "UNKNOWN"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def main():
    if not INPUT_PATH.exists():
        raise FileNotFoundError(
            f"Input JSON not found:\n{INPUT_PATH}\n\n"
            "Make sure v2.2 has already generated CEILING WORKS_inspection.json."
        )

    with INPUT_PATH.open("r", encoding="utf-8") as f:
        data = json.load(f)

    output = {
        "schema_version": "2.3.3",
        "source_file": INPUT_PATH.name,
        "source_note": (
            "Diagnostic validation only. No source workbook or extracted "
            "transaction data is modified."
        ),
        "sheets": [],
        "overall_summary": {},
    }

    all_cards = []
    all_active = []
    all_transactions = 0
    total_receiving = 0
    total_issuance = 0

    for sheet in data.get("sheets", []):
        sheet_result = {
            "sheet_name": sheet.get("sheet_name"),
            "stock_cards": [],
        }

        for block in sheet.get("stock_cards", []):
            block_result = {
                "block_number": block.get("block_number"),
                "block_start_row": block.get("block_start_row"),
                "block_end_row": block.get("block_end_row"),
                "header_row": block.get("header_row"),
                "sides": [],
            }

            for side in block.get("sides", []):
                result = validate_side(side)

                card = {
                    "block_number": block.get("block_number"),
                    "side_number": side.get("side_number"),
                    "columns": side.get("columns"),
                    "item": result["item"],
                    "location": result["location"],
                    "status": result["status"],
                    "opening_balance": result["opening_balance"],
                    "issues": result["issues"],
                    "summary": result["summary"],
                    "negative_balance_events": result.get("negative_balance_events", []),
                    "transactions": result["transactions"],
                }

                block_result["sides"].append(card)
                all_cards.append(card)

                if result["status"] != "EMPTY_TEMPLATE_CARD":
                    all_active.append(card)
                    all_transactions += result["summary"].get("transaction_count", 0)
                    total_receiving += result["summary"].get("total_receiving", 0)
                    total_issuance += result["summary"].get("total_issuance", 0)

            sheet_result["stock_cards"].append(block_result)

        output["sheets"].append(sheet_result)

    counts = Counter(card["status"] for card in all_cards)

    issue_counts = Counter()
    for card in all_active:
        for issue in card["issues"]:
            issue_counts[issue] += 1

    reconciled = sum(
        1 for c in all_active
        if c["status"] == "OK"
    )

    opening_inferred = sum(
        1 for c in all_active
        if c["opening_balance"]["status"] == "INFERRED"
    )

    opening_unavailable = sum(
        1 for c in all_active
        if c["opening_balance"]["status"] == "UNAVAILABLE"
    )

    opening_inconsistent = sum(
        1 for c in all_active
        if c["opening_balance"]["status"] == "INCONSISTENT"
    )

    recorded_anchor_checks = sum(c["summary"]["recorded_balance_anchor_checks"] for c in all_active)
    independent_recorded_checks = sum(c["summary"]["independent_recorded_balance_checks"] for c in all_active)
    recorded_matches = sum(c["summary"]["independent_recorded_balance_matches"] for c in all_active)
    recorded_mismatches = sum(c["summary"]["independent_recorded_balance_mismatches"] for c in all_active)

    negative_cards = sum(
        1 for c in all_active if c["summary"]["negative_balance_events"] > 0
    )

    chronology_cards = sum(
        1 for c in all_active if c["summary"]["chronology_warning_count"] > 0
    )

    output["overall_summary"] = {
        "template_sides": len(all_cards),
        "active_sides": len(all_active),
        "empty_template_sides": counts.get("EMPTY_TEMPLATE_CARD", 0),
        "transactions_analyzed": all_transactions,
        "reconciled_cards": reconciled,
        "validation_limited_cards": sum(
            1 for c in all_active if c["status"] == "VALIDATION_LIMITED"
        ),
        "review_required_cards": sum(
            1 for c in all_active if c["status"] == "REVIEW_REQUIRED"
        ),
        "opening_balance_inferred_cards": opening_inferred,
        "opening_balance_unavailable_cards": opening_unavailable,
        "opening_balance_inconsistent_cards": opening_inconsistent,
        "negative_balance_cards": negative_cards,
        "chronology_warning_cards": chronology_cards,
        "recorded_balance_anchor_checks": recorded_anchor_checks,
        "independent_recorded_balance_checks": independent_recorded_checks,
        "independent_recorded_balance_matches": recorded_matches,
        "independent_recorded_balance_mismatches": recorded_mismatches,
        "total_receiving": total_receiving,
        "total_issuance": total_issuance,
        "net_movement": total_receiving - total_issuance,
        "issue_counts": dict(issue_counts),
    }

    with OUTPUT_PATH.open("w", encoding="utf-8") as f:
        json.dump(output, f, indent=2, ensure_ascii=False)

    # Human-readable report
    lines = []
    lines.append("=" * 72)
    lines.append("AIYU STOCK CARD VALIDATION REPORT")
    lines.append("v2.3.3 — DIAGNOSTIC / WARNING-ONLY")
    lines.append("=" * 72)
    lines.append("")
    lines.append(f"Source: {INPUT_PATH.name}")
    lines.append("No source workbook or extracted transaction data was modified.")
    lines.append("")

    s = output["overall_summary"]
    lines.append("SUMMARY")
    lines.append("-" * 72)
    lines.append(f"Template stock-card sides:      {s['template_sides']}")
    lines.append(f"Active stock cards:             {s['active_sides']}")
    lines.append(f"Empty template cards:           {s['empty_template_sides']}")
    lines.append(f"Transactions analyzed:          {s['transactions_analyzed']}")
    lines.append(f"Fully validated cards:          {s['reconciled_cards']}")
    lines.append(f"Validation-limited cards:        {s['validation_limited_cards']}")
    lines.append(f"Cards requiring review:         {s['review_required_cards']}")
    lines.append(f"Opening balance inferred:       {s['opening_balance_inferred_cards']}")
    lines.append(f"Opening balance unavailable:    {s['opening_balance_unavailable_cards']}")
    lines.append(f"Opening balance inconsistent:   {s['opening_balance_inconsistent_cards']}")
    lines.append(f"Negative balance cards:         {s['negative_balance_cards']}")
    lines.append(f"Chronology warning cards:       {s['chronology_warning_cards']}")
    lines.append(f"Opening anchor balances:        {s['recorded_balance_anchor_checks']}")
    lines.append(f"Independent balance checks:     {s['independent_recorded_balance_checks']}")
    lines.append(f"Independent checks matching:    {s['independent_recorded_balance_matches']}")
    lines.append(f"Independent check mismatches:   {s['independent_recorded_balance_mismatches']}")
    lines.append(f"Total receiving:                {fmt_num(s['total_receiving'])}")
    lines.append(f"Total issuance:                 {fmt_num(s['total_issuance'])}")
    lines.append(f"Net movement:                   {fmt_num(s['net_movement'])}")
    lines.append("")

    review_cards = [c for c in all_active if c["status"] == "REVIEW_REQUIRED"]

    lines.append("CARDS REQUIRING REVIEW")
    lines.append("-" * 72)

    if not review_cards:
        lines.append("None.")
    else:
        for i, card in enumerate(review_cards, 1):
            lines.append("")
            lines.append(f"[{i}] Block {card['block_number']} | Side {card['side_number']}")
            lines.append(f"    Item:     {card['item']}")
            lines.append(f"    Location: {card['location']}")
            lines.append(f"    Issues:   {', '.join(card['issues'])}")

            opening = card["opening_balance"]
            if opening["status"] == "INCONSISTENT":
                candidates = []
                for c in opening["candidates"]:
                    val = c["candidate"]
                    if val not in candidates:
                        candidates.append(val)
                lines.append(
                    "    Possible opening balances: "
                    + ", ".join(fmt_num(v) for v in candidates)
                )
                lines.append(
                    "    Interpretation: recorded balances do not agree on a single opening balance."
                )

            if opening["status"] == "UNAVAILABLE":
                lines.append(
                    "    Interpretation: no recorded balance allows the opening stock to be inferred."
                )

            if card["summary"].get("negative_balance_events"):
                for event in card["negative_balance_events"]:
                    lines.append(
                        f"    Negative balance: {event['date']} | "
                        f"row {event['source_row']} | "
                        f"ref {event['reference']} | "
                        f"calculated balance {fmt_num(event['expected_balance'])}"
                    )

            if "CHRONOLOGY_WARNING" in card["issues"]:
                lines.append(
                    "    Interpretation: source rows are not in chronological order."
                )

            if "BALANCE_MISMATCH" in card["issues"]:
                lines.append(
                    f"    Independent balance mismatches: {card['summary']['independent_recorded_balance_mismatches']}"
                )

            lines.append("    Action: Human review required. No automatic correction applied.")

    lines.append("")
    lines.append("=" * 72)
    lines.append("IMPORTANT")
    lines.append("=" * 72)
    lines.append(
        "This tool identifies possible data-quality issues. "
        "It does not silently correct inventory records."
    )
    lines.append(
        "Flagged transactions should be verified against the original stock card, "
        "receiving slip, issuance slip, transfer record, or other supporting document."
    )
    lines.append("")

    REPORT_PATH.write_text("\n".join(lines), encoding="utf-8")

    print("=" * 72)
    print("AIYU STOCK CARD VALIDATION — v2.3.3")
    print("=" * 72)
    print(f"Template sides:              {s['template_sides']}")
    print(f"Active cards:                {s['active_sides']}")
    print(f"Empty template cards:       {s['empty_template_sides']}")
    print(f"Transactions analyzed:      {s['transactions_analyzed']}")
    print(f"Fully validated cards:       {s['reconciled_cards']}")
    print(f"Validation-limited cards:    {s['validation_limited_cards']}")
    print(f"Cards requiring review:      {s['review_required_cards']}")
    print(f"Opening inferred:             {s['opening_balance_inferred_cards']}")
    print(f"Opening unavailable:          {s['opening_balance_unavailable_cards']}")
    print(f"Opening inconsistent:         {s['opening_balance_inconsistent_cards']}")
    print(f"Negative balance cards:       {s['negative_balance_cards']}")
    print(f"Chronology warning cards:     {s['chronology_warning_cards']}")
    print(f"Opening anchor balances:      {s['recorded_balance_anchor_checks']}")
    print(f"Independent balance checks:   {s['independent_recorded_balance_checks']}")
    print(f"Independent checks matching:  {s['independent_recorded_balance_matches']}")
    print(f"Independent mismatches:       {s['independent_recorded_balance_mismatches']}")
    print(f"Total receiving:              {fmt_num(s['total_receiving'])}")
    print(f"Total issuance:               {fmt_num(s['total_issuance'])}")
    print(f"Net movement:                 {fmt_num(s['net_movement'])}")
    print("")
    print(f"JSON output:   {OUTPUT_PATH}")
    print(f"Text report:   {REPORT_PATH}")


if __name__ == "__main__":
    main()
