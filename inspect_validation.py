import json

INPUT = r"D:\Desktop\CEILING WORKS_balance_validation.json"

with open(INPUT, encoding="utf-8") as f:
    d = json.load(f)

for block in d["stock_cards"]:
    for side in block["sides"]:

        summary = side["summary"]
        opening = side["opening_balance"]

        # Skip completely empty template cards
        if summary["transaction_count"] == 0:
            continue

        print(
            f'Block {block["block_number"]} | '
            f'Side {side["side_number"]} | '
            f'Item: {side["item"]} | '
            f'Opening: {opening["status"]} {opening["value"]} | '
            f'End: {summary["ending_expected_balance"]} | '
            f'Negative: {summary["negative_balance_events"]} | '
            f'Chronology: {summary["chronology_warning_count"]}'
        )

        for tx in side["accounting_order_transactions"]:
            if tx["status"] not in ("MATCH", "NO_RECORDED_BALANCE"):
                print(
                    f'   ROW {tx["source_row"]} | '
                    f'{tx["date"]} | '
                    f'{tx["type"]} | '
                    f'move={tx["movement"]} | '
                    f'expected={tx["expected_balance_after"]} | '
                    f'recorded={tx["recorded_balance"]} | '
                    f'{tx["status"]}'
                )

        print()