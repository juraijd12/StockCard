import json

INPUT = r"D:\Desktop\CEILING WORKS_balance_validation.json"

with open(INPUT, encoding="utf-8") as f:
    d = json.load(f)

for block in d["stock_cards"]:
    for side in block["sides"]:

        if side["item"] == "HARDWARE CLOTH 1/8 (10M)":

            print("=" * 80)
            print(f'BLOCK {block["block_number"]} | SIDE {side["side_number"]}')
            print(f'ITEM: {side["item"]}')
            print(f'LOCATION: {side["location"]}')
            print("=" * 80)

            print()
            print("OPENING BALANCE")
            print("-" * 80)
            print(json.dumps(side["opening_balance"], indent=2))

            print()
            print("TRANSACTIONS — ACCOUNTING ORDER")
            print("-" * 80)

            for tx in side["accounting_order_transactions"]:
                print(
                    f'ROW {tx["source_row"]} | '
                    f'DATE {tx["date"]} | '
                    f'TYPE {tx["type"]} | '
                    f'REF {tx["reference"]} | '
                    f'IN {tx["quantity_in"]} | '
                    f'OUT {tx["quantity_out"]} | '
                    f'MOVE {tx["movement"]} | '
                    f'EXPECTED {tx["expected_balance_after"]} | '
                    f'RECORDED {tx["recorded_balance"]} | '
                    f'STATUS {tx["status"]}'
                )

            print()
            print("SUMMARY")
            print("-" * 80)
            print(json.dumps(side["summary"], indent=2))

            break
