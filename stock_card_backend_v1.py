from __future__ import annotations

import argparse
import json
import sqlite3
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


PROJECT_ROOT = Path(__file__).resolve().parent
DEFAULT_BATCH = PROJECT_ROOT / "04_System_Data" / "wide_batch_output"
DEFAULT_DB = PROJECT_ROOT / "04_System_Data" / "stock_card_backend.sqlite3"


def connect(db_path: Path) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def schema(conn: sqlite3.Connection) -> None:
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS materials (
            id INTEGER PRIMARY KEY,
            name TEXT NOT NULL,
            normalized_name TEXT NOT NULL,
            location TEXT,
            source_count INTEGER NOT NULL DEFAULT 0,
            transaction_count INTEGER NOT NULL DEFAULT 0,
            last_status TEXT,
            UNIQUE(normalized_name, location)
        );
        CREATE TABLE IF NOT EXISTS sources (
            id INTEGER PRIMARY KEY,
            material_id INTEGER NOT NULL REFERENCES materials(id),
            source_file TEXT NOT NULL,
            source_sheet TEXT,
            block_number INTEGER,
            side_number INTEGER,
            columns TEXT,
            status TEXT,
            issues_json TEXT NOT NULL DEFAULT '[]',
            UNIQUE(material_id, source_file, source_sheet, block_number, side_number)
        );
        CREATE TABLE IF NOT EXISTS transactions (
            id INTEGER PRIMARY KEY,
            material_id INTEGER NOT NULL REFERENCES materials(id),
            source_id INTEGER REFERENCES sources(id),
            origin TEXT NOT NULL,
            status TEXT NOT NULL,
            date TEXT,
            transaction_type TEXT,
            reference_number TEXT,
            quantity REAL,
            unit TEXT,
            quantity_in REAL,
            quantity_out REAL,
            net_movement REAL,
            supplier TEXT,
            project TEXT,
            recorded_balance REAL,
            source_row INTEGER,
            notes TEXT,
            created_at TEXT NOT NULL,
            UNIQUE(origin, source_id, source_row, reference_number, date, quantity)
        );
        CREATE TABLE IF NOT EXISTS issues (
            id INTEGER PRIMARY KEY,
            material_id INTEGER NOT NULL REFERENCES materials(id),
            source_id INTEGER REFERENCES sources(id),
            issue_code TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'OPEN',
            details TEXT,
            created_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS metadata (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS material_aliases (
            id INTEGER PRIMARY KEY,
            material_id INTEGER NOT NULL REFERENCES materials(id),
            alias TEXT NOT NULL,
            normalized_alias TEXT NOT NULL,
            UNIQUE(material_id, normalized_alias)
        );
        CREATE TABLE IF NOT EXISTS projects (
            id INTEGER PRIMARY KEY,
            name TEXT NOT NULL UNIQUE,
            normalized_name TEXT NOT NULL UNIQUE,
            created_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS documents (
            id INTEGER PRIMARY KEY,
            document_type TEXT NOT NULL,
            document_number TEXT,
            date_value TEXT NOT NULL,
            project_site TEXT,
            po_number TEXT,
            supplier TEXT,
            dr_invoice_number TEXT,
            vehicle TEXT,
            driver TEXT,
            received_by TEXT,
            released_by TEXT,
            issues TEXT,
            remarks TEXT,
            status TEXT NOT NULL DEFAULT 'DRAFT',
            created_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS document_lines (
            id INTEGER PRIMARY KEY,
            document_id INTEGER NOT NULL REFERENCES documents(id) ON DELETE CASCADE,
            material_id INTEGER NOT NULL REFERENCES materials(id),
            quantity REAL NOT NULL,
            unit TEXT,
            unit_price REAL,
            total_amount REAL,
            notes TEXT
        );
        CREATE INDEX IF NOT EXISTS idx_materials_name ON materials(normalized_name);
        CREATE INDEX IF NOT EXISTS idx_materials_location ON materials(location);
        CREATE INDEX IF NOT EXISTS idx_transactions_material_date ON transactions(material_id, date);
        CREATE INDEX IF NOT EXISTS idx_transactions_origin_status ON transactions(origin, status);
        CREATE INDEX IF NOT EXISTS idx_issues_status ON issues(status);
        CREATE INDEX IF NOT EXISTS idx_projects_name ON projects(normalized_name);
        """
    )
    columns = {row[1] for row in conn.execute("PRAGMA table_info(materials)")}
    if "description" not in columns:
        conn.execute("ALTER TABLE materials ADD COLUMN description TEXT")
    try:
        conn.execute("CREATE VIRTUAL TABLE IF NOT EXISTS materials_fts USING fts5(material_id UNINDEXED, name, location, aliases)")
    except sqlite3.OperationalError:
        pass
    now = datetime.now(timezone.utc).isoformat(timespec="seconds")
    conn.execute("INSERT OR IGNORE INTO projects(name, normalized_name, created_at) SELECT DISTINCT project, UPPER(TRIM(project)), ? FROM transactions WHERE TRIM(COALESCE(project,'')) <> ''", (now,))
    conn.execute("INSERT OR IGNORE INTO projects(name, normalized_name, created_at) SELECT DISTINCT project_site, UPPER(TRIM(project_site)), ? FROM documents WHERE TRIM(COALESCE(project_site,'')) <> ''", (now,))
    conn.commit()


def register_project(db_path: Path, name: str | None) -> None:
    clean = (name or "").strip()
    conn = connect(db_path)
    schema(conn)
    if not clean:
        conn.close()
        return
    conn.execute("INSERT OR IGNORE INTO projects(name, normalized_name, created_at) VALUES (?, ?, ?)", (clean, clean.upper(), datetime.now(timezone.utc).isoformat(timespec="seconds")))
    conn.commit()
    conn.close()


def clean_name(value: Any) -> str:
    return " ".join(str(value or "").split()).strip()


def norm_name(value: Any) -> str:
    return clean_name(value).upper()


def number(value: Any) -> float | None:
    if value is None:
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def upsert_material(conn: sqlite3.Connection, name: str | None, location: str | None) -> int:
    display = clean_name(name) or "[ITEM IDENTITY MISSING]"
    loc = clean_name(location) or None
    key = norm_name(display)
    row = conn.execute(
        "SELECT id FROM materials WHERE normalized_name=? AND location IS ?",
        (key, loc),
    ).fetchone()
    if row:
        return int(row["id"])
    cur = conn.execute(
        "INSERT INTO materials(name, normalized_name, location) VALUES (?, ?, ?)",
        (display, key, loc),
    )
    return int(cur.lastrowid)


def build_database(batch_dir: Path, db_path: Path) -> dict[str, int]:
    validation_path = batch_dir / "balance_validation.json"
    normalized_path = batch_dir / "normalized_transactions.json"
    if not validation_path.exists() or not normalized_path.exists():
        raise FileNotFoundError(
            f"Batch outputs not found in {batch_dir}. Run stock_card_batch_v1.py first."
        )
    validation = json.loads(validation_path.read_text(encoding="utf-8"))
    normalized = json.loads(normalized_path.read_text(encoding="utf-8"))
    conn = connect(db_path)
    schema(conn)
    conn.execute(
        "INSERT INTO metadata(key, value) VALUES ('source_root', ?) "
        "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
        (str(normalized.get("source_root") or ""),),
    )
    # Refresh only the imported Excel layer. Paper entries belong to the
    # backlog workflow and must survive a later workbook rescan.
    conn.execute("DELETE FROM issues")
    conn.execute("DELETE FROM transactions WHERE origin='EXCEL'")
    conn.execute("DELETE FROM sources")
    conn.execute(
        "DELETE FROM materials WHERE id NOT IN "
        "(SELECT DISTINCT material_id FROM transactions WHERE origin='PAPER')"
    )
    now = datetime.now(timezone.utc).isoformat(timespec="seconds")
    source_ids: dict[tuple[Any, ...], int] = {}
    material_ids: dict[tuple[str, str | None], int] = {}
    card_count = 0
    issue_count = 0
    for card in validation.get("cards", []):
        material_key = (norm_name(card.get("item") or "[ITEM IDENTITY MISSING]"), clean_name(card.get("location")) or None)
        material_id = material_ids.setdefault(
            material_key,
            upsert_material(conn, card.get("item"), card.get("location")),
        )
        source_key = (
            material_id, card.get("source_file"), card.get("source_sheet"),
            card.get("block_number"), card.get("side_number"),
        )
        cur = conn.execute(
            "INSERT INTO sources(material_id, source_file, source_sheet, block_number, side_number, columns, status, issues_json) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
            (material_id, card.get("source_file"), card.get("source_sheet"), card.get("block_number"), card.get("side_number"), card.get("columns"), card.get("status"), json.dumps(card.get("issues", []))),
        )
        source_id = int(cur.lastrowid)
        source_ids[source_key] = source_id
        card_count += 1
        for issue in card.get("issues", []):
            conn.execute(
                "INSERT INTO issues(material_id, source_id, issue_code, details, created_at) VALUES (?, ?, ?, ?, ?)",
                (material_id, source_id, issue, None, now),
            )
            issue_count += 1

    tx_count = 0
    for tx in normalized.get("transactions", []):
        material_key = (norm_name(tx.get("item_description") or "[ITEM IDENTITY MISSING]"), clean_name(tx.get("location")) or None)
        material_id = material_ids.setdefault(
            material_key,
            upsert_material(conn, tx.get("item_description"), tx.get("location")),
        )
        source_key = (
            material_id, tx.get("source_file"), tx.get("source_sheet"),
            tx.get("source_block_number"), tx.get("source_side_number"),
        )
        source_id = source_ids.get(source_key)
        conn.execute(
            "INSERT OR IGNORE INTO transactions(material_id, source_id, origin, status, date, transaction_type, reference_number, quantity, unit, quantity_in, quantity_out, net_movement, supplier, project, recorded_balance, source_row, notes, created_at) VALUES (?, ?, 'EXCEL', 'EXTRACTED', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (material_id, source_id, tx.get("date"), tx.get("transaction_type"), tx.get("reference_number"), number(tx.get("quantity")), tx.get("unit"), number(tx.get("quantity_in")), number(tx.get("quantity_out")), number(tx.get("net_movement")), tx.get("supplier"), tx.get("project"), number(tx.get("recorded_balance")), tx.get("source_row"), None, now),
        )
        tx_count += 1
    conn.execute("""
        UPDATE materials SET source_count=(SELECT COUNT(*) FROM sources WHERE sources.material_id=materials.id),
        transaction_count=(SELECT COUNT(*) FROM transactions WHERE transactions.material_id=materials.id),
        last_status=(SELECT status FROM sources WHERE sources.material_id=materials.id ORDER BY id DESC LIMIT 1)
    """)
    try:
        conn.execute("DELETE FROM materials_fts")
        conn.execute("""
            INSERT INTO materials_fts(material_id, name, location, aliases)
            SELECT m.id, m.name, COALESCE(m.location, ''),
                   COALESCE((SELECT group_concat(alias, ' ') FROM material_aliases a WHERE a.material_id=m.id), '')
            FROM materials m
        """)
    except sqlite3.OperationalError:
        pass
    conn.commit()
    conn.close()
    return {"materials": len(material_ids), "sources": card_count, "transactions": tx_count, "issues": issue_count}


def search(db_path: Path, term: str) -> None:
    conn = connect(db_path)
    rows = conn.execute(
        "SELECT id, name, location, source_count, transaction_count, last_status FROM materials WHERE normalized_name LIKE ? ORDER BY name",
        (f"%{norm_name(term)}%",),
    ).fetchall()
    for row in rows:
        print(f"{row['id']} | {row['name']} | {row['location'] or '[location blank]'} | sources={row['source_count']} transactions={row['transaction_count']} status={row['last_status']}")
    print(f"Matches: {len(rows)}")
    conn.close()


def show(db_path: Path, material_id: int) -> None:
    conn = connect(db_path)
    material = conn.execute("SELECT * FROM materials WHERE id=?", (material_id,)).fetchone()
    if not material:
        raise SystemExit(f"Material id not found: {material_id}")
    print(f"MATERIAL {material['id']}: {material['name']} | {material['location'] or '[location blank]'}")
    print(f"Sources: {material['source_count']} | Transactions: {material['transaction_count']} | Status: {material['last_status']}")
    print("SOURCES")
    for row in conn.execute("SELECT * FROM sources WHERE material_id=? ORDER BY id", (material_id,)):
        print(f"  {row['source_file']} | {row['source_sheet']} | block {row['block_number']} side {row['side_number']} | {row['status']} | {row['issues_json']}")
    print("TRANSACTIONS")
    for row in conn.execute("SELECT * FROM transactions WHERE material_id=? ORDER BY date, id", (material_id,)):
        print(
            f"  {row['date']} | {row['transaction_type']} | "
            f"ref={row['reference_number']} | qty={row['quantity']} | "
            f"received_from={row['supplier'] or ''} | issued_to={row['project'] or ''} | "
            f"net={row['net_movement']} | {row['origin']}:{row['status']}"
        )
    conn.close()


def add_paper(db_path: Path, material_id: int, date_value: str, tx_type: str, quantity: float, reference: str | None, notes: str | None) -> None:
    conn = connect(db_path)
    if not conn.execute("SELECT 1 FROM materials WHERE id=?", (material_id,)).fetchone():
        raise SystemExit(f"Material id not found: {material_id}")
    tx_type = tx_type.upper()
    if tx_type not in {"RECEIVING", "ISSUANCE"}:
        raise SystemExit("type must be RECEIVING or ISSUANCE")
    net = quantity if tx_type == "RECEIVING" else -quantity
    add_paper_record(db_path, material_id, date_value, tx_type, quantity, reference, notes=notes)
    print(f"Added stock-card transaction. Database only; source Excel was not changed.")
    conn.close()


def create_material(db_path: Path, name: str, location: str | None = None) -> int:
    conn = connect(db_path)
    name = clean_name(name)
    if not name:
        raise ValueError("Material name is required")
    material_id = upsert_material(conn, name, location)
    conn.commit()
    conn.close()
    return material_id


def set_material_description(db_path: Path, material_id: int, description: str | None) -> None:
    conn = connect(db_path)
    conn.execute("UPDATE materials SET description=? WHERE id=?", (clean_name(description) or None, material_id))
    conn.commit()
    conn.close()


def update_paper_record(db_path: Path, transaction_id: int, date_value: str, tx_type: str, quantity: float, reference: str | None, unit: str | None, supplier: str | None, project: str | None, notes: str | None) -> None:
    tx_type = tx_type.upper()
    if quantity <= 0: raise ValueError("Quantity must be greater than zero")
    net = quantity if tx_type in {"RECEIVING", "PULL_OUT"} else -quantity if tx_type == "ISSUANCE" else 0
    recorded_balance = quantity if tx_type in {"OPENING_BALANCE", "PHYSICAL_COUNT"} else None
    conn = connect(db_path)
    conn.execute("UPDATE transactions SET date=?, transaction_type=?, reference_number=?, quantity=?, unit=?, quantity_in=?, quantity_out=?, net_movement=?, supplier=?, project=?, recorded_balance=?, notes=? WHERE id=? AND origin='PAPER'", (date_value, tx_type, reference, quantity, unit, quantity if tx_type in {"RECEIVING", "PULL_OUT"} else None, quantity if tx_type == "ISSUANCE" else None, net, supplier, project, recorded_balance, notes, transaction_id))
    conn.commit(); conn.close()


def add_material_alias(db_path: Path, material_id: int, alias: str) -> None:
    alias = clean_name(alias)
    if not alias:
        raise ValueError("Alias is required")
    conn = connect(db_path)
    if not conn.execute("SELECT 1 FROM materials WHERE id=?", (material_id,)).fetchone():
        raise ValueError("Material id not found")
    conn.execute("INSERT OR IGNORE INTO material_aliases(material_id, alias, normalized_alias) VALUES (?, ?, ?)", (material_id, alias, norm_name(alias)))
    try:
        conn.execute("DELETE FROM materials_fts WHERE material_id=?", (str(material_id),))
        row = conn.execute("SELECT name, location FROM materials WHERE id=?", (material_id,)).fetchone()
        aliases = conn.execute("SELECT group_concat(alias, ' ') AS aliases FROM material_aliases WHERE material_id=?", (material_id,)).fetchone()["aliases"] or ""
        conn.execute("INSERT INTO materials_fts(material_id,name,location,aliases) VALUES (?,?,?,?)", (material_id, row["name"], row["location"] or "", aliases))
    except sqlite3.OperationalError:
        pass
    conn.commit()
    conn.close()


def add_paper_record(
    db_path: Path,
    material_id: int,
    date_value: str,
    tx_type: str,
    quantity: float,
    reference: str | None = None,
    notes: str | None = None,
    unit: str | None = None,
    supplier: str | None = None,
    project: str | None = None,
) -> None:
    conn = connect(db_path)
    if not conn.execute("SELECT 1 FROM materials WHERE id=?", (material_id,)).fetchone():
        raise ValueError(f"Material id not found: {material_id}")
    tx_type = tx_type.upper()
    if tx_type not in {"RECEIVING", "ISSUANCE", "PULL_OUT", "OPENING_BALANCE", "PHYSICAL_COUNT"}:
        raise ValueError("type must be RECEIVING, ISSUANCE, PULL_OUT, OPENING_BALANCE, or PHYSICAL_COUNT")
    net = quantity if tx_type in {"RECEIVING", "PULL_OUT"} else -quantity if tx_type == "ISSUANCE" else 0
    recorded_balance = quantity if tx_type in {"OPENING_BALANCE", "PHYSICAL_COUNT"} else None
    conn.execute(
        "INSERT INTO transactions(material_id, origin, status, date, transaction_type, reference_number, quantity, unit, quantity_in, quantity_out, net_movement, supplier, project, recorded_balance, notes, created_at) VALUES (?, 'PAPER', 'APPROVED', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
        (material_id, date_value, tx_type, reference, quantity, unit, quantity if tx_type in {"RECEIVING", "PULL_OUT"} else None, quantity if tx_type == "ISSUANCE" else None, net, supplier, project, recorded_balance, notes, datetime.now(timezone.utc).isoformat(timespec="seconds")),
    )
    conn.commit()
    conn.close()


def add_paper_batch(
    db_path: Path, date_value: str, tx_type: str, reference: str | None,
    site: str | None, lines: list[dict[str, Any]], supplier: str | None = None,
    project: str | None = None, issues: str | None = None, remarks: str | None = None,
) -> int:
    """Save one multi-material paper form as separate linked transactions."""
    conn = connect(db_path)
    tx_type = tx_type.upper()
    if tx_type not in {"RECEIVING", "ISSUANCE", "PULL_OUT", "OPENING_BALANCE", "PHYSICAL_COUNT"}:
        raise ValueError("Unsupported batch transaction type")
    if not lines:
        raise ValueError("Add at least one material line")
    batch_ref = clean_name(reference) or None
    saved = 0
    now = datetime.now(timezone.utc).isoformat(timespec="seconds")
    for line in lines:
        material_id = int(line["material_id"]); quantity = float(line["quantity"])
        if quantity <= 0 or not conn.execute("SELECT 1 FROM materials WHERE id=?", (material_id,)).fetchone():
            raise ValueError("Each line needs a valid material and quantity")
        net = quantity if tx_type in {"RECEIVING", "PULL_OUT"} else -quantity if tx_type == "ISSUANCE" else 0
        balance = quantity if tx_type in {"OPENING_BALANCE", "PHYSICAL_COUNT"} else None
        batch_note = " | ".join(x for x in (f"Batch site: {clean_name(site)}" if site else None, f"Issues: {clean_name(issues)}" if issues else None, f"Remarks: {clean_name(remarks)}" if remarks else None) if x)
        conn.execute("INSERT INTO transactions(material_id,origin,status,date,transaction_type,reference_number,quantity,unit,quantity_in,quantity_out,net_movement,supplier,project,recorded_balance,notes,created_at) VALUES (?, 'PAPER','APPROVED',?,?,?,?,?,?,?,?,?,?,?,?,?)", (material_id,date_value,tx_type,batch_ref,quantity,line.get("unit"),quantity if tx_type in {"RECEIVING","PULL_OUT"} else None,quantity if tx_type == "ISSUANCE" else None,net,supplier,project,balance,batch_note or None,now))
        saved += 1
    conn.commit(); conn.close(); return saved


def create_document(db_path: Path, document_type: str, date_value: str, header: dict[str, Any], lines: list[dict[str, Any]]) -> int:
    document_type = document_type.upper()
    if document_type not in {"RECEIVING", "ISSUANCE"}:
        raise ValueError("Document type must be RECEIVING or ISSUANCE")
    if not lines:
        raise ValueError("Add at least one document line")
    conn = connect(db_path); now = datetime.now(timezone.utc).isoformat(timespec="seconds")
    cur = conn.execute("INSERT INTO documents(document_type,document_number,date_value,project_site,po_number,supplier,dr_invoice_number,vehicle,driver,received_by,released_by,issues,remarks,status,created_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,'DRAFT',?)", (document_type, header.get("document_number"), date_value, header.get("project_site"), header.get("po_number"), header.get("supplier"), header.get("dr_invoice_number"), header.get("vehicle"), header.get("driver"), header.get("received_by"), header.get("released_by"), header.get("issues"), header.get("remarks"), now))
    doc_id = int(cur.lastrowid)
    for line in lines:
        qty = float(line["quantity"]); price = float(line["unit_price"]) if line.get("unit_price") not in (None, "") else None
        conn.execute("INSERT INTO document_lines(document_id,material_id,quantity,unit,unit_price,total_amount,notes) VALUES (?,?,?,?,?,?,?)", (doc_id, int(line["material_id"]), qty, line.get("unit"), price, qty * price if price is not None else None, line.get("notes")))
    conn.commit(); conn.close(); return doc_id


def post_document(db_path: Path, document_id: int) -> int:
    """Post a saved document's lines to the material stock-card transaction ledger."""
    conn = connect(db_path)
    doc = conn.execute("SELECT * FROM documents WHERE id=?", (document_id,)).fetchone()
    if not doc:
        raise ValueError("Document not found")
    if doc["status"] == "POSTED":
        return 0
    tx_type = doc["document_type"]
    inserted = 0
    for line in conn.execute("SELECT * FROM document_lines WHERE document_id=?", (document_id,)):
        qty = float(line["quantity"])
        net = qty if tx_type == "RECEIVING" else -qty
        conn.execute("INSERT INTO transactions(material_id,origin,status,date,transaction_type,reference_number,quantity,unit,quantity_in,quantity_out,net_movement,supplier,project,notes,created_at) VALUES (?, 'PAPER','APPROVED',?,?,?,?,?,?,?,?,?,?,?,?)", (line["material_id"], doc["date_value"], tx_type, doc["document_number"], qty, line["unit"], qty if tx_type == "RECEIVING" else None, qty if tx_type == "ISSUANCE" else None, net, doc["supplier"], doc["project_site"], f"Document #{document_id}", datetime.now(timezone.utc).isoformat(timespec="seconds")))
        inserted += 1
    conn.execute("UPDATE documents SET status='POSTED' WHERE id=?", (document_id,))
    conn.commit(); conn.close(); return inserted


def set_transaction_status(db_path: Path, transaction_id: int, status: str) -> None:
    if status not in {"APPROVED", "REJECTED", "PENDING_REVIEW"}:
        raise ValueError("Invalid transaction status")
    conn = connect(db_path)
    conn.execute("UPDATE transactions SET status=? WHERE id=? AND origin='PAPER'", (status, transaction_id))
    conn.commit()
    conn.close()


def main() -> int:
    parser = argparse.ArgumentParser(description="Material-centric stock-card backend MVP")
    parser.add_argument("--batch", type=Path, default=DEFAULT_BATCH)
    parser.add_argument("--db", type=Path, default=DEFAULT_DB)
    sub = parser.add_subparsers(dest="command", required=True)

    build = sub.add_parser("build", help="Build SQLite database from batch outputs")
    build.set_defaults(func=lambda args: print(build_database(args.batch, args.db)))

    find = sub.add_parser("search", help="Search materials")
    find.add_argument("term")
    find.set_defaults(func=lambda args: search(args.db, args.term))

    detail = sub.add_parser("show", help="Show one material and its transactions")
    detail.add_argument("material_id", type=int)
    detail.set_defaults(func=lambda args: show(args.db, args.material_id))

    paper = sub.add_parser("add-paper", help="Add a paper transaction pending review")
    paper.add_argument("material_id", type=int)
    paper.add_argument("date")
    paper.add_argument("type")
    paper.add_argument("quantity", type=float)
    paper.add_argument("--reference")
    paper.add_argument("--notes")
    paper.set_defaults(func=lambda args: add_paper(args.db, args.material_id, args.date, args.type, args.quantity, args.reference, args.notes))

    argv = sys.argv[1:]
    if len(argv) == 4 and argv[0] in {"-batch", "--batch"} and argv[3] == "build":
        argv = ["--batch", argv[1], "--db", argv[2], "build"]
    args = parser.parse_args(argv)
    args.func(args)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
