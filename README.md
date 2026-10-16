# AIYU Warehouse

AIYU Warehouse is a Windows desktop tool for encoding stock-card backlogs and managing warehouse receiving and issuance records.

## Current purpose

The system helps one warehouse staff member process many stock-card transactions quickly while keeping the original Excel source files unchanged.

It supports:

- Searching and viewing stock cards.
- Importing stock-card Excel files from a selected source folder.
- Recording receiving, issuance, pull-out, opening-balance, and physical-count entries.
- Creating brand-specific materials when a material does not already exist.
- Batch entry for multiple material lines under one project/site.
- Project, supplier, PO, RR/IS, driver, vehicle, and staff suggestions.
- Editing manually entered stock-card records.
- Exporting one stock card or all stock cards.
- Generating receiving and issuance Excel forms.
- Database backup and Light/Dark theme settings.

## Main workflow

```text
Select source folder
        ↓
Refresh inventory
        ↓
Search material or add backlog entry
        ↓
Save transaction to the local database
        ↓
Update stock card and optionally generate a printable Excel form
```

## Installation

Install `AIYU Warehouse Setup.exe` on the staff computer. The installer includes:

- AIYU Warehouse application.
- Stock-card, receiving, and issuance templates.
- Warehouse logo.
- Inventory scanner.
- Database builder.

The installer does not include the current development database or staff source files.

## First-time setup

1. Open AIYU Warehouse.
2. Open **Settings**.
3. Select the folder containing the stock-card Excel files as the **Source folder**.
4. Select a separate **Export folder**.
5. Click **Refresh inventory**.

The application creates its SQLite database automatically and builds the inventory from the selected source folder.

Keep the export folder outside the source folder so generated files are not scanned as source stock cards.

## Refresh behavior and cache

Refresh Inventory reads supported Excel files recursively and builds the imported inventory layer. A source signature is stored in the database. If no supported source file has changed, the application skips the scan and reports that the inventory is current.

Manual entries are preserved when the source inventory is refreshed. Source Excel files are read-only from the application's perspective.

## Stock-card entry

Single entry fields include:

- Date.
- Transaction type.
- PO / RR / IS number or other warehouse reference.
- Quantity and unit.
- Supplier or received-from value.
- Project/site.
- Issues or remarks.

Batch entry supports multiple material lines and a scrollable line-entry area. Unknown materials can be created as new stock cards, which allows separate brand-specific records.

Manual entries are saved directly and remain editable. Imported Excel records remain read-only.

## Receiving and issuance forms

Receiving and issuance actions use the updated Excel templates:

- `TemplateForms/RCVNG EXCEL.xlsx`
- `TemplateForms/ISSUANCE EXCEL.xlsx`

The forms share the database with the stock-card workflow. Their document references, PO numbers, projects, suppliers, material lines, and staff details can be searched and reused through suggestions.

The workflow is intentionally flexible. A PO, receiving record, and issuance record may be connected, but the program does not require every document to exist or force a fixed date sequence.

## Stock-card exports

**Export selected stock card** creates a completed copy of `TemplateForms/STOCK CARD.xlsx` for the selected material.

**Export all stock cards** creates one stock-card workbook per material in a selected destination folder. That folder can be transferred to a fresh installation and used as a new source folder.

## Database and backups

The local SQLite database stores materials, imported transactions, manual transactions, projects, documents, metadata, and source information.

Use **Settings → Backup database** regularly. Before updating or uninstalling the program:

1. Close AIYU Warehouse.
2. Create a database backup.
3. Keep the backup in a safe location.

The exported stock-card files are useful as a new source, but a database backup preserves more information, including projects, document links, and manual-entry details.

## Operational notes

- The first source-folder scan may take time because Excel files must be read and validated.
- Refresh displays reading and database-building status messages.
- Do not place exports inside the source folder.
- Use a clean folder when exporting all stock cards for transfer.
- Keep regular database backups.
- The current system prioritizes fast backlog entry; more advanced PO-to-receiving-to-issuance workflow features can be added later without replacing the current database.

## Development files

- `stock_card_gui_v1.py` — main application interface.
- `stock_card_backend_v1.py` — SQLite schema and transaction functions.
- `stock_card_batch_v1.py` — source-folder scanner and normalizer.
- `stock_card_document_print.py` — Excel form generation.
- `TemplateForms/` — stock-card, receiving, and issuance templates.
- `asset/warehouse_logo.jpg` — application branding.
- `AIYU Warehouse Installer.nsi` — NSIS installer configuration.
