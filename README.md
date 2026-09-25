# Excel/CSV → MySQL importer

Reads an `.xlsx` or `.csv` file and creates `{DB_NAME}.{table}` with an
auto-increment `Row` column plus one `TEXT` column per header column.

## Setup (one time, per developer)

```
pip install -r requirements.txt
copy .env.example .env      (then put in YOUR MySQL user/password/database)
```

Your MySQL user needs CREATE, DROP, INSERT and SELECT on the target database.
`.env` holds your own credentials and is git-ignored; only `.env.example` is shared.

The created tables are normal InnoDB tables (not TEMPORARY, no views/triggers/
DEFINER), so they don't depend on who created them or on the session. Any user
with SELECT on the database can read them after the import finishes.

DB settings (`DB_HOST`, `DB_PORT`, `DB_USER`, `DB_PASSWORD`, `DB_NAME`) are read
from `.env` next to the script. A real environment variable with the same name
overrides `.env` (useful for CI/servers).
The database is created automatically if it does not exist (needs CREATE privilege).

## Usage

```
python import_data.py Sample.xlsx tmt003_store      -> billy.tmt003_store
python import_data.py users.csv   tmt_user          -> billy.tmt_user
python import_data.py Sample.xlsx tmt003_store --sheet Sheet2
python import_data.py users.csv   tmt_user --encoding cp932
python import_data.py Sample.xlsx tmt003_store --dry-run   (show SQL only)
```

## Rules

- The table name is used exactly as you type it (no suffix, no case change). Max 64 chars.
- Row 1 is the header: its values become the column names. Rows 2+ are inserted as data.
- Blank headers become `col_N`; duplicate headers (and a header named `Row`) get `_2`, `_3`, ...
- If the table already exists it is dropped and recreated.
- Blank cells are stored as NULL; fully empty rows are skipped.
- Excel: first sheet unless `--sheet` is given. Formulas import their last calculated value.
- CSV: UTF-8 by default, falls back to cp932 (Shift-JIS) automatically.
- A value larger than 65,535 bytes (TEXT limit, ~16K Japanese / 65K ASCII chars) stops the import with its row/column.
- TEXT is stored off-row, so wide files work; InnoDB's hard cap is 1,016 data columns
  (very wide files, roughly 200+ columns, may still hit InnoDB's per-page row limit).
- If inserting fails, the new table is dropped so no half-loaded table remains.
