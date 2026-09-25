# Excel/CSV → MySQL importer

Reads an `.xlsx` or `.csv` file and creates `{DB_NAME}.{parameter}_tbl` with an
auto-increment `Row` column plus one `VARCHAR(512)` column per header column.

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
python import_data.py Sample.xlsx "store master"      -> billy.store_master_tbl
python import_data.py users.csv   "user master"       -> billy.user_master_tbl
python import_data.py Sample.xlsx "store master" --sheet Sheet2
python import_data.py users.csv   "user master" --encoding cp932
python import_data.py Sample.xlsx "store master" --dry-run   (show SQL only)
```

## Rules

- Parameter → table name: lowercase, non-alphanumerics become `_`, `_tbl` appended.
- Row 1 is the header: its values become the column names. Rows 2+ are inserted as data.
- Blank headers become `col_N`; duplicate headers (and a header named `Row`) get `_2`, `_3`, ...
- If the table already exists it is dropped and recreated.
- Blank cells are stored as NULL; fully empty rows are skipped.
- Excel: first sheet unless `--sheet` is given. Formulas import their last calculated value.
- CSV: UTF-8 by default, falls back to cp932 (Shift-JIS) automatically.
- A value longer than 512 characters stops the import with its row/column.
- Max 31 columns (MySQL row-size limit for VARCHAR(512) utf8mb4).
- If inserting fails, the new table is dropped so no half-loaded table remains.
