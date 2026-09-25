# Excel/CSV → MySQL importer

Reads an `.xlsx` or `.csv` file and creates `billy.{parameter}_tbl` with an
auto-increment `Row` column plus `col_1`, `col_2`, ... (`VARCHAR(512)`) — one per file column.

## Setup (one time, per developer)

```
pip install -r requirements.txt
copy .env.example .env      (then put in YOUR MySQL user/password/database)
```

Your MySQL user needs CREATE, DROP, INSERT and SELECT on the target database.
`.env` and `config.ini` hold your own credentials and are git-ignored; only the
`.example` files are shared.

The created tables are normal InnoDB tables (not TEMPORARY, no views/triggers/
DEFINER), so they don't depend on who created them or on the session. Any user
with SELECT on the database can read them after the import finishes.

DB settings are read in this order (first one found wins):

1. Environment variables `DB_HOST`, `DB_PORT`, `DB_USER`, `DB_PASSWORD`, `DB_NAME`
2. `.env` file next to the script (see `.env.example`)
3. `config.ini` `[mysql]` section

Example: `set DB_NAME=billy_test` then run the import to target another database.
The database must already exist.

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
- There is no header row: every non-empty row in the file is inserted as data.
- If the table already exists it is dropped and recreated.
- Blank cells are stored as NULL; fully empty rows are skipped.
- Excel: first sheet unless `--sheet` is given. Formulas import their last calculated value.
- CSV: UTF-8 by default, falls back to cp932 (Shift-JIS) automatically.
- A value longer than 512 characters stops the import with its row/column.
- Max 31 columns (MySQL row-size limit for VARCHAR(512) utf8mb4).
- If inserting fails, the new table is dropped so no half-loaded table remains.
