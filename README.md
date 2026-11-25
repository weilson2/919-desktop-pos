# 919 Desktop POS

Open `START_DESKTOP_POS.cmd` to run the installed Windows POS.

- `app\919pos.exe`: installed desktop application and bundled runtime.
- `app\pos.db`: the existing live database shared with Website POS.
- `source\test.py`: desktop application's Python source.
- `source\category_list_prices.py`: pricing helper, also present in Website POS.
- `source\919pos.spec`: Windows application build definition.
- `source\tests`: desktop regression tests, including two cross-interface checks.
- `Backups`: preserved database backups.
- `Archive`: previous builds, reports, original transfer notes, and POS ZIP.

Keep `app\919pos.exe`, its `_internal` folder, and `app\pos.db` together.
The Website POS launcher in the sibling project continues to use this database.

## Development

From the `source` folder, create a Python 3.12 environment and install dependencies:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
.\.venv\Scripts\python.exe -m PyInstaller --clean --noconfirm 919pos.spec
```

Cross-interface tests need the sibling `Website POS\source` folder. A source
launch of `test.py` uses a separate `source\pos.db` unless explicitly configured
in code; use the installed launcher for normal work with the existing live data.
Do not copy a development database over the live database.

The original source environment remains in `Website POS\source\.venv`. It can
also run these desktop tests from the desktop `source` working directory.

Before GitHub publication, review source and history for credentials or business
data. The project `.gitignore` excludes the runtime, database, backups, archive,
environments, and build output.
