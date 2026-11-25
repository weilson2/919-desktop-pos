import ast
import os
import sqlite3
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "test.py"
tree = ast.parse(SOURCE.read_text(encoding="utf-8-sig"))


def signature(fn):
    args = fn.args
    names = [a.arg for a in args.posonlyargs + args.args]
    if args.vararg:
        names.append("*" + args.vararg.arg)
    names.extend(a.arg for a in args.kwonlyargs)
    if args.kwarg:
        names.append("**" + args.kwarg.arg)
    return "(" + ", ".join(names) + ")"


def description(fn):
    doc = ast.get_docstring(fn)
    return doc.strip().splitlines()[0] if doc else ""


GROUPS = [
    (1218, 1721, "Application setup, localization, and shared UI"),
    (1722, 1925, "Authentication, authorization, and navigation"),
    (1926, 2754, "Action hub and administration"),
    (2755, 3430, "Main POS interface, categories, and product browsing"),
    (3431, 3794, "Ordering, variants, quantities, and totals"),
    (3795, 4621, "Tables, held orders, and recent orders"),
    (4622, 5273, "Checkout, payments, stock, and consistency"),
    (5274, 5597, "Shared utilities, audit, repair, and dialogs"),
    (5598, 6239, "Inventory and item management"),
    (6240, 7432, "Sales history, editing, and reconciliation"),
    (7433, 8683, "Reports and analytics"),
    (8684, 9408, "Cashier cash and rent workflows"),
    (9409, 9892, "Tenant and stall administration"),
    (9893, 10708, "Cash closing and daily export"),
    (10709, 11112, "Profit center"),
    (11113, 11648, "Monthly and accountant exports"),
    (11649, 11971, "Refunds"),
    (11972, 99999, "Text and combobox helpers"),
]


def group_for(line):
    return next((name for lo, hi, name in GROUPS if lo <= line <= hi), "Other")


def add_function(lines, fn):
    desc = description(fn)
    lines.append(f"- `{fn.name}{signature(fn)}` — lines {fn.lineno}–{fn.end_lineno}" + (f": {desc}" if desc else ""))
    nested = sorted(
        (n for n in ast.walk(fn) if isinstance(n, ast.FunctionDef) and n is not fn),
        key=lambda n: n.lineno,
    )
    if nested:
        lines.append("  - Nested callbacks/helpers: " + ", ".join(
            f"`{n.name}{signature(n)}` (line {n.lineno})" for n in nested
        ))


lines = [
    "# 919 Restaurant POS — Complete Function Structure", "",
    "> Generated from `test.py` for restructuring. Line numbers refer to the current source.", "",
    "## Current architecture", "",
    "- `test.py`: application, database, business logic, Tkinter UI, reports, and exports.",
    "- `pos.db`: live SQLite business database.",
    "- `919pos.spec`: PyInstaller build definition.",
    "- `tests/`: backend regression tests.", "",
    "## Function counts", "",
    f"- Top-level classes: {sum(isinstance(n, ast.ClassDef) for n in tree.body)}",
    f"- Top-level utility functions: {sum(isinstance(n, ast.FunctionDef) for n in tree.body)}",
    f"- Direct class methods: {sum(isinstance(x, ast.FunctionDef) for n in tree.body if isinstance(n, ast.ClassDef) for x in n.body)}",
    f"- All functions including nested callbacks: {sum(isinstance(n, ast.FunctionDef) for n in ast.walk(tree))}", "",
    "## Structure tree", "", "### Top-level utilities", "",
]

for node in tree.body:
    if isinstance(node, ast.FunctionDef):
        add_function(lines, node)

for cls in (n for n in tree.body if isinstance(n, ast.ClassDef) and n.name != "POS"):
    lines.extend(["", f"### Class `{cls.name}` — lines {cls.lineno}–{cls.end_lineno}", ""])
    for fn in (n for n in cls.body if isinstance(n, ast.FunctionDef)):
        add_function(lines, fn)

pos = next(n for n in tree.body if isinstance(n, ast.ClassDef) and n.name == "POS")
lines.extend(["", f"### Class `POS` — lines {pos.lineno}–{pos.end_lineno}", "",
              "> Main restructuring target. All direct methods and nested callbacks follow.", ""])
last_group = None
for fn in (n for n in pos.body if isinstance(n, ast.FunctionDef)):
    group = group_for(fn.lineno)
    if group != last_group:
        lines.extend([f"#### {group}", ""])
        last_group = group
    add_function(lines, fn)

lines.extend(["", "## Database tables", ""])
database = Path(os.environ.get("POS_DB_PATH", ROOT.parent / "app" / "pos.db"))
with sqlite3.connect(f"{database.resolve().as_uri()}?mode=ro", uri=True) as conn:
    for (name,) in conn.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name"):
        columns = ", ".join(row[1] for row in conn.execute(f'PRAGMA table_info("{name}")'))
        lines.append(f"- `{name}`: {columns}")
lines.append("- `stock_movements` (new migration): id, datetime, item_id, item_name, quantity_delta, stock_before, stock_after, source_type, reference_id, recorded_by, notes")

lines.extend(["", "## Recommended target module tree", "", "```text",
"kopitiam_pos/", "├── app.py", "├── config.py", "├── database/", "│   ├── connection.py", "│   ├── schema.py", "│   ├── migrations.py", "│   └── repositories/", "├── domain/", "│   ├── money.py", "│   ├── business_date.py", "│   ├── orders.py", "│   ├── inventory.py", "│   ├── reporting.py", "│   └── permissions.py", "├── services/", "│   ├── checkout_service.py", "│   ├── refund_service.py", "│   ├── inventory_service.py", "│   ├── cashflow_service.py", "│   ├── report_service.py", "│   ├── backup_service.py", "│   └── printing_service.py", "├── ui/", "│   ├── theme.py", "│   ├── widgets.py", "│   ├── main_window.py", "│   ├── screens/", "│   └── dialogs/", "├── exports/", "│   ├── daily_excel.py", "│   ├── monthly_excel.py", "│   └── accountant_excel.py", "└── tests/", "    ├── unit/", "    └── integration/", "```", "",
"## Suggested extraction order", "",
"1. Extract and test business-date, money, and reporting functions.",
"2. Extract database connection, schema, and migrations.",
"3. Create repositories for sales, inventory, cashflow, staff, tables, and rent.",
"4. Move checkout/refund/inventory operations into transactional services.",
"5. Split each Tkinter screen into its own class/module.",
"6. Move Excel generation behind the report service.",
"7. Reduce `POS` to navigation/shared state and rename `test.py` to `app.py`.", ""])

(ROOT / "SYSTEM_STRUCTURE.md").write_text("\n".join(lines), encoding="utf-8")
