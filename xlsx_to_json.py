import json
import sys
from pathlib import Path

import pandas as pd

raw = input("Введите название файла: ").strip().strip('"').strip("'")
path = Path("table")
if raw:
    path = Path(raw)
if not path.suffix:
    path = path.with_suffix(".xlsx")
if not path.exists():
    sys.exit(f"Файл не найден: {path}")

content_path = r"content.json"
with open(content_path, encoding="utf-8") as f:
    content = json.load(f)


def classify_sheet(sheet_name, by_field_keys):
    name = sheet_name.lower()
    if "гайд" in name or "guide" in name:
        return "guides"
    if any(w in name for w in ("еда", "food", "столов", "буфет", "питан")):
        return "food"
    keys = set(by_field_keys)
    if "text" in keys:
        return "guides"
    if keys & {"price", "pro", "con"}:
        return "food"
    return None


# ---------- 1. Читаем все листы, раскладываем по секциям ----------
xl = pd.ExcelFile(path)

sheets_data = {}
for sheet_name in xl.sheet_names:
    df = xl.parse(sheet_name, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    key_col = df.columns[0]
    entities = list(df.columns[1:])
    by_field = {}
    for _, row in df.iterrows():
        if pd.isna(row[key_col]):
            continue
        field = str(row[key_col]).strip()
        if not field or field == "lunchbox_tip":
            continue
        by_field[field] = [
            None if pd.isna(row[e]) else str(row[e]).strip()
            for e in entities
        ]
    section = classify_sheet(sheet_name, by_field.keys())
    if section is None:
        continue
    items = []
    for i in range(len(entities)):
        items.append({field: by_field[field][i] for field in by_field})
    sheets_data.setdefault(section, []).extend(items)

for key, value in sheets_data.items():
    if key == "food":
        if "food" in key:
            content["food"]["places"] = value
    else:
        content[key] = value
print(content)

# # ---------- 2. Перезапись ----------
with open(content_path, "w", encoding="utf-8") as f:
    json.dump(content, f, ensure_ascii=False, indent=2)