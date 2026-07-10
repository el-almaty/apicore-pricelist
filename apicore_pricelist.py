"""
Скрипт собирает прайс-листы от нескольких дистрибьюторов apicore.kz:
товары + закупочные цены + остатки -> один Excel-файл, каждый
дистрибьютор на отдельном листе, и заливает результат на Диск Bitrix24.

Все секреты (ключи, ID) берутся из переменных окружения:
- APICORE_API_KEY
- APICORE_CATALOG_CODE   (необязательно, по умолчанию "main")
- APICORE_DISTRIBUTORS   -- JSON-список дистрибьюторов, например:
    [{"id": "3abb4152", "name": "Дистрибьютор А"},
     {"id": "b4d0d204", "name": "Дистрибьютор Б"}]
- BITRIX_WEBHOOK
- BITRIX_FOLDER_ID

Для GitHub Actions эти переменные передаются через Secrets репозитория
(см. .github/workflows/update-pricelist.yml).
"""

import os
import re
import json
import base64
import requests
import pandas as pd
import time
from datetime import datetime
from zoneinfo import ZoneInfo
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

API_KEY = os.environ["APICORE_API_KEY"]
CATALOG_CODE = os.environ.get("APICORE_CATALOG_CODE", "main")
DISTRIBUTORS = json.loads(os.environ["APICORE_DISTRIBUTORS"])

BASE_URL = "https://api.apicore.one/dealer"
HEADERS = {
    "Api-Key": API_KEY,
    "Content-Type": "application/json",
    "Accept": "application/json",
}

# --- Bitrix24 ---
BITRIX_WEBHOOK = os.environ["BITRIX_WEBHOOK"]  # напр. https://xxx.bitrix24.kz/rest/xx/xxxxxxxx/
BITRIX_FOLDER_ID = os.environ["BITRIX_FOLDER_ID"]


def call(version, method, body):
    """Делает POST-запрос к apicore и возвращает JSON-ответ."""
    url = f"{BASE_URL}/{version}/{method}"
    resp = requests.post(url, headers=HEADERS, json=body, timeout=30)
    resp.raise_for_status()
    return resp.json()


def category_path(cat_id, categories_by_id):
    """Возвращает (категория_1_уровня, категория_2_уровня_и_глубже) для товара.
    Идёт вверх по parent_id, пока не дойдёт до корня (parent_id == 0)."""
    if cat_id not in categories_by_id or cat_id == 0:
        return "Без категории", ""

    chain = []
    node = categories_by_id.get(cat_id)
    seen = set()
    while node is not None and node["id"] not in seen:
        seen.add(node["id"])
        chain.append(node["name"])
        parent_id = node.get("parent_id", 0)
        if not parent_id:
            break
        node = categories_by_id.get(parent_id)

    chain.reverse()  # теперь от корня к текущей категории
    top = chain[0]
    sub = " > ".join(chain[1:])  # всё, что глубже 1-го уровня
    return top, sub


def fetch_distributor_df(distributor_id, run_started_at):
    """Собирает полный прайс-лист (категории + товары + цены + остатки)
    для одного дистрибьютора и возвращает готовый DataFrame."""
    body = {"distributor_id": distributor_id, "catalog_code": CATALOG_CODE}

    print("  Получаю список категорий...")
    categories_resp = call("v1", "distrib.category.list", body)
    categories_by_id = {c["id"]: c for c in categories_resp["categories"]}
    print(f"    категорий получено: {len(categories_by_id)}")
    time.sleep(0.5)

    print("  Получаю список товаров...")
    products_resp = call("v1", "distrib.product.list", body)
    products = products_resp["products"]
    print(f"    товаров получено: {len(products)}")
    time.sleep(0.5)

    print("  Получаю цены...")
    prices_resp = call("v2", "distrib.product.prices", body)
    prices_by_id = {}
    price_updated_by_id = {}
    for p in prices_resp["products"]:
        prices_by_id[p["product_id"]] = p["purchase"]["price"]
        if "date_update" in p:
            price_updated_by_id[p["product_id"]] = p["date_update"]
    time.sleep(0.5)

    print("  Получаю остатки...")
    qty_resp = call("v1", "distrib.product.quantities", body)
    qty_by_id = {}
    qty_updated_by_id = {}
    for p in qty_resp["products"]:
        qty_by_id[p["product_id"]] = p["quantity"]
        if "date_update" in p:
            qty_updated_by_id[p["product_id"]] = p["date_update"]
    time.sleep(0.5)

    rows = []
    for p in products:
        pid = p["id"]
        cat_id = p.get("category_id", 0)
        top_cat, sub_cat = category_path(cat_id, categories_by_id)
        rows.append({
            "Категория 1 ур.": top_cat,
            "Категория 2 ур.": sub_cat,
            "category_id (для проверки)": cat_id,
            "ID": pid,
            "Наименование": p["name"],
            "Производитель": p.get("vendor", ""),
            "Артикул": p.get("vendor_code", ""),
            "Штрихкод": p.get("barcode", ""),
            "Цена, KZT": prices_by_id.get(pid),
            "Наличие, шт": qty_by_id.get(pid, 0),
            "Обновление цены (apicore)": price_updated_by_id.get(pid, ""),
            "Обновление остатка (apicore)": qty_updated_by_id.get(pid, ""),
            "Данные получены": run_started_at,
        })

    df = pd.DataFrame(rows)
    sort_cols = ["Категория 1 ур.", "Категория 2 ур.", "Наименование"]
    df = df.sort_values(by=sort_cols, key=lambda col: col.str.lower())
    return df


def sanitize_sheet_name(name):
    """Названия листов Excel не могут содержать [ ] : * ? / \\ и длиннее 31 символа."""
    clean = re.sub(r"[\[\]:*?/\\]", "", name).strip()
    return clean[:31] if clean else "Лист"


def format_worksheet(ws, df):
    """Оформляет один лист как Excel-таблицу (Table): фильтры, жирная шапка,
    полосатая заливка, подходит как источник для сводных таблиц.
    Плюс подбирает ширину столбцов по содержимому и формат чисел."""
    n_rows = ws.max_row
    n_cols = ws.max_column
    last_col_letter = get_column_letter(n_cols)

    table_name = re.sub(r"[^A-Za-z0-9_]", "_", ws.title) or "Table"
    table = Table(displayName=f"T_{table_name}"[:255], ref=f"A1:{last_col_letter}{n_rows}")
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium9", showRowStripes=True, showColumnStripes=False
    )
    ws.add_table(table)

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"

    for col_idx in range(1, n_cols + 1):
        letter = get_column_letter(col_idx)
        max_len = max(
            len(str(cell.value)) if cell.value is not None else 0
            for cell in ws[letter]
        )
        ws.column_dimensions[letter].width = min(max_len + 2, 60)

    price_col = df.columns.get_loc("Цена, KZT") + 1
    qty_col = df.columns.get_loc("Наличие, шт") + 1
    for row in range(2, n_rows + 1):
        ws.cell(row=row, column=price_col).number_format = "#,##0"
        ws.cell(row=row, column=qty_col).number_format = "#,##0"


def find_existing_file_id(folder_id, filename):
    """Ищет в папке файл с указанным именем. Возвращает его ID или None, если не найден."""
    url = f"{BITRIX_WEBHOOK}disk.folder.getchildren"
    resp = requests.post(url, data={"id": folder_id}, timeout=60)
    resp.raise_for_status()
    result = resp.json()
    for item in result.get("result", []):
        if item.get("TYPE") == "file" and item.get("NAME") == filename:
            return item["ID"]
    return None


def upload_to_bitrix_disk(local_file_path, folder_id, filename):
    """Загружает файл в указанную папку на Диске Bitrix24.
    Если файл с таким именем уже существует -- обновляет его СОДЕРЖИМОЕ
    через disk.file.uploadversion (передаётся сразу base64, одним запросом),
    благодаря чему ID файла и ссылка на него не меняются от запуска к запуску.
    Если файла ещё нет -- создаёт его через disk.folder.uploadfile
    (двухшаговая схема с uploadUrl, только в самый первый раз)."""
    existing_file_id = find_existing_file_id(folder_id, filename)

    if existing_file_id:
        print(f"Файл '{filename}' уже существует (ID {existing_file_id}) -- обновляю версию.")
        with open(local_file_path, "rb") as f:
            content_b64 = base64.b64encode(f.read()).decode("ascii")

        data = {
            "id": existing_file_id,
            "fileContent[0]": filename,
            "fileContent[1]": content_b64,
        }
        resp = requests.post(f"{BITRIX_WEBHOOK}disk.file.uploadversion", data=data, timeout=120)
        if not resp.ok:
            print("Bitrix24 вернул ошибку, тело ответа:", resp.text)
        resp.raise_for_status()
        result = resp.json()
        print("Ответ Bitrix24:", result)
        if "result" not in result:
            raise RuntimeError(f"Bitrix24 вернул ошибку при обновлении версии: {result}")
        print("Файл на Диске Bitrix24 обновлён.")
        return result["result"]

    print(f"Файла '{filename}' ещё нет -- создаю новый.")
    step1_url = f"{BITRIX_WEBHOOK}disk.folder.uploadfile"
    data = {
        "id": folder_id,
        "data[NAME]": filename,
    }
    resp1 = requests.post(step1_url, data=data, timeout=60)
    if not resp1.ok:
        print("Bitrix24 (шаг 1) вернул ошибку, тело ответа:", resp1.text)
    resp1.raise_for_status()
    step1_result = resp1.json()
    print("Шаг 1 (получение uploadUrl):", step1_result)

    if "result" not in step1_result or "uploadUrl" not in step1_result["result"]:
        raise RuntimeError(f"Не удалось получить uploadUrl: {step1_result}")

    upload_url = step1_result["result"]["uploadUrl"]

    with open(local_file_path, "rb") as f:
        files = {"file": (filename, f)}
        resp2 = requests.post(upload_url, files=files, timeout=120)
    if not resp2.ok:
        print("Bitrix24 (шаг 2) вернул ошибку, тело ответа:", resp2.text)
    resp2.raise_for_status()
    step2_result = resp2.json()
    print("Шаг 2 (загрузка содержимого):", step2_result)

    if "result" not in step2_result:
        raise RuntimeError(f"Bitrix24 вернул ошибку при загрузке файла: {step2_result}")

    print("Файл загружен на Диск Bitrix24.")
    return step2_result["result"]


def main():
    run_started_at = datetime.now(ZoneInfo("Asia/Almaty")).strftime("%d.%m.%Y %H:%M:%S")
    out_file = "pricelist.xlsx"

    dfs_by_sheet = {}
    for d in DISTRIBUTORS:
        print(f"=== Дистрибьютор: {d['name']} ({d['id']}) ===")
        df = fetch_distributor_df(d["id"], run_started_at)
        sheet_name = sanitize_sheet_name(d["name"])
        # на случай совпадения названий после обрезки/очистки -- не даём листам дублироваться
        base_name, i = sheet_name, 2
        while sheet_name in dfs_by_sheet:
            suffix = f" ({i})"
            sheet_name = base_name[: 31 - len(suffix)] + suffix
            i += 1
        dfs_by_sheet[sheet_name] = df
        print(f"  готово: {len(df)} строк")

    with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
        for sheet_name, df in dfs_by_sheet.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name)

    wb = load_workbook(out_file)
    for sheet_name, df in dfs_by_sheet.items():
        format_worksheet(wb[sheet_name], df)
    wb.save(out_file)

    total_rows = sum(len(df) for df in dfs_by_sheet.values())
    print(f"Готово. Файл сохранён: {out_file} ({len(dfs_by_sheet)} листов, {total_rows} строк всего)")

    if BITRIX_FOLDER_ID:
        upload_to_bitrix_disk(out_file, BITRIX_FOLDER_ID, "Прайс-лист.xlsx")
    else:
        print("BITRIX_FOLDER_ID не задан -- пропускаю загрузку на Bitrix24 Диск.")


if __name__ == "__main__":
    main()
