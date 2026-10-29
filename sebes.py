# Offline cost-column exporter. No SmartPricer/Ozon API calls are made.
import argparse
import json
from decimal import Decimal, InvalidOperation
import pandas as pd
from pathlib import Path

# Legacy example filenames for help only; the CLI requires explicit paths/config.
OZON_PRICES          = Path("OZON ЦЕНЫ.xlsm")
ARTICUL_PROFIT       = Path("Спр Артикулы Доходность.xlsx")
SMARTPRICER_TEMPLATE = Path("Шаблон загрузки цен SmartPricer.xlsx")
RESULT_FILE          = Path("Шаблон загрузки цен SmartPricer_заполнено.xlsx")

OZON_REQUIRED   = ("НОМЕР СТРОКИ", "СЕБЕСТОИМОСТЬ")   # из OZON ЦЕНЫ
# из Спр Артикулы Доходность нужен только номер строки и Артикул Ozon
PROFIT_REQUIRED = ("НОМЕР СТРОКИ",)                   # вторая колонка ищем динамически

# ---------- утилиты ----------------------------------------------------------
def std(col: str) -> str:            # нормализуем имена
    return str(col).strip().replace("\xa0", " ").upper()

def load(path: Path, sheet=0, header=0) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=sheet, header=header, dtype=object)
    df.columns = [std(c) for c in df.columns]
    return df

def find_sheet_by_cols(xls: pd.ExcelFile, need: tuple[str, ...], hdr: int = 0):
    for sh in xls.sheet_names:
        cols = [std(c) for c in xls.parse(sh, nrows=2, header=hdr).columns]
        if all(n in cols for n in need):
            return sh
    return None

def find_sheet_by_keywords(xls: pd.ExcelFile,
                           kw_map: dict[str, str],
                           hdr: int = 5):
    """
    kw_map = {"art": "АРТИКУЛ", "cost": "ЗАКУП"}
    → (sheet_name, {key: real_column_name})
    """
    for sh in xls.sheet_names:
        sample = xls.parse(sh, nrows=2, header=hdr)
        norm = {std(c): c for c in sample.columns}     # {CAPS: original}
        found = {}
        for key, frag in kw_map.items():
            col = next((orig for caps, orig in norm.items() if frag in caps), None)
            if col:
                found[key] = col
        if len(found) == len(kw_map):
            return sh, found
    return None, {}

# ---------- основной скрипт --------------------------------------------------
def _row_key(value):
    if pd.isna(value) or isinstance(value, bool):
        raise ValueError("НОМЕР СТРОКИ must contain nonempty integer identifiers")
    try:
        number = Decimal(str(value).strip())
    except InvalidOperation as exc:
        raise ValueError("НОМЕР СТРОКИ must contain integer identifiers") from exc
    if not number.is_finite() or number != number.to_integral_value():
        raise ValueError("НОМЕР СТРОКИ must contain integer identifiers")
    return int(number)


def generate_cost_template(ozon_prices: Path, articles: Path, template: Path, output: Path, *, overwrite=False) -> int:
    ozon_prices, articles, template, output = map(Path, (ozon_prices, articles, template, output))
    if output.suffix.lower() != ".xlsx":
        raise ValueError("Output must be a new .xlsx workbook")
    inputs = (ozon_prices, articles, template)
    if any(output.resolve() == path.resolve() for path in inputs):
        raise ValueError("Output must not replace an input workbook")
    if output.exists() and any(path.exists() and output.samefile(path) for path in inputs):
        raise ValueError("Output must not alias an input workbook")
    if output.exists() and not overwrite:
        raise FileExistsError("Output exists; choose another path or explicitly pass --overwrite")
    # 1.  OZON ЦЕНЫ — себестоимость
    with pd.ExcelFile(ozon_prices) as oz_xls:
        oz_sheet = find_sheet_by_cols(oz_xls, OZON_REQUIRED)
        if not oz_sheet:
            raise ValueError("Нет листа с колонками 'НОМЕР СТРОКИ' и 'СЕБЕСТОИМОСТЬ'")
        oz = load(ozon_prices, sheet=oz_sheet)

    # 2.  Спр Артикулы Доходность — ищем «Артикул Ozon»
    pr = load(articles)
    art_col = next((c for c in pr.columns if "АРТИКУЛ" in c and "OZON" in c), None)
    if art_col is None:
        raise ValueError("В «Спр Артикулы Доходность» не найден столбец 'Артикул Ozon'")
    if "НОМЕР СТРОКИ" not in pr.columns:
        raise ValueError("Нет колонки 'НОМЕР СТРОКИ' в «Спр Артикулы Доходность»")

    # 3.  merge по номеру строки
    oz = oz[list(OZON_REQUIRED)].dropna(how="all").copy()
    pr = pr[["НОМЕР СТРОКИ", art_col]].dropna(how="all").copy()
    oz["НОМЕР СТРОКИ"] = oz["НОМЕР СТРОКИ"].map(_row_key)
    pr["НОМЕР СТРОКИ"] = pr["НОМЕР СТРОКИ"].map(_row_key)
    merged = oz.merge(pr, on="НОМЕР СТРОКИ", how="inner", validate="one_to_one")
    print(f"✅ совпадений по номеру строки: {len(merged)}")

    # 4.  Шаблон SmartPricer — ищем колонки «артикул» и «закуп»
    kw = {"art": "АРТИКУЛ", "cost": "ЗАКУП"}
    with pd.ExcelFile(template) as sp_xls:
        sp_sheet, cols = find_sheet_by_keywords(sp_xls, kw, hdr=5)
        if not sp_sheet:
            raise ValueError("Не найден лист с колонками *артикул* и *закуп* в шаблоне")
    col_art, col_cost = cols["art"], cols["cost"]

    # 5.  формируем блок для вставки
    block = pd.DataFrame({
        col_art:  merged[art_col],
        col_cost: merged["СЕБЕСТОИМОСТЬ"],
    }).dropna(subset=[col_art]).reset_index(drop=True)

    # 6.  сохраняем
    block.to_excel(output, index=False)
    print(f"Saved {output.name} | rows: {len(block)}")
    return len(block)


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Create an offline article/purchase-cost XLSX block; no prices are uploaded.")
    parser.add_argument("--config", type=Path, help="Explicit private JSON path; no automatic config search")
    parser.add_argument("--ozon-prices", type=Path, help=f"Cost workbook (legacy name: {OZON_PRICES.name})")
    parser.add_argument("--articles", type=Path, help=f"Row/article mapping (legacy name: {ARTICUL_PROFIT.name})")
    parser.add_argument("--template", type=Path, help=f"Header sample (legacy name: {SMARTPRICER_TEMPLATE.name})")
    parser.add_argument("--out", type=Path, help=f"New output XLSX (legacy name: {RESULT_FILE.name})")
    parser.add_argument("--overwrite", action="store_true", help="Allow replacing an existing output, never an input")
    args = parser.parse_args(argv)
    config = {}
    if args.config:
        config = json.loads(args.config.read_text(encoding="utf-8"))
        if not isinstance(config, dict) or set(config) - {"ozon_prices", "articles", "template", "out"}:
            parser.error("Config must contain only ozon_prices/articles/template/out path fields")
    paths = {}
    for key in ("ozon_prices", "articles", "template", "out"):
        value = getattr(args, key)
        if value is None and config.get(key):
            if not isinstance(config[key], str):
                parser.error("Config path values must be strings")
            value = Path(config[key])
            if not value.is_absolute():
                value = args.config.parent / value
        if value is None:
            parser.error(f"Provide --{key.replace('_', '-')} or a nonempty explicit config value")
        paths[key] = value
    generate_cost_template(paths["ozon_prices"], paths["articles"], paths["template"], paths["out"], overwrite=args.overwrite)
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
