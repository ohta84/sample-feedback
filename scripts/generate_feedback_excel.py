#!/usr/bin/env python3
"""
Bカート受注CSV -> メーカー別サンプルフィードバックExcel 生成スクリプト

使い方:
  python generate_feedback_excel.py --csv bcart_order.csv --maker-name "○○" \
      --maker-code "121" --output "○○様サンプルフィードバック.xlsx"

--maker-code は省略可（CSVが既に1メーカー分に絞られている場合）。
複数コードはカンマ区切り: --maker-code "121,122"
"""
import argparse
import re
import sys
from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SET_NAME_COL_INDEX = 68  # BR列 = 69列目 (0-indexed 68)
EXCLUDE_COMPANY_KEYWORDS = ["ナゴミヤ", "なごみや"]

HEADER_KEYWORDS = {
    "納品日": ["納品日"],
    "業態": ["業態"],
    "会社名": ["会社名", "お届け先名", "届け先名", "店舗名"],
    "都道府県": ["都道府県"],
    "商品名": ["商品名"],
}

FONT_NAME = "Arial"
GREY_FILL = PatternFill("solid", start_color="D9D9D9", end_color="D9D9D9")
GREEN_FILL = PatternFill("solid", start_color="E2EFDA", end_color="E2EFDA")
THIN_GREY = Side(style="thin", color="BFBFBF")
BORDER_ALL = Border(left=THIN_GREY, right=THIN_GREY, top=THIN_GREY, bottom=THIN_GREY)

COLUMN_WIDTHS = {
    "A": 14.6, "B": 15.1, "C": 31.6, "D": 11.9,
    "E": 37.0, "F": 35.2, "G": 71.4, "H": 14.0,
}


def read_csv_flexible(path):
    for enc in ("cp932", "utf-8-sig", "utf-8", "shift_jis"):
        try:
            df = pd.read_csv(path, encoding=enc, dtype=str, keep_default_na=False)
            return df
        except (UnicodeDecodeError, UnicodeError):
            continue
    raise RuntimeError("CSVの文字コードを判定できませんでした（cp932 / utf-8 を試しましたが失敗）")


def find_column(df, keywords):
    for col in df.columns:
        for kw in keywords:
            if kw in str(col):
                return col
    return None


def get_set_name_series(df):
    if df.shape[1] > SET_NAME_COL_INDEX:
        return df.iloc[:, SET_NAME_COL_INDEX].astype(str)
    # フォールバック: 列名に「セット名」を含む列を探す
    col = find_column(df, ["セット名"])
    if col is not None:
        return df[col].astype(str)
    raise RuntimeError(
        f"BR列（69列目）が見つかりません。CSVの列数は{df.shape[1]}列でした。"
        "セット名の列を手動で指定してください。"
    )


def extract_maker_code(product_name):
    m = re.match(r"^\s*\[(\d+)\]", str(product_name))
    return m.group(1) if m else None


def parse_date(value):
    value = str(value).strip()
    for fmt in ("%Y/%m/%d", "%Y-%m-%d", "%Y/%m/%d %H:%M:%S", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--csv", required=True)
    ap.add_argument("--maker-name", required=True)
    ap.add_argument("--maker-code", default=None, help="カンマ区切りで複数指定可")
    ap.add_argument("--output", required=True)
    args = ap.parse_args()

    df = read_csv_flexible(args.csv)

    col_date = find_column(df, HEADER_KEYWORDS["納品日"])
    col_type = find_column(df, HEADER_KEYWORDS["業態"])
    col_company = find_column(df, HEADER_KEYWORDS["会社名"])
    col_pref = find_column(df, HEADER_KEYWORDS["都道府県"])
    col_product = find_column(df, HEADER_KEYWORDS["商品名"])

    missing = [name for name, col in [
        ("納品日", col_date), ("会社名", col_company),
        ("都道府県", col_pref), ("商品名", col_product),
    ] if col is None]
    if missing:
        sys.stderr.write(
            "以下の列が見つかりませんでした: " + ", ".join(missing) + "\n"
            "CSVの実際のヘッダー一覧:\n" + "\n".join(f"  {i}: {c}" for i, c in enumerate(df.columns)) + "\n"
        )
        sys.exit(1)

    df["_set_name"] = get_set_name_series(df)
    df["_is_sample"] = df["_set_name"].str.contains("サンプル", na=False)

    # 自社テスト除外
    df = df[~df[col_company].astype(str).apply(
        lambda v: any(kw in v for kw in EXCLUDE_COMPANY_KEYWORDS)
    )]

    # メーカーコードで絞り込み
    if args.maker_code:
        codes = [c.strip() for c in args.maker_code.split(",") if c.strip()]
        df = df[df[col_product].apply(lambda p: extract_maker_code(p) in codes)]

    if df.empty:
        sys.stderr.write("絞り込み後、対象データが0件でした。メーカーコードや会社名の条件を確認してください。\n")
        sys.exit(1)

    df_sample = df[df["_is_sample"]].copy()
    df_normal = df[~df["_is_sample"]].copy()

    if df_sample.empty:
        sys.stderr.write("サンプル注文（セット名に「サンプル」を含む行）が見つかりませんでした。\n")
        sys.exit(1)

    companies_with_order = set(df_normal[col_company].astype(str))

    df_sample["_date_parsed"] = df_sample[col_date].apply(parse_date)
    df_sample_sorted = df_sample.sort_values(
        by=[col_company, col_date], key=lambda s: s if s.name != col_date else s
    )
    df_sample_sorted = df_sample.assign(_company=df_sample[col_company].astype(str)).sort_values(
        ["_company", "_date_parsed"]
    )

    provided_companies = sorted(set(df_sample_sorted["_company"]))
    n_provided = len(provided_companies)
    n_converted = sum(1 for c in provided_companies if c in companies_with_order)

    all_dates = [d for d in df_sample_sorted["_date_parsed"] if d is not None]
    period_str = ""
    if all_dates:
        period_str = f"{min(all_dates):%Y/%m/%d}〜{max(all_dates):%Y/%m/%d}"

    # 会社名+納品日でグループ化して行ブロックを作る
    groups = []
    for (company, date_val), g in df_sample_sorted.groupby(["_company", col_date], sort=False):
        groups.append((company, date_val, g))
    groups.sort(key=lambda x: (x[0], parse_date(x[1]) or datetime.max))

    wb = Workbook()
    ws = wb.active
    ws.title = "サンプルフィードバック"

    def set_row(row_idx, text, bold=False, size=11, color=None, wrap=False):
        cell = ws.cell(row=row_idx, column=1, value=text)
        cell.font = Font(name=FONT_NAME, size=size, bold=bold, color=color)
        cell.alignment = Alignment(wrap_text=wrap)
        return cell

    set_row(1, f"{args.maker_name}様サンプルフィードバック", bold=True, size=18)
    set_row(2, f"サンプル提供社数：{n_provided}社")
    set_row(3, f"注文転換社数：{n_converted}社")

    rate = (n_converted / n_provided) if n_provided else 0
    cell4 = ws.cell(row=4, column=1, value=f"注文転換率：{rate*100:.1f}%")
    cell4.font = Font(name=FONT_NAME, size=11, bold=True, color="FF0000")
    cell4.alignment = Alignment(wrap_text=False)

    ws.cell(row=6, column=1, value=period_str).font = Font(name=FONT_NAME, size=11)
    ws.cell(row=6, column=1).alignment = Alignment(wrap_text=False)
    updater_cell = ws.cell(row=6, column=8, value="更新日時：なごみや太田")
    updater_cell.font = Font(name=FONT_NAME, size=11)
    updater_cell.alignment = Alignment(wrap_text=False, horizontal="right")

    headers = ["納品日", "業態", "会社名", "配送先都道府県", "商品名", "サンプル用途", "サンプル評価結果", "注文有無"]
    header_row = 7
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=header_row, column=i, value=h)
        c.font = Font(name=FONT_NAME, size=11, bold=True)
        c.fill = GREY_FILL
        c.alignment = Alignment(horizontal="center", wrap_text=True)
        c.border = BORDER_ALL

    row_ptr = header_row + 1
    for company, date_val, g in groups:
        products = list(g[col_product])
        n_rows = max(len(products), 1)
        start_row = row_ptr
        end_row = row_ptr + n_rows - 1

        date_display = date_val
        d_parsed = parse_date(date_val)
        if d_parsed:
            date_display = f"{d_parsed:%Y/%m/%d}"

        has_order = company in companies_with_order
        order_text = "✅ 注文あり" if has_order else "－ なし"

        first = g.iloc[0]
        merged_values = {
            "A": date_display,
            "B": first[col_type] if col_type else "",
            "C": company,
            "D": first[col_pref],
            "F": "",
            "G": "",
            "H": order_text,
        }

        for col_letter, value in merged_values.items():
            col_idx = "ABCDEFGH".index(col_letter) + 1
            cell = ws.cell(row=start_row, column=col_idx, value=value)
            cell.font = Font(name=FONT_NAME, size=11,
                              bold=(col_letter == "H" and has_order),
                              color=("375623" if (col_letter == "H" and has_order)
                                     else ("999999" if col_letter == "H" else None)))
            if col_letter == "H" and has_order:
                cell.fill = GREEN_FILL
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if end_row > start_row:
                ws.merge_cells(start_row=start_row, start_column=col_idx,
                                end_row=end_row, end_column=col_idx)
            for r in range(start_row, end_row + 1):
                ws.cell(row=r, column=col_idx).border = BORDER_ALL

        for offset, product in enumerate(products):
            r = start_row + offset
            cell = ws.cell(row=r, column=5, value=product)
            cell.font = Font(name=FONT_NAME, size=11)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.border = BORDER_ALL

        row_ptr = end_row + 1

    for col_letter, width in COLUMN_WIDTHS.items():
        ws.column_dimensions[col_letter].width = width

    wb.save(args.output)
    print(f"OK: {args.output} (提供社数={n_provided}, 転換社数={n_converted}, 転換率={rate*100:.1f}%)")


if __name__ == "__main__":
    main()
