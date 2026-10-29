#!/usr/bin/env python3
"""
Bカート受注CSV(1本) -> メーカー別サンプルフィードバックExcel 生成スクリプト (v3)

受注CSVには 業態・配送先都道府県・商品名・会社名・納品日・受注番号・セット名・
サンプル用途・サンプル評価結果・評価都合／採用数・受注数・小計 が全て含まれているため、
このCSV1本だけで完結する。

使い方:
  python generate_feedback_excel.py --order-csv bcart_order.csv

--maker-name / --maker-code / --output はすべて省略可。
省略時はCSVの商品名からメーカーコードを自動検出し、data/maker_codes.csv から
メーカー名を自動解決して「{コード}{会社名}様サンプルフィードバック.xlsx」を生成する。
"""
import argparse
import os
import re
import sys
from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.page import PageMargins

MAKER_CODES_CSV = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data", "maker_codes.csv")


def load_maker_code_map(path=MAKER_CODES_CSV):
    if not os.path.exists(path):
        return {}
    df = pd.read_csv(path, dtype=str, keep_default_na=False)
    return dict(zip(df.iloc[:, 0].str.strip(), df.iloc[:, 1].str.strip()))


EXCLUDE_COMPANY_KEYWORDS = ["ナゴミヤ", "なごみや"]

HEADER_KEYWORDS = {
    "納品日": ["納品日"],
    "業態": ["業態"],
    "会社名": ["会社名"],
    "都道府県": ["配送先 都道府県", "配送先都道府県", "都道府県"],
    "商品名": ["商品名"],
    "セット名": ["セット名"],
    "受注番号": ["受注番号"],
    "サンプル用途": ["サンプル用途"],
    "評価結果": ["サンプル評価結果"],
    "採用数": ["評価都合／採用数", "評価都合/採用数"],
    "受注数": ["受注数"],
    "小計": ["小計"],
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
            return pd.read_csv(path, encoding=enc, dtype=str, keep_default_na=False)
        except (UnicodeDecodeError, UnicodeError):
            continue
    raise RuntimeError(f"{path} の文字コードを判定できませんでした")


def find_column(df, keywords):
    for kw in keywords:
        for col in df.columns:
            if kw in str(col):
                return col
    return None


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


# 社内の対応ログ行（例「8:30 2024/06/28　未回収メール済み　石崎」）の判定
LOG_LINE_RE = re.compile(r"^\s*\d{1,2}:\d{2}\s+\d{4}/\d{1,2}/\d{1,2}")
# 「評価がまだ回収できていない」ことを表す言い回し（短い文だけを「未回収」に置き換える）
UNCOLLECTED_HINTS = ("未回収", "回収できていない", "まだ試", "まだ食べ", "まだ召し上", "試せていない",
                     "食べていない", "試食していない", "未試食", "不出", "未提出", "未返信", "未返答")
EVALUATION_WORDS = ("OK", "NG", "採用", "検討", "見送", "注文", "発注", "導入")
UNCOLLECTED_MAX_LEN = 30


def clean_result_text(text):
    """メーカー向けの評価結果表示に整える。
    - 社内の対応ログ行（時刻 日付 …担当者名）は除く（時間・担当者名はメーカーに不要）
    - 評価の本文が無く、ログ行だけの場合は「未回収」と表示する
    - 短い文で「まだ試せていない」など評価未回収の内容だけのものも「未回収」にする"""
    text = str(text).replace("\r\n", "\n").replace("\r", "\n").strip()
    if not text:
        return ""
    lines = text.split("\n")
    if any(LOG_LINE_RE.match(l) for l in lines):
        text = "\n".join(l for l in lines if not LOG_LINE_RE.match(l)).strip()
        if not text:
            return "未回収"
    if (len(text) <= UNCOLLECTED_MAX_LEN
            and any(h in text for h in UNCOLLECTED_HINTS)
            and not any(w in text for w in EVALUATION_WORDS)):
        return "未回収"
    return text


def count_keyword(text, kw):
    return str(text).count(kw)


def to_number(value):
    s = str(value).strip().replace(",", "")
    if s in ("", "nan", "None"):
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


CORP_PREFIXES = ["株式会社", "有限会社", "合同会社", "（株）", "(株)", "（有）", "(有)"]


def clean_company_name(name):
    s = str(name).strip()
    for p in CORP_PREFIXES:
        s = s.replace(p, "")
    return s.strip()


def strip_maker_code(product_name):
    """商品名冒頭の [141] のようなメーカーコード表記を除去する（提出先はメーカー自身なので不要）"""
    return re.sub(r"^\s*\[\d+\]\s*", "", str(product_name)).strip()


def detect_maker_code(series_of_product_names):
    codes = [extract_maker_code(p) for p in series_of_product_names]
    codes = [c for c in codes if c]
    if not codes:
        return None
    return pd.Series(codes).value_counts().idxmax()


FREE_TEXT_NG_KEYWORDS = ["不採用", "見送り", "見送る", "お断り", "断念", "不要", "非導入"]
FREE_TEXT_OK_KEYWORDS = ["採用", "注文します", "注文したい", "発注", "導入", "欲しい", "希望"]


def classify_evaluation(text, adopt_raw, n_products):
    """
    1つの受注（評価コメント1件）についてOK/NG/検討中の件数を判定する。
    戻り値: (ok_count, ng_count, kento_count)

    優先順位:
    1. 文章内にOK/NG/検討中の明示的な表記がある -> その出現回数をそのまま使う
    2. 「評価都合／採用数」が数字（スラッシュを含まない）-> 採用数=OK、
       残り（同一受注内の商品行数－採用数）=NG
    3. 「7/8」のようにスラッシュを含む値（日付の書き間違い、無視する）-> 4へ
    4. 自由記述のみ -> 採用を示す語があればOK、不採用を示す語があればNG、
       それ以外は検討中（該当した1区分のみ+1、他は0）
    """
    text = text or ""
    if any(kw in text for kw in ("OK", "NG", "検討中")):
        return (count_keyword(text, "OK"), count_keyword(text, "NG"), count_keyword(text, "検討中"))

    if adopt_raw and "/" not in adopt_raw and adopt_raw.strip().isdigit():
        adopted = int(adopt_raw.strip())
        not_adopted = max(n_products - adopted, 0)
        return (adopted, not_adopted, 0)

    # 自由記述のみ（採用数が無い、または日付らしきスラッシュ表記で無視）
    if any(kw in text for kw in FREE_TEXT_NG_KEYWORDS):
        return (0, 1, 0)
    if any(kw in text for kw in FREE_TEXT_OK_KEYWORDS):
        return (1, 0, 0)
    return (0, 0, 1)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--order-csv", required=True, nargs="+",
                     help="Bカート受注CSV（1本で完結。500件制限などで複数ファイルに分かれる場合はスペース区切りで複数指定可。同じ内容を結合して処理する）")
    ap.add_argument("--maker-name", default=None, help="省略時はCSVから自動検出したメーカーコードとdata/maker_codes.csvから自動解決")
    ap.add_argument("--maker-code", default=None, help="省略時はCSVの商品名から自動検出。カンマ区切りで複数指定可")
    ap.add_argument("--output", default=None, help="省略時は '{コード}{会社名}様サンプルフィードバック.xlsx' を自動生成")
    args = ap.parse_args()

    df_parts = [read_csv_flexible(p) for p in args.order_csv]
    df_order = pd.concat(df_parts, ignore_index=True) if len(df_parts) > 1 else df_parts[0]
    # 複数ファイルにまたがる場合、受注番号+商品名で完全重複する行があれば1件に統一する
    if len(df_parts) > 1:
        df_order = df_order.drop_duplicates(ignore_index=True)

    o_date = find_column(df_order, HEADER_KEYWORDS["納品日"])
    o_type = find_column(df_order, HEADER_KEYWORDS["業態"])
    o_company = find_column(df_order, HEADER_KEYWORDS["会社名"])
    o_pref = find_column(df_order, HEADER_KEYWORDS["都道府県"])
    o_product = find_column(df_order, HEADER_KEYWORDS["商品名"])
    o_set = find_column(df_order, HEADER_KEYWORDS["セット名"])
    o_orderno = find_column(df_order, HEADER_KEYWORDS["受注番号"])
    o_qty = find_column(df_order, HEADER_KEYWORDS["受注数"])
    o_subtotal = find_column(df_order, HEADER_KEYWORDS["小計"])
    o_usage = find_column(df_order, HEADER_KEYWORDS["サンプル用途"])
    o_result = find_column(df_order, HEADER_KEYWORDS["評価結果"])
    o_adopt = find_column(df_order, HEADER_KEYWORDS["採用数"])

    missing = [n for n, c in [
        ("納品日", o_date), ("会社名", o_company), ("配送先都道府県", o_pref),
        ("商品名", o_product), ("セット名", o_set), ("受注番号", o_orderno),
    ] if c is None]
    if missing:
        sys.stderr.write("以下の列が見つかりませんでした: " + ", ".join(missing) + "\n")
        sys.exit(1)

    # メーカーコードの決定（未指定ならCSVの商品名から自動検出）
    if args.maker_code:
        maker_code_for_name = args.maker_code.split(",")[0].strip()
    else:
        detected = detect_maker_code(df_order[o_product])
        if detected:
            args.maker_code = detected
            maker_code_for_name = detected
        else:
            maker_code_for_name = None

    # メーカー名の決定（未指定ならコード表から自動解決、法人格は除去）
    if not args.maker_name:
        if not maker_code_for_name:
            sys.stderr.write("メーカーコードをCSVから自動検出できませんでした。--maker-name か --maker-code を指定してください。\n")
            sys.exit(1)
        code_map = load_maker_code_map()
        resolved = code_map.get(maker_code_for_name)
        if not resolved:
            sys.stderr.write(f"メーカーコード {maker_code_for_name} が data/maker_codes.csv に見つかりません。--maker-name を指定してください。\n")
            sys.exit(1)
        args.maker_name = f"{maker_code_for_name}{clean_company_name(resolved)}"

    if not args.output:
        args.output = f"{args.maker_name}様サンプルフィードバック.xlsx"

    # 自社テスト除外
    df_order = df_order[~df_order[o_company].astype(str).apply(
        lambda v: any(kw in v for kw in EXCLUDE_COMPANY_KEYWORDS)
    )]

    # メーカーコードで絞り込み
    if args.maker_code:
        codes = [c.strip() for c in args.maker_code.split(",") if c.strip()]
        df_order = df_order[df_order[o_product].apply(lambda p: extract_maker_code(p) in codes)]

    if df_order.empty:
        sys.stderr.write("絞り込み後、対象データが0件でした。メーカーコードを確認してください。\n")
        sys.exit(1)

    df_order["_is_sample"] = df_order[o_set].astype(str).str.contains("サンプル", na=False)
    df = df_order[df_order["_is_sample"]].copy()

    if df.empty:
        sys.stderr.write("サンプル注文（セット名に「サンプル」を含む行）が見つかりませんでした。\n")
        sys.exit(1)

    # 通常注文（サンプルでない注文）の「会社×商品名」ごとの納品日一覧（商品名は完全一致で突き合わせる）
    normal_df = df_order[~df_order["_is_sample"]]
    normal_dates = {}
    for c_, p_, d_ in zip(normal_df[o_company].astype(str),
                          normal_df[o_product].astype(str).str.strip(),
                          normal_df[o_date].apply(parse_date)):
        if d_ is not None:
            normal_dates.setdefault((c_, p_), []).append(d_)

    df["_company"] = df[o_company].astype(str)
    df["_date_parsed"] = df[o_date].apply(parse_date)

    df["_product"] = df[o_product].astype(str).str.strip()
    df["_pair"] = list(zip(df["_company"], df["_product"]))
    provided_companies = sorted(set(df["_company"]))
    n_companies = len(provided_companies)
    sample_pairs = set(df["_pair"])
    n_provided = len(sample_pairs)  # サンプル提供数（会社×商品）

    # 転換 = サンプル納品後に、同じ会社が同じ商品を通常注文（サンプルでない注文）している
    # 判定は「通常注文の納品日 > そのサンプルの（最も早い）納品日」。サンプル納品日が不明なら注文の有無のみで判定
    first_sample_date = {}
    for pr, d_ in zip(df["_pair"], df["_date_parsed"]):
        if d_ is not None and (pr not in first_sample_date or d_ < first_sample_date[pr]):
            first_sample_date[pr] = d_

    def is_converted(pr):
        dates = normal_dates.get(pr)
        if not dates:
            return False
        sd = first_sample_date.get(pr)
        return True if sd is None else any(d_ > sd for d_ in dates)

    converted_pairs = {pr for pr in sample_pairs if is_converted(pr)}
    companies_converted = {pr[0] for pr in converted_pairs}
    n_converted = len(converted_pairs)
    rate = (n_converted / n_provided) if n_provided else 0
    n_converted_co = len(companies_converted)
    rate_co = (n_converted_co / n_companies) if n_companies else 0

    all_dates = [d for d in df["_date_parsed"] if d is not None]
    period_str = f"{min(all_dates):%Y/%m/%d}〜{max(all_dates):%Y/%m/%d}" if all_dates else ""

    df_sorted = df.sort_values(["_company", "_date_parsed"])

    groups = []
    for (company, date_val), g in df_sorted.groupby(["_company", o_date], sort=False):
        groups.append((company, date_val, g))
    groups.sort(key=lambda x: (parse_date(x[1]) or datetime.max, x[0]))

    # ---------------- 業態別集計（同じシートの下に表示する） ----------------
    company_type = {}
    for _, row in df.iterrows():
        c = row["_company"]
        t = row.get(o_type)
        if c not in company_type and pd.notna(t) and str(t).strip():
            company_type[c] = str(t).strip()

    # ---------------- Excel生成 ----------------
    wb = Workbook()
    ws = wb.active
    ws.title = "サンプル結果"

    def set_row(row_idx, text, bold=False, size=11, color=None):
        cell = ws.cell(row=row_idx, column=1, value=text)
        cell.font = Font(name=FONT_NAME, size=size, bold=bold, color=color)
        cell.alignment = Alignment(vertical="center", wrap_text=False)
        return cell

    set_row(1, f"{args.maker_name}様サンプルフィードバック", bold=True, size=18)
    set_row(2, f"サンプル提供数：{n_provided}件（商品単位）")
    set_row(3, f"注文転換数：{n_converted}件（サンプル納品後に同じ商品が通常注文された数）")
    cell4 = set_row(4, f"注文転換率：{rate*100:.1f}%", bold=True, color="FF0000")
    # 会社単位（そのメーカーの商品を1件でも注文した会社）はE列に並べる
    def set_row_e(row_idx, text, bold=False, color=None):
        c = ws.cell(row=row_idx, column=5, value=text)
        c.font = Font(name=FONT_NAME, size=11, bold=bold, color=color)
        c.alignment = Alignment(vertical="center", wrap_text=False)
    set_row_e(2, f"サンプル提供社数：{n_companies}社（会社単位）")
    set_row_e(3, f"注文転換社数：{n_converted_co}社（サンプル納品後に同じ商品を通常注文）")
    set_row_e(4, f"注文転換率（会社単位）：{rate_co*100:.1f}%", bold=True, color="FF0000")

    set_row(6, period_str)
    today = datetime.now()
    updater_cell = ws.cell(row=6, column=8, value=f"{today.year}/{today.month}/{today.day}　なごみや")
    updater_cell.font = Font(name=FONT_NAME, size=11)
    updater_cell.alignment = Alignment(vertical="center", wrap_text=False, horizontal="right")

    headers = ["納品日", "業態", "会社名", "配送先都道府県", "商品名", "サンプル用途", "サンプル評価結果", "注文有無"]
    header_row = 7
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=header_row, column=i, value=h)
        c.font = Font(name=FONT_NAME, size=11, bold=True)
        c.fill = GREY_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER_ALL

    row_ptr = header_row + 1
    for company, date_val, g in groups:
        products = list(g[o_product])
        n_rows = max(len(products), 1)
        start_row = row_ptr
        end_row = row_ptr + n_rows - 1

        d_parsed = parse_date(date_val)
        date_display = f"{d_parsed:%Y/%m/%d}" if d_parsed else date_val

        first = g.iloc[0]
        result_text = str(first[o_result]).strip() if o_result else ""
        # 「評価都合／採用数」はメーカーに不要なので表示しない（集計には内部で使う）
        g_value = clean_result_text(result_text)

        merged_values = {
            "A": date_display,
            "B": first.get(o_type, "") if o_type else "",
            "C": company,
            "D": first.get(o_pref, ""),
            "F": str(first[o_usage]).strip() if o_usage else "",
            "G": g_value,
        }

        for col_letter, value in merged_values.items():
            col_idx = "ABCDEFGH".index(col_letter) + 1
            cell = ws.cell(row=start_row, column=col_idx, value=value)
            cell.font = Font(name=FONT_NAME, size=11,
                              bold=False)
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

            # 注文有無は商品ごと（サンプルと同じ商品が同じ会社から注文されたか）
            has_order = (company, str(product).strip()) in converted_pairs
            hcell = ws.cell(row=r, column=8, value="✅ 注文あり" if has_order else "－ なし")
            hcell.font = Font(name=FONT_NAME, size=11, bold=has_order,
                              color=("375623" if has_order else "999999"))
            if has_order:
                hcell.fill = GREEN_FILL
            hcell.alignment = Alignment(vertical="center", wrap_text=True)
            hcell.border = BORDER_ALL

        row_ptr = end_row + 1

    for col_letter, width in COLUMN_WIDTHS.items():
        ws.column_dimensions[col_letter].width = width

    # ---------------- 業態別（提供社数・注文社数）: 同じシートの下に置く ----------------
    biz_stats = {}
    for c in provided_companies:
        t = company_type.get(c, "(不明)")
        st = biz_stats.setdefault(t, {"提供": 0, "転換": 0})
        st["提供"] += 1
        if c in companies_converted:
            st["転換"] += 1

    r = row_ptr + 1
    t_cell = ws.cell(row=r, column=1, value="業態別（提供社数・注文社数）")
    t_cell.font = Font(name=FONT_NAME, size=11, bold=True)
    t_cell.alignment = Alignment(vertical="center", wrap_text=False)
    r += 1

    def biz_row(row_idx, label, n_prov, n_conv, header=False, bold=False):
        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=2)
        vals = {1: label, 3: n_prov, 4: n_conv}
        for col in (1, 2, 3, 4):
            cell = ws.cell(row=row_idx, column=col, value=vals.get(col))
            cell.font = Font(name=FONT_NAME, size=11, bold=(header or bold))
            cell.border = BORDER_ALL
            cell.alignment = Alignment(vertical="center", wrap_text=True,
                                       horizontal=("center" if (header or col > 2) else None))
            if header:
                cell.fill = GREY_FILL

    biz_row(r, "業態", "提供社数", "注文社数", header=True)
    r += 1
    for biz, st in sorted(biz_stats.items(), key=lambda kv: kv[1]["提供"], reverse=True):
        biz_row(r, biz, st["提供"], st["転換"])
        r += 1
    biz_row(r, "合計", n_companies, n_converted_co, bold=True)

    # ---------------- 印刷設定: A4横・1ページ幅に収める・印刷範囲はA1〜H最終行 ----------------
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0  # 縦は自動（複数ページ可）
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_area = f"A1:H{r}"
    ws.print_title_rows = f"{header_row}:{header_row}"  # 2ページ目以降もヘッダー行を繰り返す
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.6, header=0.3, footer=0.3)
    ws.print_options.horizontalCentered = True
    ws.oddFooter.center.text = "&P / &N"

    wb.save(args.output)
    print(f"OK: {args.output} (提供数={n_provided}件, 転換数={n_converted}件, 転換率={rate*100:.1f}% / 会社単位 {n_converted_co}/{n_companies}社 {rate_co*100:.1f}%)")


if __name__ == "__main__":
    main()
