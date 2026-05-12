"""
테라로사 자사몰 주문취합 처리 스크립트
"""

import re
import sys
from datetime import datetime
from copy import copy

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

REMOVE_STRINGS = [
    "/구매 안함", "/플러스", "[플러스] ", "/불필요", "/필요",
    "불필요", "필요", "/상자 없음", "/포장 없음",
    "테라로사 시그니처 ", "[Online Exclusive] ", "[Online Exclusive/플러스] ",
    "[C.O.E/플러스] ",
]

TEXT_REPLACE = {
    "중간 분쇄(드립용)": "드립용",
    "가는 분쇄(에스프레소용)": "에스프레소용",
}

COLOR_DRIP = "E2EFDA"
COLOR_SCOOP = "FFF2CC"
COLOR_HEADER = "D9D9D9"
COLOR_WHITE = "FFFFFF"
COL_WIDTHS = {"A": 42, "B": 10, "C": 35, "D": 8, "E": 18}

THIN_BORDER = Border(
    left=Side(style="thin", color="BFBFBF"),
    right=Side(style="thin", color="BFBFBF"),
    top=Side(style="thin", color="BFBFBF"),
    bottom=Side(style="thin", color="BFBFBF"),
)


def load_order_data(filepath):
    df = pd.read_excel(filepath, sheet_name="취합용", header=0, dtype=str)
    df = df.iloc[:, :2].copy()
    df.columns = ["품목명_원본", "수량"]
    df.dropna(subset=["품목명_원본"], inplace=True)
    df["수량"] = pd.to_numeric(df["수량"], errors="coerce").fillna(0).astype(int)
    return df.reset_index(drop=True)


def load_code_data(filepath):
    df = pd.read_excel(filepath, header=0, dtype=str)
    df.fillna("", inplace=True)
    return df


def clean_item_name(name):
    for s in REMOVE_STRINGS:
        name = name.replace(s, "")
    for old, new in TEXT_REPLACE.items():
        name = name.replace(old, new)
    if "어센틱" in name and "정기 배송" in name:
        if "_" in name:
            option = name.split("_", 1)[1]
            name = "어센틱 에스프레소 블렌드_" + option
        else:
            name = "어센틱 에스프레소 블렌드"
    return name.strip()


def apply_sos_weight(name, weight):
    if "S.O.S" in name and weight == "300g":
        return "150g"
    return weight


WEIGHT_PATTERN = re.compile(r"(\d+(?:\.\d+)?\s*(?:kg|g))", re.IGNORECASE)


def extract_weight(text):
    m = WEIGHT_PATTERN.search(text)
    if m:
        w = m.group(1).replace(" ", "")
        rest = (text[:m.start()] + text[m.end():]).strip().strip("/").strip()
        return w, rest
    return "", text


def split_item(raw_name):
    if "[첫 구매 찬스]" in raw_name:
        name = raw_name.replace("[첫 구매 찬스] ", "").replace("250g", "").strip()
        return name, "250g", ""
    if "무료원두 쿠폰" in raw_name:
        return "무료원두 쿠폰 250g", "250g", "증정 원두"
    if "이 달의 킹콩" in raw_name or "이달의 킹콩" in raw_name:
        return raw_name, "500g", "플러스쿠폰"
    if "이 달의 드립백" in raw_name or "이달의 드립백" in raw_name:
        if "_" in raw_name:
            parts = raw_name.split("_", 1)
            item_name = parts[0].strip()
            option = parts[1].strip()
            weight, option = extract_weight(option)
            return item_name, weight, option
        return raw_name, "", ""
    if "원두&커피 스쿱 세트" in raw_name or "원두 & 커피 스쿱 세트" in raw_name:
        if "_" in raw_name:
            parts = raw_name.split("_", 1)
            item_name = parts[0].strip()
            option = parts[1].strip()
            option = option.replace("갈지않음/", "").replace("(250g)", "").strip()
            return item_name, "250g", option
        return raw_name, "250g", ""
    if "_" in raw_name:
        parts = raw_name.split("_", 1)
        item_name = parts[0].strip()
        option = parts[1].strip()
        weight, option = extract_weight(option)
        return item_name, weight, option
    weight, rest = extract_weight(raw_name)
    if weight:
        return rest if rest else raw_name, weight, ""
    return raw_name, "", ""


def resolve_kingkong_name(df):
    king_rows = df[df["품목명"].str.contains(r"[Kk][Ii][Nn][Gg]콩", na=False)]
    mask = df["품목명"].str.contains("이 달의 킹콩|이달의 킹콩", na=False)
    if not mask.any():
        return df
    if not king_rows.empty:
        king_name = king_rows.iloc[0]["품목명"]
    else:
        month = datetime.today().month
        king_name = f"[{month}월 KING콩]"
    df.loc[mask, "품목명"] = king_name
    df.loc[mask, "중량"] = "500g"
    df.loc[mask, "옵션"] = "플러스쿠폰"
    return df


def clean_kingkong_options(df):
    mask = df["품목명"].str.contains(r"[Kk][Ii][Nn][Gg]콩", na=False)
    for keyword in ["테라로사 바리스타", "에티오피아 농부", "멕시코 농장주"]:
        df.loc[mask, "옵션"] = df.loc[mask, "옵션"].str.replace(
            r"\s*/{1,2}\s*" + keyword, "", regex=True
        ).str.strip()
    df.loc[mask, "옵션"] = df.loc[mask, "옵션"].str.strip("/").str.strip()
    return df


def merge_gratitude_month(df):
    mask = df["품목명"].str.contains("감사의 달", na=False)
    if not mask.any():
        return df
    gdf = df[mask].copy()
    others = df[~mask].copy()
    merged2_rows = []
    gdf_with_opt = gdf[gdf["옵션"].str.strip() != ""].copy()
    gdf_no_opt = gdf[gdf["옵션"].str.strip() == ""].copy()
    if not gdf_with_opt.empty:
        gdf_with_opt["_key"] = gdf_with_opt["옵션"].str[:5]
        for key, g in gdf_with_opt.groupby("_key", sort=False):
            row = {
                "품목명": "[감사의 달] 2026 선물대전",
                "중량": g.iloc[0]["중량"],
                "옵션": min(g["옵션"].values, key=len),
                "수량": g["수량"].sum(),
            }
            merged2_rows.append(row)
    if not gdf_no_opt.empty:
        for name, g in gdf_no_opt.groupby("품목명", sort=False):
            row = {
                "품목명": name,
                "중량": g.iloc[0]["중량"],
                "옵션": "",
                "수량": g["수량"].sum(),
            }
            merged2_rows.append(row)
    merged2 = pd.DataFrame(merged2_rows)
    return pd.concat([others, merged2], ignore_index=True)


def classify(row):
    name = row["품목명"]
    weight = str(row["중량"])
    if "드립백" in name:
        return "드립백"
    if "원두&커피 스쿱 세트" in name or "원두 & 커피 스쿱 세트" in name:
        return "스쿱세트"
    if "세트" in name:
        return "세트"
    if re.search(r"\d+\s*(?:g|kg)", weight, re.IGNORECASE):
        return "원두"
    return "기타"


GROUP_ORDER = {"세트": 0, "기타": 1, "드립백": 2, "스쿱세트": 3, "원두": 4}


def weight_to_gram(w):
    w = str(w).strip()
    m = re.match(r"([\d.]+)\s*(kg|g)", w, re.IGNORECASE)
    if not m:
        return 0
    val = float(m.group(1))
    return val * 1000 if m.group(2).lower() == "kg" else val


def aggregate_and_sort(df):
    df = df.copy()
    df["_group"] = df.apply(classify, axis=1)

    def agg_key(row):
        g = row["_group"]
        if g in ("세트", "드립백", "스쿱세트"):
            return (row["품목명"], row["옵션"])
        return (row["품목명"], row["중량"], row["옵션"])

    df["_key"] = df.apply(agg_key, axis=1)
    df = df.groupby("_key", sort=False).agg(
        품목명=("품목명", "first"),
        중량=("중량", "first"),
        옵션=("옵션", "first"),
        수량=("수량", "sum"),
        _group=("_group", "first"),
    ).reset_index(drop=True)
    df["_g_order"] = df["_group"].map(GROUP_ORDER)
    df["_w_gram"] = df["중량"].apply(weight_to_gram)
    df.sort_values(["_g_order", "품목명", "_w_gram", "옵션"], ascending=[True, True, False, True], inplace=True)
    df.reset_index(drop=True, inplace=True)
    return df


def match_code(row, code_df):
    name = str(row["품목명"]).strip()
    weight = str(row["중량"]).strip()
    option = str(row["옵션"]).strip()
    c_code = code_df.columns[0]
    c_name = code_df.columns[1]
    c_opt = code_df.columns[2]

    def eq(col, val):
        return code_df[col].str.strip() == val.strip()

    if re.search(r"[Kk][Ii][Nn][Gg]콩", name):
        res = code_df[eq(c_name, name) & (code_df[c_opt].str.strip() == "500g")]
        if not res.empty:
            return str(res.iloc[0][c_code])
        res = code_df[eq(c_name, name)]
        if not res.empty:
            return str(res.iloc[0][c_code])
    if "S.O.S" in name:
        res = code_df[eq(c_name, name) & eq(c_opt, weight)]
        if not res.empty:
            return str(res.iloc[0][c_code])
    if "스쿱 세트" in name or "스쿱세트" in name:
        opt_with_weight = option + "(250g)"
        res = code_df[eq(c_opt, opt_with_weight)]
        if not res.empty:
            return str(res.iloc[0][c_code])
    if "TO-GO" in name or "to-go" in name.lower():
        opt_no_color = re.sub(r"/블랙|/투명|/화이트|/레드", "", option).strip()
        res = code_df[eq(c_name, name) & eq(c_opt, opt_no_color)]
        if not res.empty:
            return str(res.iloc[0][c_code])
    if weight:
        res = code_df[eq(c_name, name) & eq(c_opt, weight)]
        if not res.empty:
            return str(res.iloc[0][c_code])
    res = code_df[eq(c_name, name) & eq(c_opt, option)]
    if not res.empty:
        return str(res.iloc[0][c_code])
    sorted_opt = "+".join(sorted(option.split("+")))
    name_rows = code_df[code_df[c_name].str.strip() == name]
    if not name_rows.empty:
        match = name_rows[name_rows[c_opt].apply(lambda x: "+".join(sorted(str(x).split("+"))) == sorted_opt)]
        if not match.empty:
            return str(match.iloc[0][c_code])
    res = code_df[(code_df[c_name].str.strip() == name) & (code_df[c_opt].str.strip() == "")]
    if not res.empty:
        return str(res.iloc[0][c_code])
    if "옥스포드" in name:
        res = code_df[eq(c_name, name) & eq(c_opt, option)]
        if not res.empty:
            return str(res.iloc[0][c_code])
    return ""


def build_sheet3(raw_df):
    targets = {
        "테라로사 바리스타": "/ 테라로사 바리스타",
        "에티오피아 농부": "/ 에티오피아 농부",
        "멕시코 농장주": "/ 멕시코 농장주",
    }
    rows = []
    for label, keyword in targets.items():
        mask = raw_df["품목명_원본"].str.contains(keyword, na=False)
        qty = raw_df.loc[mask, "수량"].sum()
        rows.append({
            "품목명": "옥스포드 피규어",
            "빈칸": "",
            "이름": label,
            "수량": qty if qty > 0 else "-",
        })
    return pd.DataFrame(rows)


def build_sheet2(main_df):
    bean_mask = main_df["_group"] == "원두"
    scoop_mask = main_df["_group"] == "스쿱세트"
    rows = {}
    for _, r in main_df[bean_mask].iterrows():
        name = r["품목명"]
        g = weight_to_gram(r["중량"]) * r["수량"]
        rows[name] = rows.get(name, 0) + g
    for _, r in main_df[scoop_mask].iterrows():
        name = r["옵션"]
        g = 250 * r["수량"]
        rows[name] = rows.get(name, 0) + g
    result = []
    for name, total_g in rows.items():
        result.append({"품목명": name, "중량(kg)": round(total_g / 1000, 3)})
    return pd.DataFrame(result)


def apply_style(ws, df_with_groups):
    header_fill = PatternFill("solid", fgColor=COLOR_HEADER)
    drip_fill = PatternFill("solid", fgColor=COLOR_DRIP)
    scoop_fill = PatternFill("solid", fgColor=COLOR_SCOOP)
    white_fill = PatternFill("solid", fgColor=COLOR_WHITE)
    header_font = Font(name="Arial", size=10, bold=True)
    body_font = Font(name="Arial", size=10)

    # 1행 빈행
    ws.cell(row=1, column=1, value="")

    # 2행 헤더
    headers = ["품목명", "중량", "옵션", "수량", "자사몰상품코드"]
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = THIN_BORDER

    row_num = 3  # 데이터는 3행부터
    prev_name = None

    for _, r in df_with_groups.iterrows():
        cur_name = r["품목명"]
        group = r.get("_group", "")

        if prev_name is not None and cur_name != prev_name:
            for col_idx in range(1, 6):
                ws.cell(row=row_num, column=col_idx).border = THIN_BORDER
            row_num += 1

        if group == "드립백":
            fill = drip_fill
        elif group == "스쿱세트":
            fill = scoop_fill
        else:
            fill = white_fill

        values = [cur_name, r["중량"], r["옵션"], r["수량"], r.get("상품코드", "")]
        for col_idx, val in enumerate(values, 1):
            cell = ws.cell(row=row_num, column=col_idx, value=val)
            cell.font = body_font
            cell.fill = fill
            cell.border = THIN_BORDER

        prev_name = cur_name
        row_num += 1

    for col_letter, width in COL_WIDTHS.items():
        ws.column_dimensions[col_letter].width = width


def write_simple_sheet(ws, df, title_row):
    header_fill = PatternFill("solid", fgColor=COLOR_HEADER)
    header_font = Font(name="Arial", size=10, bold=True)
    body_font = Font(name="Arial", size=10)
    for col_idx, h in enumerate(title_row, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = THIN_BORDER
    for r_idx, row in df.iterrows():
        for col_idx, val in enumerate(row.values, 1):
            cell = ws.cell(row=r_idx + 2, column=col_idx, value=val)
            cell.font = body_font
            cell.border = THIN_BORDER


def insert_sheet3_into_sheet1(wb):
    ws1 = wb["주문취합"]
    ws3 = wb["바리스타·농부·농장주"]

    max_row_3 = ws3.max_row
    insert_count = max_row_3 - 1  # 헤더 제외 (데이터만)
    if insert_count <= 0:
        return

    # 2행에 삽입 (1행 빈행 유지, +1은 구분 빈행)
    ws1.insert_rows(2, amount=insert_count + 1)

    # 시트3 데이터(2행~) → 시트1 2행~에 복사
    for src_row_idx in range(2, max_row_3 + 1):
        dest_row_idx = src_row_idx
        for col_idx in range(1, ws3.max_column + 1):
            src_cell = ws3.cell(row=src_row_idx, column=col_idx)
            dst_cell = ws1.cell(row=dest_row_idx, column=col_idx)
            dst_cell.value = src_cell.value
            if src_cell.has_style:
                dst_cell.font = copy(src_cell.font)
                dst_cell.fill = copy(src_cell.fill)
                dst_cell.border = copy(src_cell.border)
                dst_cell.alignment = copy(src_cell.alignment)

    # 데이터 삽입 후 구분 빈행
    blank_row = 1 + insert_count + 1
    for col_idx in range(1, 6):
        ws1.cell(row=blank_row, column=col_idx).border = THIN_BORDER


def main(order_file, code_file, output_file=None):
    raw_df = load_order_data(order_file)
    code_df = load_code_data(code_file)
    sheet3_df = build_sheet3(raw_df)
    raw_df["품목명_정리"] = raw_df["품목명_원본"].apply(clean_item_name)
    split_result = raw_df["품목명_정리"].apply(split_item)
    raw_df[["품목명", "중량", "옵션"]] = pd.DataFrame(split_result.tolist(), index=raw_df.index)
    raw_df["중량"] = raw_df.apply(lambda r: apply_sos_weight(r["품목명"], r["중량"]), axis=1)
    raw_df = resolve_kingkong_name(raw_df)
    raw_df = clean_kingkong_options(raw_df)
    raw_df = merge_gratitude_month(raw_df)
    main_df = aggregate_and_sort(raw_df)
    main_df["상품코드"] = main_df.apply(lambda r: match_code(r, code_df), axis=1)
    sheet2_df = build_sheet2(main_df)

    wb = Workbook()
    ws1 = wb.active
    ws1.title = "주문취합"
    apply_style(ws1, main_df)

    ws2 = wb.create_sheet("원두 중량 합산")
    write_simple_sheet(ws2, sheet2_df, ["품목명", "중량(kg)"])

    ws3 = wb.create_sheet("바리스타·농부·농장주")
    write_simple_sheet(ws3, sheet3_df, ["품목명", "빈칸", "이름", "수량"])

    insert_sheet3_into_sheet1(wb)

    today = datetime.today().strftime("%Y%m%d")
    output_path = output_file if output_file else f"자사몰주문취합_{today}.xlsx"
    wb.save(output_path)
    print(f"저장 완료: {output_path}")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("사용법: python terarosa_order_processor.py <주문취합.xlsx> <자사몰상품코드.xlsx>")
        sys.exit(1)
    main(sys.argv[1], sys.argv[2])
