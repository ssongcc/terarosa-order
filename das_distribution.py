"""분배모드 원클릭 실행 (로컬 전용 — 개인정보 외부 전송 없음)
사용법:
  python 분배모드_실행.py 발송양식파일.xlsx
  또는 인자 없이 실행하면 같은 폴더에서 최신 '발송양식' xlsx를 자동으로 찾음
출력 (원본은 절대 수정하지 않음):
  ①업로드용_정렬본_*.xlsx : 원본 열 그대로 + 순서만 변경 → 그대로 우체국 접수
  ②작업지시서_*.xlsx     : 줄포장 / 씨딩지시 / 요약
"""
import sys, re, glob, os, datetime
import pandas as pd
from collections import defaultdict, Counter

BATCH_SIZE = 24  # 선반 칸수 바뀌면 여기만 수정
ITEM_COL = '상품명(확정)+옵션(확정)+수량(조합용)'
ORD_COLS = ['주문번호(쇼핑몰)', '주문번호(사방넷)']  # 쇼핑몰 번호가 합포(주문 묶음) 기준

def _parse_item(s):
    s = str(s or '').strip()
    m = re.search(r'_(\d+)(?:ea)?\s*$', s)  # _2ea, _2 두 형식 모두 지원
    return (s[:m.start()], int(m.group(1))) if m else (s, 1)

def run(df, batch_size=BATCH_SIZE):
    df = df.reset_index(drop=True)
    def okey(row):
        for c in ORD_COLS:
            v = row.get(c)
            if v is not None and str(v).strip() and str(v).lower() != 'nan':
                if isinstance(v, float) and v == int(v):
                    v = int(v)
                return str(v).strip()
        return None
    df['_okey'] = df.apply(okey, axis=1)
    df['_okey'] = df['_okey'].fillna(pd.Series('행' + df.index.astype(str), index=df.index))

    orders = defaultdict(list)
    for i, k in df['_okey'].items():
        orders[k].append(i)
    skus = {k: [_parse_item(df.at[i, ITEM_COL]) for i in idx] for k, idx in orders.items()}
    singles = {k: v for k, v in skus.items() if len(v) == 1}
    multis  = {k: v for k, v in skus.items() if len(v) > 1}

    sku_tot = Counter()
    for k, v in singles.items():
        sku_tot[v[0][0]] += v[0][1]
    single_sorted = sorted(singles, key=lambda k: (-sku_tot[singles[k][0][0]], singles[k][0][0], -singles[k][0][1], k))
    multi_sorted = sorted(multis, key=lambda k: tuple(sorted(s for s, q in multis[k])))
    assign = {k: (i // batch_size + 1, i % batch_size + 1) for i, k in enumerate(multi_sorted)}

    seq_keys = single_sorted + multi_sorted
    rows_out, serials, serial = [], {}, 0
    for k in seq_keys:
        serial += 1
        serials[k] = serial
        rows_out.extend(orders[k])
    upload = df.loc[rows_out].drop(columns=['_okey']).reset_index(drop=True)

    line_rows = [{'송장순번': serials[k], 'SKU': singles[k][0][0], '수량': singles[k][0][1], '주문번호': k}
                 for k in single_sorted]
    seed_rows = []
    for b in sorted({assign[k][0] for k in multi_sorted}):
        keys_b = [k for k in multi_sorted if assign[k][0] == b]
        sku_slots = defaultdict(list)
        for k in keys_b:
            for s, q in multis[k]:
                sku_slots[s].append((assign[k][1], q))
        for s in sorted(sku_slots, key=lambda x: -sum(q for _, q in sku_slots[x])):
            slots = sorted(sku_slots[s])
            seed_rows.append({'배치': b, 'SKU': s, '총수량': sum(q for _, q in slots),
                              '분배': ', '.join(f'{c}번칸×{q}' if q > 1 else f'{c}번칸' for c, q in slots)})
        for k in keys_b:
            seed_rows.append({'배치': b, 'SKU': f'[칸배정] {assign[k][1]}번칸', '총수량': len(multis[k]),
                              '분배': f'송장순번 {serials[k]} / 주문 {k}'})
    stats = {'총주문': len(orders), '단일': len(singles), '복합': len(multis),
             '배치수': max((assign[k][0] for k in multi_sorted), default=0),
             '씨딩SKU방문': sum(1 for r in seed_rows if not str(r['SKU']).startswith('[칸배정]'))}
    return upload, pd.DataFrame(line_rows), pd.DataFrame(seed_rows), stats, df
