"""
테라로사 DAS 작업 화면 (태블릿/모바일용)
씨딩 작업 + 배치 품목 리스트
"""
import streamlit as st
from github_storage import gh_load

st.set_page_config(
    page_title="DAS 작업",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
section[data-testid="stSidebar"] { display: none; }
.block-container { padding: 3rem 0.8rem 1rem 0.8rem !important; max-width: 100% !important; }
.sku-card {
    background: #FAF3F0; border: 3px solid #8B3A2A;
    border-radius: 16px; padding: 20px 24px; margin-bottom: 16px;
}
.sku-name { color: #8B3A2A; font-size: 1.8rem; font-weight: 800; line-height: 1.3; margin-bottom: 6px; }
.sku-total { color: #6B6056; font-size: 1.1rem; margin-bottom: 16px; }
.slot-grid { display: flex; flex-wrap: wrap; gap: 10px; margin-top: 6px; }
.slot-chip {
    background: #8B3A2A; color: white; font-size: 1.3rem; font-weight: 700;
    padding: 10px 18px; border-radius: 10px; white-space: nowrap;
}
.progress-bar-wrap {
    background: #EDE5DC; border-radius: 8px; height: 12px; margin: 6px 0 4px 0; overflow: hidden;
}
.progress-bar-fill { background: #8B3A2A; height: 100%; border-radius: 8px; }
.batch-badge {
    background: #C4644A; color: white; font-size: 0.95rem; font-weight: 700;
    padding: 4px 14px; border-radius: 20px; display: inline-block; margin-bottom: 10px;
}
.batch-title {
    background: #8B3A2A; color: white; font-size: 1rem; font-weight: 700;
    padding: 10px 16px; border-radius: 10px 10px 0 0; margin-bottom: 0;
}
.done-box {
    background: #E8F5E9; border: 3px solid #388E3C;
    border-radius: 16px; padding: 32px; text-align: center;
}
.stButton > button {
    font-size: 1.2rem !important; font-weight: 700 !important;
    padding: 16px 0 !important; border-radius: 12px !important;
    border: none !important; width: 100%;
}
.btn-prev > button   { background: #EDE5DC !important; color: #2C2C2C !important; }
.btn-next > button   { background: #8B3A2A !important; color: white !important; }
.btn-batch > button  { background: #EDE5DC !important; color: #8B3A2A !important;
                       font-size: 0.95rem !important; padding: 9px 0 !important; }
.btn-restart > button{ background: #6B6056 !important; color: white !important; }
.btn-load > button   { background: #C4644A !important; color: white !important;
                       font-size: 0.9rem !important; padding: 9px 0 !important; }
.btn-sm > button     { font-size: 1rem !important; padding: 12px 0 !important; }
@media (max-width: 600px) {
    .sku-name  { font-size: 1.2rem !important; }
    .slot-chip { font-size: 1rem !important; padding: 8px 12px !important; }
    .stButton > button { font-size: 1rem !important; padding: 13px 0 !important; }
}
</style>
""", unsafe_allow_html=True)

# ── 세션 초기화 ──
defaults = {
    "all_batches": {}, "batch_list": [], "cur_batch": None, "seed_idx": 0, "loaded": False,
    "all_prep": {}, "prep_cur_batch": None, "prep_page": 0,
    "auto_loaded": False,
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

ITEMS_PER_PAGE = 8  # 배치 시트는 한 화면에 더 많이 표시

def load_all():
    data = gh_load("das_data.json", None)
    if data and isinstance(data, dict) and "batches" in data:
        st.session_state.all_batches = {str(k): v for k, v in data["batches"].items()}
        st.session_state.batch_list  = data.get("batch_list", [])
        st.session_state.cur_batch   = str(data["batch_list"][0]) if data.get("batch_list") else None
        st.session_state.seed_idx    = 0
        st.session_state.loaded      = True
        if "all_prep" in data:
            st.session_state.all_prep       = {str(k): v for k, v in data["all_prep"].items()}
            first = str(data["batch_list"][0]) if data.get("batch_list") else None
            st.session_state.prep_cur_batch = first
            st.session_state.prep_page      = 0
        return True
    return False

def parse_slots(분배_str):
    slots = []
    for part in str(분배_str).split(","):
        part = part.strip()
        if "×" in part:
            칸, 수량 = part.split("×")
            slots.append((칸.strip(), int(수량.strip())))
        elif "번칸" in part:
            slots.append((part.strip(), 1))
    return slots

# ── 자동 불러오기 ──
if not st.session_state.auto_loaded:
    load_all()
    st.session_state.auto_loaded = True

# ── 헤더 ──
col_h1, col_h2 = st.columns([5, 3])
with col_h1:
    st.markdown("## 📦 DAS 작업 화면")
with col_h2:
    st.markdown('<div class="btn-load">', unsafe_allow_html=True)
    if st.button("🔄 불러오기", key="btn_load_all", use_container_width=True):
        if load_all():
            st.rerun()
        else:
            st.error("데이터 없음. PC에서 전송하세요.")
    st.markdown('</div>', unsafe_allow_html=True)

st.divider()
tab_seed, tab_prep = st.tabs(["📦 씨딩 작업", "📋 배치 품목 리스트"])

# ════ 탭1: 씨딩 작업 ════
with tab_seed:
    if not st.session_state.loaded or not st.session_state.all_batches:
        st.info("👆 '불러오기'를 눌러주세요.")
        st.stop()

    batch_list = st.session_state.batch_list
    cur_batch  = st.session_state.cur_batch

    st.markdown("**배치 선택**")
    cols_b = st.columns(len(batch_list))
    for i, b in enumerate(batch_list):
        with cols_b[i]:
            is_cur = str(b) == str(cur_batch)
            st.markdown('<div class="btn-batch">', unsafe_allow_html=True)
            if st.button(f"{'✅ ' if is_cur else ''}배치 {b}", key=f"sel_b_{b}", use_container_width=True):
                st.session_state.cur_batch = str(b)
                st.session_state.seed_idx  = 0
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)

    st.divider()
    skus  = st.session_state.all_batches.get(str(cur_batch), [])
    idx   = st.session_state.seed_idx
    total = len(skus)

    if total == 0:
        st.warning("이 배치에 씨딩할 SKU가 없습니다.")
        st.stop()

    if idx >= total:
        st.markdown(f"""
<div class="done-box">
  <div style="font-size:2.5rem">🎉</div>
  <div style="font-size:1.6rem;font-weight:800;color:#388E3C;margin:10px 0">배치 {cur_batch} 씨딩 완료!</div>
  <div style="font-size:1rem;color:#6B6056">총 {total}개 SKU 처리</div>
</div>""", unsafe_allow_html=True)
        st.markdown("")
        st.markdown('<div class="btn-restart">', unsafe_allow_html=True)
        if st.button("↩️ 처음부터 다시", key="btn_restart"):
            st.session_state.seed_idx = 0; st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
        st.stop()

    sku   = skus[idx]
    slots = parse_slots(sku.get("분배", ""))
    pct   = int(idx / total * 100)

    st.markdown(f'<div class="batch-badge">배치 {cur_batch}</div>', unsafe_allow_html=True)
    st.markdown(f"""
<div style="display:flex;justify-content:space-between;color:#6B6056;font-size:0.9rem;">
  <span>진행 {idx} / {total}</span><span>{pct}%</span>
</div>
<div class="progress-bar-wrap"><div class="progress-bar-fill" style="width:{pct}%"></div></div>
""", unsafe_allow_html=True)

    slot_chips = "".join(
        f'<div class="slot-chip">{칸} &nbsp;×{수량}</div>' if 수량 > 1
        else f'<div class="slot-chip">{칸}</div>'
        for 칸, 수량 in slots
    )
    st.markdown(f"""
<div class="sku-card">
  <div class="sku-name">🫘 {sku.get("SKU","")}</div>
  <div class="sku-total">총 {sku.get("총수량","")}개 뿌리기</div>
  <div class="slot-grid">{slot_chips}</div>
</div>""", unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown('<div class="btn-prev">', unsafe_allow_html=True)
        if st.button("← 이전 SKU", key="btn_prev", disabled=(idx == 0)):
            st.session_state.seed_idx -= 1; st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    with col2:
        label = "완료 → 다음 SKU ✅" if idx < total - 1 else "완료 🎉 마지막!"
        st.markdown('<div class="btn-next">', unsafe_allow_html=True)
        if st.button(label, key="btn_next"):
            st.session_state.seed_idx += 1; st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

# ════ 탭2: 배치 품목 리스트 ════
with tab_prep:
    if not st.session_state.all_prep:
        st.info("👆 '불러오기'를 눌러주세요.")
        st.stop()

    prep_batch_list = st.session_state.batch_list
    prep_cur        = st.session_state.prep_cur_batch

    # 배치 선택
    if len(prep_batch_list) > 1:
        st.markdown("**배치 선택**")
        cols_pb = st.columns(len(prep_batch_list))
        for i, b in enumerate(prep_batch_list):
            with cols_pb[i]:
                is_cur = str(b) == str(prep_cur)
                st.markdown('<div class="btn-batch">', unsafe_allow_html=True)
                if st.button(f"{'✅ ' if is_cur else ''}배치 {b}", key=f"prep_sel_{b}", use_container_width=True):
                    st.session_state.prep_cur_batch = str(b)
                    st.session_state.prep_page      = 0
                    st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)
        st.markdown("")

    batch_data = st.session_state.all_prep.get(str(prep_cur), {})
    items      = batch_data.get('items', []) if isinstance(batch_data, dict) else []
    order_cnt  = batch_data.get('order_cnt', '') if isinstance(batch_data, dict) else ''
    serial_rng = batch_data.get('serial_range', '') if isinstance(batch_data, dict) else ''
    total_qty  = batch_data.get('total_qty', '') if isinstance(batch_data, dict) else ''
    total_i    = len(items)

    if total_i == 0:
        st.warning("이 배치에 품목 데이터가 없습니다.")
        st.stop()

    # 타이틀 (배치 시트 1행과 동일)
    st.markdown(
        f'<div class="batch-title">배치 {prep_cur} &nbsp;|&nbsp; 복합주문 {order_cnt}건 &nbsp;|&nbsp; 송장순번 {serial_rng}</div>',
        unsafe_allow_html=True
    )

    # 테이블 헤더
    st.markdown(
        '<div style="display:grid;grid-template-columns:5fr 1.5fr 2.5fr 1fr;'
        'background:#EDE5DC;padding:8px 12px;font-size:0.85rem;font-weight:700;'
        'color:#2C2C2C;border:1px solid #D6CEC8;margin-top:0;">'
        '<div>품목명</div><div style="text-align:center;">중량</div>'
        '<div>옵션</div><div style="text-align:center;">수량</div>'
        '</div>',
        unsafe_allow_html=True
    )

    # 페이지네이션
    page    = st.session_state.prep_page
    total_p = (total_i + ITEMS_PER_PAGE - 1) // ITEMS_PER_PAGE
    start   = page * ITEMS_PER_PAGE
    end     = min(start + ITEMS_PER_PAGE, total_i)

    for i, item in enumerate(items[start:end]):
        name   = item.get('품목명', '')
        weight = item.get('중량', '')
        opt    = item.get('옵션', '')
        qty    = item.get('수량', '')
        is_bag = '쇼핑백' in name
        bg     = '#F5E8D8' if is_bag else ('#FBF8F5' if i % 2 == 0 else '#F5F0EB')
        nc     = '#A0622A' if is_bag else '#2C2C2C'
        fw     = '700' if is_bag else '400'
        st.markdown(
            f'<div style="display:grid;grid-template-columns:5fr 1.5fr 2.5fr 1fr;'
            f'background:{bg};padding:9px 12px;font-size:0.95rem;'
            f'border:1px solid #D6CEC8;border-top:none;align-items:center;">'
            f'<div style="color:{nc};font-weight:{fw};word-break:keep-all;line-height:1.35;">{name}</div>'
            f'<div style="color:#6B6056;text-align:center;">{weight}</div>'
            f'<div style="color:#6B6056;font-size:0.85rem;">{opt}</div>'
            f'<div style="color:#8B3A2A;font-weight:700;text-align:center;font-size:1.1rem;">{qty}</div>'
            f'</div>',
            unsafe_allow_html=True
        )

    # 합계 행
    st.markdown(
        f'<div style="display:grid;grid-template-columns:5fr 1.5fr 2.5fr 1fr;'
        f'background:#EDE5DC;padding:9px 12px;font-size:0.95rem;font-weight:700;'
        f'border:1px solid #D6CEC8;border-top:none;">'
        f'<div style="color:#8B3A2A;">합계</div>'
        f'<div></div><div></div>'
        f'<div style="color:#8B3A2A;text-align:center;">{total_qty}</div>'
        f'</div>',
        unsafe_allow_html=True
    ) if end == total_i else None  # 마지막 페이지에만 합계 표시

    # 페이지 정보
    st.markdown(
        f'<div style="text-align:center;color:#6B6056;font-size:0.85rem;padding:8px 0;">'
        f'{start+1}~{end} / {total_i}개 &nbsp;|&nbsp; {page+1} / {total_p} 페이지</div>',
        unsafe_allow_html=True
    )

    if total_p > 1:
        col1, col2, col3 = st.columns([3, 4, 3])
        with col1:
            st.markdown('<div class="btn-sm">', unsafe_allow_html=True)
            if st.button("← 이전", key="btn_prep_prev", disabled=(page == 0), use_container_width=True):
                st.session_state.prep_page -= 1; st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)
        with col2:
            msg = "✅ 마지막" if page == total_p - 1 else f"{page+1} / {total_p}"
            clr = "#388E3C" if page == total_p - 1 else "#6B6056"
            st.markdown(
                f'<div style="text-align:center;color:{clr};font-weight:700;padding:14px 0;">{msg}</div>',
                unsafe_allow_html=True
            )
        with col3:
            st.markdown('<div class="btn-sm">', unsafe_allow_html=True)
            if st.button("다음 →", key="btn_prep_next", disabled=(page == total_p - 1), use_container_width=True):
                st.session_state.prep_page += 1; st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)
