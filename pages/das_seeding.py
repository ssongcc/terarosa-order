"""
테라로사 DAS 씨딩 화면 (태블릿/모바일용)
Galaxy Tab A8 + 8인치 태블릿 + 휴대폰 최적화
"""
import streamlit as st
from github_storage import gh_load, gh_save

st.set_page_config(
    page_title="씨딩 작업",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
section[data-testid="stSidebar"] { display: none; }
.block-container { padding: 3.5rem 0.8rem 1rem 0.8rem !important; max-width: 100% !important; }
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
.done-box {
    background: #E8F5E9; border: 3px solid #388E3C;
    border-radius: 16px; padding: 32px; text-align: center;
}
.stButton > button {
    font-size: 1.2rem !important; font-weight: 700 !important;
    padding: 16px 0 !important; border-radius: 12px !important;
    border: none !important; width: 100%;
}
.btn-prev > button  { background: #EDE5DC !important; color: #2C2C2C !important; }
.btn-next > button  { background: #8B3A2A !important; color: white !important; }
.btn-batch > button { background: #EDE5DC !important; color: #8B3A2A !important;
                      font-size: 0.95rem !important; padding: 9px 0 !important; }
.btn-restart > button { background: #6B6056 !important; color: white !important; }
.btn-load > button  { background: #C4644A !important; color: white !important;
                      font-size: 0.9rem !important; padding: 9px 0 !important; }
.btn-sm > button    { font-size: 1rem !important; padding: 12px 0 !important; }
.btn-check > button { background: #EDE5DC !important; color: #6B6056 !important;
                      font-size: 1rem !important; padding: 12px 0 !important; }
.btn-checked > button { background: #388E3C !important; color: white !important;
                        font-size: 1rem !important; padding: 12px 0 !important; }
.btn-filter > button { background: #EDE5DC !important; color: #8B3A2A !important;
                       font-size: 0.95rem !important; padding: 10px 0 !important; }
.btn-filter-active > button { background: #8B3A2A !important; color: white !important;
                               font-size: 0.95rem !important; padding: 10px 0 !important; }
@media (max-width: 600px) {
    .sku-name  { font-size: 1.2rem !important; }
    .slot-chip { font-size: 1rem !important; padding: 8px 12px !important; }
    .sku-total { font-size: 0.95rem !important; }
    .stButton > button { font-size: 1rem !important; padding: 13px 0 !important; }
}
</style>
""", unsafe_allow_html=True)

# ── 세션 초기화 ──
defaults = {
    "all_batches": {}, "batch_list": [], "cur_batch": None, "seed_idx": 0, "loaded": False,
    "all_prep": {}, "prep_batch_list": [], "prep_cur_batch": None,
    "prep_items": [], "prep_page": 0, "prep_loaded": False,
    "prep_checked": {},
    "prep_show_unprepared": False,
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

ITEMS_PER_PAGE = 5

def get_checked(batch: str) -> set:
    if batch not in st.session_state.prep_checked:
        st.session_state.prep_checked[batch] = set()
    return st.session_state.prep_checked[batch]

def load_all():
    data = gh_load("das_data.json", None)
    if data and isinstance(data, dict) and "batches" in data and "all_prep" in data:
        # 씨딩
        st.session_state.all_batches = {str(k): v for k, v in data["batches"].items()}
        st.session_state.batch_list  = data.get("batch_list", [])
        st.session_state.cur_batch   = str(data["batch_list"][0]) if data.get("batch_list") else None
        st.session_state.seed_idx    = 0
        st.session_state.loaded      = True
        # 품목 준비
        st.session_state.all_prep        = {str(k): v for k, v in data["all_prep"].items()}
        st.session_state.prep_batch_list = data.get("batch_list", [])
        first = str(data["batch_list"][0]) if data.get("batch_list") else None
        st.session_state.prep_cur_batch  = first
        st.session_state.prep_items      = st.session_state.all_prep.get(first, []) if first else []
        st.session_state.prep_page       = 0
        st.session_state.prep_checked    = {}
        st.session_state.prep_show_unprepared = False
        st.session_state.prep_loaded     = True
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
tab_seed, tab_prep = st.tabs(["📦 씨딩 작업", "📋 품목 준비"])

# ════ 탭1: 씨딩 작업 ════
with tab_seed:
    if not st.session_state.loaded or not st.session_state.all_batches:
        st.info("👆 '씨딩 불러오기'를 눌러주세요.")
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

# ════ 탭2: 품목 준비 ════
with tab_prep:
    if not st.session_state.prep_loaded or not st.session_state.all_prep:
        st.info("👆 '품목 불러오기'를 눌러주세요.")
        st.stop()

    prep_batch_list = st.session_state.prep_batch_list
    prep_cur        = st.session_state.prep_cur_batch

    if len(prep_batch_list) > 1:
        st.markdown("**배치 선택**")
        cols_pb = st.columns(len(prep_batch_list))
        for i, b in enumerate(prep_batch_list):
            with cols_pb[i]:
                is_cur = str(b) == str(prep_cur)
                st.markdown('<div class="btn-batch">', unsafe_allow_html=True)
                if st.button(f"{'✅ ' if is_cur else ''}배치 {b}", key=f"prep_sel_{b}", use_container_width=True):
                    st.session_state.prep_cur_batch = str(b)
                    st.session_state.prep_items     = st.session_state.all_prep.get(str(b), [])
                    st.session_state.prep_page      = 0
                    st.session_state.prep_show_unprepared = False
                    st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)
        st.markdown("")

    items   = st.session_state.prep_items
    total_i = len(items)
    if total_i == 0:
        st.warning("이 배치에 품목 데이터가 없습니다.")
        st.stop()

    checked    = get_checked(str(prep_cur))
    done_count = len(checked)

    st.markdown(f'<div class="batch-badge">배치 {prep_cur} — 품목 준비</div>', unsafe_allow_html=True)
    pct_done  = int(done_count / total_i * 100) if total_i else 0
    s_color   = "#388E3C" if done_count == total_i else "#8B3A2A"
    s_label   = f"✅ 전체 완료! ({done_count}/{total_i})" if done_count == total_i else f"준비 완료 {done_count} / {total_i}"
    bar_color = "#388E3C" if done_count == total_i else "#8B3A2A"
    st.markdown(f"""
<div style="display:flex;justify-content:space-between;color:{s_color};font-size:0.95rem;font-weight:700;margin-top:6px;">
  <span>{s_label}</span><span>{pct_done}%</span>
</div>
<div class="progress-bar-wrap">
  <div style="background:{bar_color};height:100%;border-radius:8px;width:{pct_done}%;"></div>
</div>
""", unsafe_allow_html=True)

    show_unprepared  = st.session_state.prep_show_unprepared
    unprepared_count = total_i - done_count

    col_f1, col_f2 = st.columns(2)
    with col_f1:
        cls = "btn-filter-active" if not show_unprepared else "btn-filter"
        st.markdown(f'<div class="{cls}">', unsafe_allow_html=True)
        if st.button(f"전체 보기 ({total_i})", key="btn_view_all", use_container_width=True):
            st.session_state.prep_show_unprepared = False
            st.session_state.prep_page = 0; st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    with col_f2:
        cls = "btn-filter-active" if show_unprepared else "btn-filter"
        st.markdown(f'<div class="{cls}">', unsafe_allow_html=True)
        if st.button(f"⚠️ 미준비 ({unprepared_count})", key="btn_view_unprepared", use_container_width=True):
            st.session_state.prep_show_unprepared = True
            st.session_state.prep_page = 0; st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("")

    if show_unprepared:
        display_items = [(i, items[i]) for i in range(total_i) if i not in checked]
        if not display_items:
            st.markdown("""
<div class="done-box">
  <div style="font-size:2rem">🎉</div>
  <div style="font-size:1.4rem;font-weight:800;color:#388E3C;margin:8px 0">모든 품목 준비 완료!</div>
</div>""", unsafe_allow_html=True)
            st.stop()
        page = 0; total_p = 1
        page_items_indexed = display_items
    else:
        page    = st.session_state.prep_page
        total_p = (total_i + ITEMS_PER_PAGE - 1) // ITEMS_PER_PAGE
        start   = page * ITEMS_PER_PAGE
        end     = min(start + ITEMS_PER_PAGE, total_i)
        page_items_indexed = [(start + j, items[start + j]) for j in range(end - start)]
        page_done = sum(1 for gi, _ in page_items_indexed if gi in checked)
        st.markdown(f"""
<div style="display:flex;justify-content:space-between;color:#6B6056;font-size:0.85rem;margin-bottom:8px;">
  <span>{page+1} / {total_p} 페이지 ({start+1}~{end} / {total_i}개)</span>
  <span>이 페이지 {page_done}/{len(page_items_indexed)} 완료</span>
</div>""", unsafe_allow_html=True)

    # ── 품목 카드 + 완료 버튼 ──
    for global_idx, item in page_items_indexed:
        name   = item.get("품목명", "")
        weight = item.get("중량", "")
        opt    = item.get("옵션", "")
        qty    = item.get("수량", "")
        sub    = " / ".join(x for x in [weight, opt] if x)
        is_bag     = "쇼핑백" in name
        is_checked = global_idx in checked

        if is_checked:
            bg = "#E8F5E9"; border = "#388E3C"; nc = "#256029"; qbg = "#388E3C"
        elif is_bag:
            bg = "#F5E8D8"; border = "#A0622A"; nc = "#A0622A"; qbg = "#A0622A"
        else:
            bg = "#FAF3F0"; border = "#D6CEC8"; nc = "#2C2C2C"; qbg = "#8B3A2A"

        sub_part = f'<div style="color:#6B6056;font-size:0.85rem;margin-top:3px;">{sub}</div>' if sub else ''
        strike   = "text-decoration:line-through;opacity:0.55;" if is_checked else ""
        cm_html  = '<div style="color:#388E3C;font-size:1.3rem;margin-left:8px;flex-shrink:0;">✅</div>' if is_checked else ''

        st.markdown(
            f'<div style="background:{bg};border:2px solid {border};border-radius:12px;'
            f'padding:14px 16px;margin-bottom:4px;display:flex;align-items:center;justify-content:space-between;">'
            f'<div style="flex:1;min-width:0;">'
            f'<div style="color:{nc};font-size:1.05rem;font-weight:700;word-break:keep-all;line-height:1.4;{strike}">{name}</div>'
            f'{sub_part}'
            f'</div>'
            f'{cm_html}'
            f'<div style="background:{qbg};color:white;font-size:1.4rem;font-weight:800;'
            f'border-radius:10px;text-align:center;padding:8px 14px;margin-left:10px;'
            f'white-space:nowrap;min-width:48px;">{qty}</div>'
            f'</div>',
            unsafe_allow_html=True
        )
        btn_css   = "btn-checked" if is_checked else "btn-check"
        btn_label = "✅ 완료" if is_checked else "◻ 완료"
        st.markdown(f'<div class="{btn_css}" style="margin-bottom:12px;">', unsafe_allow_html=True)
        if st.button(btn_label, key=f"chk_{prep_cur}_{global_idx}", use_container_width=True):
            if is_checked:
                checked.discard(global_idx)
            else:
                checked.add(global_idx)
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    if not show_unprepared and total_p > 1:
        st.markdown("")
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
