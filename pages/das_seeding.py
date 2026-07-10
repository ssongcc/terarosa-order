"""
테라로사 DAS 씨딩 화면 (태블릿용)
Galaxy Tab A8 (10.5") 최적화
탭1: 씨딩 작업 / 탭2: 품목 준비
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
.block-container { padding: 3.5rem 1.5rem 1rem 1.5rem !important; max-width: 100% !important; }

.sku-card {
    background: #FAF3F0; border: 3px solid #8B3A2A;
    border-radius: 16px; padding: 24px 32px; margin-bottom: 16px;
}
.sku-name { color: #8B3A2A; font-size: 2rem; font-weight: 800; line-height: 1.3; margin-bottom: 8px; }
.sku-total { color: #6B6056; font-size: 1.2rem; margin-bottom: 20px; }
.slot-grid { display: flex; flex-wrap: wrap; gap: 12px; margin-top: 8px; }
.slot-chip {
    background: #8B3A2A; color: white; font-size: 1.4rem; font-weight: 700;
    padding: 12px 20px; border-radius: 10px; white-space: nowrap;
}
.prep-card {
    background: #FAF3F0; border: 2px solid #D6CEC8;
    border-radius: 12px; padding: 16px 20px; margin-bottom: 10px;
    display: flex; align-items: center; justify-content: space-between;
}
.prep-name { color: #2C2C2C; font-size: 1.3rem; font-weight: 700; }
.prep-sub { color: #6B6056; font-size: 0.95rem; margin-top: 2px; }
.prep-qty {
    background: #8B3A2A; color: white; font-size: 1.6rem; font-weight: 800;
    padding: 8px 20px; border-radius: 10px; min-width: 70px; text-align: center;
}
.prep-bag {
    background: #F5E8D8; border: 2px solid #A0622A;
    border-radius: 12px; padding: 16px 20px; margin-bottom: 10px;
    display: flex; align-items: center; justify-content: space-between;
}
.prep-bag .prep-name { color: #A0622A; }
.prep-bag .prep-qty { background: #A0622A; }
.progress-bar-wrap {
    background: #EDE5DC; border-radius: 8px; height: 14px;
    margin: 8px 0 4px 0; overflow: hidden;
}
.progress-bar-fill { background: #8B3A2A; height: 100%; border-radius: 8px; }
.batch-badge {
    background: #C4644A; color: white; font-size: 1rem; font-weight: 700;
    padding: 4px 14px; border-radius: 20px; display: inline-block; margin-bottom: 12px;
}
.done-box {
    background: #E8F5E9; border: 3px solid #388E3C;
    border-radius: 16px; padding: 40px; text-align: center;
}
.stButton > button {
    font-size: 1.3rem !important; font-weight: 700 !important;
    padding: 18px 0 !important; border-radius: 12px !important;
    border: none !important; width: 100%;
}
.btn-prev > button { background: #EDE5DC !important; color: #2C2C2C !important; }
.btn-next > button { background: #8B3A2A !important; color: white !important; }
.btn-batch > button { background: #EDE5DC !important; color: #8B3A2A !important; font-size: 1rem !important; padding: 10px 0 !important; }
.btn-restart > button { background: #6B6056 !important; color: white !important; }
.btn-load > button { background: #C4644A !important; color: white !important; font-size: 1rem !important; padding: 10px 0 !important; }
.btn-sm > button { font-size: 1rem !important; padding: 10px 0 !important; }
</style>
""", unsafe_allow_html=True)

# ── 세션 초기화 ──
defaults = {
    "all_batches": {}, "batch_list": [], "cur_batch": None,
    "seed_idx": 0, "loaded": False,
    "prep_items": [], "prep_batch": None, "prep_page": 0, "prep_loaded": False,
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

ITEMS_PER_PAGE = 5

def load_seed():
    data = gh_load("das_session.json", None)
    if data and isinstance(data, dict) and "batches" in data:
        st.session_state.all_batches = {str(k): v for k, v in data["batches"].items()}
        st.session_state.batch_list  = data.get("batch_list", [])
        st.session_state.cur_batch   = str(st.session_state.batch_list[0]) if st.session_state.batch_list else None
        st.session_state.seed_idx    = 0
        st.session_state.loaded      = True
        return True
    return False

def load_prep():
    data = gh_load("das_prep.json", None)
    if data and isinstance(data, dict) and "items" in data:
        st.session_state.prep_items = data["items"]
        st.session_state.prep_batch = data.get("batch")
        st.session_state.prep_page  = 0
        st.session_state.prep_loaded = True
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
col_h1, col_h2, col_h3 = st.columns([4, 2, 2])
with col_h1:
    st.markdown("## 📦 DAS 작업 화면")
with col_h2:
    st.markdown('<div class="btn-load">', unsafe_allow_html=True)
    if st.button("🔄 씨딩 데이터 불러오기", key="btn_load_seed"):
        if load_seed():
            st.rerun()
        else:
            st.error("씨딩 데이터 없음. PC에서 전송하세요.")
    st.markdown('</div>', unsafe_allow_html=True)
with col_h3:
    st.markdown('<div class="btn-load">', unsafe_allow_html=True)
    if st.button("🔄 품목 데이터 불러오기", key="btn_load_prep"):
        if load_prep():
            st.rerun()
        else:
            st.error("품목 데이터 없음. PC에서 배치_SKU 전송하세요.")
    st.markdown('</div>', unsafe_allow_html=True)

st.divider()

# ── 탭 ──
tab_seed, tab_prep = st.tabs(["📦 씨딩 작업", "📋 품목 준비"])

# ════ 탭1: 씨딩 작업 ════
with tab_seed:
    if not st.session_state.loaded or not st.session_state.all_batches:
        st.info("👆 '씨딩 데이터 불러오기'를 누르면 작업이 나타납니다.")
        st.stop()

    skus_all   = st.session_state.all_batches
    batch_list = st.session_state.batch_list
    cur_batch  = st.session_state.cur_batch

    # 배치 선택
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

    skus  = skus_all.get(str(cur_batch), [])
    idx   = st.session_state.seed_idx
    total = len(skus)

    if total == 0:
        st.warning("이 배치에 씨딩할 SKU가 없습니다.")
        st.stop()

    if idx >= total:
        st.markdown(f"""
        <div class="done-box">
            <div style="font-size:3rem">🎉</div>
            <div style="font-size:1.8rem;font-weight:800;color:#388E3C;margin:12px 0">배치 {cur_batch} 씨딩 완료!</div>
            <div style="font-size:1.1rem;color:#6B6056">총 {total}개 SKU 처리</div>
        </div>""", unsafe_allow_html=True)
        st.markdown("")
        st.markdown('<div class="btn-restart">', unsafe_allow_html=True)
        if st.button("↩️ 처음부터 다시", key="btn_restart"):
            st.session_state.seed_idx = 0
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
        st.stop()

    sku   = skus[idx]
    slots = parse_slots(sku.get("분배",""))
    pct   = int(idx / total * 100)

    st.markdown(f'<div class="batch-badge">배치 {cur_batch}</div>', unsafe_allow_html=True)
    st.markdown(f"""
    <div style="display:flex;justify-content:space-between;color:#6B6056;font-size:0.95rem;">
        <span>진행 {idx} / {total}</span><span>{pct}%</span>
    </div>
    <div class="progress-bar-wrap">
        <div class="progress-bar-fill" style="width:{pct}%"></div>
    </div>""", unsafe_allow_html=True)

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
        if st.button("← 이전 SKU", key="btn_prev", disabled=(idx==0)):
            st.session_state.seed_idx -= 1; st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    with col2:
        label = "완료 → 다음 SKU ✅" if idx < total-1 else "완료 🎉 마지막!"
        st.markdown('<div class="btn-next">', unsafe_allow_html=True)
        if st.button(label, key="btn_next"):
            st.session_state.seed_idx += 1; st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

# ════ 탭2: 품목 준비 ════
with tab_prep:
    if not st.session_state.prep_loaded or not st.session_state.prep_items:
        st.info("👆 PC에서 '배치N_SKU 전송' 버튼을 누른 후 '품목 데이터 불러오기'를 누르세요.")
        st.stop()

    items    = st.session_state.prep_items
    page     = st.session_state.prep_page
    total_i  = len(items)
    total_p  = (total_i + ITEMS_PER_PAGE - 1) // ITEMS_PER_PAGE
    start    = page * ITEMS_PER_PAGE
    end      = min(start + ITEMS_PER_PAGE, total_i)
    page_items = items[start:end]

    st.markdown(f'<div class="batch-badge">배치 {st.session_state.prep_batch} — 품목 준비</div>', unsafe_allow_html=True)
    pct_p = int((page+1)/total_p*100)
    st.markdown(f"""
    <div style="display:flex;justify-content:space-between;color:#6B6056;font-size:0.95rem;">
        <span>{page+1} / {total_p} 페이지 (품목 {start+1}~{end} / {total_i})</span><span>{pct_p}%</span>
    </div>
    <div class="progress-bar-wrap">
        <div class="progress-bar-fill" style="width:{pct_p}%"></div>
    </div>""", unsafe_allow_html=True)
    st.markdown("")

    for item in page_items:
        name  = item.get("품목명","")
        weight= item.get("중량","")
        opt   = item.get("옵션","")
        qty   = item.get("수량","")
        sub   = " / ".join(x for x in [weight, opt] if x)
        is_bag = "쇼핑백" in name
        card_cls = "prep-bag" if is_bag else "prep-card"
        sub_html = f'<div class="prep-sub">{sub}</div>' if sub else ""
        st.markdown(f"""
        <div class="{card_cls}">
            <div>
                <div class="prep-name">{name}</div>
                {sub_html}
            </div>
            <div class="prep-qty">{qty}</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("")
    col1, col2, col3 = st.columns([2, 3, 2])
    with col1:
        st.markdown('<div class="btn-sm">', unsafe_allow_html=True)
        if st.button("← 이전", key="btn_prep_prev", disabled=(page==0), use_container_width=True):
            st.session_state.prep_page -= 1; st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    with col2:
        if page == total_p - 1:
            st.markdown(f'<div style="text-align:center;color:#388E3C;font-weight:700;font-size:1.1rem;padding:14px 0">✅ 마지막 페이지</div>', unsafe_allow_html=True)
        else:
            st.markdown(f'<div style="text-align:center;color:#6B6056;padding:14px 0">{page+1} / {total_p}</div>', unsafe_allow_html=True)
    with col3:
        st.markdown('<div class="btn-sm">', unsafe_allow_html=True)
        if st.button("다음 →", key="btn_prep_next", disabled=(page==total_p-1), use_container_width=True):
            st.session_state.prep_page += 1; st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
