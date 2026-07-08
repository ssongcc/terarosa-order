"""
테라로사 DAS 씨딩 화면 (태블릿용)
Galaxy Tab A8 (10.5") 최적화
"""
import json
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
/* 태블릿 터치 최적화 */
section[data-testid="stSidebar"] { display: none; }
.block-container { padding: 1rem 1.5rem !important; max-width: 100% !important; }

.sku-card {
    background: #FAF3F0;
    border: 3px solid #8B3A2A;
    border-radius: 16px;
    padding: 24px 32px;
    margin-bottom: 16px;
}
.sku-name {
    color: #8B3A2A;
    font-size: 2rem;
    font-weight: 800;
    line-height: 1.3;
    margin-bottom: 8px;
}
.sku-total {
    color: #6B6056;
    font-size: 1.2rem;
    margin-bottom: 20px;
}
.slot-grid {
    display: flex;
    flex-wrap: wrap;
    gap: 12px;
    margin-top: 8px;
}
.slot-chip {
    background: #8B3A2A;
    color: white;
    font-size: 1.4rem;
    font-weight: 700;
    padding: 12px 20px;
    border-radius: 10px;
    white-space: nowrap;
}
.progress-bar-wrap {
    background: #EDE5DC;
    border-radius: 8px;
    height: 14px;
    margin: 8px 0 4px 0;
    overflow: hidden;
}
.progress-bar-fill {
    background: #8B3A2A;
    height: 100%;
    border-radius: 8px;
    transition: width 0.3s;
}
.batch-badge {
    background: #C4644A;
    color: white;
    font-size: 1rem;
    font-weight: 700;
    padding: 4px 14px;
    border-radius: 20px;
    display: inline-block;
    margin-bottom: 12px;
}
.done-box {
    background: #E8F5E9;
    border: 3px solid #388E3C;
    border-radius: 16px;
    padding: 40px;
    text-align: center;
}
.stButton > button {
    font-size: 1.3rem !important;
    font-weight: 700 !important;
    padding: 18px 0 !important;
    border-radius: 12px !important;
    border: none !important;
    width: 100%;
}
.btn-prev > button { background: #EDE5DC !important; color: #2C2C2C !important; }
.btn-next > button { background: #8B3A2A !important; color: white !important; }
.btn-next > button:hover { background: #C4644A !important; }
.btn-restart > button { background: #6B6056 !important; color: white !important; }
</style>
""", unsafe_allow_html=True)

# ── 세션 초기화 ──
for k, v in [("seed_skus", []), ("seed_batch", 0), ("seed_idx", 0), ("seed_loaded", False)]:
    if k not in st.session_state:
        st.session_state[k] = v

def load_session():
    data = gh_load("das_session.json", None)
    if data and isinstance(data, dict) and "skus" in data:
        st.session_state.seed_skus   = data["skus"]
        st.session_state.seed_batch  = data.get("batch", 0)
        st.session_state.seed_idx    = 0
        st.session_state.seed_loaded = True
        return True
    return False

def parse_slots(분배_str):
    """'3번칸×2, 7번칸' 형태 파싱 → [(칸번호, 수량), ...]"""
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
col_h1, col_h2 = st.columns([5, 2])
with col_h1:
    st.markdown("## 📦 씨딩 작업 화면")
with col_h2:
    if st.button("🔄 최신 데이터 불러오기", key="btn_load"):
        if load_session():
            st.rerun()
        else:
            st.error("저장된 씨딩 데이터가 없습니다. PC에서 분배모드를 먼저 실행해주세요.")

st.divider()

# ── 데이터 없을 때 ──
if not st.session_state.seed_loaded or not st.session_state.seed_skus:
    st.info("👆 PC에서 분배모드 처리 후 '씨딩 시작' 버튼을 누르면 여기에 작업이 나타납니다.")
    st.stop()

skus    = st.session_state.seed_skus
batch   = st.session_state.seed_batch
idx     = st.session_state.seed_idx
total   = len(skus)

# ── 전체 완료 ──
if idx >= total:
    st.markdown(f"""
    <div class="done-box">
        <div style="font-size:3rem">🎉</div>
        <div style="font-size:1.8rem; font-weight:800; color:#388E3C; margin:12px 0">배치 {batch} 씨딩 완료!</div>
        <div style="font-size:1.1rem; color:#6B6056">총 {total}개 SKU 처리</div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("")
    with st.container():
        st.markdown('<div class="btn-restart">', unsafe_allow_html=True)
        if st.button("↩️ 처음부터 다시", key="btn_restart"):
            st.session_state.seed_idx = 0
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    st.stop()

# ── 현재 SKU ──
sku = skus[idx]
slots = parse_slots(sku.get("분배", ""))
pct = int(idx / total * 100)

# 배치 배지
st.markdown(f'<div class="batch-badge">배치 {batch}</div>', unsafe_allow_html=True)

# 진행률
st.markdown(f"""
<div style="display:flex; justify-content:space-between; color:#6B6056; font-size:0.95rem;">
    <span>진행 {idx} / {total}</span><span>{pct}%</span>
</div>
<div class="progress-bar-wrap">
    <div class="progress-bar-fill" style="width:{pct}%"></div>
</div>
""", unsafe_allow_html=True)

# SKU 카드
slot_chips = "".join(
    f'<div class="slot-chip">{칸} &nbsp;×{수량}</div>' if 수량 > 1
    else f'<div class="slot-chip">{칸}</div>'
    for 칸, 수량 in slots
)
st.markdown(f"""
<div class="sku-card">
    <div class="sku-name">🫘 {sku['SKU']}</div>
    <div class="sku-total">총 {sku['총수량']}개 뿌리기</div>
    <div class="slot-grid">{slot_chips}</div>
</div>
""", unsafe_allow_html=True)

# ── 버튼 ──
col1, col2 = st.columns(2)
with col1:
    st.markdown('<div class="btn-prev">', unsafe_allow_html=True)
    if st.button("← 이전 SKU", key="btn_prev", disabled=(idx == 0)):
        st.session_state.seed_idx -= 1
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)
with col2:
    label = "완료 → 다음 SKU ✅" if idx < total - 1 else "완료 🎉 마지막!"
    st.markdown('<div class="btn-next">', unsafe_allow_html=True)
    if st.button(label, key="btn_next"):
        st.session_state.seed_idx += 1
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)
