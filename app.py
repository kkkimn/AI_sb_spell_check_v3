import os
import json
import base64
import streamlit as st
import pandas as pd
from pptx import Presentation
import core
import io
from dotenv import load_dotenv

load_dotenv() # .env 파일이 존재하면 로컬 환경변수로 불러옴

# 기본 키 (금고 st.secrets 또는 .env 환경변수에서 우선 가져오기)
API_KEY_DEFAULT = ""
try:
    API_KEY_DEFAULT = st.secrets.get("OPENAI_API_KEY", "")
except Exception:
    pass

if not API_KEY_DEFAULT:
    API_KEY_DEFAULT = os.environ.get("OPENAI_API_KEY", "")

st.set_page_config(page_title="AI 품질관리 시스템(원고, 스토리보드 검토)", page_icon="✨", layout="wide")

# 상단 여백 최소화
st.markdown(
    """
    <style>
    .block-container,
    .stMainBlockContainer,
    [data-testid="stAppViewBlockContainer"] {
        padding-top: 2rem !important;
        padding-bottom: 2rem !important;
    }
    [data-testid="stHeader"] {
        height: 2.5rem !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

st.title("✨AI 품질관리 시스템(원고, 스토리보드 검토)")
st.markdown("압도적 성능의 **OpenAI (GPT-5.4)** AI를 사용하여 PPT 문맥을 파악하고 맞춤법을 전수 검사합니다.")

# 로고 (사이드바 열림/닫힘 모두 표시)
_logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ARASoft로고.png")
if os.path.exists(_logo_path):
    st.markdown(
        """
        <style>
        .st-emotion-cache-4xtz07 {
            height: 3rem !important;
            margin-top: 2.25rem !important;
        }
        .st-emotion-cache-1h1td79 hr {
            margin: 0rem 0px !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    try:
        st.logo(_logo_path, size="large")
    except TypeError:
        # size 파라미터 미지원 버전
        st.logo(_logo_path)
    except AttributeError:
        # st.logo 자체 미지원 구버전 fallback
        _logo_b64 = base64.b64encode(open(_logo_path, "rb").read()).decode()
        st.markdown(
            f"""
            <style>
            [data-testid="stToolbar"]::before {{
                content: "";
                display: inline-block;
                background-image: url("data:image/png;base64,{_logo_b64}");
                background-size: contain;
                background-repeat: no-repeat;
                background-position: center;
                width: 120px;
                height: 34px;
                vertical-align: middle;
                margin-right: 8px;
            }}
            </style>
            """,
            unsafe_allow_html=True,
        )

# 사이드바

with st.sidebar:
    st.divider()
    st.subheader("⚙️ AI 모델 선택")


    model_choice = st.radio(
        "정확도와 속도/비용 사이에서 선택하세요.",
        options=["꼼꼼 모드 (gpt-5.4)", "빠른 모드 (gpt-5.4-mini)"],
        index=0,
        help="꼼꼼 모드는 한국어 맞춤법·띄어쓰기·외래어 표기를 훨씬 정확하게 잡아냅니다. "
             "빠른 모드는 5~10배 저렴하지만 정확도가 떨어집니다."
    )
    selected_model = "gpt-5.4" if "gpt-5.4)" in model_choice else "gpt-5.4-mini"

    st.divider()
    st.subheader("🧠 AI 사전 학습 (지식 베이스)")
    
    kb_file_path = "knowledge_base.json"
    knowledge_base = {}
    if os.path.exists(kb_file_path):
        with open(kb_file_path, "r", encoding="utf-8") as f:
            try:
                knowledge_base = json.load(f)
            except Exception:
                pass
                
    new_keyword = st.text_input("학습할 주제/키워드 명 (예: 소형무인기논문)", placeholder="키워드 입력")
    kb_file = st.file_uploader("학습할 문서 업로드 (선택, PPTX/PDF/TXT/DOCX)", type=["pptx", "pdf", "txt", "docx"])
    
    if st.button("🚀 지식 학습 시작"):
        target_keyword = new_keyword.strip()
        if kb_file and not target_keyword:
            target_keyword = os.path.splitext(kb_file.name)[0]
            
        if not target_keyword:
            st.error("주제명(키워드)을 입력하거나 문서를 업로드해주세요.")
        elif not API_KEY_DEFAULT or not API_KEY_DEFAULT.startswith("sk-"):
            st.error("OpenAI API 키가 설정되어 있지 않습니다.")
        else:
            with st.spinner(f"'{target_keyword}'에 대한 전문 지식 생성 중..."):
                kb_data = None
                if kb_file:
                    ext = os.path.splitext(kb_file.name)[1].lower()
                    file_text = ""
                    try:
                        if ext == ".pdf":
                            import fitz
                            doc = fitz.open(stream=kb_file.read(), filetype="pdf")
                            file_text = core.extract_full_text_pdf(doc)
                        elif ext == ".pptx":
                            doc = Presentation(kb_file)
                            file_text = core.extract_full_text_pptx(doc)
                        elif ext == ".txt":
                            file_text = kb_file.read().decode("utf-8", errors="ignore")
                        elif ext == ".docx":
                            import docx
                            doc = docx.Document(kb_file)
                            file_text = "\n".join([p.text for p in doc.paragraphs])
                    except ImportError:
                        if ext == ".docx":
                            st.error("Word(.docx) 처리를 위해 python-docx 패키지가 필요합니다.")
                        else:
                            st.error("모듈 불러오기 실패.")
                    except Exception as e:
                        st.error(f"파일 읽기 오류: {e}")
                        
                    if file_text.strip():
                        kb_data = core.generate_knowledge_from_text(file_text, API_KEY_DEFAULT, model="gpt-4o")
                    else:
                        st.error("문서에서 텍스트를 추출하지 못했습니다.")
                else:
                    kb_data = core.generate_knowledge(target_keyword, API_KEY_DEFAULT, model="gpt-4o")
                    
                if kb_data:
                    knowledge_base[target_keyword] = kb_data
                    with open(kb_file_path, "w", encoding="utf-8") as f:
                        json.dump(knowledge_base, f, ensure_ascii=False, indent=2)
                    st.success(f"'{target_keyword}' 학습 완료 및 저장됨!")
                else:
                    if kb_file and file_text.strip():
                        st.error("지식 생성에 실패했습니다.")
                    elif not kb_file:
                        st.error("지식 생성에 실패했습니다.")
    
    if knowledge_base:
        with st.expander("📚 현재 학습된 지식 목록 보기", expanded=False):
            for kw in list(knowledge_base.keys()):
                if st.session_state.get(f"edit_mode_{kw}", False):
                    new_name = st.text_input("새 이름", value=kw, key=f"new_name_{kw}", label_visibility="collapsed")
                    col_s1, col_s2, col_s3 = st.columns([7.5, 1.2, 1.3], vertical_alignment="center")
                    with col_s2:
                        if st.button("💾", key=f"save_{kw}", help="저장", type="tertiary"):
                            if new_name and new_name != kw:
                                knowledge_base[new_name] = knowledge_base.pop(kw)
                                with open(kb_file_path, "w", encoding="utf-8") as f:
                                    json.dump(knowledge_base, f, ensure_ascii=False, indent=2)
                            st.session_state[f"edit_mode_{kw}"] = False
                            st.rerun()
                    with col_s3:
                        if st.button("❌", key=f"cancel_{kw}", help="취소", type="tertiary"):
                            st.session_state[f"edit_mode_{kw}"] = False
                            st.rerun()
                else:
                    col1, col2, col3 = st.columns([7.5, 1.2, 1.3], vertical_alignment="center")
                    with col1:
                        st.caption(f"- {kw} ({len(knowledge_base[kw].get('terms', []))}개 용어)")
                    with col2:
                        if st.button("✏️", key=f"edit_{kw}", help=f"'{kw}' 이름 수정", type="tertiary"):
                            st.session_state[f"edit_mode_{kw}"] = True
                            st.rerun()
                    with col3:
                        if st.button("🗑️", key=f"del_{kw}", help=f"'{kw}' 지식 삭제", type="tertiary"):
                            del knowledge_base[kw]
                            with open(kb_file_path, "w", encoding="utf-8") as f:
                                json.dump(knowledge_base, f, ensure_ascii=False, indent=2)
                            st.rerun()

    st.divider()
    st.subheader("📖 사용자 맞춤법 사전")
    
    sp_dict_file_path = "custom_spelling_dicts.json"
    spelling_dicts = {}
    
    def _save_all_spelling_dicts(dicts_to_save):
        with open(sp_dict_file_path, "w", encoding="utf-8") as f:
            json.dump(dicts_to_save, f, ensure_ascii=False, indent=2)
        # 하위 호환성을 위해 모든 사전의 단어를 맞춤법사전.txt에 통합 저장
        all_words = []
        for words in dicts_to_save.values():
            all_words.extend(words)
        unique_words = sorted(list(set(all_words)))
        with open("맞춤법사전.txt", "w", encoding="utf-8") as f:
            f.write("\n".join(unique_words))

    # 데이터 로드 및 마이그레이션
    if os.path.exists(sp_dict_file_path):
        with open(sp_dict_file_path, "r", encoding="utf-8") as f:
            try:
                spelling_dicts = json.load(f)
            except Exception:
                pass
    else:
        # 기존 맞춤법사전.txt가 있으면 가져와서 '기본 사전'으로 마이그레이션
        old_dict_path = "맞춤법사전.txt"
        if os.path.exists(old_dict_path):
            try:
                with open(old_dict_path, "r", encoding="utf-8") as f:
                    old_text = f.read()
                raw_words = old_text.replace('\n', ',').split(',')
                words_list = [w.strip() for w in raw_words if w.strip()]
                if words_list:
                    spelling_dicts["기본 사전"] = words_list
                    _save_all_spelling_dicts(spelling_dicts)
            except Exception:
                pass

    new_dict_name = st.text_input("새 맞춤법 사전 이름", placeholder="예: IT 용어 사전")
    new_dict_words = st.text_area(
        "예외 단어 입력 (쉼표(,)나 줄바꿈으로 구분)",
        height=100,
        placeholder="단어1\n단어2"
    )

    if st.button("➕ 맞춤법 사전 등록"):
        target_name = new_dict_name.strip()
        if not target_name:
            st.error("사전 이름을 입력해주세요.")
        else:
            raw_w = new_dict_words.replace('\n', ',').split(',')
            w_list = [w.strip() for w in raw_w if w.strip()]
            spelling_dicts[target_name] = w_list
            _save_all_spelling_dicts(spelling_dicts)
            st.success(f"'{target_name}' 사전 등록 완료!")
            st.rerun()

    if spelling_dicts:
        with st.expander("📖 등록된 맞춤법 사전 목록 보기", expanded=False):
            for dn in list(spelling_dicts.keys()):
                words_str = "\n".join(spelling_dicts[dn])

                if st.session_state.get(f"edit_sp_mode_{dn}", False):
                    new_dn = st.text_input("새 사전 이름", value=dn, key=f"new_dn_{dn}", label_visibility="collapsed")
                    new_words_val = st.text_area("단어 편집", value=words_str, key=f"edit_words_{dn}", height=120)

                    col_s1, col_s2, col_s3 = st.columns([7.5, 1.2, 1.3], vertical_alignment="center")
                    with col_s2:
                        if st.button("💾", key=f"save_sp_{dn}", help="저장", type="tertiary"):
                            raw_w = new_words_val.replace('\n', ',').split(',')
                            w_list = [w.strip() for w in raw_w if w.strip()]
                            if new_dn and new_dn != dn:
                                spelling_dicts.pop(dn)
                                spelling_dicts[new_dn] = w_list
                            else:
                                spelling_dicts[dn] = w_list
                            _save_all_spelling_dicts(spelling_dicts)
                            st.session_state[f"edit_sp_mode_{dn}"] = False
                            st.rerun()
                    with col_s3:
                        if st.button("❌", key=f"cancel_sp_{dn}", help="취소", type="tertiary"):
                            st.session_state[f"edit_sp_mode_{dn}"] = False
                            st.rerun()
                else:
                    col1, col2, col3 = st.columns([7.5, 1.2, 1.3], vertical_alignment="center")
                    with col1:
                        st.caption(f"- {dn} ({len(spelling_dicts[dn])}개 단어)")
                    with col2:
                        if st.button("✏️", key=f"edit_sp_{dn}", help=f"'{dn}' 이름 및 단어 수정", type="tertiary"):
                            st.session_state[f"edit_sp_mode_{dn}"] = True
                            st.rerun()
                    with col3:
                        if st.button("🗑️", key=f"del_sp_{dn}", help=f"'{dn}' 사전 삭제", type="tertiary"):
                            spelling_dicts.pop(dn)
                            _save_all_spelling_dicts(spelling_dicts)
                            st.rerun()

# ==========================================
# 점수 대시보드 렌더링 함수
# ==========================================
def render_score_dashboard(sr):
    """가중치 기반 문서 품질 점수 대시보드를 HTML로 렌더링한다."""
    score = sr['score']
    grade_label = sr['grade_label']
    grade_color = sr['grade_color']
    total_words = sr['total_words']
    total_errors = sr['total_errors']
    ec = sr['error_counts']
    wsum = sr['weighted_error_sum']

    # 점수에 따른 배경 그라디언트 색상
    bg_start = "#1a1a2e"
    bg_end   = "#16213e"

    html = f"""
    <style>
      @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;700;900&display=swap');
      .score-card {{
        font-family: 'Inter', sans-serif;
        background: linear-gradient(135deg, {bg_start} 0%, {bg_end} 100%);
        border-radius: 18px;
        padding: 28px 32px;
        margin: 8px 0 20px 0;
        border: 1px solid rgba(255,255,255,0.1);
        box-shadow: 0 8px 40px rgba(0,0,0,0.4);
      }}
      .score-title {{
        color: rgba(255,255,255,0.65);
        font-size: 15px;
        font-weight: 700;
        letter-spacing: 1px;
        text-transform: uppercase;
        margin: 0 0 18px 0;
      }}
      .score-main {{
        display: flex;
        align-items: center;
        gap: 28px;
        margin-bottom: 22px;
      }}
      .score-number {{
        font-size: 80px;
        font-weight: 900;
        color: {grade_color};
        line-height: 1;
        text-shadow: 0 0 30px {grade_color}66;
        min-width: 140px;
        text-align: center;
      }}
      .score-unit {{
        color: rgba(255,255,255,0.4);
        font-size: 16px;
        text-align: center;
        margin-top: 4px;
      }}
      .score-right {{ flex: 1; }}
      .grade-label {{
        font-size: 24px;
        font-weight: 700;
        color: {grade_color};
        margin-bottom: 14px;
      }}
      .bar-bg {{
        background: rgba(255,255,255,0.12);
        border-radius: 10px;
        height: 14px;
        overflow: hidden;
      }}
      .bar-fill {{
        background: linear-gradient(90deg, {grade_color}99, {grade_color});
        height: 100%;
        width: {score}%;
        border-radius: 10px;
      }}
      .bar-labels {{
        display: flex;
        justify-content: space-between;
        color: rgba(255,255,255,0.3);
        font-size: 11px;
        margin-top: 5px;
      }}
      .stat-grid {{
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        gap: 12px;
        margin-bottom: 16px;
      }}
      .stat-box {{
        background: rgba(255,255,255,0.05);
        border-radius: 12px;
        padding: 14px 10px;
        text-align: center;
        border: 1px solid rgba(255,255,255,0.07);
      }}
      .stat-label {{
        color: rgba(255,255,255,0.45);
        font-size: 11px;
        margin-bottom: 6px;
        letter-spacing: 0.3px;
      }}
      .stat-value {{
        font-size: 22px;
        font-weight: 800;
      }}
      .stat-weight {{
        color: rgba(255,255,255,0.25);
        font-size: 10px;
        margin-top: 3px;
      }}
      .score-footnote {{
        padding: 10px 14px;
        background: rgba(255,255,255,0.03);
        border-radius: 8px;
        border-left: 3px solid {grade_color};
        color: rgba(255,255,255,0.4);
        font-size: 11.5px;
        line-height: 1.6;
      }}
    </style>

    <div class="score-card">
      <p class="score-title">📊 문서 품질 점수 (가중치 기반)</p>

      <div class="score-main">
        <div>
          <div class="score-number">{score}</div>
          <div class="score-unit">/ 100점</div>
        </div>
        <div class="score-right">
          <div class="grade-label">{grade_label}</div>
          <div class="bar-bg"><div class="bar-fill"></div></div>
          <div class="bar-labels"><span>0</span><span>50</span><span>100</span></div>
        </div>
      </div>

      <div class="stat-grid">
        <div class="stat-box">
          <div class="stat-label">총 어절 수</div>
          <div class="stat-value" style="color:#fff;">{total_words:,}</div>
          <div class="stat-weight">분석 대상</div>
        </div>
        <div class="stat-box">
          <div class="stat-label">맞춤법 / 오타</div>
          <div class="stat-value" style="color:#E74C3C;">{ec['spelling']}건</div>
          <div class="stat-weight">가중치 ×2.0 (심각)</div>
        </div>
        <div class="stat-box">
          <div class="stat-label">외래어 표기</div>
          <div class="stat-value" style="color:#F39C12;">{ec['foreign']}건</div>
          <div class="stat-weight">가중치 ×1.5 (보통)</div>
        </div>
        <div class="stat-box">
          <div class="stat-label">띄어쓰기</div>
          <div class="stat-value" style="color:#3498DB;">{ec['spacing']}건</div>
          <div class="stat-weight">가중치 ×1.0 (경미)</div>
        </div>
      </div>

      <div class="score-footnote">
        💡 <b>점수 산출 공식</b>: (1 − 가중 오류 합계 / 총 어절 수) × 100 &nbsp;|&nbsp;
        가중 오류 합계: <b>{wsum}</b>점 &nbsp;|&nbsp;
        실제 오류 발생 횟수 합계: <b>{total_errors}건</b><br>
        맞춤법·오타는 2배, 외래어 표기는 1.5배, 띄어쓰기는 1배로 감점됩니다.
      </div>
    </div>
    """
    st.markdown(html, unsafe_allow_html=True)


# ==========================================
# 등급별 색상 안내 범례
# ==========================================
def render_grade_legend():
    grades = [
        ("S", "95~100점", "#FFD700", "🏆 최우수"),
        ("A", "85~94점",  "#2ECC71", "✅ 우수"),
        ("B", "70~84점",  "#3498DB", "🔵 양호"),
        ("C", "50~69점",  "#F39C12", "⚠️ 미흡"),
        ("D", "0~49점",   "#E74C3C", "🔴 불량"),
    ]
    cols = st.columns(5)
    for col, (g, rng, color, label) in zip(cols, grades):
        col.markdown(
            f"<div style='text-align:center; background:rgba(255,255,255,0.05);"
            f"border-radius:10px; padding:10px 4px; border:1px solid {color}44;'>"
            f"<div style='font-size:22px; font-weight:900; color:{color};'>{g}</div>"
            f"<div style='font-size:12px; color:rgba(255,255,255,0.6);'>{rng}</div>"
            f"<div style='font-size:11px; color:{color}; margin-top:3px;'>{label}</div>"
            f"</div>",
            unsafe_allow_html=True
        )


# 메인 영역
st.markdown(
    """
    <div style='background-color: rgba(128, 128, 128, 0.08); padding: 15px; border-radius: 10px; border-left: 5px solid #FF00E5; margin-bottom: 20px; color: inherit;'>
        <h4 style='margin-top: 0; color: #FF00E5;'>💡 파일 업로드 가이드</h4>
        <ul style='margin-bottom: 0; padding-left: 20px; font-size: 15px; line-height: 1.6;'>
            <li>한글 파일 : DOCX 파일로 변화 하여 업로드하는 것을 추천(한글 파일 업로드 시 DOCX 파일로 변환되어 추출되나 표 등 깨짐 현상 있음)</li>
            <li>PDF 파일 : PDF로 추출되나 수정은 안되고 수정해야 될 부분이 체크되서 추출됨</li>
            <li>PPT 파일 : 교정된 형태로 PPT 추출됨</li>
        </ul>
    </div>
    """,
    unsafe_allow_html=True
)

st.subheader("📁 1. 파일 업로드")

# 지식 및 사전 선택 (가로 배치)
col_kb, col_sp = st.columns(2)
with col_kb:
    kb_options = ["선택 안함"] + list(knowledge_base.keys()) if 'knowledge_base' in locals() else ["선택 안함"]
    selected_kb_keyword = st.selectbox("검사에 적용할 사전 학습 지식 (선택)", options=kb_options)

with col_sp:
    # 맞춤법 사전 선택 (단일 선택)
    sp_options = ["선택 안함"] + list(spelling_dicts.keys()) if 'spelling_dicts' in locals() else ["선택 안함"]
    selected_sp_dict = st.selectbox("검사에 적용할 사용자 맞춤법 사전 (선택)", options=sp_options)

# 엑셀 이미지 포함 옵션 추가
export_images = st.checkbox("엑셀 다운로드용 슬라이드/페이지 이미지 추출 (LibreOffice/PowerPoint COM 작동, 수십 초 소요)", value=False, help="체크하면 엑셀 파일에 슬라이드 이미지가 삽입되지만, 검사 속도가 느려집니다. 체크 해제 시 이미지 없이 빠르게 다운로드 가능합니다.")


if "uploader_id" not in st.session_state:
    st.session_state.uploader_id = 0

uploaded_file = st.file_uploader(
    "검사할 문서를 올려주세요.", 
    type=["pptx", "hwp", "hwpx", "docx", "pdf"],
    key=f"file_uploader_{st.session_state.uploader_id}"
)

if uploaded_file is not None:
    st.success(f"'{uploaded_file.name}' 업로드 성공!")
    
    # 세션 상태 초기화
    for key in ['corrections', 'script_text', 'full_text', 'score_result']:
        if key not in st.session_state:
            st.session_state[key] = None
        
    file_ext = os.path.splitext(uploaded_file.name)[1].lower()
    
    # 업로드된 파일을 메모리 기반 객체로 로드
    doc_obj = None
    hwp_text_content = ""
    
    # 파일 포인터를 처음으로 돌려줍니다. (BadZipFile 에러 예방)
    uploaded_file.seek(0)
    
    if file_ext == '.pdf':
        import fitz
        file_bytes = uploaded_file.read()
        doc_obj = fitz.open(stream=file_bytes, filetype="pdf")
    elif file_ext == '.pptx':
        doc_obj = Presentation(uploaded_file)
    elif file_ext == '.docx':
        import docx
        doc_obj = docx.Document(uploaded_file)
    elif file_ext == '.hwp':
        file_bytes = uploaded_file.read()
        hwp_text_content = core.extract_text_hwp(file_bytes)
        doc_obj = file_bytes
    elif file_ext == '.hwpx':
        file_bytes = uploaded_file.read()
        hwp_text_content = core.extract_text_hwpx(file_bytes)
        doc_obj = file_bytes
        
    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button("🚀 AI 분석 및 텍스트 스캔 시작", use_container_width=True):
            # 이전 결과 초기화
            st.session_state.corrections = None
            st.session_state.script_text = None
            st.session_state.full_text = None
            st.session_state.score_result = None
            st.session_state.img_cache = {}  # 이미지 캐시 초기화
            
            with st.spinner("문서를 스캔하고 대본을 추출하는 중..."):
                if file_ext == '.pdf':
                    script_text = core.extract_narrations_pdf(doc_obj)
                    full_text   = core.extract_full_text_pdf(doc_obj)
                elif file_ext == '.pptx':
                    script_text = core.extract_narrations(doc_obj)
                    full_text   = core.extract_full_text_pptx(doc_obj)
                elif file_ext == '.docx':
                    script_text = {}
                    full_text   = core.extract_full_text_docx(doc_obj)
                elif file_ext in ('.hwp', '.hwpx'):
                    script_text = {}
                    full_text   = hwp_text_content
                st.session_state.script_text = script_text
                st.session_state.full_text   = full_text
                
            st.success(f"대본 추출 완료! 이제 문서 검사에 진입합니다.")
            
            if not API_KEY_DEFAULT or not API_KEY_DEFAULT.startswith("sk-"):
                st.error("서버에 올바른 OpenAI API 환경변수 비밀키가 설정되어 있지 않습니다!")
                st.stop()
                
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            def update_progress(current, total):
                progress = int((current / total) * 100)
                progress_bar.progress(progress)
                status_text.markdown(f"**진행 상황 (맞춤법 스캔):** {current}/{total} 페이지/슬라이드 스캔 완료... ({selected_model} 사용 중)")
            
            # 선택된 맞춤법 사전들로부터 단어 취합 (다중 선택 - 숨김 처리)
            # custom_dict_list = []
            # if 'selected_sp_dicts' in locals() and selected_sp_dicts:
            #     for dn in selected_sp_dicts:
            #         custom_dict_list.extend(spelling_dicts.get(dn, []))
            
            # 선택된 맞춤법 사전으로부터 단어 취합 (단일 선택)
            custom_dict_list = []
            if 'selected_sp_dict' in locals() and selected_sp_dict != "선택 안함":
                custom_dict_list.extend(spelling_dicts.get(selected_sp_dict, []))
            
            # 선택된 지식 베이스가 있다면 용어 목록을 맞춤법 예외 사전에 병합
            active_kb_data = None
            if selected_kb_keyword != "선택 안함" and 'knowledge_base' in locals():
                active_kb_data = knowledge_base.get(selected_kb_keyword)
                if active_kb_data:
                    kb_terms = active_kb_data.get("terms", [])
                    custom_dict_list.extend(kb_terms)
                    
            with st.spinner(f"OpenAI 맞춤법 스캔 중 (1단계) ({selected_model})..."):
                if file_ext == '.pdf':
                    corrections, locations = core.get_openai_corrections_by_page_pdf(
                        doc_obj, 
                        API_KEY_DEFAULT, 
                        is_paid_tier=True,
                        custom_dict=custom_dict_list,
                        progress_callback=update_progress,
                        model=selected_model
                    )
                elif file_ext == '.pptx':
                    corrections, locations = core.get_openai_corrections_by_slide(
                        doc_obj, 
                        API_KEY_DEFAULT, 
                        is_paid_tier=True,
                        custom_dict=custom_dict_list,
                        progress_callback=update_progress,
                        model=selected_model
                    )
                elif file_ext == '.docx':
                    corrections, locations = core.get_openai_corrections_docx(
                        doc_obj,
                        API_KEY_DEFAULT,
                        is_paid_tier=True,
                        custom_dict=custom_dict_list,
                        progress_callback=update_progress,
                        model=selected_model
                    )
                elif file_ext in ('.hwp', '.hwpx'):
                    corrections, locations = core.get_openai_corrections_hwp_text(
                        full_text,
                        API_KEY_DEFAULT,
                        is_paid_tier=True,
                        custom_dict=custom_dict_list,
                        progress_callback=update_progress,
                        model=selected_model
                    )
                st.session_state.corrections = corrections
                st.session_state.locations = locations

            # 2단계: 내용 검토 (선택된 지식이 있을 때만 수행)
            st.session_state.content_reviews = {}
            if active_kb_data and file_ext == '.pptx':
                with st.spinner(f"OpenAI 내용 검토 스캔 중 (2단계) ({selected_model})..."):
                    slide_contents = []
                    for slide in doc_obj.slides:
                        parts = []
                        for shape in core.iter_shapes(slide.shapes):
                            t = core._safe_shape_text(shape).strip()
                            if t: parts.append(t)
                            if shape.has_table:
                                for row in shape.table.rows:
                                    for cell in row.cells:
                                        ct = cell.text.strip()
                                        if ct: parts.append(ct)
                        slide_contents.append("\n".join(parts))
                        
                    progress_bar_rev = st.progress(0)
                    status_text_rev = st.empty()
                    
                    def update_progress_rev(current, total):
                        progress = int((current / total) * 100)
                        progress_bar_rev.progress(progress)
                        status_text_rev.markdown(f"**진행 상황 (내용 검토):** {current}/{total} 슬라이드 검토 완료... ({selected_model} 사용 중)")
                        
                    st.session_state.content_reviews = core.get_openai_content_reviews_by_slide_batch(
                        slide_contents,
                        active_kb_data,
                        API_KEY_DEFAULT,
                        progress_callback=update_progress_rev,
                        model=selected_model
                    )
                    progress_bar_rev.progress(100)
                    status_text_rev.markdown("**✅ 내용 검토 완료!**")

            # ── 점수 계산 ──────────────────────────────────
            if st.session_state.full_text:
                st.session_state.score_result = core.calculate_score(
                    corrections,
                    st.session_state.full_text
                )
                
            progress_bar.progress(100)
            status_text.markdown("**✅ AI 분석 완료!**")

    # ──────────────────────────────────────────────
    # 점수 대시보드 표시
    # ──────────────────────────────────────────────
    if st.session_state.score_result is not None:
        if st.button("🔄 검사 결과 초기화 (새 파일 올리기)", use_container_width=True):
            st.session_state.uploader_id += 1
            st.session_state.corrections = None
            st.session_state.script_text = None
            st.session_state.full_text = None
            st.session_state.score_result = None
            st.session_state.locations = None
            st.session_state.content_reviews = {}
            st.session_state.img_cache = {}  # 이미지 캐시 초기화
            st.rerun()

        st.subheader("🏅 문서 품질 점수")
        render_score_dashboard(st.session_state.score_result)
        with st.expander("📘 등급 기준표 보기"):
            render_grade_legend()

    if st.session_state.corrections is not None:
        st.subheader("📋 2. 수정 전 / 수정 후 검토")
        
        c_dict = st.session_state.corrections
        loc_dict = st.session_state.get('locations', {})
        if len(c_dict) == 0:
            st.info("AI가 변경할 곳을 찾지 못했습니다. 문장이 이미 완벽하거나 수정할 내용이 없습니다.")
        else:
            # 엑셀용 이미지 추출 (미리 캐싱)
            img_cache = {}
            if export_images:
                if "img_cache" not in st.session_state or st.session_state.img_cache is None:
                    st.session_state.img_cache = {}
                
                unique_locs = set()
                for old in c_dict.keys():
                    locs = loc_dict.get(old, [])
                    if locs:
                        unique_locs.add(locs[0])
                
                # 캐시되지 않은 위치의 이미지만 추출하여 캐싱
                missing_locs = [loc for loc in unique_locs if loc not in st.session_state.img_cache]
                
                if missing_locs:
                    with st.spinner("엑셀 다운로드를 위한 원본 이미지 준비 중 (수 초 소요될 수 있습니다)..."):
                        if file_ext == '.pdf':
                            for loc in missing_locs:
                                st.session_state.img_cache[loc] = core.get_pdf_page_image_bytes(doc_obj, loc)
                        elif file_ext == '.pptx':
                            uploaded_file.seek(0)
                            pptx_bytes = uploaded_file.read()
                            new_imgs = core.get_pptx_slide_images(pptx_bytes, missing_locs)
                            st.session_state.img_cache.update(new_imgs)
                
                img_cache = st.session_state.img_cache

            # 오류 유형 컬럼 추가
            rows = []
            image_mappings = []
            seen_locs = set()
            
            for old, new in c_dict.items():
                err_type = core.classify_error(old, new)
                label_map = {'spelling': '맞춤법/오타', 'foreign': '외래어 표기', 'spacing': '띄어쓰기'}
                
                locs = loc_dict.get(old, [])
                loc_str = ", ".join(map(str, locs))
                if loc_str:
                    if file_ext == '.pdf':
                        loc_str += " 페이지"
                    elif file_ext == '.pptx':
                        loc_str += " 슬라이드"
                
                if loc_str and loc_str not in seen_locs:
                    img_bytes = img_cache.get(locs[0]) if locs else None
                    seen_locs.add(loc_str)
                else:
                    img_bytes = None
                
                image_mappings.append(img_bytes)
                
                rows.append({
                    "발생 위치": loc_str,
                    "원본 이미지": "",
                    "수정 전(원본)": old,
                    "수정 후(AI 제안)": new,
                    "오류 유형": label_map.get(err_type, '기타')
                })
            df = pd.DataFrame(rows)
            # 화면에는 이미지 컬럼을 빼고 보여줌
            st.dataframe(df.drop(columns=["원본 이미지"]), use_container_width=True, hide_index=True)
            
            # 내용 검토 피드백 렌더링
            if st.session_state.get('content_reviews'):
                st.subheader("💡 내용 검토 피드백")
                review_rows = []
                for s_num, feedback in st.session_state.content_reviews.items():
                    review_rows.append({"슬라이드 번호": f"{s_num} 슬라이드", "피드백 내용": feedback})
                st.dataframe(pd.DataFrame(review_rows), use_container_width=True, hide_index=True)
            
            # 엑셀 다운로드 버튼 추가
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='교정결과')
                workbook = writer.book
                worksheet = writer.sheets['교정결과']
                
                # 서식 정의
                bg_color_1 = '#FFFFFF'
                bg_color_2 = '#F4F8FC'
                
                fmt_c1 = workbook.add_format({'bg_color': bg_color_1, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center', 'border': 1, 'border_color': '#D9D9D9'})
                fmt_l1 = workbook.add_format({'bg_color': bg_color_1, 'text_wrap': True, 'valign': 'vcenter', 'align': 'left', 'border': 1, 'border_color': '#D9D9D9'})
                
                fmt_c2 = workbook.add_format({'bg_color': bg_color_2, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center', 'border': 1, 'border_color': '#D9D9D9'})
                fmt_l2 = workbook.add_format({'bg_color': bg_color_2, 'text_wrap': True, 'valign': 'vcenter', 'align': 'left', 'border': 1, 'border_color': '#D9D9D9'})
                
                # 열 너비 설정
                worksheet.set_column('A:A', 14)
                worksheet.set_column('B:B', 60)
                worksheet.set_column('C:C', 40)
                worksheet.set_column('D:D', 40)
                worksheet.set_column('E:E', 15)
                
                # 그룹별로 서식 적용 및 병합
                groups = []
                current_loc = None
                start_idx = 0
                for i in range(len(df)):
                    loc = df.iloc[i, 0] # 발생 위치
                    if loc != current_loc:
                        if current_loc is not None:
                            groups.append((start_idx, i - 1, current_loc))
                        current_loc = loc
                        start_idx = i
                if len(df) > 0:
                    groups.append((start_idx, len(df) - 1, current_loc))
                    
                img_scale = 0.17 if file_ext == '.pdf' else 0.32
                
                for group_idx, (s_idx, e_idx, loc_str) in enumerate(groups):
                    group_size = e_idx - s_idx + 1
                    is_even = (group_idx % 2 == 0)
                    fc = fmt_c1 if is_even else fmt_c2
                    fl = fmt_l1 if is_even else fmt_l2
                    
                    # 행 높이: 이미지가 잘리지 않도록 단일 항목일 때는 높이를 충분히 크게(190), 여러 개일 때는 골고루 분배
                    if group_size == 1:
                        row_h = 190
                    else:
                        row_h = max(190 // group_size, 45)
                        
                    for r in range(s_idx, e_idx + 1):
                        worksheet.set_row(r + 1, row_h)
                        
                    # 발생 위치, 원본 이미지 병합 (A, B열)
                    if group_size > 1:
                        worksheet.merge_range(s_idx + 1, 0, e_idx + 1, 0, loc_str, fc)
                        worksheet.merge_range(s_idx + 1, 1, e_idx + 1, 1, "", fc)
                    else:
                        worksheet.write(s_idx + 1, 0, loc_str, fc)
                        worksheet.write(s_idx + 1, 1, "", fc)
                        
                    # 이미지 삽입
                    img_bytes = image_mappings[s_idx]
                    if img_bytes:
                        # 이미지는 병합된 블록의 시작 셀(s_idx + 1)에 삽입
                        worksheet.insert_image(s_idx + 1, 1, f"img_{s_idx}.png", {
                            'image_data': io.BytesIO(img_bytes),
                            'x_scale': img_scale,
                            'y_scale': img_scale,
                            'x_offset': 5,
                            'y_offset': 5,
                            'object_position': 1
                        })
                        
                    # 나머지 열 데이터 쓰기 (C, D, E)
                    for r in range(s_idx, e_idx + 1):
                        worksheet.write(r + 1, 2, df.iloc[r, 2], fl) # 수정 전
                        worksheet.write(r + 1, 3, df.iloc[r, 3], fl) # 수정 후
                        worksheet.write(r + 1, 4, df.iloc[r, 4], fc) # 오류 유형
                
                # 내용 검토 시트 추가
                if st.session_state.get('content_reviews'):
                    review_df = pd.DataFrame([{"슬라이드 번호": f"{s_num} 슬라이드", "피드백 내용": fb} for s_num, fb in st.session_state.content_reviews.items()])
                    review_df.to_excel(writer, index=False, sheet_name='내용검토')
                    review_ws = writer.sheets['내용검토']
                    review_ws.set_column('A:A', 20)
                    review_ws.set_column('B:B', 100)
                    for r in range(len(review_df)):
                        review_ws.set_row(r + 1, 60)
                        review_ws.write(r + 1, 0, review_df.iloc[r, 0], fmt_c1)
                        review_ws.write(r + 1, 1, review_df.iloc[r, 1], fmt_l1)
                        
            excel_data = output.getvalue()
            
            st.download_button(
                label="📊 교정 결과 엑셀 다운로드",
                data=excel_data,
                file_name=f"교정결과_{uploaded_file.name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            if file_ext == '.pdf':
                st.warning("위 변경 사항들은 완성본 다운로드 시 '핑크색(FF00E5) 형광펜 (메모 코멘트)' 형태로 PDF에 표시됩니다.")
            elif file_ext == '.pptx':
                st.warning("위 변경 사항들은 완성본 다운로드 시 '핑크색(FF00E5)' 서식으로 PPT에 일괄 덮어씌워집니다. "
                           "(부분 굵게/색상 등 일부 인라인 서식은 초기화될 수 있습니다.)")
            elif file_ext == '.docx':
                st.warning("위 변경 사항들은 완성본 다운로드 시 '핑크색(FF00E5)' 서식으로 워드(Word) 파일에 덮어씌워집니다.")
            elif file_ext in ('.hwp', '.hwpx'):
                st.warning("위 변경 사항들은 완성본 다운로드 시 교정된 내용이 반영된 워드(.docx) 문서 파일로 자동 변환되어 다운로드됩니다.")
            
        st.subheader("📥 3. 완성본 다운로드")
        
        with st.spinner("수정 및 덧그리기 작업 중입니다..."):
            out_stream = io.BytesIO()
            if file_ext == '.pdf':
                core.apply_corrections_to_pdf(doc_obj, st.session_state.corrections)
                doc_obj.save(out_stream)
                doc_obj.close()
                mime_type = "application/pdf"
                btn_label = "💖 교정 하이라이트 PDF 다운로드"
                download_name = f"완료_{uploaded_file.name}"
            elif file_ext == '.pptx':
                core.apply_corrections_to_ppt(doc_obj, st.session_state.corrections)
                doc_obj.save(out_stream)
                mime_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                btn_label = "💖 핑크색 교정 반영본 PPTX 다운로드"
                download_name = f"완료_{uploaded_file.name}"
            elif file_ext == '.docx':
                core.apply_corrections_to_docx(doc_obj, st.session_state.corrections)
                doc_obj.save(out_stream)
                mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                btn_label = "💖 핑크색 교정 반영본 DOCX 다운로드"
                download_name = f"완료_{uploaded_file.name}"
            elif file_ext in ('.hwp', '.hwpx'):
                try:
                    full_txt = st.session_state.full_text or ""
                    corrections = st.session_state.corrections or {}
                    # 한글 본문 텍스트와 교정 사전을 바탕으로 Word 문서(.docx) 생성
                    docx_doc = core.create_docx_from_hwp_text(full_txt, corrections)
                    docx_doc.save(out_stream)
                    mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    btn_label = "💖 워드(.docx)로 변환된 교정 반영본 다운로드"
                    base_name = os.path.splitext(uploaded_file.name)[0]
                    download_name = f"완료_{base_name}.docx"
                except Exception as e:
                    # 예외 발생 시 최종 백업으로 텍스트 파일 제공
                    corrected_text = core.apply_corrections_to_text(st.session_state.full_text or "", st.session_state.corrections or {})
                    out_stream.write(corrected_text.encode('utf-8'))
                    mime_type = "text/plain"
                    btn_label = "💖 교정 반영본 텍스트 파일(TXT) 다운로드 (대체)"
                    base_name = os.path.splitext(uploaded_file.name)[0]
                    download_name = f"완료_{base_name}.txt"
                
            download_data = out_stream.getvalue()
            
        st.download_button(
            label=btn_label,
            data=download_data,
            file_name=download_name,
            mime=mime_type,
            use_container_width=True
        )

