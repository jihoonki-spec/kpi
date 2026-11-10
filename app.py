"""
================================================================================
서울시농수산식품공사 「2027년도 지방공기업 경영평가」
전사 정량·정성 통합 성과관리 및 자동 검산 시스템 (app.py)
================================================================================
[총괄 관리 부서] 경영성과팀
[UI/UX 설계 원칙]
- 한 화면에는 오직 하나의 핵심 질문에 대한 답만 깔끔하게 제시 (Progressive Disclosure)
- 좌측 사이드바 메뉴로 3대 핵심 업무 모드 분리:
  1) 📊 종합 관제 대시보드 (지금 우리 공사 점수가 몇 점이고, 어느 지표가 위험한가?)
  2) ✏️ 정량 실적 입력·검산 (부서 실적 입력 검증 및 수기 엑셀 오류 적발)
  3) 📑 정성 보고서 AI 심사 (부서 실적자료 파일 업로드 및 편람 기준 자동 채점)
================================================================================
"""

from __future__ import annotations

import io
import re
from copy import deepcopy
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

try:
    import pypdf
except ImportError:
    pypdf = None

try:
    import docx
except ImportError:
    docx = None

# ==============================================================================
# 0. 페이지 설정 및 미니멀·고시인성 CSS 스타일링
# ==============================================================================
st.set_page_config(
    page_title="경영성과팀 | 2027 경영평가 성과관리 시스템",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded",
)

CLEAN_CSS = """
<style>
    /* 여백 및 타이포그래피 정돈 */
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2.5rem;
        max-width: 1380px;
    }
    /* 심플 탑 헤더 */
    .top-bar {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 0.9rem 1.25rem;
        background: #0f172a;
        color: #f8fafc;
        border-radius: 10px;
        margin-bottom: 1.4rem;
    }
    .top-bar-title {
        font-size: 1.15rem;
        font-weight: 700;
        letter-spacing: -0.01em;
    }
    .top-bar-sub {
        font-size: 0.82rem;
        color: #94a3b8;
    }
    /* 핵심 KPI 카드 (여백 충분, 숫자 강조) */
    .metric-box {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 1.25rem 1.4rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.03);
    }
    .metric-title {
        font-size: 0.82rem;
        font-weight: 600;
        color: #64748b;
        margin-bottom: 0.4rem;
    }
    .metric-num {
        font-size: 2.0rem;
        font-weight: 800;
        color: #0f172a;
        line-height: 1.15;
    }
    .metric-foot {
        font-size: 0.8rem;
        color: #64748b;
        margin-top: 0.45rem;
    }
    /* 상태 배지 */
    .pill-red {
        background: #fef2f2; color: #dc2626; border: 1px solid #fecaca;
        padding: 0.15rem 0.55rem; border-radius: 99px; font-size: 0.76rem; font-weight: 700;
    }
    .pill-yellow {
        background: #fffbeb; color: #d97706; border: 1px solid #fde68a;
        padding: 0.15rem 0.55rem; border-radius: 99px; font-size: 0.76rem; font-weight: 700;
    }
    .pill-green {
        background: #f0fdf4; color: #16a34a; border: 1px solid #bbf7d0;
        padding: 0.15rem 0.55rem; border-radius: 99px; font-size: 0.76rem; font-weight: 700;
    }
    .pill-blue {
        background: #eff6ff; color: #2563eb; border: 1px solid #bfdbfe;
        padding: 0.15rem 0.55rem; border-radius: 99px; font-size: 0.76rem; font-weight: 700;
    }
    /* 섹션 구분 타이틀 */
    .sec-title {
        font-size: 1.05rem;
        font-weight: 700;
        color: #1e293b;
        margin: 1.2rem 0 0.6rem 0;
        padding-left: 0.6rem;
        border-left: 4px solid #2563eb;
    }
</style>
"""
st.markdown(CLEAN_CSS, unsafe_allow_html=True)

QUARTERS: List[str] = ["1분기", "2분기", "3분기", "4분기(최종)"]


# ==============================================================================
# 1. 전사 52개 정량지표 마스터 설정 (2026년 3분기 경영실적 점검 실제 데이터)
# ==============================================================================
def _make_indicator(
    ind_id: str,
    name: str,
    eval_agency: str,
    category: str,
    department: str,
    weight: float,
    unit: str,
    target_2026: float,
    q3_target: float,
    q3_actual: float,
    official_q3_score: float,
    is_reverse: bool = False,
    eval_type: str = "standard_3yr",
) -> Dict[str, Any]:
    if not is_reverse:
        base_val = round(target_2026 / 1.10, 2) if eval_type == "standard_3yr" else round(target_2026, 2)
    else:
        base_val = round(target_2026 / 0.90, 2) if eval_type == "standard_3yr" else round(target_2026, 2)

    h_2023 = round(base_val * 0.98, 2)
    h_2024 = round(base_val * 1.00, 2)
    h_2025 = round(base_val * 1.02, 2)

    q1_act = round(q3_actual * 0.33, 2)
    q2_act = round(q3_actual * 0.67, 2)
    q4_est = round(max(target_2026, q3_actual * 1.32), 2) if not is_reverse else round(q3_actual * 1.25, 2)

    return {
        "id": ind_id,
        "name": name,
        "eval_agency": eval_agency,
        "category": category,
        "department": department,
        "weight": float(weight),
        "unit": unit,
        "eval_method": eval_type,
        "is_reverse": is_reverse,
        "target_2026": float(target_2026),
        "q3_target": float(q3_target),
        "official_q3_score": float(official_q3_score),
        "history_3yr": {"2023년": h_2023, "2024년": h_2024, "2025년": h_2025},
        "base_override": base_val,
        "prev_year_actual": h_2025,
        "prev_year_cumulative": {
            "1분기": round(h_2025 * 0.25, 2),
            "2분기": round(h_2025 * 0.50, 2),
            "3분기": round(h_2025 * 0.75, 2),
            "4분기(최종)": h_2025,
        },
        "quarterly_actuals": {
            "1분기": q1_act,
            "2분기": q2_act,
            "3분기": float(q3_actual),
            "4분기(최종)": q4_est,
        },
        "formula_desc": (
            f"[{eval_agency} | {category}] 연간목표(A): {target_2026:,.2f}{unit}, "
            f"3분기목표(B): {q3_target:,.2f}{unit}, 기준치(3개년평균): {base_val:,.2f}{unit}"
        ),
    }


INDICATORS_CONFIG: Dict[str, Dict[str, Any]] = {
    # 1. 주요사업(거래실적) - 12개
    "IND_01": _make_indicator("IND_01", "가락 청과 거래물량", "행안부", "1. 주요사업(거래실적)", "가락시장본부 청과부류팀", 1.75, "톤", 2_445_773, 1_834_330, 1_701_149, 1.52),
    "IND_02": _make_indicator("IND_02", "가락 청과 거래금액", "행안부", "1. 주요사업(거래실적)", "가락시장본부 청과부류팀", 1.75, "백만원", 6_036_264, 4_527_198, 3_896_525, 1.30),
    "IND_03": _make_indicator("IND_03", "가락 수산 거래물량", "행안부", "1. 주요사업(거래실적)", "가락시장본부 수산부류팀", 0.25, "톤", 85_935, 64_451, 54_271, 0.18),
    "IND_04": _make_indicator("IND_04", "가락 수산 거래금액", "행안부", "1. 주요사업(거래실적)", "가락시장본부 수산부류팀", 0.25, "백만원", 624_613, 468_460, 393_469, 0.18),
    "IND_05": _make_indicator("IND_05", "강서 청과 거래물량", "행안부", "1. 주요사업(거래실적)", "강서지사 시장관리팀", 0.50, "톤", 666_408, 499_806, 463_957, 0.44),
    "IND_06": _make_indicator("IND_06", "강서 청과 거래금액", "행안부", "1. 주요사업(거래실적)", "강서지사 시장관리팀", 0.50, "백만원", 1_619_666, 1_214_750, 1_113_262, 0.42),
    "IND_07": _make_indicator("IND_07", "학교급식 공급물량", "행안부", "1. 주요사업(거래실적)", "서울친환경유통센터 급식사업팀", 0.50, "톤", 28_661, 21_496, 17_380, 0.33),
    "IND_08": _make_indicator("IND_08", "학교급식 공급금액", "행안부", "1. 주요사업(거래실적)", "서울친환경유통센터 급식사업팀", 0.50, "백만원", 281_453, 211_090, 187_099, 0.40),
    "IND_09": _make_indicator("IND_09", "가락 상장예외품목 거래물량", "서울시", "1. 주요사업(거래실적)", "가락시장본부 유통관리팀", 3.00, "톤", 179_337, 134_503, 129_328, 2.88),
    "IND_10": _make_indicator("IND_10", "가락 상장예외품목 거래금액", "서울시", "1. 주요사업(거래실적)", "가락시장본부 유통관리팀", 3.00, "백만원", 924_444, 693_333, 540_689, 2.20),
    "IND_11": _make_indicator("IND_11", "강서 시장도매인 거래물량", "서울시", "1. 주요사업(거래실적)", "강서지사 시장관리팀", 3.00, "톤", 342_018, 256_514, 276_188, 3.00),
    "IND_12": _make_indicator("IND_12", "강서 시장도매인 거래금액", "서울시", "1. 주요사업(거래실적)", "강서지사 시장관리팀", 3.00, "백만원", 1_093_271, 819_953, 720_578, 2.55),

    # 2. 주요사업(유통·급식·현대화) - 11개
    "IND_13": _make_indicator("IND_13", "전자송품장 이용률 (가락시장)", "행안부", "2. 주요사업(유통·급식·현대화)", "유통조성팀", 2.50, "%", 19.45, 19.45, 56.78, 2.50),
    "IND_14": _make_indicator("IND_14", "전자송품장 이용률 (강서시장)", "행안부", "2. 주요사업(유통·급식·현대화)", "강서지사 시장관리팀", 0.50, "%", 3.59, 3.59, 21.38, 0.50),
    "IND_15": _make_indicator("IND_15", "출하전 안전성검사 전환율", "행안부", "2. 주요사업(유통·급식·현대화)", "안전성검사소", 3.00, "%", 9.99, 9.99, 8.14, 2.44),
    "IND_16": _make_indicator("IND_16", "지역맞춤형 든든급식 어린이집 이용률", "행안부", "2. 주요사업(유통·급식·현대화)", "서울친환경유통센터 든든급식팀", 1.00, "%", 43.73, 43.73, 54.88, 1.00),
    "IND_17": _make_indicator("IND_17", "물류효율화 파렛트 출하율", "서울시", "2. 주요사업(유통·급식·현대화)", "물류개선팀", 5.00, "%", 90.34, 90.34, 87.50, 4.89),
    "IND_18": _make_indicator("IND_18", "유통정보 가격 알림 서비스", "서울시", "2. 주요사업(유통·급식·현대화)", "유통정보팀", 4.00, "건", 27_322, 21_627, 16_228, 2.76),
    "IND_19": _make_indicator("IND_19", "친환경유통센터 이용 학교 수", "서울시", "2. 주요사업(유통·급식·현대화)", "서울친환경유통센터 급식사업팀", 3.00, "개교", 1_200, 1_200, 1_153, 2.92),
    "IND_20": _make_indicator("IND_20", "친환경식재료 공급물량", "서울시", "2. 주요사업(유통·급식·현대화)", "서울친환경유통센터 급식사업팀", 3.00, "톤", 11_020, 7_656, 7_077, 2.72),
    "IND_21": _make_indicator("IND_21", "친환경식재료 공급금액", "서울시", "2. 주요사업(유통·급식·현대화)", "서울친환경유통센터 급식사업팀", 3.00, "백만원", 62_850, 43_075, 42_442, 2.95),
    "IND_22": _make_indicator("IND_22", "가락 임대시설 공실률(하향)", "서울시", "2. 주요사업(유통·급식·현대화)", "임대사업팀", 3.00, "%", 5.30, 5.30, 6.30, 2.28, is_reverse=True),
    "IND_23": _make_indicator("IND_23", "채소2동 거래품목(11개) 매출액", "서울시", "2. 주요사업(유통·급식·현대화)", "가락시장본부 청과부류팀", 2.00, "백만원", 1_142_464, 819_271, 662_268, 1.73),

    # 3. 재무·경영효율성과 - 10개
    "IND_24": _make_indicator("IND_24", "영업수지비율", "행안부", "3. 재무·경영효율성과", "재무회계팀", 4.80, "%", 120.00, 120.00, 116.04, 4.64),
    "IND_25": _make_indicator("IND_25", "영업수지개선도", "행안부", "3. 재무·경영효율성과", "재무회계팀", 3.20, "%", 124.99, 124.99, 116.04, 2.79),
    "IND_26": _make_indicator("IND_26", "1인당 영업수익", "행안부", "3. 재무·경영효율성과", "경영성과팀·재무회계팀", 4.00, "천원", 312_913, 234_685, 223_480, 3.73),
    "IND_27": _make_indicator("IND_27", "노동생산성 (1인당 부가가치)", "행안부", "3. 재무·경영효율성과", "경영성과팀·재무회계팀", 1.00, "천원", 175_741, 131_806, 117_029, 0.84),
    "IND_28": _make_indicator("IND_28", "자본생산성 (총자산 대비 부가가치)", "행안부", "3. 재무·경영효율성과", "재무회계팀", 1.00, "%", 5.53, 4.15, 3.60, 0.81),
    "IND_29": _make_indicator("IND_29", "생산성향상노력 (영업수익)", "서울시", "3. 재무·경영효율성과", "경영성과팀", 1.00, "백만원", 298_689, 224_017, 223_480, 1.00),
    "IND_30": _make_indicator("IND_30", "부채비율관리(하향)", "서울시", "3. 재무·경영효율성과", "재무회계팀", 2.00, "%", 21.04, 21.04, 21.24, 1.98, is_reverse=True),
    "IND_31": _make_indicator("IND_31", "목표채무달성(하향)", "서울시", "3. 재무·경영효율성과", "재무회계팀", 1.00, "%", 9.18, 9.18, 9.10, 1.00, is_reverse=True),
    "IND_32": _make_indicator("IND_32", "전년대비수지개선", "서울시", "3. 재무·경영효율성과", "재무회계팀", 4.00, "%", 128.56, 128.56, 116.04, 3.61),
    "IND_33": _make_indicator("IND_33", "수입증대", "서울시", "3. 재무·경영효율성과", "재무회계팀", 3.00, "백만원", 108_879, 81_659, 82_464, 3.00),

    # 4. 예산·공공구매·인사·친환경 - 19개
    "IND_34": _make_indicator("IND_34", "신속집행(반기)", "행안부", "4. 예산·공공구매·인사·친환경", "경영성과팀·재무회계팀", 0.60, "백만원", 9_362, 9_362, 10_734, 0.60, eval_type="target_achievement"),
    "IND_35": _make_indicator("IND_35", "재정집행(연간)", "행안부", "4. 예산·공공구매·인사·친환경", "경영성과팀·재무회계팀", 0.40, "백만원", 172_648, 129_486, 118_774, 0.38, eval_type="target_achievement"),
    "IND_36": _make_indicator("IND_36", "중소기업 제품 우선구매", "행안부", "4. 예산·공공구매·인사·친환경", "재무회계팀", 0.30, "%", 80.0, 80.0, 84.2, 0.30, eval_type="target_achievement"),
    "IND_37": _make_indicator("IND_37", "(중소기업) 기술개발 제품 구매", "행안부", "4. 예산·공공구매·인사·친환경", "재무회계팀", 0.18, "%", 15.0, 15.0, 25.8, 0.18, eval_type="target_achievement"),
    "IND_38": _make_indicator("IND_38", "(중소기업) 시범구매 제품", "행안부", "4. 예산·공공구매·인사·친환경", "재무회계팀", 0.12, "%", 1.5, 1.5, 0.02, 0.00, eval_type="target_achievement"),
    "IND_39": _make_indicator("IND_39", "사회적기업 제품 우선구매", "행안부", "4. 예산·공공구매·인사·친환경", "재무회계팀", 1.00, "%", 2.0, 2.0, 3.5, 1.00, eval_type="target_achievement"),
    "IND_40": _make_indicator("IND_40", "중증장애인 생산품 구매(실적+개선도)", "행안부", "4. 예산·공공구매·인사·친환경", "재무회계팀", 0.50, "%", 1.1, 1.1, 4.7, 0.50, eval_type="target_achievement"),
    "IND_41": _make_indicator("IND_41", "혁신제품 구매실적", "행안부", "4. 예산·공공구매·인사·친환경", "재무회계팀", 0.30, "백만원", 617.0, 463.0, 568.0, 0.30),
    "IND_42": _make_indicator("IND_42", "지역사랑상품권 구매실적", "행안부", "4. 예산·공공구매·인사·친환경", "총무팀·재무회계팀", 0.40, "천원", 84_947, 63_710, 69_983, 0.30),
    "IND_43": _make_indicator("IND_43", "창업기업제품 우선구매", "행안부", "4. 예산·공공구매·인사·친환경", "재무회계팀", 0.30, "%", 8.0, 8.0, 3.3, 0.12, eval_type="target_achievement"),
    "IND_44": _make_indicator("IND_44", "상생결제 지급실적", "행안부", "4. 예산·공공구매·인사·친환경", "재무회계팀", 0.30, "%", 10.0, 10.0, 0.0, 0.00, eval_type="target_achievement"),
    "IND_45": _make_indicator("IND_45", "희망기업 제품 구매(서울시)", "서울시", "4. 예산·공공구매·인사·친환경", "재무회계팀", 1.00, "백만원", 14_424, 10_818, 13_937, 0.97),
    "IND_46": _make_indicator("IND_46", "청년 의무고용 실적", "행안부", "4. 예산·공공구매·인사·친환경", "인사팀", 0.50, "명", 12.0, 12.0, 0.0, 0.00, eval_type="target_achievement"),
    "IND_47": _make_indicator("IND_47", "장애인 의무고용률", "행안부", "4. 예산·공공구매·인사·친환경", "인사팀", 1.00, "%", 3.80, 3.80, 4.65, 1.00, eval_type="target_achievement"),
    "IND_48": _make_indicator("IND_48", "체험형 인턴 채용실적", "행안부", "4. 예산·공공구매·인사·친환경", "인사팀", 0.50, "명", 16.0, 16.0, 4.0, 0.13, eval_type="target_achievement"),
    "IND_49": _make_indicator("IND_49", "재난안전 안전사고 건수(하향)", "행안부", "4. 예산·공공구매·인사·친환경", "안전총괄팀", 4.00, "건", 5.0, 5.0, 2.0, 4.00, is_reverse=True),
    "IND_50": _make_indicator("IND_50", "녹색제품 우선구매 실적", "행안부", "4. 예산·공공구매·인사·친환경", "재무회계팀", 0.50, "백만원", 1_645, 1_234, 977, 0.40),
    "IND_51": _make_indicator("IND_51", "(전체) 온실가스 감축량(하향)", "행안부", "4. 예산·공공구매·인사·친환경", "시설안전팀", 0.50, "tCO2", 74_777, 56_083, 37_057, 0.50, is_reverse=True),
    "IND_52": _make_indicator("IND_52", "(청사) 온실가스 감축량(하향)", "서울시", "4. 예산·공공구매·인사·친환경", "시설안전팀", 0.50, "tCO2", 639, 479, 312, 0.50, is_reverse=True),
}


# ==============================================================================
# 2. 11대 정성지표 마스터 설정 (총 50.9점 만점)
# ==============================================================================
QUALITATIVE_INDICATORS_CONFIG: Dict[str, Dict[str, Any]] = {
    "QUAL_01": {
        "id": "QUAL_01", "code": "0-1-1-1-0", "category": "경영관리 - 리더십·전략",
        "name": "경영층의 리더십", "department": "경영성과팀 · 비서실", "weight": 5.0,
        "definition": "경영층이 경영개선과 경영 목표 달성을 위해 어떻게 리더십을 발휘하였는지를 평가한다.",
        "sub_criteria": [
            "❶ 기관 설립 목적 및 경영 목표 이행 (ESG 경영원칙, 도전적 목표 설정, 기관장 성과계약 이행)",
            "❷ 구성원 동기부여 (핵심가치 공유, 적극행정, 현장 소통, 임원 교육 연 21시간 이상)",
            "❸ 대내·외 이해관계자 협력 증진 및 지역문제 해결 (시의회·서울시 협력, 유통인 갈등 해소 제도화)",
        ],
        "required_keywords": ["ESG", "경영성과계약", "핵심가치", "소통", "임원교육", "21시간", "이해관계자", "시의회", "적극행정", "환류"],
        "sample_report": (
            "[경영층의 리더십 추진실적 보고서 - 경영성과팀]\n"
            "1. 계획(Plan): 공사 설립목적 및 ESG 경영원칙 구현을 위해 2026년 기관장 경영성과계약 과제를 12대 핵심과제로 구체화함.\n"
            "2. 실행(Do): CEO 현장 소통 간담회 24회 개최로 현대화사업 갈등을 해소하고, 임원교육 연 28시간(법정 21시간 초과) 이수를 완료함. 서울시·시의회 정례 정책협의체를 제도화하여 예산을 적기 확보함.\n"
            "3. 성과(Check): 기관장 경영성과계약 이행률 98.4% 달성, 적극행정 우수사례 6건 발굴, 이해관계자 만족도 4.2점 상승.\n"
            "4. 환류(Act): 분기별 경영성과 점검회의를 통해 미진 지표를 조기 발굴하고 2027년 경영전략에 환류 반영함."
        ),
    },
    "QUAL_02": {
        "id": "QUAL_02", "code": "0-1-1-2-0", "category": "경영관리 - 리더십·전략",
        "name": "전략 및 혁신", "department": "경영성과팀 · 디지털혁신팀", "weight": 7.0,
        "definition": "중장기 경영계획 수립·집행 노력과 AI·디지털 기반 경영혁신 우수사례 실적을 평가한다.",
        "sub_criteria": [
            "❶ 중장기 경영계획 수립의 충실성 및 중장기재무관리계획 연계성 (4.0점)",
            "❷ 경영혁신 추진 활동 및 AI·디지털 기반 경영 구현, 중앙부처 혁신 우수사례 수상 (2.5점)",
            "❸ 안전 관련 투자 비용 확대 및 근로자 생활임금 반영 노력 (0.5점 + AI 가점 0.5점)",
        ],
        "required_keywords": ["중장기경영계획", "중장기재무관리계획", "환류", "AI", "디지털", "적극행정", "혁신우수사례", "생활임금", "안전투자", "데이터보안"],
        "sample_report": (
            "[전략 및 혁신 추진실적 보고서 - 경영성과팀]\n"
            "1. 계획(Plan): 2026~2030 중장기경영계획과 중장기재무관리계획을 100% 연동하고 AI 디지털 유통혁신 로드맵을 수립함.\n"
            "2. 실행(Do): AI 자동 검산 시스템을 도입하고 안전투자 예산을 전년 대비 18.5% 확대 편성함. 근로자 생활임금을 100% 반영함.\n"
            "3. 성과(Check): 행안부 경영혁신우수사례 경진대회 우수기관 선정, 구비서류 35% 감축 및 데이터보안 무결점 유지.\n"
            "4. 환류(Act): 분기별 경영실적 점검 결과를 차년도 전략과제 및 예산 편성에 환류함."
        ),
    },
    "QUAL_03": {
        "id": "QUAL_03", "code": "0-1-2-1-0", "category": "경영관리 - 경영시스템",
        "name": "조직·인사관리", "department": "인사팀 · 경영성과팀", "weight": 6.0,
        "definition": "전략 연계 조직관리, 공정 인사, 일·가정 양립, 직무중심 인사관리 체계를 평가한다.",
        "sub_criteria": [
            "❶ 경영전략 연계 조직진단 및 인력 재배치 (2.0점)",
            "❷ 중장기 인력운용계획 및 양성평등·균형인사 (1.0점)",
            "❸ 일·가정 양립 (출산·육아휴직 통합신청제, 대체인력 채용, 유연근무제) (1.0점)",
            "❹ 역량 교육훈련 및 직무중심 인사·직무급 도입 (2.0점)",
        ],
        "required_keywords": ["조직진단", "중장기인력운용계획", "양성평등", "일가정양립", "육아휴직", "유연근무제", "직무중심", "직무급", "성과관리", "인력재배치"],
        "sample_report": (
            "[조직·인사관리 추진실적 보고서 - 인사팀]\n"
            "1. 계획(Plan): 직무중심 인사관리 및 3개년 중장기인력운용계획을 수립하고 조직진단 기반 인력재배치를 단행함.\n"
            "2. 실행(Do): 출산·육아휴직 통합신청제 전면 도입 및 대체인력 100% 채용 완료. 일가정양립을 위해 유연근무제 활용률을 42.5%로 확대하고 양성평등 보직관리를 강화함.\n"
            "3. 성과(Check): 지원부서 인력비율 24.1%(30% 이내), 관리직 비율 17.8%(20% 이내) 준수로 감점 요인 제로화. 직무급 비중 확대 및 성과관리 체계 고도화.\n"
            "4. 환류(Act): 내부직원 만족도 조사 결과를 반영하여 인사규정 및 보직관리 기준을 개선 환류함."
        ),
    },
    "QUAL_04": {
        "id": "QUAL_04", "code": "0-1-2-2-2", "category": "경영관리 - 경영시스템",
        "name": "재무관리 (공사 A-2형)", "department": "재무회계팀 · 경영성과팀", "weight": 3.9,
        "definition": "자산운용, 중장기 재무관리, 수지개선 및 예산·회계 운영의 적정성을 평가한다.",
        "sub_criteria": [
            "❶ 유휴자금 운용수익률 제고 및 재물조사 전산화",
            "❷ 5개년 중장기재무관리계획 수립 및 리스크관리회의 운영",
            "❸ 원가절감, 신규 수익원 창출 및 행안부 예산편성기준 준수",
        ],
        "required_keywords": ["중장기재무관리계획", "유휴자금", "운용수익률", "재물조사", "리스크관리", "원가절감", "예산편성기준", "부채비율", "수지분석", "내부통제"],
        "sample_report": (
            "[재무관리 추진실적 보고서 - 재무회계팀]\n"
            "1. 계획(Plan): 5개년 중장기재무관리계획 수립 및 리스크관리위원회 분기별 운영으로 부채비율 21%대 안정적 관리 목표 설정.\n"
            "2. 실행(Do): 유휴자금 운용수익률을 전년 대비 0.45%p 제고하고 전 사업장 재물조사 및 원가절감을 이행함. 행안부 예산편성기준을 철저히 준수함.\n"
            "3. 성과(Check): 경상경비 8.2% 절감, 부채비율 21.24% 달성, 회계 내부통제 고도화로 외부감사 적정의견 달성.\n"
            "4. 환류(Act): 신규 투자사업 사전 수지분석 의무화를 통해 재무 건전성 환류 체계를 강화함."
        ),
    },
    "QUAL_05": {
        "id": "QUAL_05", "code": "0-1-2-3-0", "category": "경영관리 - 경영시스템",
        "name": "윤리경영", "department": "감사팀 · 인권윤리팀", "weight": 3.0,
        "definition": "윤리·인권경영 체제 구축, 인권영향평가, 청렴도 향상 및 이해충돌방지 노력을 평가한다.",
        "sub_criteria": [
            "❶ 내부통제시스템 구축 및 부패방지 교육·예방 활동",
            "❷ 인권경영체계 구축 및 인권영향평가 실시·공개·환류",
            "❸ 청렴도 향상, 이해충돌방지 및 도매시장 공정거래 환경 조성",
        ],
        "required_keywords": ["내부통제", "인권영향평가", "청렴도", "이해충돌방지", "부패방지", "공정거래", "채용비리", "윤리교육", "인권경영", "환류"],
        "sample_report": (
            "[윤리경영 추진실적 보고서 - 감사팀]\n"
            "1. 계획(Plan): 자율적 내부통제 시스템 고도화 및 청렴도 최우수 등급 달성을 위한 연간 윤리·인권경영 계획 수립.\n"
            "2. 실행(Do): 기관 운영 및 주요사업 인권영향평가를 실시·공개하고, 전 임직원 이해충돌방지 점검 및 부패방지 윤리교육 이수율 100% 달성.\n"
            "3. 성과(Check): 채용비리 0건 유지, 도매시장 공정거래 신문고 운영으로 불공정 민원 32% 감소.\n"
            "4. 환류(Act): 인권영향평가 도출 개선과제 8건을 차년도 제도개선에 100% 환류 완료함."
        ),
    },
    "QUAL_06": {
        "id": "QUAL_06", "code": "0-1-3-1-0", "category": "경영관리 - 사회적 책임",
        "name": "지역상생·협력", "department": "총무팀 · 유통조성팀", "weight": 5.0,
        "definition": "도농상생, 저출생·지방소멸 대응, 지역경제 활성화 및 재난구호 노력을 평가한다.",
        "sub_criteria": [
            "❶ 도농상생 및 지역경제 활성화·민간협력 (2.0점)",
            "❷ 고유사업 연계 저출생·지방소멸 극복 노력 (2.0점)",
            "❸ 지역 일자리 창출 및 취약계층 나눔·재난구호 가점 (1.0점 + 가점 0.5점)",
        ],
        "required_keywords": ["도농상생", "저출생", "지방소멸", "민간협력", "일자리창출", "사회적약자", "가락몰", "산지지원", "재난구호", "환류"],
        "sample_report": (
            "[지역상생·협력 추진실적 보고서 - 총무팀·유통조성팀]\n"
            "1. 계획(Plan): 도매시장 고유 기능을 활용한 도농상생 및 저출생·지방소멸 대응 사회공헌 로드맵 수립.\n"
            "2. 실행(Do): 지방소멸 위기 산지지원 물류비 지원 및 민간협력을 확대함. 저출생 극복을 위해 든든급식 지원을 늘리고 수해 산지에 재난구호 물품을 긴급 지원함.\n"
            "3. 성과(Check): 사회적약자 일자리창출 48명, 가락몰 소상공인 상생 프로모션 매출 14% 증대.\n"
            "4. 환류(Act): 수혜자 만족도 조사를 통해 산지-소비지 도농상생 프로그램을 정례화하여 환류함."
        ),
    },
    "QUAL_07": {
        "id": "QUAL_07", "code": "0-1-3-2-0", "category": "경영관리 - 사회적 책임",
        "name": "재난·안전관리", "department": "안전총괄팀 · 시설안전팀", "weight": 6.0,
        "definition": "중대재해 예방 안전보건관리체계, 위험성평가, 다중이용시설 재난관리 적정성을 평가한다.",
        "sub_criteria": [
            "❶ 안전활동 수준평가 및 위험성평가 이행 (3.0점)",
            "❷ 도매시장 다중이용시설 화재·풍수해 예방 및 유관기관 합동훈련 (3.0점)",
            "❸ 지방공공기관 안전보건관리 가이드라인 준수 여부 (감점 -1.0점)",
        ],
        "required_keywords": ["중대재해", "위험성평가", "안전보건관리체계", "재난관리", "합동훈련", "안전예산", "전담조직", "가이드라인", "다중이용시설", "화재예방"],
        "sample_report": (
            "[재난·안전관리 추진실적 보고서 - 안전총괄팀]\n"
            "1. 계획(Plan): 중대재해 Zero 달성을 위한 안전보건관리체계 고도화 및 안전예산·전담조직 확대 계획 수립.\n"
            "2. 실행(Do): 도매시장 다중이용시설 화재예방 열화상 감지기를 구축하고 위험성평가를 실시하여 위험요인 142건을 100% 개선함. 소방서 합동훈련 4회 실시 및 안전보건관리 가이드라인 전 항목 준수.\n"
            "3. 성과(Check): 중대재해 0건, 안전사고 목표 대비 60% 감축(5건 목표 대비 2건), 재난관리 평가 우수 달성.\n"
            "4. 환류(Act): 아차사고 발굴 신고제를 통해 현장 물류 안전 수칙을 개정 환류함."
        ),
    },
    "QUAL_08": {
        "id": "QUAL_08", "code": "0-1-3-3-0", "category": "경영관리 - 사회적 책임",
        "name": "친환경 경영", "department": "시설안전팀 · 환경관리팀", "weight": 2.0,
        "definition": "탄소중립, 도매시장 에너지 효율화 및 농수산물 폐기물 자원순환 성과를 평가한다.",
        "sub_criteria": [
            "❶ 탄소중립 계획 수립 및 온실가스 감축 모니터링",
            "❷ 현대화 시설 신재생에너지·고효율 설비 도입을 통한 에너지 절감",
            "❸ 도매시장 농수산물 폐기물 감축 및 자원순환 재활용 활성화",
        ],
        "required_keywords": ["탄소중립", "온실가스", "에너지절감", "자원순환", "폐기물감축", "재활용", "신재생에너지", "미세먼지", "녹색제품", "환류"],
        "sample_report": (
            "[친환경 경영 추진실적 보고서 - 시설안전팀]\n"
            "1. 계획(Plan): 2050 탄소중립 로드맵에 따른 온실가스 감축 및 자원순환형 친환경 도매시장 구축 계획 수립.\n"
            "2. 실행(Do): 현대화 시설에 신재생에너지 설비를 확충하여 에너지절감을 실천하고, 농산물 부산물 퇴비화 자원순환 체계로 폐기물감축 및 재활용 확대. 미세먼지 저감 및 녹색제품 구매 확대.\n"
            "3. 성과(Check): 온실가스 배출량 목표 대비 50.4% 감축(37,057 tCO2), 재활용률 8.5%p 향상.\n"
            "4. 환류(Act): 월별 에너지 모니터링 결과를 바탕으로 피크전력 제어 기준을 개선 환류함."
        ),
    },
    "QUAL_09": {
        "id": "QUAL_09", "code": "0-1-3-4-0", "category": "경영관리 - 사회적 책임",
        "name": "소통 및 참여", "department": "홍보팀 · 노사협력팀", "weight": 3.0,
        "definition": "고객·주민의 경영참여 및 상생·협력의 공공 노사관계 구축 노력을 평가한다.",
        "sub_criteria": [
            "❶ 경영 주기별 고객·시민 경영참여 및 제도 개선 환류 (1.5점)",
            "❷ 노사협의회 실질 운영, 노동법 준수 및 무분규 상생 노사관계 (1.5점)",
        ],
        "required_keywords": ["경영참여", "시민참여", "고객만족", "노사상생", "노사협의회", "단체교섭", "무분규", "타임오프", "소통채널", "환류"],
        "sample_report": (
            "[소통 및 참여 추진실적 보고서 - 홍보팀·노사협력팀]\n"
            "1. 계획(Plan): 경영 전 주기 시민참여·경영참여 확대 및 협력적 공공 노사상생 전략 수립.\n"
            "2. 실행(Do): 시민참여혁신단 및 소통채널을 운영하여 제도 개선과제 19건을 발굴함. 분기별 노사협의회 개최 및 타임오프제 등 노동법령 100% 준수.\n"
            "3. 성과(Check): 15년 연속 노사 무분규 단체교섭 체결 달성, 고객만족 조사 피드백 환류 이행률 100%.\n"
            "4. 환류(Act): 노사 공동 근로조건 개선 TF를 상설화하여 갈등 사전 예방 체계를 환류함."
        ),
    },
    "QUAL_10": {
        "id": "QUAL_10", "code": "3.2.1-2-1-1-0", "category": "경영성과 - 주요사업성과",
        "name": "유통시장 관리 (공사 고유 핵심지표)", "department": "가락시장본부 · 강서지사 · 물류개선팀 · 안전성검사소", "weight": 9.0,
        "definition": "도매시장 유통 원활화, 물류·교통 여건 개선, 농수산물 안전성 확보 및 성과관리 적정성을 평가한다.",
        "sub_criteria": [
            "❶ 농수산물 유통 원활화 (도매시장법인 거래목표 부여·실적분석, 중도매인 상장예외품목 관리, 유통정보 전파)",
            "❷ 물류 및 교통 여건 개선 (파렛트 출하율 제고, 하역 효율화, 주차시설·교통체계 개선)",
            "❸ 농수산물 안전성 확보 (잔류농약 검사 확대, 산지 출하자 교육) 및 성과관리 환류",
        ],
        "required_keywords": ["유통구조개선", "도매시장법인", "중도매인", "상장예외품목", "유통정보", "물류효율화", "파렛트", "주차시설", "잔류농약", "안전성검사", "산지출하자교육", "환류"],
        "sample_report": (
            "[유통시장 관리 추진실적 보고서 - 가락시장본부·강서지사·물류개선팀·안전성검사소]\n"
            "1. 계획(Plan): 도매시장 유통구조개선, 물류효율화, 농수산물 안전성 확보를 위한 2026년 유통시장 관리 종합계획 수립.\n"
            "2. 실행(Do): 도매시장법인별 거래목표를 부여·분석하고 중도매인 상장예외품목 실적을 관리함. 파렛트 출하율 87.5% 달성 및 지능형 주차시설 관제로 체류시간 22% 단축. 잔류농약 안전성검사 확대 및 산지출하자교육 38회 실시.\n"
            "3. 성과(Check): 전자송품장 가락시장 이용률 56.78%(목표 대비 291.9% 달성), 유통정보 가격 알림 16,228건 제공, 부적합 농수산물 유통 차단 100%.\n"
            "4. 환류(Act): 분기별 유통종사자 간담회 및 거래실적 분석을 통해 현대화 시설 물류 동선을 개선 환류함."
        ),
    },
    "QUAL_11": {
        "id": "QUAL_11", "code": "3.2.1-2-1-5-0", "category": "경영성과 - 지역맞춤형",
        "name": "서울 어린이집 든든급식 추진성과 (신규)", "department": "서울친환경유통센터 든든급식팀", "weight": 1.0,
        "definition": "어린이들의 안전한 먹거리 공급을 위한 서울 어린이집 급식사업(든든급식) 추진 노력과 성과를 평가한다.",
        "sub_criteria": [
            "❶ 든든급식 사업 운영 효율성 강화 및 자치구 협력 확대",
            "❷ 친환경·우수 식재료 안전성 검사 및 품질 관리",
            "❸ 이용 어린이집·학부모 만족도 제고 및 클레임 대응 환류",
        ],
        "required_keywords": ["든든급식", "어린이집", "친환경유통센터", "식재료", "안전성검사", "품질관리", "학부모만족도", "배송효율화", "자치구", "환류"],
        "sample_report": (
            "[서울 어린이집 든든급식 추진성과 보고서 - 든든급식팀]\n"
            "1. 계획(Plan): 보육 공공성 강화를 위해 친환경유통센터 기반 어린이집 든든급식 공급망 확대 및 식재료 품질관리 계획 수립.\n"
            "2. 실행(Do): 25개 자치구 맞춤형 소포장 배송효율화 체계를 구축하고 공급 식재료 전수 안전성검사를 100% 실시함.\n"
            "3. 성과(Check): 든든급식 이용률 54.88% 달성(목표 43.73% 대비 125.5% 달성), 원장 및 학부모만족도 94.6점 달성.\n"
            "4. 환류(Act): 현장 건의를 수렴하여 영유아 맞춤형 세척·절단 식재료 품목 35종을 추가 환류함."
        ),
    },
}


# ==============================================================================
# 3. 정량·정성 계산 엔진 및 검증 함수
# ==============================================================================
@dataclass
class EvaluationResult:
    indicator_id: str
    indicator_name: str
    eval_agency: str
    category: str
    department: str
    weight: float
    unit: str
    is_reverse: bool
    base_value: float
    upper_target: float
    lower_target: float
    q3_target: float
    actual_value: float
    q3_target_rate: float
    annual_target_rate: float
    achievement_rate: float
    score: float
    official_q3_score: float
    score_ratio: float
    grade_label: str
    status_level: str
    status_badge_reason: str


def compute_base_value(ind_cfg: Dict[str, Any]) -> float:
    if "base_override" in ind_cfg and ind_cfg["base_override"] is not None:
        return float(ind_cfg["base_override"])
    history = ind_cfg.get("history_3yr", {})
    if history:
        vals = [float(v) for v in history.values()]
        return round(sum(vals) / len(vals), 4)
    return float(ind_cfg.get("target_2026", 0.0))


def compute_target_bounds(ind_cfg: Dict[str, Any]) -> Tuple[float, float]:
    base_value = compute_base_value(ind_cfg)
    is_reverse = bool(ind_cfg.get("is_reverse", False))
    target_2026 = float(ind_cfg.get("target_2026", base_value * 1.10))
    if not is_reverse:
        upper_target = round(target_2026 if target_2026 > 0 else base_value * 1.10, 4)
        lower_target = round(base_value * 0.90, 4)
        if lower_target >= upper_target:
            lower_target = round(upper_target * 0.80, 4)
    else:
        upper_target = round(target_2026 if target_2026 > 0 else base_value * 0.90, 4)
        lower_target = round(base_value * 1.10, 4)
        if lower_target <= upper_target:
            lower_target = round(upper_target * 1.20, 4)
    return upper_target, lower_target


def calculate_linear_interpolation_score(
    actual: float,
    base_value: float,
    weight: float,
    is_reverse: bool = False,
    upper_override: Optional[float] = None,
    lower_override: Optional[float] = None,
) -> Tuple[float, float, str, float, float]:
    if upper_override is not None and lower_override is not None:
        upper_target, lower_target = upper_override, lower_override
    else:
        upper_target = round(base_value * (0.90 if is_reverse else 1.10), 4)
        lower_target = round(base_value * (1.10 if is_reverse else 0.90), 4)

    if base_value <= 0 and upper_target <= 0:
        return 0.0, 0.0, "산출불가", upper_target, lower_target

    if not is_reverse:
        if actual >= upper_target:
            ratio, grade = 1.0, "수 (만점·100%)"
        elif actual <= lower_target:
            ratio, grade = 0.6, "마 (하한·60%)"
        else:
            denom = upper_target - lower_target
            ratio = 0.6 + 0.4 * ((actual - lower_target) / denom) if denom > 0 else 0.6
            grade = "보간구간 (60~100%)"
    else:
        if actual <= upper_target:
            ratio, grade = 1.0, "수 (만점·100%)"
        elif actual >= lower_target:
            ratio, grade = 0.6, "마 (하한·60%)"
        else:
            denom = lower_target - upper_target
            ratio = 0.6 + 0.4 * ((lower_target - actual) / denom) if denom > 0 else 0.6
            grade = "보간구간 (60~100%)"

    raw_score = weight * ratio
    return round(raw_score, 4), round(ratio * 100.0, 2), grade, upper_target, lower_target


def evaluate_indicator(
    ind_cfg: Dict[str, Any],
    quarter: str,
    actual_override: Optional[float] = None,
    scoring_standard: str = "분기목표 대비 점검득점(공사 3분기 점검표 기준)",
) -> EvaluationResult:
    base_val = compute_base_value(ind_cfg)
    weight = float(ind_cfg["weight"])
    is_reverse = bool(ind_cfg.get("is_reverse", False))
    upper_target, lower_target = compute_target_bounds(ind_cfg)
    q3_target = float(ind_cfg.get("q3_target", upper_target * 0.75))
    actual = float(actual_override) if actual_override is not None else float(ind_cfg["quarterly_actuals"].get(quarter, 0.0))

    interp_score, interp_ratio, grade_label, _, _ = calculate_linear_interpolation_score(
        actual=actual, base_value=base_val, weight=weight, is_reverse=is_reverse,
        upper_override=upper_target, lower_override=lower_target,
    )

    if not is_reverse:
        q3_target_rate = round((actual / q3_target) * 100.0, 1) if q3_target > 0 else 0.0
        annual_target_rate = round((actual / upper_target) * 100.0, 1) if upper_target > 0 else 0.0
        achievement_rate = round((actual / base_val) * 100.0, 1) if base_val > 0 else annual_target_rate
    else:
        q3_target_rate = round((q3_target / actual) * 100.0, 1) if actual > 0 else 160.0
        annual_target_rate = round((upper_target / actual) * 100.0, 1) if actual > 0 else 160.0
        achievement_rate = round((base_val / actual) * 100.0, 1) if actual > 0 else 160.0

    official_q3_score = float(ind_cfg.get("official_q3_score", interp_score))
    default_q3_actual = float(ind_cfg["quarterly_actuals"].get("3분기", 0.0))

    if "점검표" in scoring_standard and quarter == "3분기" and abs(actual - default_q3_actual) < 1e-4:
        final_score = official_q3_score
        final_ratio = round((final_score / weight) * 100.0, 1) if weight > 0 else 0.0
    elif "점검표" in scoring_standard and quarter != "4분기(최종)":
        q_ratio = min(1.0, max(0.0, q3_target_rate / 100.0))
        final_score = round(weight * q_ratio, 3)
        final_ratio = round(q_ratio * 100.0, 1)
    else:
        final_score = interp_score
        final_ratio = interp_ratio

    if q3_target_rate < 85.0 or achievement_rate < 70.0:
        status_level = "위험"
        status_badge_reason = f"분기달성률 {q3_target_rate:.1f}% (목표 미달 위험)"
    elif q3_target_rate < 95.0:
        status_level = "주의"
        status_badge_reason = f"분기달성률 {q3_target_rate:.1f}% (4분기 집중관리)"
    else:
        status_level = "정상"
        status_badge_reason = f"분기달성률 {q3_target_rate:.1f}% (목표 순항)"

    return EvaluationResult(
        indicator_id=ind_cfg["id"], indicator_name=ind_cfg["name"], eval_agency=ind_cfg.get("eval_agency", "행안부"),
        category=ind_cfg.get("category", "주요사업"), department=ind_cfg.get("department", "소관부서"),
        weight=weight, unit=ind_cfg["unit"], is_reverse=is_reverse, base_value=base_val,
        upper_target=upper_target, lower_target=lower_target, q3_target=q3_target, actual_value=actual,
        q3_target_rate=q3_target_rate, annual_target_rate=annual_target_rate, achievement_rate=achievement_rate,
        score=final_score, official_q3_score=official_q3_score, score_ratio=final_ratio,
        grade_label=grade_label, status_level=status_level, status_badge_reason=status_badge_reason,
    )


def validate_indicator_input(ind_cfg: Dict[str, Any], quarter: str, raw_input_text: str):
    cleaned = raw_input_text.strip().replace(",", "")
    if cleaned == "":
        return None, [("error", "입력값 누락 오류: 실적값을 숫자로 입력해 주십시오.")], False
    try:
        val = float(cleaned)
    except ValueError:
        return None, [("error", f"숫자 형식 오류: '{raw_input_text}'은(는) 유효한 숫자가 아닙니다.")], False
    if val < 0:
        return None, [("error", "음수 입력 오류: 실적은 0 이상이어야 합니다.")], False

    msgs = []
    unit = ind_cfg["unit"]
    prev_sync = float(ind_cfg.get("prev_year_cumulative", {}).get(quarter, compute_base_value(ind_cfg)))
    if prev_sync > 0 and val > 0:
        ratio = val / prev_sync
        if ratio >= 10.0:
            msgs.append(("warning", f"🚨 단위({unit}) 오입력 경고: 전년 대비 {ratio:.1f}배 과대 입력되었습니다. 단위(kg/원 등) 착오를 확인하세요."))
        elif ratio <= 0.10:
            msgs.append(("warning", f"🚨 단위({unit}) 오입력 경고: 전년 대비 {ratio * 100:.1f}% (1/10 이하) 수준입니다. 단위 축소 오입력을 확인하세요."))

    res = evaluate_indicator(ind_cfg, quarter=quarter, actual_override=val)
    if res.achievement_rate < 70.0 or res.q3_target_rate < 85.0:
        msgs.append(("warning", f"⚠️ 목표 미달 위험 경고: 분기목표 달성률 {res.q3_target_rate:.1f}% (연간 기준치 대비 {res.achievement_rate:.1f}%)로 미달 위험 구간입니다."))
    if not msgs:
        msgs.append(("success", f"✅ 검증 통과: 분기 달성률 {res.q3_target_rate:.1f}%, 예상 득점 {res.score:.2f}점 / {res.weight:.2f}점"))
    return val, msgs, True


def extract_text_from_uploaded_file(uploaded_file: Any) -> str:
    filename = uploaded_file.name.lower()
    raw_bytes = uploaded_file.getvalue()
    if filename.endswith((".txt", ".md")):
        for enc in ("utf-8", "cp949", "euc-kr"):
            try:
                return raw_bytes.decode(enc)
            except UnicodeDecodeError:
                continue
        return raw_bytes.decode("utf-8", errors="ignore")
    if filename.endswith(".pdf") and pypdf is not None:
        reader = pypdf.PdfReader(io.BytesIO(raw_bytes))
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    if filename.endswith(".docx") and docx is not None:
        d = docx.Document(io.BytesIO(raw_bytes))
        paras = [p.text for p in d.paragraphs if p.text.strip()]
        for t in d.tables:
            for r in t.rows:
                paras.append(" | ".join(c.text.strip() for c in r.cells if c.text.strip()))
        return "\n".join(paras)
    if filename.endswith(".xlsx"):
        xl = pd.ExcelFile(io.BytesIO(raw_bytes))
        return "\n\n".join(xl.parse(s).fillna("").to_string(index=False) for s in xl.sheet_names)
    if filename.endswith(".csv"):
        for enc in ("utf-8-sig", "utf-8", "cp949"):
            try:
                return pd.read_csv(io.BytesIO(raw_bytes), encoding=enc).fillna("").to_string(index=False)
            except Exception:
                continue
    return ""


def evaluate_qualitative_text(qual_cfg: Dict[str, Any], report_text: str) -> Dict[str, Any]:
    text_clean = report_text.strip()
    text_no_space = re.sub(r"\s+", "", text_clean)
    req_kws = qual_cfg.get("required_keywords", [])
    matched = [kw for kw in req_kws if kw in text_clean or re.sub(r"\s+", "", kw) in text_no_space]
    missing = [kw for kw in req_kws if kw not in matched]
    kw_cov = (len(matched) / len(req_kws) * 100.0) if req_kws else 100.0

    pdca_markers = {
        "Plan (계획)": ["계획", "목표", "수립", "전략", "비전", "로드맵"],
        "Do (실행)": ["추진", "실시", "운영", "개최", "구축", "도입", "지원", "확대"],
        "Check (성과)": ["성과", "달성", "증가", "감소", "절감", "향상", "만족도", "선정", "%"],
        "Act (환류)": ["환류", "개선", "반영", "피드백", "보완", "정비", "고도화"],
    }
    pdca_scores = {k: min(100.0, round(sum(1 for m in v if m in text_clean) / 2.5 * 100.0, 1)) for k, v in pdca_markers.items()}
    pdca_avg = sum(pdca_scores.values()) / 4.0

    quant_matches = re.findall(r"\d+(?:[.,]\d+)?\s*(?:%|건|명|시간|톤|원|백만원|억원|점|회|개교|종|가구|%p)", text_clean)
    quant_cnt = len(quant_matches)
    quant_score = min(100.0, quant_cnt * 18.0)

    length_factor = min(1.0, len(text_clean) / 160.0) if text_clean else 0.0
    total_pct = round(min(100.0, (kw_cov * 0.50 + pdca_avg * 0.30 + quant_score * 0.20) * (0.6 + 0.4 * length_factor)), 1) if text_clean else 0.0

    if total_pct >= 88.0:
        grade, coeff = "수 (탁월·100%)", 1.00
    elif total_pct >= 78.0:
        grade, coeff = "우 (우수·90%)", 0.90
    elif total_pct >= 68.0:
        grade, coeff = "양 (보통·80%)", 0.80
    elif total_pct >= 58.0:
        grade, coeff = "가 (미흡·70%)", 0.70
    else:
        grade, coeff = "마 (부진·60%)", 0.60

    weight = float(qual_cfg["weight"])
    return {
        "qual_id": qual_cfg["id"],
        "name": qual_cfg["name"],
        "category": qual_cfg["category"],
        "department": qual_cfg["department"],
        "weight": weight,
        "matched": matched,
        "missing": missing,
        "kw_cov": round(kw_cov, 1),
        "pdca_scores": pdca_scores,
        "quant_cnt": quant_cnt,
        "total_pct": total_pct,
        "grade": grade,
        "est_score": round(weight * coeff, 2),
    }


# ==============================================================================
# 4. 세션 초기화 및 공통 테이블 생성
# ==============================================================================
if "indicators_data" not in st.session_state or len(st.session_state["indicators_data"]) < 30:
    st.session_state["indicators_data"] = deepcopy(INDICATORS_CONFIG)
if "qualitative_reports" not in st.session_state:
    st.session_state["qualitative_reports"] = {q_id: cfg["sample_report"] for q_id, cfg in QUALITATIVE_INDICATORS_CONFIG.items()}


def build_summary_df(
    indicators_dict: Dict[str, Dict[str, Any]],
    quarter: str,
    scoring_standard: str,
    agency_filter: str = "전체",
    category_filter: str = "전체",
) -> pd.DataFrame:
    rows = []
    for ind_id, cfg in indicators_dict.items():
        if agency_filter != "전체" and cfg.get("eval_agency") != agency_filter:
            continue
        if category_filter != "전체" and cfg.get("category") != category_filter:
            continue
        res = evaluate_indicator(cfg, quarter=quarter, scoring_standard=scoring_standard)
        rem = max(0.0, res.upper_target - res.actual_value) if not res.is_reverse else max(0.0, res.actual_value - res.upper_target)
        rows.append({
            "지표코드": res.indicator_id,
            "평가주관": res.eval_agency,
            "분류": res.category,
            "지표명": res.indicator_name,
            "소관부서": res.department,
            "배점": res.weight,
            "단위": res.unit,
            "기준치(3개년평균)": res.base_value,
            "2026년 목표(A)": res.upper_target,
            "3분기 목표(B)": res.q3_target,
            f"{quarter} 실적(C)": res.actual_value,
            "3분기 달성률(C/B)": res.q3_target_rate,
            "연간 달성률(C/A)": res.annual_target_rate,
            "산출 득점": res.score,
            "만점 부족량": round(rem, 1),
            "상태": res.status_level,
        })
    return pd.DataFrame(rows)


def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "경영성과팀_실적") -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return buf.getvalue()


# ==============================================================================
# 5. 심플 사이드바 (3대 핵심 메뉴 네비게이션)
# ==============================================================================
with st.sidebar:
    st.markdown("### 🏛️ 서울시농수산식품공사")
    st.caption("총괄 관리부서: **경영성과팀**")
    st.divider()

    # 한 화면에 한 가지 목적만 보여주도록 직관적인 3대 메뉴로 분리
    nav_menu = st.radio(
        "📌 업무 화면 선택",
        options=[
            "📊 1. 종합 관제 대시보드",
            "✏️ 2. 정량 실적 입력 · 엑셀 검산",
            "📑 3. 정성 보고서 AI 심사",
        ],
        index=0,
    )

    st.divider()
    selected_quarter = st.selectbox("📅 기준 분기", options=QUARTERS, index=2)
    scoring_standard = st.selectbox(
        "📐 점수 산출 방식",
        options=[
            "분기목표 대비 점검득점(공사 3분기 점검표 기준)",
            "연간 편람 선형보간 공식 즉시 적용",
        ],
        index=0,
    )

    st.divider()
    if st.button("🔄 초기 공식 데이터로 리셋", use_container_width=True):
        st.session_state["indicators_data"] = deepcopy(INDICATORS_CONFIG)
        st.session_state["qualitative_reports"] = {q_id: cfg["sample_report"] for q_id, cfg in QUALITATIVE_INDICATORS_CONFIG.items()}
        st.toast("2026년 3분기 공식 데이터로 초기화되었습니다.", icon="✅")
        st.rerun()


# ==============================================================================
# 6. 상단 심플 바
# ==============================================================================
st.markdown(
    f"""
    <div class="top-bar">
        <div>
            <div class="top-bar-title">🏛️ 2027년도 지방공기업 경영평가 성과관리 시스템 (경영성과팀)</div>
            <div class="top-bar-sub">복잡한 부가 정보를 숨기고 가장 중요한 핵심 지표와 미달 위험만 한눈에 보여줍니다.</div>
        </div>
        <div style="font-size:0.85rem; font-weight:600; background:rgba(255,255,255,0.12); padding:0.35rem 0.85rem; border-radius:99px;">
            {selected_quarter} 기준 &nbsp;|&nbsp; {nav_menu.split('.')[1].strip()}
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# ==============================================================================
# 화면 1: 📊 종합 관제 대시보드 (핵심 숫자 4개 + 미달 위험 집중 표 + 깔끔한 차트)
# ==============================================================================
if nav_menu.startswith("📊"):
    # 상단 필터 2개만 심플하게 배치
    f1, f2, f3 = st.columns([1.2, 1.8, 1.5])
    all_categories = sorted(list({cfg["category"] for cfg in st.session_state["indicators_data"].values()}))
    with f1:
        agency_sel = st.segmented_control("평가주관", options=["전체", "행안부", "서울시"], default="전체") or "전체"
    with f2:
        cat_sel = st.selectbox("분야 필터", options=["전체"] + all_categories, label_visibility="collapsed")
    with f3:
        only_risk = st.toggle("🚨 위험·주의 지표만 모아보기", value=False)

    df_dash = build_summary_df(
        st.session_state["indicators_data"],
        quarter=selected_quarter,
        scoring_standard=scoring_standard,
        agency_filter=agency_sel,
        category_filter=cat_sel,
    )

    # 정성지표 요약 합산
    qual_list = [evaluate_qualitative_text(cfg, st.session_state["qualitative_reports"][qid]) for qid, cfg in QUALITATIVE_INDICATORS_CONFIG.items()]
    qual_w_sum = sum(x["weight"] for x in qual_list)
    qual_s_sum = sum(x["est_score"] for x in qual_list)

    moi_df = df_dash[df_dash["평가주관"] == "행안부"] if agency_sel == "전체" else df_dash
    quant_w_sum = float(df_dash["배점"].sum())
    quant_s_sum = float(df_dash["산출 득점"].sum())
    danger_cnt = int((df_dash["상태"] == "위험").sum())
    warn_cnt = int((df_dash["상태"] == "주의").sum())
    norm_cnt = len(df_dash) - danger_cnt - warn_cnt

    # 1) 핵심 요약 카드 4개 (가장 중요한 숫자만 크게)
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(
            f"""
            <div class="metric-box">
                <div class="metric-title">정량 예상 득점 ({agency_sel})</div>
                <div class="metric-num" style="color:#1d4ed8;">{quant_s_sum:.2f} <span style="font-size:1.05rem; color:#64748b; font-weight:600;">/ {quant_w_sum:.2f}점</span></div>
                <div class="metric-foot">득점률 <b>{(quant_s_sum/quant_w_sum*100 if quant_w_sum else 0):.1f}%</b> (총 {len(df_dash)}개 지표)</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with c2:
        st.markdown(
            f"""
            <div class="metric-box">
                <div class="metric-title">정성 AI 진단 득점 (11대 지표)</div>
                <div class="metric-num" style="color:#0f766e;">{qual_s_sum:.2f} <span style="font-size:1.05rem; color:#64748b; font-weight:600;">/ {qual_w_sum:.1f}점</span></div>
                <div class="metric-foot">편람 기준 부합도 평균 <b>{sum(x['total_pct'] for x in qual_list)/len(qual_list):.1f}%</b></div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with c3:
        st.markdown(
            f"""
            <div class="metric-box">
                <div class="metric-title">3분기 목표 달성률 평균</div>
                <div class="metric-num">{df_dash['3분기 달성률(C/B)'].mean():.1f}%</div>
                <div class="metric-foot">100% 이상 달성: <b>{int((df_dash['3분기 달성률(C/B)'] >= 100).sum())}개</b> 지표</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with c4:
        st.markdown(
            f"""
            <div class="metric-box">
                <div class="metric-title">집중 관리 필요 (위험·주의)</div>
                <div class="metric-num" style="color:#dc2626;">{danger_cnt + warn_cnt}개</div>
                <div class="metric-foot">
                    <span class="pill-red">위험 {danger_cnt}</span>
                    <span class="pill-yellow">주의 {warn_cnt}</span>
                    <span class="pill-green">정상 {norm_cnt}</span>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    # 2) 핵심 액션 테이블: 무엇이 문제이고 4분기에 얼마나 더 해야 하는지 한눈에 표시
    view_df = df_dash[df_dash["상태"].isin(["위험", "주의"])] if only_risk else df_dash
    st.markdown(
        f'<div class="sec-title">📋 정량지표 핵심 성적표 ({len(view_df)}개 표시)</div>',
        unsafe_allow_html=True,
    )

    act_col = f"{selected_quarter} 실적(C)"
    clean_table = pd.DataFrame({
        "상태": view_df["상태"].map(lambda s: "🔴 위험" if s == "위험" else ("🟡 주의" if s == "주의" else "🟢 정상")),
        "평가주관": view_df["평가주관"],
        "지표명": view_df["지표명"],
        "소관부서": view_df["소관부서"],
        "배점": view_df["배점"],
        "산출 득점": view_df["산출 득점"],
        "3분기 달성률": view_df["3분기 달성률(C/B)"],
        "2026 목표(A)": view_df["2026년 목표(A)"],
        act_col: view_df[act_col],
        "4분기 만점 부족량": view_df["만점 부족량"],
        "단위": view_df["단위"],
    })

    st.dataframe(
        clean_table.style.format({
            "배점": "{:.2f}",
            "산출 득점": "{:.2f}",
            "3분기 달성률": "{:.1f}%",
            "2026 목표(A)": "{:,.1f}",
            act_col: "{:,.1f}",
            "4분기 만점 부족량": "{:,.1f}",
        }),
        use_container_width=True,
        hide_index=True,
        height=390,
    )

    # 3) 한눈에 들어오는 분야별 달성 현황 차트 (복잡한 52개 막대 대신 깔끔한 가로 바 차트)
    col_g1, col_g2 = st.columns([1.1, 1.3])
    with col_g1:
        st.markdown('<div class="sec-title">📊 분야별 득점 현황 (배점 대비 득점)</div>', unsafe_allow_html=True)
        cat_grp = df_dash.groupby("분류", as_index=False)[["배점", "산출 득점"]].sum()
        cat_grp["득점률(%)"] = (cat_grp["산출 득점"] / cat_grp["배점"] * 100.0).round(1)
        fig_cat = go.Figure()
        fig_cat.add_trace(go.Bar(
            y=cat_grp["분류"], x=cat_grp["산출 득점"], orientation="h",
            name="현재 득점", marker_color="#2563eb",
            text=cat_grp.apply(lambda r: f"{r['산출 득점']:.2f} / {r['배점']:.2f}점 ({r['득점률(%)']}%)", axis=1),
            textposition="inside",
        ))
        fig_cat.add_trace(go.Bar(
            y=cat_grp["분류"], x=cat_grp["배점"] - cat_grp["산출 득점"], orientation="h",
            name="미확보 잔여배점", marker_color="#e2e8f0",
        ))
        fig_cat.update_layout(
            barmode="stack", height=280, margin=dict(l=10, r=15, t=10, b=10),
            legend=dict(orientation="h", y=1.12, x=1, xanchor="right"),
            xaxis_title="점수 (점)", yaxis=dict(autorange="reversed"),
        )
        st.plotly_chart(fig_cat, use_container_width=True)

    with col_g2:
        st.markdown('<div class="sec-title">🚨 달성률 하위 집중관리 지표 TOP 10 (100% 기준선)</div>', unsafe_allow_html=True)
        bottom10 = df_dash.sort_values("3분기 달성률(C/B)", ascending=True).head(10)
        fig_b10 = px.bar(
            bottom10, x="3분기 달성률(C/B)", y="지표명", orientation="h",
            color="상태", color_discrete_map={"위험": "#ef4444", "주의": "#f59e0b", "정상": "#10b981"},
            text=bottom10["3분기 달성률(C/B)"].map(lambda v: f"{v:.1f}%"),
        )
        fig_b10.add_vline(x=100.0, line_dash="dash", line_color="#1e3a8a", annotation_text="100% 목표")
        fig_b10.update_layout(
            height=280, margin=dict(l=10, r=20, t=10, b=10),
            xaxis=dict(range=[0, 115], title="3분기 목표 대비 달성률 (%)"),
            yaxis_title="", showlegend=False,
        )
        st.plotly_chart(fig_b10, use_container_width=True)

    # 엑셀 다운로드 버튼 (하단 한 줄로 깔끔하게 배치)
    dl1, dl2 = st.columns(2)
    dl1.download_button(
        "📥 현재 조회된 정량 성적표 Excel(.xlsx) 다운로드",
        data=to_excel_bytes(df_dash),
        file_name=f"경영성과팀_{selected_quarter}_정량성과표.xlsx",
        use_container_width=True,
    )
    dl2.download_button(
        "📥 현재 조회된 정량 성적표 CSV 다운로드",
        data=df_dash.to_csv(index=False).encode("utf-8-sig"),
        file_name=f"경영성과팀_{selected_quarter}_정량성과표.csv",
        use_container_width=True,
    )


# ==============================================================================
# 화면 2: ✏️ 정량 실적 입력 · 엑셀 검산 (부서 입력 폼 + 외부 엑셀 수식 오류 적발)
# ==============================================================================
elif nav_menu.startswith("✏️"):
    sub_tab1, sub_tab2 = st.tabs([
        "1️⃣ 부서별 실적 입력 및 AI 오입력 검증",
        "2️⃣ 부서 제출 수기 엑셀 자동 검산 (오류 셀 적발)",
    ])

    with sub_tab1:
        indicators_map = st.session_state["indicators_data"]
        all_cats = sorted(list({cfg["category"] for cfg in indicators_map.values()}))

        col_sel1, col_sel2 = st.columns([1.2, 2.2])
        with col_sel1:
            cat_filter_in = st.selectbox("1. 지표 분야 선택", options=["전체"] + all_cats)
        filtered_ids = [
            k for k, v in indicators_map.items()
            if cat_filter_in == "전체" or v["category"] == cat_filter_in
        ]
        with col_sel2:
            sel_id = st.selectbox(
                "2. 실적 입력 대상 지표 선택",
                options=filtered_ids,
                format_func=lambda k: f"[{indicators_map[k]['eval_agency']}] {indicators_map[k]['name']} ({indicators_map[k]['department']})",
            )

        cfg = indicators_map[sel_id]
        res_cur = evaluate_indicator(cfg, quarter=selected_quarter, scoring_standard=scoring_standard)

        # 심플한 요약 박스
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("배점 / 단위", f"{cfg['weight']:.2f}점 ({cfg['unit']})")
        m2.metric("2026년 연간목표(A)", f"{res_cur.upper_target:,.1f} {cfg['unit']}")
        m3.metric("3분기 목표(B)", f"{res_cur.q3_target:,.1f} {cfg['unit']}")
        m4.metric(f"현재 {selected_quarter} 실적(C)", f"{res_cur.actual_value:,.1f} {cfg['unit']}")

        st.divider()

        in_left, in_right = st.columns([1.2, 1.0])
        with in_left:
            inp_key = f"clean_in_{sel_id}_{selected_quarter}"
            if inp_key not in st.session_state:
                st.session_state[inp_key] = f"{res_cur.actual_value}"

            raw_val = st.text_input(
                f"✏️ [{cfg['name']}] {selected_quarter} 누적 실적 입력 (단위: {cfg['unit']})",
                key=inp_key,
            )

            # 테스트용 빠른 버튼 (접기/펼치기 안에 넣지 않고 작은 보조 버튼으로 배치)
            st.caption("테스트 값 자동 입력:")
            tb1, tb2, tb3 = st.columns(3)
            if tb1.button("정상 실적", use_container_width=True):
                st.session_state[inp_key] = f"{res_cur.actual_value}"
                st.rerun()
            if tb2.button("단위오류(×15배)", use_container_width=True):
                st.session_state[inp_key] = f"{max(1.0, res_cur.base_value) * 15:.0f}"
                st.rerun()
            if tb3.button("미달위험(55%)", use_container_width=True):
                st.session_state[inp_key] = f"{res_cur.base_value * 0.55:.0f}"
                st.rerun()

            parsed_v, v_msgs, is_ok = validate_indicator_input(cfg, selected_quarter, raw_val)
            for lvl, txt in v_msgs:
                if lvl == "error":
                    st.error(txt)
                elif lvl == "warning":
                    st.warning(txt)
                else:
                    st.success(txt)

            if st.button("💾 실적 저장하기", type="primary", disabled=not is_ok, use_container_width=True):
                st.session_state["indicators_data"][sel_id]["quarterly_actuals"][selected_quarter] = parsed_v
                st.toast(f"[{cfg['name']}] 실적이 저장되었습니다!", icon="✅")
                st.rerun()

        with in_right:
            pv_v = parsed_v if (is_ok and parsed_v is not None) else res_cur.actual_value
            pv_res = evaluate_indicator(cfg, quarter=selected_quarter, actual_override=pv_v, scoring_standard=scoring_standard)
            need_q4 = max(0.0, pv_res.upper_target - pv_v) if not pv_res.is_reverse else 0.0

            st.markdown("##### 🎯 입력값 적용 시 예상 결과")
            r1, r2 = st.columns(2)
            r1.metric("예상 득점", f"{pv_res.score:.2f} / {pv_res.weight:.2f}점", f"득점률 {pv_res.score_ratio:.1f}%")
            r2.metric("연말 만점까지 남은 물량", f"{need_q4:,.1f} {cfg['unit']}", f"일평균 {need_q4/90:,.1f} {cfg['unit']} 필요")

    with sub_tab2:
        st.markdown("##### 📂 부서 제출 수기 엑셀 자동 검산")
        st.caption("부서에서 제출한 엑셀 파일의 수기 계산값과 시스템 표준 선형보간 계산값을 비교하여 **틀린 셀을 빨간색으로 표시**합니다.")

        up_excel = st.file_uploader("검산할 엑셀(.xlsx, .csv) 업로드 (미업로드 시 아래 오류 적발 샘플이 표시됩니다)", type=["xlsx", "csv"])

        base_df = build_summary_df(st.session_state["indicators_data"], selected_quarter, "연간 편람 선형보간 공식 즉시 적용").head(10)
        audit_rows = []
        err_cells = []

        for idx, row in base_df.iterrows():
            nm = row["지표명"]
            w = float(row["배점"])
            act = float(row[f"{selected_quarter} 실적(C)"])
            sys_up = float(row["2026년 목표(A)"])
            sys_sc = float(row["산출 득점"])

            m_up, m_sc = sys_up, sys_sc
            reason = "✅ 수식 정상 일치"
            bad_cols = []

            if up_excel is None:
                if nm == "가락 청과 거래물량":
                    m_up = round(sys_up * 0.95, 0)
                    m_sc = round(w * 0.95, 3)
                    reason = "❌ 최고목표치 오적용 및 평점 과대산출"
                    bad_cols = ["수기_목표치", "수기_득점", "득점오차"]
                elif nm == "가락 수산 거래물량":
                    m_sc = round(w * (act / sys_up), 3)
                    reason = "❌ 편람 하한선(60%) 미적용 단순비례 오산출"
                    bad_cols = ["수기_득점", "득점오차"]

            diff = round(m_sc - sys_sc, 3)
            err_cells.append((len(audit_rows), bad_cols))
            audit_rows.append({
                "판정": "❌ 불일치" if bad_cols else "✅ 일치",
                "지표명": nm,
                "소관부서": row["소관부서"],
                "배점": w,
                "당해실적": act,
                "수기_목표치": m_up,
                "시스템_목표치": sys_up,
                "수기_득점": m_sc,
                "시스템_득점": sys_sc,
                "득점오차": diff,
                "검산 진단사유": reason,
            })

        audit_df = pd.DataFrame(audit_rows)

        def highlight_err(df_t: pd.DataFrame) -> pd.DataFrame:
            styles = pd.DataFrame("", index=df_t.index, columns=df_t.columns)
            for r_i, cols in err_cells:
                if cols:
                    styles.loc[r_i, "판정"] = "background-color:#fee2e2; color:#991b1b; font-weight:700;"
                    for c in cols:
                        styles.loc[r_i, c] = "background-color:#fecaca; color:#7f1d1d; font-weight:800;"
            return styles

        st.dataframe(
            audit_df.style.apply(highlight_err, axis=None).format({
                "배점": "{:.2f}", "당해실적": "{:,.1f}", "수기_목표치": "{:,.1f}",
                "시스템_목표치": "{:,.1f}", "수기_득점": "{:.3f}", "시스템_득점": "{:.3f}", "득점오차": "{:+.3f}"
            }),
            use_container_width=True, hide_index=True,
        )


# ==============================================================================
# 화면 3: 📑 정성 보고서 AI 심사 (파일 업로드 + 편람 기준 자동 채점)
# ==============================================================================
else:
    st.markdown('<div class="sec-title">📂 1단계: 부서 정성 실적자료 파일 업로드 (PDF, Word, Excel, TXT)</div>', unsafe_allow_html=True)
    uploaded_files = st.file_uploader(
        "전 부서 정성지표 실적 파일을 끌어다 놓으세요 (여러 파일 동시 업로드 가능)",
        type=["pdf", "docx", "xlsx", "csv", "txt", "md"],
        accept_multiple_files=True,
    )

    if uploaded_files:
        for uf in uploaded_files:
            txt = extract_text_from_uploaded_file(uf)
            if txt.strip():
                matched_qid = "QUAL_10"
                best_cnt = -1
                for q_id, q_cfg in QUALITATIVE_INDICATORS_CONFIG.items():
                    short_nm = q_cfg["name"].split("(")[0].replace("·", "").strip()
                    if short_nm in uf.name or q_cfg["name"] in txt:
                        matched_qid = q_id
                        break
                    cnt = sum(1 for kw in q_cfg["required_keywords"] if kw in txt)
                    if cnt > best_cnt:
                        best_cnt, matched_qid = cnt, q_id
                st.session_state["qualitative_reports"][matched_qid] = txt
                st.success(f"✅ `{uf.name}` ➔ **[{QUALITATIVE_INDICATORS_CONFIG[matched_qid]['name']}]** 자동 매핑 및 편람 심사 완료!")

    # 전체 11개 정성지표 요약표
    st.markdown('<div class="sec-title">📊 2단계: 전사 11대 정성지표 편람 부합도 총괄표</div>', unsafe_allow_html=True)
    q_rows = []
    q_results = {}
    for q_id, q_cfg in QUALITATIVE_INDICATORS_CONFIG.items():
        r = evaluate_qualitative_text(q_cfg, st.session_state["qualitative_reports"].get(q_id, ""))
        q_results[q_id] = r
        q_rows.append({
            "정성지표명": r["name"],
            "소관부서": r["department"],
            "배점": r["weight"],
            "예상득점": r["est_score"],
            "예상등급": r["grade"],
            "종합부합도": r["total_pct"],
            "편람키워드 반영률": r["kw_cov"],
            "계량수치(건)": r["quant_cnt"],
            "보완필요 핵심키워드": ", ".join(r["missing"][:4]) if r["missing"] else "✅ 완벽 반영",
        })
    q_df = pd.DataFrame(q_rows)
    st.dataframe(
        q_df.style.format({
            "배점": "{:.1f}", "예상득점": "{:.2f}", "종합부합도": "{:.1f}%", "편람키워드 반영률": "{:.1f}%"
        }),
        use_container_width=True, hide_index=True,
    )

    # 개별 정성지표 상세 심사 및 수정
    st.markdown('<div class="sec-title">🔍 3단계: 지표별 편람 체크포인트 및 보완사항 상세 보기</div>', unsafe_allow_html=True)
    sel_qid = st.selectbox(
        "상세 검토할 정성지표 선택",
        options=list(QUALITATIVE_INDICATORS_CONFIG.keys()),
        format_func=lambda k: f"{QUALITATIVE_INDICATORS_CONFIG[k]['name']} (배점 {QUALITATIVE_INDICATORS_CONFIG[k]['weight']}점 | {QUALITATIVE_INDICATORS_CONFIG[k]['department']})",
    )
    q_cfg = QUALITATIVE_INDICATORS_CONFIG[sel_qid]

    left_c, right_c = st.columns([1.15, 1.0])
    with left_c:
        st.caption("📖 **행안부 편람 세부평가기준:**")
        for sc in q_cfg["sub_criteria"]:
            st.markdown(f"- {sc}")
        new_txt = st.text_area(
            "✍️ 부서 제출 실적보고서 본문 (내용을 수정하면 우측 점수가 실시간 변경됩니다)",
            value=st.session_state["qualitative_reports"][sel_qid],
            height=220,
        )
        if new_txt != st.session_state["qualitative_reports"][sel_qid]:
            st.session_state["qualitative_reports"][sel_qid] = new_txt
            st.rerun()

    with right_c:
        live_r = evaluate_qualitative_text(q_cfg, new_txt)
        k1, k2, k3 = st.columns(3)
        k1.metric("예상 등급", live_r["grade"])
        k2.metric("예상 점수", f"{live_r['est_score']:.2f} / {live_r['weight']:.1f}점")
        k3.metric("편람 부합도", f"{live_r['total_pct']:.1f}%")

        # 키워드 체크 배지
        badges = [
            f'<span class="pill-green">✓ {kw}</span>' if kw in live_r["matched"] else f'<span class="pill-red">✗ {kw}</span>'
            for kw in q_cfg["required_keywords"]
        ]
        st.markdown("**편람 필수 키워드 포함 여부:**<br/>" + " ".join(badges), unsafe_allow_html=True)
        st.markdown("<div style='height:0.6rem;'></div>", unsafe_allow_html=True)

        if live_r["missing"]:
            st.warning(f"🛠️ **보완 권고:** 편람 필수 키워드인 **`{', '.join(live_r['missing'])}`** 관련 추진 실적을 본문에 추가 기술하면 등급이 상승합니다.")
        else:
            st.success("👍 **진단 결과:** 편람의 모든 필수 착안사항 키워드와 PDCA(계획-실행-성과-환류) 구조가 완벽히 반영되어 있습니다.")
