"""
================================================================================
서울시농수산식품공사 「2027년도 지방공기업 경영평가」
분기별 정량 성과관리 및 자동 검산 웹 애플리케이션 (app.py)
================================================================================
[종속성 설치 (requirements.txt)]
pip install streamlit>=1.35.0 pandas>=2.0.0 plotly>=5.20.0 openpyxl>=3.1.0

[실행 방법]
streamlit run app.py

[모듈화 및 유지보수 안내]
- 추후 신규 정량지표가 추가되거나 행정안전부 지방공기업 경영평가 편람의 배점·목표비율·
  역산 여부(하향지표) 등이 변경될 경우, 아래 `INDICATORS_CONFIG` 딕셔너리(JSON 구조)만
  수정하거나 JSON 파일로 불러오면 백엔드 계산 엔진, UI 입력 폼, 대시보드, 시뮬레이터,
  엑셀 자동 검산 모듈 전체에 즉시 반영되도록 설계되어 있습니다.
================================================================================
"""

from __future__ import annotations

import io
import json
from copy import deepcopy
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ==============================================================================
# 0. 페이지 기본 설정 및 커스텀 스타일링
# ==============================================================================
st.set_page_config(
    page_title="서울시농수산식품공사 | 2027 경영평가 정량 성과관리·자동검산 시스템",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

CUSTOM_CSS = """
<style>
    /* 전체 폰트 및 배경 톤 정돈 */
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2.5rem;
    }
    /* 상단 헤더 배너 */
    .hero-banner {
        background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 60%, #1d4ed8 100%);
        color: #ffffff;
        padding: 1.35rem 1.75rem;
        border-radius: 14px;
        margin-bottom: 1.25rem;
        box-shadow: 0 6px 20px rgba(15, 23, 42, 0.14);
        display: flex;
        justify-content: space-between;
        align-items: center;
        flex-wrap: wrap;
        gap: 1rem;
    }
    .hero-title {
        font-size: 1.45rem;
        font-weight: 800;
        letter-spacing: -0.02em;
        margin: 0 0 0.3rem 0;
    }
    .hero-subtitle {
        font-size: 0.92rem;
        color: #dbeafe;
        margin: 0;
        font-weight: 400;
    }
    .hero-badge {
        background: rgba(255, 255, 255, 0.14);
        border: 1px solid rgba(255, 255, 255, 0.28);
        padding: 0.45rem 0.9rem;
        border-radius: 999px;
        font-size: 0.84rem;
        font-weight: 600;
        color: #f8fafc;
        backdrop-filter: blur(6px);
    }
    /* KPI 카드 스타일 */
    .kpi-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 1.15rem 1.25rem;
        box-shadow: 0 2px 8px rgba(15, 23, 42, 0.04);
        height: 100%;
    }
    .kpi-label {
        font-size: 0.83rem;
        font-weight: 600;
        color: #64748b;
        margin-bottom: 0.35rem;
        text-transform: uppercase;
        letter-spacing: 0.02em;
    }
    .kpi-value {
        font-size: 1.75rem;
        font-weight: 800;
        color: #0f172a;
        line-height: 1.2;
        margin-bottom: 0.35rem;
    }
    .kpi-sub {
        font-size: 0.82rem;
        color: #475569;
    }
    /* 상태 배지 */
    .badge-red {
        display: inline-block;
        background: #fee2e2;
        color: #b91c1c;
        border: 1px solid #fecaca;
        font-weight: 700;
        padding: 0.18rem 0.6rem;
        border-radius: 999px;
        font-size: 0.78rem;
        margin-right: 0.35rem;
    }
    .badge-yellow {
        display: inline-block;
        background: #fef3c7;
        color: #b45309;
        border: 1px solid #fde68a;
        font-weight: 700;
        padding: 0.18rem 0.6rem;
        border-radius: 999px;
        font-size: 0.78rem;
        margin-right: 0.35rem;
    }
    .badge-green {
        display: inline-block;
        background: #dcfce7;
        color: #15803d;
        border: 1px solid #bbf7d0;
        font-weight: 700;
        padding: 0.18rem 0.6rem;
        border-radius: 999px;
        font-size: 0.78rem;
        margin-right: 0.35rem;
    }
    .badge-blue {
        display: inline-block;
        background: #dbeafe;
        color: #1d4ed8;
        border: 1px solid #bfdbfe;
        font-weight: 700;
        padding: 0.18rem 0.6rem;
        border-radius: 999px;
        font-size: 0.78rem;
    }
    /* 지표 메타 카드 */
    .meta-card {
        background: #f8fafc;
        border: 1px solid #cbd5e1;
        border-left: 5px solid #1d4ed8;
        border-radius: 10px;
        padding: 1rem 1.2rem;
        margin-bottom: 1rem;
    }
    .meta-title {
        font-size: 1.05rem;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 0.4rem;
    }
    .meta-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
        gap: 0.65rem;
        margin-top: 0.65rem;
    }
    .meta-item {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 0.55rem 0.75rem;
    }
    .meta-item-label {
        font-size: 0.75rem;
        color: #64748b;
        font-weight: 600;
    }
    .meta-item-val {
        font-size: 0.96rem;
        color: #0f172a;
        font-weight: 700;
        margin-top: 0.15rem;
    }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# ==============================================================================
# 1. 샘플 데이터셋 및 지표 마스터 설정 (JSON/Dict 기반 모듈화)
# ==============================================================================
# ※ 추후 지표가 추가되거나 편람 배점, 3개년 실적, 목표산식(목표부여 vs 목표비율 달성도),
#   하향지표(역산식 적용 여부)가 변경될 때 아래 INDICATORS_CONFIG 딕셔너리만 수정하면
#   애플리케이션 전체(계산 엔진, 입력 검증, 대시보드, 시뮬레이터, 엑셀 검산)에 자동 반영됩니다.

QUARTERS: List[str] = ["1분기", "2분기", "3분기", "4분기(최종)"]

# 분기별 권장 진도율(누적 기준) - 추세 이상 감지 및 가이드라인에 활용
QUARTER_PROGRESS_RATIO: Dict[str, float] = {
    "1분기": 0.25,
    "2분기": 0.50,
    "3분기": 0.75,
    "4분기(최종)": 1.00,
}

INDICATORS_CONFIG: Dict[str, Dict[str, Any]] = {
    "IND_01": {
        "id": "IND_01",
        "name": "가락 청과 거래실적",
        "department": "가락시장본부 청과부류팀",
        "weight": 3.5,
        "unit": "톤",
        "eval_method": "standard_3yr",  # 직전 3개년 평균 대비 ±10% 목표부여(편람 표준)
        "is_reverse": False,            # False: 상향지표(실적이 클수록 고득점), True: 하향지표(역산식)
        "history_3yr": {
            "2023년": 2_190_000.0,
            "2024년": 2_220_000.0,
            "2025년": 2_260_290.0,
        },
        "base_override": 2_223_430.0,   # 직전 3개년 단순 평균: 2,223,430톤 (2026 목표치 2,445,773톤의 100% 기준선)
        "prev_year_actual": 2_260_290.0,
        "prev_year_cumulative": {
            "1분기": 550_000.0,
            "2분기": 1_130_000.0,
            "3분기": 1_685_000.0,
            "4분기(최종)": 2_260_290.0,
        },
        "quarterly_actuals": {
            "1분기": 556_000.0,
            "2분기": 1_145_000.0,
            "3분기": 1_701_149.0,       # ★ 2026년 3분기 경영실적 점검 실제 실적 (연간 달성률 76.5%, 분기목표 대비 92.7%)
            "4분기(최종)": 2_340_000.0,
        },
        "formula_desc": (
            "기준치(3개년평균) = 2,223,430톤 | "
            "최고목표(2026 목표, 110%) = 2,445,773톤 | "
            "최저목표(하한, 90%) = 2,001,087톤 | "
            "선형보간법: 배점(3.5) × [0.6 + 0.4 × ((실적 - 최저) / (최고 - 최저))]"
        ),
    },
    "IND_02": {
        "id": "IND_02",
        "name": "강서 청과 거래실적",
        "department": "강서지사 시장관리팀",
        "weight": 1.0,
        "unit": "톤",
        "eval_method": "standard_3yr",
        "is_reverse": False,
        "history_3yr": {
            "2023년": 595_000.0,
            "2024년": 605_000.0,
            "2025년": 617_476.5,
        },
        "base_override": 605_825.5,     # 직전 3개년 단순 평균: 605,825.5톤 (2026 목표치 666,408톤의 100% 기준선)
        "prev_year_actual": 617_476.5,
        "prev_year_cumulative": {
            "1분기": 150_000.0,
            "2분기": 308_000.0,
            "3분기": 458_000.0,
            "4분기(최종)": 617_476.5,
        },
        "quarterly_actuals": {
            "1분기": 152_000.0,
            "2분기": 311_000.0,
            "3분기": 463_957.0,         # ★ 2026년 3분기 경영실적 점검 실제 실적 (연간 달성률 76.6%, 분기목표 대비 92.8%)
            "4분기(최종)": 635_000.0,
        },
        "formula_desc": (
            "기준치(3개년평균) = 605,825.5톤 | "
            "최고목표(2026 목표, 110%) = 666,408톤 | "
            "최저목표(하한, 90%) = 545,243톤 | "
            "선형보간법: 배점(1.0) × [0.6 + 0.4 × ((실적 - 최저) / (최고 - 최저))]"
        ),
    },
    "IND_03": {
        "id": "IND_03",
        "name": "가락 수산 거래실적",
        "department": "가락시장본부 수산부류팀",
        "weight": 0.5,
        "unit": "톤",
        "eval_method": "standard_3yr",
        "is_reverse": False,
        "history_3yr": {
            "2023년": 76_500.0,
            "2024년": 78_200.0,
            "2025년": 79_668.1,
        },
        "base_override": 78_122.7,      # 직전 3개년 단순 평균: 78,122.7톤 (2026 목표치 85,935톤의 100% 기준선)
        "prev_year_actual": 79_668.1,
        "prev_year_cumulative": {
            "1분기": 19_500.0,
            "2분기": 39_800.0,
            "3분기": 58_200.0,
            "4분기(최종)": 79_668.1,
        },
        "quarterly_actuals": {
            "1분기": 18_200.0,
            "2분기": 36_900.0,
            "3분기": 54_271.0,          # ★ 2026년 3분기 경영실적 점검 실제 실적 (연간 달성률 69.5%, 분기목표 대비 84.2%)
            "4분기(최종)": 76_000.0,
        },
        "formula_desc": (
            "기준치(3개년평균) = 78,122.7톤 | "
            "최고목표(2026 목표, 110%) = 85,935톤 | "
            "최저목표(하한, 90%) = 70,310톤 | "
            "선형보간법: 배점(0.5) × [0.6 + 0.4 × ((실적 - 최저) / (최고 - 최저))]"
        ),
    },
    "IND_04": {
        "id": "IND_04",
        "name": "학교급식 공급실적",
        "department": "서울친환경유통센터 급식사업팀",
        "weight": 1.0,
        "unit": "톤",
        "eval_method": "standard_3yr",
        "is_reverse": False,
        "history_3yr": {
            "2023년": 25_200.0,
            "2024년": 26_100.0,
            "2025년": 26_866.5,
        },
        "base_override": 26_055.5,      # 직전 3개년 단순 평균: 26,055.5톤 (2026 목표치 28,661톤의 100% 기준선)
        "prev_year_actual": 26_866.5,
        "prev_year_cumulative": {
            "1분기": 6_200.0,
            "2분기": 13_100.0,
            "3분기": 19_500.0,
            "4분기(최종)": 26_866.5,
        },
        "quarterly_actuals": {
            "1분기": 5_600.0,
            "2분기": 11_800.0,
            "3분기": 17_380.0,          # ★ 2026년 3분기 경영실적 점검 실제 실적 (연간 달성률 66.7%, 분기목표 대비 80.9%)
            "4분기(최종)": 24_500.0,
        },
        "formula_desc": (
            "기준치(3개년평균) = 26,055.5톤 | "
            "최고목표(2026 목표, 110%) = 28,661톤 | "
            "최저목표(하한, 90%) = 23,450톤 | "
            "선형보간법: 배점(1.0) × [0.6 + 0.4 × ((실적 - 최저) / (최고 - 최저))]"
        ),
    },
    "IND_05": {
        "id": "IND_05",
        "name": "혁신제품 구매실적",
        "department": "기획조정실 재무회계팀",
        "weight": 0.3,
        "unit": "백만원",
        "eval_method": "standard_3yr",
        "is_reverse": False,
        "history_3yr": {
            "2023년": 540.0,
            "2024년": 560.0,
            "2025년": 582.7,
        },
        "base_override": 560.9,         # 직전 3개년 단순 평균: 560.9백만원 (2026 목표치 617백만원의 100% 기준선)
        "prev_year_actual": 582.7,
        "prev_year_cumulative": {
            "1분기": 130.0,
            "2분기": 280.0,
            "3분기": 420.0,
            "4분기(최종)": 582.7,
        },
        "quarterly_actuals": {
            "1분기": 145.0,
            "2분기": 320.0,
            "3분기": 568.0,             # ★ 2026년 3분기 경영실적 점검 실제 실적 (연간 달성률 101.3%, 분기목표 대비 122.7%)
            "4분기(최종)": 625.0,
        },
        "formula_desc": (
            "기준치(3개년평균) = 560.9백만원 | "
            "최고목표(2026 목표, 110%) = 617백만원 | "
            "최저목표(하한, 90%) = 504.8백만원 | "
            "선형보간법: 배점(0.3) × [0.6 + 0.4 × ((실적 - 최저) / (최고 - 최저))]"
        ),
    },
}


# ==============================================================================
# 2. 평가편람 표준 계산 엔진 (핵심 백엔드 알고리즘)
# ==============================================================================
@dataclass
class EvaluationResult:
    """단일 지표의 평가편람 산출 결과를 담는 데이터 클래스"""
    indicator_id: str
    indicator_name: str
    department: str
    weight: float
    unit: str
    is_reverse: bool
    base_value: float          # 기준치(Base): 직전 3개년 실적 단순 평균
    upper_target: float        # 최고목표치(상한): 상향지표 기준치×1.10 / 하향지표 기준치×0.90
    lower_target: float        # 최저목표치(하한): 상향지표 기준치×0.90 / 하향지표 기준치×1.10
    actual_value: float        # 당해(선택 분기) 실적
    achievement_rate: float    # 기준치 대비 달성률 (%)
    max_target_rate: float     # 최고목표치(만점 기준) 대비 달성률 (%)
    score: float               # 선형보간법 산출 평점 (배점 반영)
    score_ratio: float         # 배점 대비 득점률 (60.0% ~ 100.0%)
    grade_label: str           # 평어 등급 구간 ('수(만점)', '우/양/가(보간)', '마(하한)')
    status_level: str          # 관제 상태 ('정상', '주의', '위험')
    status_badge_reason: str   # 상태 판정 사유


def compute_base_value(ind_cfg: Dict[str, Any]) -> float:
    """
    1. 기준치(Base) 산출: 직전 3개년 실적의 단순 평균
       - 설정에 history_3yr가 있으면 3개년 값의 산술평균을 계산하며,
         명시적 base_override가 있으면 이를 검증용 기본값으로 사용합니다.
    """
    history = ind_cfg.get("history_3yr", {})
    if history and len(history) > 0:
        vals = [float(v) for v in history.values()]
        avg_val = sum(vals) / len(vals)
        return round(avg_val, 4)
    return float(ind_cfg.get("base_override", 0.0))


def compute_target_bounds(base_value: float, is_reverse: bool = False) -> Tuple[float, float]:
    """
    2. 목표치 구간 산출:
       [일반 상향지표 (is_reverse=False)]
         - 최고목표치(상한, 100% 만점 기준) = 기준치 × 1.10 (110%)
         - 최저목표치(하한, 60% '마' 등급 기준) = 기준치 × 0.90 (90%)
       [하향지표 - 온실가스·에너지 감축 등 (is_reverse=True)]
         - 실적이 낮을수록 우수하므로 역산 목표 구간 적용:
         - 최고목표치(상한, 만점 감축목표) = 기준치 × 0.90 (90% 이하 배출 시 만점)
         - 최저목표치(하한, 마 등급 한계선) = 기준치 × 1.10 (110% 이상 배출 시 60%)
    """
    if not is_reverse:
        upper_target = round(base_value * 1.10, 4)
        lower_target = round(base_value * 0.90, 4)
    else:
        upper_target = round(base_value * 0.90, 4)
        lower_target = round(base_value * 1.10, 4)
    return upper_target, lower_target


def calculate_linear_interpolation_score(
    actual: float,
    base_value: float,
    weight: float,
    is_reverse: bool = False,
) -> Tuple[float, float, str, float, float]:
    """
    3. 평점 산출 공식 (행정안전부 지방공기업 경영평가 편람 선형보간법):
       - 상향지표(is_reverse=False):
         * 당해 실적 >= 최고목표치(기준치×1.10): 배점 만점 부여 (100% 득점)
         * 당해 실적 <= 최저목표치(기준치×0.90): 배점의 60% 부여 (평어 '마' 등급 기준선)
         * 최저목표치 < 당해 실적 < 최고목표치:
           득점 = 배점 × [ 0.6 + 0.4 × ( (실적 - 최저목표치) / (최고목표치 - 최저목표치) ) ]

       - 하향지표(is_reverse=True, 온실가스·에너지 등 역산식 플래그 적용):
         * 당해 실적 <= 최고목표치(기준치×0.90): 배점 만점 부여 (100% 득점)
         * 당해 실적 >= 최저목표치(기준치×1.10): 배점의 60% 부여 (평어 '마' 등급 기준선)
         * 최고목표치 < 당해 실적 < 최저목표치:
           득점 = 배점 × [ 0.6 + 0.4 × ( (최저목표치 - 실적) / (최저목표치 - 최고목표치) ) ]
           (수치적으로 (실적 - 최저목표치) / (최고목표치 - 최저목표치) 와 완벽히 동일)
    """
    upper_target, lower_target = compute_target_bounds(base_value, is_reverse=is_reverse)

    if base_value <= 0:
        return 0.0, 0.0, "산출불가", upper_target, lower_target

    if not is_reverse:
        if actual >= upper_target:
            ratio = 1.0
            grade = "수 (만점·100%)"
        elif actual <= lower_target:
            ratio = 0.6
            grade = "마 (하한·60%)"
        else:
            ratio = 0.6 + 0.4 * ((actual - lower_target) / (upper_target - lower_target))
            grade = "보간구간 (60~100%)"
    else:
        # 하향지표(온실가스·에너지 감축 등): 낮을수록 고득점
        if actual <= upper_target:
            ratio = 1.0
            grade = "수 (만점·100%)"
        elif actual >= lower_target:
            ratio = 0.6
            grade = "마 (하한·60%)"
        else:
            ratio = 0.6 + 0.4 * ((lower_target - actual) / (lower_target - upper_target))
            grade = "보간구간 (60~100%)"

    raw_score = weight * ratio
    return round(raw_score, 4), round(ratio * 100.0, 2), grade, upper_target, lower_target


def evaluate_indicator(
    ind_cfg: Dict[str, Any],
    quarter: str,
    actual_override: Optional[float] = None,
) -> EvaluationResult:
    """지표 설정과 선택된 분기(또는 직접 입력한 실적값)를 받아 표준 평가 결과를 반환합니다."""
    base_val = compute_base_value(ind_cfg)
    weight = float(ind_cfg["weight"])
    is_reverse = bool(ind_cfg.get("is_reverse", False))

    if actual_override is not None:
        actual = float(actual_override)
    else:
        actual = float(ind_cfg["quarterly_actuals"].get(quarter, 0.0))

    score, score_ratio, grade_label, upper_target, lower_target = (
        calculate_linear_interpolation_score(
            actual=actual,
            base_value=base_val,
            weight=weight,
            is_reverse=is_reverse,
        )
    )

    # 달성률 계산
    # - 기준치 대비 달성률(%): 상향지표는 (실적 / 기준치)*100, 하향지표는 (기준치 / 실적)*100
    # - 최고목표(만점선) 대비 달성률(%): 상향지표는 (실적 / 최고목표)*100
    if base_val > 0:
        if not is_reverse:
            achievement_rate = round((actual / base_val) * 100.0, 2)
            max_target_rate = round((actual / upper_target) * 100.0, 2) if upper_target > 0 else 0.0
        else:
            achievement_rate = round((base_val / actual) * 100.0, 2) if actual > 0 else 100.0
            max_target_rate = round((upper_target / actual) * 100.0, 2) if actual > 0 else 100.0
    else:
        achievement_rate = 0.0
        max_target_rate = 0.0

    # 관제 상태 판정 (위험 / 주의 / 정상)
    # 1) 누적 달성률(기준치 대비)이 70% 미만이거나 최저목표치 이하(60% 하한선)인 경우 -> 위험(Red)
    # 2) 누적 달성률이 70% 이상 ~ 90% 미만(또는 만점 미달 보간구간 초중반) -> 주의(Yellow)
    # 3) 누적 달성률이 90% 이상(또는 분기 권장 진도율 상회) -> 정상(Green)
    if achievement_rate < 70.0 or score_ratio <= 60.0:
        status_level = "위험"
        if achievement_rate < 70.0:
            status_badge_reason = f"누적 달성률 {achievement_rate:.1f}% (70% 미만 목표미달 위험)"
        else:
            status_badge_reason = f"최저목표치 미달 (평어 '마' 60% 하한 구간)"
    elif achievement_rate < 82.0 or score_ratio < 80.0:
        status_level = "주의"
        status_badge_reason = f"달성률 {achievement_rate:.1f}% (만점 수렴을 위한 4분기 집중 관리 필요)"
    else:
        status_level = "정상"
        status_badge_reason = f"달성률 {achievement_rate:.1f}% (목표 순항 중)"

    return EvaluationResult(
        indicator_id=ind_cfg["id"],
        indicator_name=ind_cfg["name"],
        department=ind_cfg.get("department", "소관부서"),
        weight=weight,
        unit=ind_cfg["unit"],
        is_reverse=is_reverse,
        base_value=base_val,
        upper_target=upper_target,
        lower_target=lower_target,
        actual_value=actual,
        achievement_rate=achievement_rate,
        max_target_rate=max_target_rate,
        score=score,
        score_ratio=score_ratio,
        grade_label=grade_label,
        status_level=status_level,
        status_badge_reason=status_badge_reason,
    )


# ==============================================================================
# 3. AI 입력값 유효성 검증 (Smart Validation Engine)
# ==============================================================================
@dataclass
class ValidationMessage:
    level: str   # "error" | "warning" | "info" | "success"
    title: str
    detail: str


def validate_indicator_input(
    ind_cfg: Dict[str, Any],
    quarter: str,
    raw_input_text: str,
) -> Tuple[Optional[float], List[ValidationMessage], bool]:
    """
    부서 실적 입력값에 대한 AI 스마트 유효성 검증 로직:
    1. 비어있거나 숫자가 아니거나 음수(< 0)일 경우 -> Error (저장 차단)
    2. 단위 오류 감지: 전년 동기(또는 전년도 실적) 대비 10배 이상 크거나 1/10 이하로 작으면
       -> "단위(톤/천원) 오입력 경고" 출력 (Warning)
    3. 추세 이상 감지: 3분기(또는 선택 분기) 누적 달성률이 70% 미만일 경우
       -> "목표 미달 위험 경고" 출력 (Warning)
    """
    messages: List[ValidationMessage] = []

    # 1) 공백 체크
    cleaned = raw_input_text.strip().replace(",", "")
    if cleaned == "":
        messages.append(
            ValidationMessage(
                level="error",
                title="입력값 누락 오류 (Empty Value Error)",
                detail="실적 입력란이 비어 있습니다. 당해 분기 누적 실적을 숫자로 입력해 주십시오.",
            )
        )
        return None, messages, False

    # 숫자 변환 체크
    try:
        val = float(cleaned)
    except ValueError:
        messages.append(
            ValidationMessage(
                level="error",
                title="숫자 형식 오류 (Invalid Number Format)",
                detail=f"입력된 값('{raw_input_text}')은 유효한 숫자가 아닙니다. 쉼표(,)를 제외한 수치만 입력해 주십시오.",
            )
        )
        return None, messages, False

    # 음수 체크
    if val < 0:
        messages.append(
            ValidationMessage(
                level="error",
                title="음수 입력 오류 (Negative Value Error)",
                detail=f"입력된 실적({val:,.2f} {ind_cfg['unit']})이 음수(-)입니다. 정량 거래·공급·구매 실적은 0 이상이어야 합니다.",
            )
        )
        return None, messages, False

    if val == 0:
        messages.append(
            ValidationMessage(
                level="warning",
                title="실적 0 입력 주의",
                detail="입력된 실적이 0입니다. 실적 미집계 상태인지 확인해 주십시오.",
            )
        )

    # 2) 단위 오류 감지 (전년 동기 누적 실적 및 전년도 연간 실적과 비교하여 10배 이상 또는 1/10 이하 감지)
    unit = ind_cfg["unit"]
    prev_sync = float(
        ind_cfg.get("prev_year_cumulative", {}).get(
            quarter, ind_cfg.get("prev_year_actual", compute_base_value(ind_cfg))
        )
    )
    prev_annual = float(ind_cfg.get("prev_year_actual", compute_base_value(ind_cfg)))

    if prev_sync > 0 and val > 0:
        ratio_sync = val / prev_sync
        ratio_annual = val / prev_annual if prev_annual > 0 else ratio_sync
        if ratio_sync >= 10.0 or ratio_annual >= 10.0:
            messages.append(
                ValidationMessage(
                    level="warning",
                    title=f"🚨 단위({unit}/천원·kg·원) 오입력 경고 [전년 대비 10배 이상 과대]",
                    detail=(
                        f"입력값 **{val:,.1f} {unit}**은(는) 전년 동기 실적(**{prev_sync:,.1f} {unit}**) 대비 "
                        f"**{ratio_sync:,.1f}배** (전년 연간 실적 대비 {ratio_annual:,.1f}배) 큽니다. "
                        f"`{unit}` 단위를 `kg`이나 `원` 단위로 환산 착오하여 1,000배 과대 입력하지 않았는지 반드시 확인하십시오!"
                    ),
                )
            )
        elif ratio_sync <= 0.10 or ratio_annual <= 0.10:
            messages.append(
                ValidationMessage(
                    level="warning",
                    title=f"🚨 단위(톤/천원) 오입력 경고 [전년 대비 1/10 이하 과소]",
                    detail=(
                        f"입력값 **{val:,.1f} {unit}**은(는) 전년 동기 실적(**{prev_sync:,.1f} {unit}**)의 "
                        f"**{ratio_sync * 100:.2f}% (1/10 이하)** 수준입니다. "
                        f"`{unit}` 단위를 `천톤`이나 `백만원·억원` 단위로 잘못 축소 입력하지 않았는지 확인하십시오!"
                    ),
                )
            )

    # 3) 추세 이상 감지 (기준치 대비 누적 달성률이 70% 미만일 경우 "목표 미달 위험 경고" 출력)
    eval_res = evaluate_indicator(ind_cfg, quarter=quarter, actual_override=val)
    if eval_res.achievement_rate < 70.0:
        messages.append(
            ValidationMessage(
                level="warning",
                title="⚠️ 추세 이상 감지: 목표 미달 위험 경고 (누적 달성률 70% 미만)",
                detail=(
                    f"[{quarter}] 현재 입력 실적 기준 누적 달성률이 **{eval_res.achievement_rate:.2f}%**로 "
                    f"안전 임계선(70.0%)에 미달합니다. (기준치: {eval_res.base_value:,.1f} {unit}, "
                    f"최저목표치: {eval_res.lower_target:,.1f} {unit}). "
                    f"현 추세 지속 시 연말 평가에서 **평어 '마' 등급(배점의 60%, {eval_res.weight * 0.6:.2f}점)**으로 "
                    f"하락할 위험이 높으므로 부서 비상 대책 수립이 필요합니다."
                ),
            )
        )
    elif eval_res.actual_value < eval_res.lower_target and not eval_res.is_reverse:
        messages.append(
            ValidationMessage(
                level="info",
                title="ℹ️ 최저목표치(90%) 도달 전 구간 안내",
                detail=(
                    f"현재 누적 달성률은 **{eval_res.achievement_rate:.2f}%**로 70% 이상이나, "
                    f"선형보간 가점이 시작되는 최저목표치(**{eval_res.lower_target:,.1f} {unit}**, 90%)에는 "
                    f"**{eval_res.lower_target - eval_res.actual_value:,.1f} {unit}** 부족합니다."
                ),
            )
        )

    if not messages:
        messages.append(
            ValidationMessage(
                level="success",
                title="✅ AI 스마트 유효성 검증 통과 (정상 범위)",
                detail=(
                    f"단위 정합성(전년 동기 대비 {val / prev_sync:.2f}배) 및 누적 달성률({eval_res.achievement_rate:.2f}%)이 "
                    f"모두 정상 범위 내에 있습니다. 예상 산출 평점: **{eval_res.score:.3f}점 / {eval_res.weight:.1f}점**"
                ),
            )
        )

    return val, messages, True


# ==============================================================================
# 4. 세션 상태(Session State) 초기화 및 관리
# ==============================================================================
def init_session_state() -> None:
    if "indicators_data" not in st.session_state:
        st.session_state["indicators_data"] = deepcopy(INDICATORS_CONFIG)
    if "audit_logs" not in st.session_state:
        st.session_state["audit_logs"] = [
            {
                "일시": "2026-10-15 09:00",
                "분기": "3분기",
                "지표명": "시스템 초기 데이터 로드",
                "변경전": "-",
                "변경후": "5개 주요 정량지표 탑재 완료",
                "예상득점": "5.148점 / 6.300점",
            }
        ]


init_session_state()


# ==============================================================================
# 5. 엑셀(CSV/XLSX) 변환 및 외부 수기 엑셀 자동 검산 유틸리티
# ==============================================================================
def build_summary_dataframe(
    indicators_dict: Dict[str, Dict[str, Any]],
    quarter: str,
) -> pd.DataFrame:
    """현재 선택된 분기의 전체 정량지표 달성 현황 및 산출 평점 DataFrame 생성"""
    rows: List[Dict[str, Any]] = []
    for ind_id, cfg in indicators_dict.items():
        res = evaluate_indicator(cfg, quarter=quarter)
        remaining_for_max = (
            max(0.0, res.upper_target - res.actual_value)
            if not res.is_reverse
            else max(0.0, res.actual_value - res.upper_target)
        )
        rows.append(
            {
                "지표코드": res.indicator_id,
                "지표명": res.indicator_name,
                "소관부서": res.department,
                "배점": res.weight,
                "단위": res.unit,
                "지표성격": "하향(역산)" if res.is_reverse else "상향(일반)",
                "기준치(3개년평균)": res.base_value,
                "최저목표치(90%)": res.lower_target,
                "최고목표치(110%)": res.upper_target,
                "전년도실적": float(cfg.get("prev_year_actual", res.base_value)),
                "전년동기실적": float(cfg.get("prev_year_cumulative", {}).get(quarter, res.base_value * 0.75)),
                f"{quarter} 실적": res.actual_value,
                "기준치 대비 달성률(%)": res.achievement_rate,
                "최고목표 대비 달성률(%)": res.max_target_rate,
                "산출 득점": res.score,
                "득점률(%)": res.score_ratio,
                "평어구간": res.grade_label,
                "만점 잔여실적": round(remaining_for_max, 2),
                "상태": res.status_level,
                "상태사유": res.status_badge_reason,
            }
        )
    return pd.DataFrame(rows)


def generate_sample_audit_excel_df(
    indicators_dict: Dict[str, Dict[str, Any]],
    quarter: str,
    inject_human_errors: bool = True,
) -> pd.DataFrame:
    """
    외부 수기 엑셀 업로드 자동 검산 기능을 즉시 테스트할 수 있도록
    수기 엑셀 양식 DataFrame을 생성합니다.
    inject_human_errors=True 인 경우, 현업에서 자주 발생하는
    ① 선형보간 수식 오류, ② 최고/최저목표 참조 왜곡, ③ 단위/반올림 착오를 의도적으로 포함시킵니다.
    """
    df = build_summary_dataframe(indicators_dict, quarter=quarter)
    audit_rows: List[Dict[str, Any]] = []
    for idx, row in df.iterrows():
        base_val = float(row["기준치(3개년평균)"])
        lower_val = float(row["최저목표치(90%)"])
        upper_val = float(row["최고목표치(110%)"])
        actual_val = float(row[f"{quarter} 실적"])
        score_val = float(row["산출 득점"])
        rate_val = float(row["기준치 대비 달성률(%)"])

        error_note = "정상 (수식 일치)"
        if inject_human_errors:
            if row["지표명"] == "가락 청과 거래실적":
                # 오류 유형 1: 최고목표치를 110%가 아닌 105%로 잘못 입력하여 평점이 과대 산출된 수기 엑셀 오류
                upper_val = round(base_val * 1.05, 2)
                score_val = round(row["배점"] * 0.92, 4)
                error_note = "[수기오류 시뮬레이션] 최고목표치(105% 오적용) 및 평점 수식 왜곡"
            elif row["지표명"] == "가락 수산 거래실적":
                # 오류 유형 2: 최저목표치 미달(60% 하한선)인데 단순 비례식(실적/기준치*배점)을 잘못 적용한 오류
                score_val = round(row["배점"] * (actual_val / base_val), 4)
                error_note = "[수기오류 시뮬레이션] 편람 하한선(60%) 미적용·단순비례 오산출"
            elif row["지표명"] == "혁신제품 구매실적":
                # 오류 유형 3: 기준치 대비 달성률 분모를 최고목표치로 잘못 참조한 엑셀 셀 참조 밀림 오류
                rate_val = round((actual_val / upper_val) * 100.0, 2)
                score_val = round(score_val + 0.035, 4)
                error_note = "[수기오류 시뮬레이션] 달성률 분모 셀 참조 왜곡 및 평점 오차"

        audit_rows.append(
            {
                "지표명": row["지표명"],
                "배점": row["배점"],
                "단위": row["단위"],
                "지표성격": row["지표성격"],
                "수기_기준치": base_val,
                "수기_최저목표치": lower_val,
                "수기_최고목표치": upper_val,
                "수기_당해실적": actual_val,
                "수기_달성률(%)": rate_val,
                "수기_산출득점": score_val,
                "비고(수기검증메모)": error_note,
            }
        )
    return pd.DataFrame(audit_rows)


def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "경영평가_정량실적") -> bytes:
    """DataFrame을 메모리의 .xlsx 바이너리로 변환"""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return buffer.getvalue()


# ==============================================================================
# 6. 좌측 사이드바 (분기 선택, 모드 전환, 역산 플래그/설정 관리)
# ==============================================================================
with st.sidebar:
    st.markdown("### 🏛️ 서울시농수산식품공사")
    st.caption("2027년도 지방공기업 경영평가 대응 (2026년 실적)")
    st.divider()

    # [사이드바 필수 요소 1] 분기 선택
    selected_quarter: str = st.selectbox(
        "📅 평가 분기 선택",
        options=QUARTERS,
        index=2,  # 기본값: 3분기 (요구사항의 3분기 실적 입력 및 4분기 잔여 예측에 최적화)
        help=" 조회 및 실적을 입력할 분기를 선택하십시오. 기본값은 3분기입니다.",
    )

    # [사이드바 필수 요소 2] 모드 전환
    operation_mode: str = st.radio(
        "🧭 운영 모드 전환",
        options=[
            "기획팀 총괄 관제 대시보드",
            "부서 실적 입력 폼",
        ],
        index=0,
        help="'부서 실적 입력 폼'은 현업 부서 담당자용 스마트 검증 중심 모드이며, '기획팀 총괄 관제 대시보드'는 전사 KPI 관제 및 시뮬레이션 중심 모드입니다.",
    )

    st.divider()

    # 사이드바 실시간 요약 위젯
    summary_df_sidebar = build_summary_dataframe(st.session_state["indicators_data"], selected_quarter)
    total_weight_sb = summary_df_sidebar["배점"].sum()
    total_score_sb = summary_df_sidebar["산출 득점"].sum()
    score_pct_sb = (total_score_sb / total_weight_sb * 100.0) if total_weight_sb > 0 else 0.0
    danger_cnt_sb = int((summary_df_sidebar["상태"] == "위험").sum())
    warn_cnt_sb = int((summary_df_sidebar["상태"] == "주의").sum())

    st.markdown(f"#### 📌 [{selected_quarter}] 실시간 요약")
    col_sb1, col_sb2 = st.columns(2)
    col_sb1.metric("총 배점", f"{total_weight_sb:.1f}점")
    col_sb2.metric("예상 득점", f"{total_score_sb:.2f}점", f"{score_pct_sb:.1f}%")

    st.markdown(
        f"""
        <div style="margin-top:0.3rem; margin-bottom:0.6rem;">
            <span class="badge-red">위험 {danger_cnt_sb}건</span>
            <span class="badge-yellow">주의 {warn_cnt_sb}건</span>
            <span class="badge-green">정상 {len(summary_df_sidebar) - danger_cnt_sb - warn_cnt_sb}건</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.divider()
    with st.expander("⚙️ 편람 마스터 설정 (JSON 모듈화)", expanded=False):
        st.caption(
            "지표 추가·배점 변경·하향지표(온실가스 등 역산식) 플래그 테스트를 지원합니다."
        )
        if st.button("🔄 초기 샘플 데이터(Mock) 복원", use_container_width=True):
            st.session_state["indicators_data"] = deepcopy(INDICATORS_CONFIG)
            st.toast("초기 5대 정량지표 샘플 데이터로 복원되었습니다.", icon="✅")
            st.rerun()


# ==============================================================================
# 7. 상단 헤더 배너 및 모드 안내
# ==============================================================================
mode_badge_text = (
    "🏢 기획팀 총괄 관제 모드 (전사 모니터링·검산)"
    if operation_mode == "기획팀 총괄 관제 대시보드"
    else "✍️ 현업 부서 실적 입력 모드 (스마트 유효성 검증)"
)

st.markdown(
    f"""
    <div class="hero-banner">
        <div>
            <div class="hero-title">📊 2027년도 지방공기업 경영평가 정량 성과관리 및 자동 검산 시스템</div>
            <p class="hero-subtitle">
                서울시농수산식품공사 | 편람 선형보간법 표준 엔진 · AI 단위/추세 이상 감지 · 외부 엑셀 수식 자동 검산
            </p>
        </div>
        <div class="hero-badge">
            기준 분기: {selected_quarter} &nbsp;|&nbsp; {mode_badge_text}
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# 모드 전환에 따른 상단 컨텍스트 알림바
if operation_mode == "부서 실적 입력 폼":
    st.info(
        f"💡 **[부서 실적 입력 폼 모드]** 현재 **{selected_quarter}** 부서별 실적 등록 및 AI 유효성 검증 화면이 활성화되어 있습니다. "
        "하단 **[탭 1: 부서 실적 입력 및 스마트 검증]**에서 소관 지표를 선택하여 실적을 입력·저장하십시오."
    )
else:
    if danger_cnt_sb > 0 or warn_cnt_sb > 0:
        st.warning(
            f"📢 **[기획팀 총괄 관제 알림 | {selected_quarter}]** 전체 {len(summary_df_sidebar)}개 주요 정량지표 중 "
            f"**위험 {danger_cnt_sb}건**, **주의 {warn_cnt_sb}건**이 감지되었습니다. "
            "종합 성과 관제탑(탭 2) 및 4분기 잔여 목표 시뮬레이터(탭 3)에서 만점 수렴 필요량을 확인하십시오."
        )
    else:
        st.success(
            f"✅ **[기획팀 총괄 관제 알림 | {selected_quarter}]** 전 지표가 목표 궤도 내에서 순항 중입니다."
        )


# ==============================================================================
# 8. 메인 화면 3개 핵심 탭 구성
#    - 탭 1: 부서 실적 입력 및 스마트 검증 (Data Input & Validation)
#    - 탭 2: 종합 성과 관제탑 (Dashboard)
#    - 탭 3: 편람 산식 시뮬레이터 & 엑셀 연동
# ==============================================================================
tab1, tab2, tab3 = st.tabs(
    [
        "📝 탭 1: 부서 실적 입력 및 스마트 검증 (Data Input & Validation)",
        "📡 탭 2: 종합 성과 관제탑 (Dashboard)",
        "🧮 탭 3: 편람 산식 시뮬레이터 & 엑셀 연동 (Simulator & Excel Audit)",
    ]
)


# ------------------------------------------------------------------------------
# [탭 1] 부서 실적 입력 및 스마트 검증 (Data Input & Validation)
# ------------------------------------------------------------------------------
with tab1:
    st.subheader(f"📝 부서 정량실적 입력 및 AI 스마트 유효성 검증 ({selected_quarter} 기준)")
    st.caption(
        "지표를 선택하면 평가편람의 배점, 단위, 산식 개요, 직전 3개년 및 전년도 실적이 자동 표시되며, "
        "실적 입력 시 단위 오입력(±10배)과 목표 미달 위험(달성률 70% 미만)을 실시간 검증합니다."
    )

    indicators_map: Dict[str, Dict[str, Any]] = st.session_state["indicators_data"]
    indicator_options = {
        ind_id: f"[{cfg['name']}] (배점 {cfg['weight']}점, 단위: {cfg['unit']} | {cfg.get('department', '')})"
        for ind_id, cfg in indicators_map.items()
    }

    col_sel, col_rev = st.columns([3, 1])
    with col_sel:
        selected_ind_id = st.selectbox(
            "📌 평가 지표 선택 (드롭다운)",
            options=list(indicator_options.keys()),
            format_func=lambda k: indicator_options[k],
            key="tab1_indicator_select",
        )
    selected_cfg = indicators_map[selected_ind_id]

    with col_rev:
        # 하향지표(온실가스·에너지 감축 등) 역산식 적용 플래그 스위치
        is_reverse_flag = st.toggle(
            "🔄 하향지표 역산식 적용",
            value=bool(selected_cfg.get("is_reverse", False)),
            help="온실가스 감축량·에너지 사용량 등 실적이 낮을수록 고득점인 하향지표일 경우 활성화합니다.",
            key=f"rev_toggle_{selected_ind_id}",
        )
        if is_reverse_flag != bool(selected_cfg.get("is_reverse", False)):
            st.session_state["indicators_data"][selected_ind_id]["is_reverse"] = is_reverse_flag
            selected_cfg = st.session_state["indicators_data"][selected_ind_id]
            st.rerun()

    # 선택 지표의 기준치 및 목표 구간 계산
    base_val_t1 = compute_base_value(selected_cfg)
    upper_val_t1, lower_val_t1 = compute_target_bounds(base_val_t1, is_reverse=is_reverse_flag)
    prev_annual_t1 = float(selected_cfg.get("prev_year_actual", base_val_t1))
    prev_sync_t1 = float(
        selected_cfg.get("prev_year_cumulative", {}).get(selected_quarter, base_val_t1 * 0.75)
    )
    current_saved_val = float(selected_cfg["quarterly_actuals"].get(selected_quarter, 0.0))

    # [지표 선택 시 자동 표시되는 메타데이터 카드]
    history_str = " / ".join(
        [f"{yr}: {v:,.0f}{selected_cfg['unit']}" for yr, v in selected_cfg.get("history_3yr", {}).items()]
    )
    direction_badge = (
        '<span class="badge-blue">하향지표 (역산식: 낮을수록 만점)</span>'
        if is_reverse_flag
        else '<span class="badge-green">상향지표 (정산식: 높을수록 만점)</span>'
    )

    st.markdown(
        f"""
        <div class="meta-card">
            <div class="meta-title">
                📌 [{selected_cfg['name']}] 편람 기준정보 카드 &nbsp; {direction_badge}
            </div>
            <div style="font-size:0.86rem; color:#334155; margin-bottom:0.45rem;">
                <b>• 산식 개요:</b> {selected_cfg['formula_desc']}<br/>
                <b>• 직전 3개년 실적 내역:</b> {history_str}
            </div>
            <div class="meta-grid">
                <div class="meta-item">
                    <div class="meta-item-label">평가 배점 / 단위</div>
                    <div class="meta-item-val">{selected_cfg['weight']:.1f}점 / {selected_cfg['unit']}</div>
                </div>
                <div class="meta-item">
                    <div class="meta-item-label">기준치 (3개년 평균)</div>
                    <div class="meta-item-val">{base_val_t1:,.1f} {selected_cfg['unit']}</div>
                </div>
                <div class="meta-item">
                    <div class="meta-item-label">최고목표치 (상한·만점)</div>
                    <div class="meta-item-val" style="color:#1d4ed8;">{upper_val_t1:,.1f} {selected_cfg['unit']}</div>
                </div>
                <div class="meta-item">
                    <div class="meta-item-label">최저목표치 (하한·60%)</div>
                    <div class="meta-item-val" style="color:#b45309;">{lower_val_t1:,.1f} {selected_cfg['unit']}</div>
                </div>
                <div class="meta-item">
                    <div class="meta-item-label">전년도 연간 실적</div>
                    <div class="meta-item-val">{prev_annual_t1:,.1f} {selected_cfg['unit']}</div>
                </div>
                <div class="meta-item">
                    <div class="meta-item-label">전년 동기({selected_quarter}) 실적</div>
                    <div class="meta-item-val">{prev_sync_t1:,.1f} {selected_cfg['unit']}</div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # 실적 입력 및 스마트 검증 영역
    col_in, col_preview = st.columns([1.25, 1])

    with col_in:
        st.markdown(f"#### ✏️ {selected_quarter} 누적 실적값 입력")

        # 빠른 검증 시연 버튼 (정상값 / 단위오류 10배 / 목표미달 70% 미만 / 음수 에러)
        st.caption("⚡ **AI 검증 시나리오 원클릭 테스트** (클릭 시 입력 필드에 테스트 값이 자동 채워집니다):")
        b_col1, b_col2, b_col3, b_col4 = st.columns(4)
        input_state_key = f"raw_input_{selected_ind_id}_{selected_quarter}"
        if input_state_key not in st.session_state:
            st.session_state[input_state_key] = f"{current_saved_val:.1f}".rstrip("0").rstrip(".")

        if b_col1.button("정상 실적값", key=f"btn_norm_{selected_ind_id}", use_container_width=True):
            st.session_state[input_state_key] = f"{current_saved_val:.0f}"
        if b_col2.button("단위오류(×15배)", key=f"btn_unit_{selected_ind_id}", use_container_width=True):
            st.session_state[input_state_key] = f"{prev_sync_t1 * 15.0:.0f}"
        if b_col3.button("미달위험(60%)", key=f"btn_low_{selected_ind_id}", use_container_width=True):
            st.session_state[input_state_key] = f"{base_val_t1 * 0.60:.0f}"
        if b_col4.button("음수에러(-100)", key=f"btn_neg_{selected_ind_id}", use_container_width=True):
            st.session_state[input_state_key] = "-100"

        raw_user_input = st.text_input(
            f"3분기(현재 선택: {selected_quarter}) 누적 실적값 입력 (단위: {selected_cfg['unit']})",
            key=input_state_key,
            help="숫자만 입력하거나 쉼표(,)를 포함하여 입력할 수 있습니다. 비어있거나 음수일 경우 저장이 차단됩니다.",
        )

        # 실시간 AI 유효성 검증 실행
        parsed_val, val_messages, is_valid_input = validate_indicator_input(
            ind_cfg=selected_cfg,
            quarter=selected_quarter,
            raw_input_text=raw_user_input,
        )

        st.markdown("##### 🤖 AI 입력값 유효성 검증 (Smart Validation) 결과")
        for msg in val_messages:
            if msg.level == "error":
                st.error(f"**{msg.title}**\n\n{msg.detail}")
            elif msg.level == "warning":
                st.warning(f"**{msg.title}**\n\n{msg.detail}")
            elif msg.level == "info":
                st.info(f"**{msg.title}**\n\n{msg.detail}")
            else:
                st.success(f"**{msg.title}**\n\n{msg.detail}")

        # 실적 저장 버튼
        save_clicked = st.button(
            f"💾 [{selected_cfg['name']}] {selected_quarter} 실적 저장 (세션 즉시 반영)",
            type="primary",
            disabled=not is_valid_input,
            use_container_width=True,
        )
        if save_clicked and parsed_val is not None:
            old_val = st.session_state["indicators_data"][selected_ind_id]["quarterly_actuals"][selected_quarter]
            st.session_state["indicators_data"][selected_ind_id]["quarterly_actuals"][selected_quarter] = parsed_val
            new_eval = evaluate_indicator(
                st.session_state["indicators_data"][selected_ind_id],
                quarter=selected_quarter,
            )
            st.session_state["audit_logs"].insert(
                0,
                {
                    "일시": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "분기": selected_quarter,
                    "지표명": selected_cfg["name"],
                    "변경전": f"{old_val:,.1f} {selected_cfg['unit']}",
                    "변경후": f"{parsed_val:,.1f} {selected_cfg['unit']}",
                    "예상득점": f"{new_eval.score:.3f}점 / {new_eval.weight:.1f}점 ({new_eval.achievement_rate:.1f}%)",
                },
            )
            st.toast(
                f"[{selected_cfg['name']}] {selected_quarter} 실적이 {parsed_val:,.1f} {selected_cfg['unit']}(으)로 저장되었습니다!",
                icon="✅",
            )
            st.rerun()

    with col_preview:
        st.markdown("#### 🎯 실시간 선형보간 평점 시뮬레이션 프리뷰")
        preview_actual = parsed_val if (is_valid_input and parsed_val is not None) else current_saved_val
        preview_res = evaluate_indicator(
            selected_cfg,
            quarter=selected_quarter,
            actual_override=preview_actual,
        )

        p_col1, p_col2, p_col3 = st.columns(3)
        p_col1.metric(
            "기준치 대비 달성률",
            f"{preview_res.achievement_rate:.1f}%",
            f"{preview_res.actual_value - preview_res.base_value:+,.0f} {selected_cfg['unit']}",
        )
        p_col2.metric(
            "예상 산출 평점",
            f"{preview_res.score:.3f}점",
            f"배점 {preview_res.weight:.1f}점 중 {preview_res.score_ratio:.1f}%",
        )
        p_col3.metric(
            "평어 / 관제상태",
            preview_res.status_level,
            preview_res.grade_label,
        )

        # 선형보간 구간 내 현재 실적 위치 시각화 차트
        x_min = min(preview_res.lower_target, preview_res.upper_target, preview_res.actual_value) * 0.85
        x_max = max(preview_res.lower_target, preview_res.upper_target, preview_res.actual_value) * 1.12
        if not preview_res.is_reverse:
            curve_x = [x_min, preview_res.lower_target, preview_res.upper_target, x_max]
            curve_y = [
                preview_res.weight * 0.6,
                preview_res.weight * 0.6,
                preview_res.weight * 1.0,
                preview_res.weight * 1.0,
            ]
        else:
            curve_x = [x_min, preview_res.upper_target, preview_res.lower_target, x_max]
            curve_y = [
                preview_res.weight * 1.0,
                preview_res.weight * 1.0,
                preview_res.weight * 0.6,
                preview_res.weight * 0.6,
            ]

        fig_interp = go.Figure()
        fig_interp.add_trace(
            go.Scatter(
                x=curve_x,
                y=curve_y,
                mode="lines+markers",
                name="편람 선형보간 평점 곡선",
                line=dict(color="#1d4ed8", width=3),
            )
        )
        fig_interp.add_trace(
            go.Scatter(
                x=[preview_res.actual_value],
                y=[preview_res.score],
                mode="markers+text",
                name="현재 입력 실적",
                marker=dict(color="#dc2626", size=14, symbol="diamond"),
                text=[f"현재: {preview_res.score:.3f}점"],
                textposition="top center",
            )
        )
        fig_interp.update_layout(
            height=265,
            margin=dict(l=20, r=20, t=30, b=20),
            xaxis_title=f"실적 ({selected_cfg['unit']})",
            yaxis_title="산출 득점 (점)",
            yaxis=dict(range=[preview_res.weight * 0.5, preview_res.weight * 1.08]),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        )
        st.plotly_chart(fig_interp, use_container_width=True)

    with st.expander("🕒 최근 실적 입력·변경 감사 로그 (Audit Trail)", expanded=False):
        st.dataframe(pd.DataFrame(st.session_state["audit_logs"]), use_container_width=True, hide_index=True)


# ------------------------------------------------------------------------------
# [탭 2] 종합 성과 관제탑 (Dashboard)
# ------------------------------------------------------------------------------
with tab2:
    st.subheader(f"📡 기획팀 종합 성과 관제탑 ({selected_quarter} 기준)")
    st.caption(
        "지방공기업 경영평가 편람의 선형보간법 공식을 적용한 실시간 전사 예상 평점 및 위험·주의 지표 모니터링 대시보드입니다."
    )

    dash_df = build_summary_dataframe(st.session_state["indicators_data"], quarter=selected_quarter)

    total_weight = float(dash_df["배점"].sum())
    expected_score = float(dash_df["산출 득점"].sum())
    overall_score_rate = (expected_score / total_weight * 100.0) if total_weight > 0 else 0.0
    danger_count = int((dash_df["상태"] == "위험").sum())
    warning_count = int((dash_df["상태"] == "주의").sum())
    normal_count = int((dash_df["상태"] == "정상").sum())

    # [KPI 카드 4개 필수 구현]
    kpi_c1, kpi_c2, kpi_c3, kpi_c4 = st.columns(4)

    with kpi_c1:
        st.markdown(
            f"""
            <div class="kpi-card">
                <div class="kpi-label">1. 총 배점 (주요 정량지표)</div>
                <div class="kpi-value">{total_weight:.2f}점 <span style="font-size:0.95rem; font-weight:600; color:#64748b;">만점 기준</span></div>
                <div class="kpi-sub">탑재 지표 총 {len(dash_df)}개 (하한 60% = {total_weight * 0.6:.2f}점)</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with kpi_c2:
        st.markdown(
            f"""
            <div class="kpi-card">
                <div class="kpi-label">2. 현재 예상 득점 (선형보간법)</div>
                <div class="kpi-value" style="color:#1d4ed8;">{expected_score:.3f}점</div>
                <div class="kpi-sub">만점 대비 잔여 득점 여력: <b>+{total_weight - expected_score:.3f}점</b></div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with kpi_c3:
        st.markdown(
            f"""
            <div class="kpi-card">
                <div class="kpi-label">3. 전체 득점률 (%)</div>
                <div class="kpi-value" style="color:#0f766e;">{overall_score_rate:.2f}%</div>
                <div class="kpi-sub">기준치 대비 평균 달성률: <b>{dash_df['기준치 대비 달성률(%)'].mean():.1f}%</b></div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with kpi_c4:
        st.markdown(
            f"""
            <div class="kpi-card">
                <div class="kpi-label">4. 위험 / 주의 지표 현황</div>
                <div class="kpi-value">{danger_count + warning_count}개 <span style="font-size:0.95rem; font-weight:600; color:#64748b;">집중관리 대상</span></div>
                <div class="kpi-sub">
                    <span class="badge-red">위험 {danger_count}건</span>
                    <span class="badge-yellow">주의 {warning_count}건</span>
                    <span class="badge-green">정상 {normal_count}건</span>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("<div style='height: 1rem;'></div>", unsafe_allow_html=True)

    # [지표별 달성 현황 대화형 테이블]
    # 요구사항 컬럼: (지표명, 배점, 기준치, 최고목표, 3분기 실적, 달성률, 산출 득점, 상태)
    st.markdown(f"#### 📋 지표별 달성 현황 대화형 테이블 ({selected_quarter})")

    display_table_df = pd.DataFrame(
        {
            "지표명": dash_df["지표명"],
            "소관부서": dash_df["소관부서"],
            "단위": dash_df["단위"],
            "배점": dash_df["배점"],
            "기준치": dash_df["기준치(3개년평균)"],
            "최저목표(90%)": dash_df["최저목표치(90%)"],
            "최고목표(110%)": dash_df["최고목표치(110%)"],
            "3분기 실적" if selected_quarter == "3분기" else f"{selected_quarter} 실적": dash_df[f"{selected_quarter} 실적"],
            "달성률(%)": dash_df["기준치 대비 달성률(%)"],
            "산출 득점": dash_df["산출 득점"],
            "득점률(%)": dash_df["득점률(%)"],
            "상태": dash_df["상태"].map(
                lambda s: "🔴 위험" if s == "위험" else ("🟡 주의" if s == "주의" else "🟢 정상")
            ),
        }
    )

    actual_col_label = "3분기 실적" if selected_quarter == "3분기" else f"{selected_quarter} 실적"

    def style_status_table(styler: pd.io.formats.style.Styler) -> pd.io.formats.style.Styler:
        def highlight_row(row: pd.Series) -> List[str]:
            status_val = str(row["상태"])
            if "위험" in status_val:
                return ["background-color: #fef2f2; color: #991b1b; font-weight: 600;"] * len(row)
            if "주의" in status_val:
                return ["background-color: #fffbeb; color: #92400e;"] * len(row)
            return [""] * len(row)

        return styler.apply(highlight_row, axis=1).format(
            {
                "배점": "{:.1f}",
                "기준치": "{:,.0f}",
                "최저목표(90%)": "{:,.0f}",
                "최고목표(110%)": "{:,.0f}",
                actual_col_label: "{:,.0f}",
                "달성률(%)": "{:.2f}%",
                "산출 득점": "{:.3f}",
                "득점률(%)": "{:.1f}%",
            }
        )

    st.dataframe(
        style_status_table(display_table_df.style),
        use_container_width=True,
        hide_index=True,
    )

    # [Plotly 차트: 지표별 목표 대비 실적 달성률 막대그래프 (100% 기준선 표시)]
    st.markdown("<div style='height: 0.75rem;'></div>", unsafe_allow_html=True)
    chart_col1, chart_col2 = st.columns([1.45, 1])

    with chart_col1:
        st.markdown("#### 📊 지표별 목표 대비 실적 달성률 (%) [100% 기준선 표시]")
        rate_basis = st.radio(
            "달성률 기준 선택",
            options=[
                "기준치(3개년 평균=100%) 대비 달성률",
                "최고목표치(만점 상한=100%) 대비 달성률",
            ],
            horizontal=True,
            key="dash_chart_basis",
        )
        y_col = (
            "기준치 대비 달성률(%)"
            if "기준치" in rate_basis
            else "최고목표 대비 달성률(%)"
        )

        color_map = {"위험": "#ef4444", "주의": "#f59e0b", "정상": "#10b981"}
        fig_bar = px.bar(
            dash_df,
            x="지표명",
            y=y_col,
            color="상태",
            color_discrete_map=color_map,
            text=dash_df[y_col].map(lambda v: f"{v:.1f}%"),
            hover_data=["배점", "단위", "기준치(3개년평균)", "최고목표치(110%)", f"{selected_quarter} 실적", "산출 득점"],
        )
        # 100% 기준선 및 70% 위험 경계선 추가
        fig_bar.add_hline(
            y=100.0,
            line_dash="dash",
            line_color="#1e3a8a",
            line_width=2.5,
            annotation_text="100% 목표 기준선",
            annotation_position="top right",
        )
        if "기준치" in rate_basis:
            fig_bar.add_hline(
                y=110.0,
                line_dash="dot",
                line_color="#2563eb",
                line_width=1.8,
                annotation_text="110% 최고목표(만점선)",
                annotation_position="bottom right",
            )
            fig_bar.add_hline(
                y=70.0,
                line_dash="dot",
                line_color="#dc2626",
                line_width=1.8,
                annotation_text="70% 미달 위험선",
                annotation_position="bottom left",
            )

        fig_bar.update_traces(textposition="outside")
        fig_bar.update_layout(
            height=410,
            yaxis=dict(range=[0, max(125.0, float(dash_df[y_col].max()) * 1.18)]),
            xaxis_title="공사 주요 정량지표",
            yaxis_title=y_col,
            legend_title_text="관제 상태",
            margin=dict(l=20, r=20, t=30, b=20),
        )
        st.plotly_chart(fig_bar, use_container_width=True)

    with chart_col2:
        st.markdown("#### 🎯 지표별 배점 대비 예상 득점 기여도")
        fig_score = go.Figure()
        fig_score.add_trace(
            go.Bar(
                name="현재 예상 득점",
                x=dash_df["지표명"],
                y=dash_df["산출 득점"],
                marker_color="#2563eb",
                text=dash_df["산출 득점"].map(lambda v: f"{v:.2f}점"),
                textposition="inside",
            )
        )
        fig_score.add_trace(
            go.Bar(
                name="만점 대비 미확보 점수(잔여 여력)",
                x=dash_df["지표명"],
                y=dash_df["배점"] - dash_df["산출 득점"],
                marker_color="#cbd5e1",
                text=(dash_df["배점"] - dash_df["산출 득점"]).map(
                    lambda v: f"+{v:.2f}" if v > 0.01 else ""
                ),
                textposition="outside",
            )
        )
        fig_score.update_layout(
            barmode="stack",
            height=410,
            xaxis_title="공사 주요 정량지표",
            yaxis_title="배점 및 예상 득점 (점)",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            margin=dict(l=20, r=20, t=30, b=20),
        )
        st.plotly_chart(fig_score, use_container_width=True)


# ------------------------------------------------------------------------------
# [탭 3] 편람 산식 시뮬레이터 & 엑셀 연동
# ------------------------------------------------------------------------------
with tab3:
    st.subheader("🧮 4분기 잔여 목표 역산 시뮬레이터 & 수기 엑셀 자동 검산")

    # --------------------------------------------------------------------------
    # 섹션 A: 4분기 잔여 목표 예측 및 만점 역산 시뮬레이터
    # --------------------------------------------------------------------------
    st.markdown("### 1️⃣ 4분기 잔여 목표 예측 (연말 만점 달성을 위한 최소 필요 실적 역산)")
    st.caption(
        "현재 3분기(또는 선택 분기) 누적 실적을 바탕으로, 연말 평가에서 **만점(100%, 최고목표치 110%)**을 획득하기 위해 "
        "4분기에 추가로 달성해야 하는 최소 실적량과 일평균 필요 물량을 자동 역산합니다."
    )

    sim_rows: List[Dict[str, Any]] = []
    for ind_id, cfg in st.session_state["indicators_data"].items():
        q3_actual = float(cfg["quarterly_actuals"].get("3분기", 0.0))
        res_q3 = evaluate_indicator(cfg, quarter="3분기", actual_override=q3_actual)

        if not res_q3.is_reverse:
            min_required_q4_for_max = max(0.0, res_q3.upper_target - q3_actual)
            min_required_q4_for_base = max(0.0, res_q3.base_value - q3_actual)
            min_required_q4_for_lower = max(0.0, res_q3.lower_target - q3_actual)
        else:
            # 하향지표(역산식)인 경우 남은 분기 허용 가능한 잔여 한도
            min_required_q4_for_max = max(0.0, res_q3.upper_target - q3_actual)
            min_required_q4_for_base = max(0.0, res_q3.base_value - q3_actual)
            min_required_q4_for_lower = max(0.0, res_q3.lower_target - q3_actual)

        # 전년도 4분기 단일 분기 실적 (연간 - 3분기 누적)과 비교하여 실현 가능성 진단
        prev_q4_only = max(
            1.0,
            float(cfg.get("prev_year_actual", res_q3.base_value))
            - float(cfg.get("prev_year_cumulative", {}).get("3분기", res_q3.base_value * 0.75)),
        )
        burden_ratio = (min_required_q4_for_max / prev_q4_only) * 100.0

        if min_required_q4_for_max <= 0:
            feasibility = "✅ 조기 만점 달성"
        elif burden_ratio <= 105.0:
            feasibility = f"🟢 달성 유력 (전년 4Q 대비 {burden_ratio:.0f}%)"
        elif burden_ratio <= 130.0:
            feasibility = f"🟡 총력 추진 필요 (전년 4Q 대비 {burden_ratio:.0f}%)"
        else:
            feasibility = f"🔴 고강도 비상대책 필요 (전년 4Q 대비 {burden_ratio:.0f}%)"

        sim_rows.append(
            {
                "지표명": res_q3.indicator_name,
                "배점": res_q3.weight,
                "단위": res_q3.unit,
                "기준치(100%)": res_q3.base_value,
                "최고목표치(만점·110%)": res_q3.upper_target,
                "3분기 누적실적": q3_actual,
                "3분기 기준 득점": res_q3.score,
                "★ 4분기 만점 필요 추가실적": round(min_required_q4_for_max, 1),
                "4분기 일평균 필요량(90일)": round(min_required_q4_for_max / 90.0, 1),
                "기준치(100%) 잔여실적": round(min_required_q4_for_base, 1),
                "만점 도달 난이도 진단": feasibility,
            }
        )

    sim_df = pd.DataFrame(sim_rows)
    st.dataframe(
        sim_df.style.format(
            {
                "배점": "{:.1f}",
                "기준치(100%)": "{:,.0f}",
                "최고목표치(만점·110%)": "{:,.0f}",
                "3분기 누적실적": "{:,.0f}",
                "3분기 기준 득점": "{:.3f}",
                "★ 4분기 만점 필요 추가실적": "{:,.1f}",
                "4분기 일평균 필요량(90일)": "{:,.1f}",
                "기준치(100%) 잔여실적": "{:,.1f}",
            }
        ),
        use_container_width=True,
        hide_index=True,
    )

    # 대화형 What-If 4분기 추가 실적 시뮬레이터
    with st.expander("🎛️ [What-If 시뮬레이터] 4분기 예상 추가 실적 슬라이더 조절 및 연말 최종 평점 예측", expanded=True):
        sim_ind_col, sim_slider_col, sim_res_col = st.columns([1.1, 1.3, 1.2])
        with sim_ind_col:
            sim_ind_id = st.selectbox(
                "시뮬레이션 대상 지표 선택",
                options=list(st.session_state["indicators_data"].keys()),
                format_func=lambda k: st.session_state["indicators_data"][k]["name"],
                key="sim_ind_select",
            )
            sim_cfg = st.session_state["indicators_data"][sim_ind_id]
            q3_val_sim = float(sim_cfg["quarterly_actuals"].get("3분기", 0.0))
            res_q3_sim = evaluate_indicator(sim_cfg, quarter="3분기", actual_override=q3_val_sim)
            req_for_max_sim = max(0.0, res_q3_sim.upper_target - q3_val_sim)

            st.markdown(
                f"""
                - **3분기 누적 실적:** `{q3_val_sim:,.0f} {sim_cfg['unit']}`
                - **연말 만점 목표치(110%):** `{res_q3_sim.upper_target:,.0f} {sim_cfg['unit']}`
                - **만점 필요 4분기 추가량:** **`{req_for_max_sim:,.0f} {sim_cfg['unit']}`**
                """
            )

        with sim_slider_col:
            max_slider_val = float(max(req_for_max_sim * 1.5, res_q3_sim.base_value * 0.5, 1000.0))
            default_slider_val = float(min(req_for_max_sim, max_slider_val))
            q4_added_sim = st.slider(
                f"4분기 예상 추가 달성 실적 ({sim_cfg['unit']})",
                min_value=0.0,
                max_value=round(max_slider_val, -1),
                value=round(default_slider_val, -1),
                step=max(100.0, round(max_slider_val / 100.0, -1)),
                key=f"sim_slider_{sim_ind_id}",
            )
            projected_annual = q3_val_sim + q4_added_sim
            projected_res = evaluate_indicator(sim_cfg, quarter="4분기(최종)", actual_override=projected_annual)

        with sim_res_col:
            st.metric(
                "예측 연말 최종 실적",
                f"{projected_annual:,.0f} {sim_cfg['unit']}",
                f"최고목표 대비 {projected_annual - res_q3_sim.upper_target:+,.0f} {sim_cfg['unit']}",
            )
            st.metric(
                "예측 연말 최종 평점",
                f"{projected_res.score:.3f}점 / {projected_res.weight:.1f}점",
                f"3분기 대비 +{projected_res.score - res_q3_sim.score:.3f}점 ({projected_res.grade_label})",
            )

    st.divider()

    # --------------------------------------------------------------------------
    # 섹션 B: 현재 데이터 엑셀(CSV / XLSX) 다운로드
    # --------------------------------------------------------------------------
    st.markdown("### 2️⃣ 현재 데이터 엑셀(CSV / XLSX) 다운로드")
    export_df = build_summary_dataframe(st.session_state["indicators_data"], quarter=selected_quarter)

    dl_col1, dl_col2, dl_col3 = st.columns(3)
    csv_bytes = export_df.to_csv(index=False).encode("utf-8-sig")
    xlsx_bytes = to_excel_bytes(export_df, sheet_name=f"{selected_quarter}_정량성과")

    with dl_col1:
        st.download_button(
            label=f"📥 현재 데이터 CSV 다운로드 ({selected_quarter}, UTF-8 BOM)",
            data=csv_bytes,
            file_name=f"서울시농수산식품공사_2027경영평가_{selected_quarter}_정량실적.csv",
            mime="text/csv",
            use_container_width=True,
            type="primary",
        )
    with dl_col2:
        st.download_button(
            label=f"📊 현재 데이터 Excel(.xlsx) 다운로드 ({selected_quarter})",
            data=xlsx_bytes,
            file_name=f"서울시농수산식품공사_2027경영평가_{selected_quarter}_정량실적.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with dl_col3:
        sample_err_df = generate_sample_audit_excel_df(
            st.session_state["indicators_data"],
            quarter=selected_quarter,
            inject_human_errors=True,
        )
        st.download_button(
            label="🧪 [검산 테스트용] 수기오류 포함 샘플 엑셀(.xlsx) 다운로드",
            data=to_excel_bytes(sample_err_df, sheet_name="수기작성_실적표"),
            file_name=f"수기작성_경영평가실적표_{selected_quarter}_오류샘플.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.divider()

    # --------------------------------------------------------------------------
    # 섹션 C: 외부 엑셀 업로드 시 자동 검산 (수기 계산값 vs 시스템 표준 엔진 비교 하이라이트)
    # --------------------------------------------------------------------------
    st.markdown("### 3️⃣ 외부 엑셀 업로드 시 자동 검산 (수기 엑셀 수식 오류·참조 왜곡 적발)")
    st.caption(
        "현업 부서에서 제출한 수기 엑셀(.xlsx 또는 .csv)을 업로드하면, 파일 내 **수기 계산값(최저·최고목표치, 달성률, 산출득점)**과 "
        "**시스템 표준 선형보간 엔진 계산값**을 셀 단위로 대조하여 **차이가 나는 셀을 빨간색으로 하이라이트**하고 오류 원인을 진단합니다."
    )

    audit_mode_col1, audit_mode_col2 = st.columns([1.4, 1])
    with audit_mode_col1:
        uploaded_file = st.file_uploader(
            "📂 검산할 외부 엑셀 파일(.xlsx, .csv)을 업로드하십시오",
            type=["xlsx", "csv"],
            key="audit_excel_uploader",
        )
    with audit_mode_col2:
        use_builtin_audit_demo = st.checkbox(
            "🔍 파일 업로드 없이 '수기 엑셀 오류 샘플'로 즉시 자동 검산 시연 보기",
            value=(uploaded_file is None),
            help="체크 시, 현업에서 빈번한 수기 수식 오류 3건(가락 청과·가락 수산·혁신제품)이 포함된 데이터로 즉시 자동 검산 결과를 시연합니다.",
        )

    audit_input_df: Optional[pd.DataFrame] = None
    if uploaded_file is not None:
        try:
            if uploaded_file.name.lower().endswith(".csv"):
                audit_input_df = pd.read_csv(uploaded_file)
            else:
                audit_input_df = pd.read_excel(uploaded_file)
            st.success(f"✅ 외부 파일 `{uploaded_file.name}` 로드 완료 ({len(audit_input_df)}개 행)")
        except Exception as e:
            st.error(f"파일을 읽는 중 오류가 발생했습니다: {e}")
    elif use_builtin_audit_demo:
        audit_input_df = generate_sample_audit_excel_df(
            st.session_state["indicators_data"],
            quarter=selected_quarter,
            inject_human_errors=True,
        )

    if audit_input_df is not None and not audit_input_df.empty:
        # 컬럼 매핑 유연성 확보 (수기 엑셀 양식 또는 시스템 다운로드 양식 모두 호환)
        def pick_col(df_in: pd.DataFrame, candidates: List[str], default: Any = 0.0) -> pd.Series:
            for c in candidates:
                if c in df_in.columns:
                    return df_in[c]
            return pd.Series([default] * len(df_in))

        names_s = pick_col(audit_input_df, ["지표명", "지표", "indicator_name"], "미상지표")
        weights_s = pd.to_numeric(pick_col(audit_input_df, ["배점", "weight"], 1.0), errors="coerce").fillna(1.0)
        units_s = pick_col(audit_input_df, ["단위", "unit"], "톤")
        dir_s = pick_col(audit_input_df, ["지표성격", "direction"], "상향(일반)")
        manual_base_s = pd.to_numeric(
            pick_col(audit_input_df, ["수기_기준치", "기준치(3개년평균)", "기준치"], 0.0), errors="coerce"
        ).fillna(0.0)
        manual_low_s = pd.to_numeric(
            pick_col(audit_input_df, ["수기_최저목표치", "최저목표치(90%)", "최저목표(90%)", "최저목표치"], 0.0),
            errors="coerce",
        ).fillna(0.0)
        manual_high_s = pd.to_numeric(
            pick_col(audit_input_df, ["수기_최고목표치", "최고목표치(110%)", "최고목표(110%)", "최고목표치"], 0.0),
            errors="coerce",
        ).fillna(0.0)
        manual_actual_s = pd.to_numeric(
            pick_col(
                audit_input_df,
                [
                    "수기_당해실적",
                    f"{selected_quarter} 실적",
                    "3분기 실적",
                    "당해실적",
                    "실적",
                ],
                0.0,
            ),
            errors="coerce",
        ).fillna(0.0)
        manual_rate_s = pd.to_numeric(
            pick_col(audit_input_df, ["수기_달성률(%)", "기준치 대비 달성률(%)", "달성률(%)"], 0.0),
            errors="coerce",
        ).fillna(0.0)
        manual_score_s = pd.to_numeric(
            pick_col(audit_input_df, ["수기_산출득점", "산출 득점", "산출득점", "득점"], 0.0),
            errors="coerce",
        ).fillna(0.0)

        comparison_rows: List[Dict[str, Any]] = []
        discrepancy_cells: List[Tuple[int, List[str]]] = []

        # 시스템 마스터 지표명 매핑 사전
        master_by_name = {cfg["name"]: cfg for cfg in st.session_state["indicators_data"].values()}

        for row_idx in range(len(audit_input_df)):
            ind_name = str(names_s.iloc[row_idx])
            w_val = float(weights_s.iloc[row_idx])
            u_val = str(units_s.iloc[row_idx])
            is_rev = "하향" in str(dir_s.iloc[row_idx]) or "역산" in str(dir_s.iloc[row_idx])

            m_base = float(manual_base_s.iloc[row_idx])
            m_low = float(manual_low_s.iloc[row_idx])
            m_high = float(manual_high_s.iloc[row_idx])
            m_actual = float(manual_actual_s.iloc[row_idx])
            m_rate = float(manual_rate_s.iloc[row_idx])
            m_score = float(manual_score_s.iloc[row_idx])

            # 시스템 기준치: 마스터에 존재하는 지표면 마스터 3개년 평균 사용, 없으면 수기 기준치 사용
            if ind_name in master_by_name:
                sys_base = compute_base_value(master_by_name[ind_name])
                is_rev = bool(master_by_name[ind_name].get("is_reverse", is_rev))
            else:
                sys_base = m_base

            sys_score, _, _, sys_high, sys_low = calculate_linear_interpolation_score(
                actual=m_actual,
                base_value=sys_base,
                weight=w_val,
                is_reverse=is_rev,
            )
            sys_rate = (
                round((m_actual / sys_base) * 100.0, 2)
                if (sys_base > 0 and not is_rev)
                else (round((sys_base / m_actual) * 100.0, 2) if m_actual > 0 else 0.0)
            )

            # 오차 판정 (부동소수점 허용 오차 고려)
            diff_low = round(m_low - sys_low, 2)
            diff_high = round(m_high - sys_high, 2)
            diff_rate = round(m_rate - sys_rate, 2)
            diff_score = round(m_score - sys_score, 4)

            err_cols: List[str] = []
            reasons: List[str] = []
            if abs(diff_low) > 0.5:
                err_cols.extend(["수기_최저목표", "시스템_최저목표"])
                reasons.append(f"최저목표치 오차({diff_low:+,.1f})")
            if abs(diff_high) > 0.5:
                err_cols.extend(["수기_최고목표", "시스템_최고목표"])
                reasons.append(f"최고목표치 오차({diff_high:+,.1f})")
            if abs(diff_rate) > 0.05:
                err_cols.extend(["수기_달성률(%)", "시스템_달성률(%)"])
                reasons.append(f"달성률 참조오차({diff_rate:+.2f}%p)")
            if abs(diff_score) > 0.002:
                err_cols.extend(["수기_산출득점", "시스템_산출득점", "득점오차(수기-시스템)"])
                reasons.append(f"선형보간 평점오차({diff_score:+.4f}점)")

            discrepancy_cells.append((row_idx, err_cols))

            comparison_rows.append(
                {
                    "검산판정": "❌ 수식불일치" if err_cols else "✅ 정상일치",
                    "지표명": ind_name,
                    "배점": w_val,
                    "당해실적": m_actual,
                    "수기_최저목표": m_low,
                    "시스템_최저목표": sys_low,
                    "수기_최고목표": m_high,
                    "시스템_최고목표": sys_high,
                    "수기_달성률(%)": m_rate,
                    "시스템_달성률(%)": sys_rate,
                    "수기_산출득점": m_score,
                    "시스템_산출득점": sys_score,
                    "득점오차(수기-시스템)": diff_score,
                    "자동검산 진단사유": " / ".join(reasons) if reasons else "수기 엑셀 수식과 편람 표준 엔진 완벽 일치",
                }
            )

        comp_df = pd.DataFrame(comparison_rows)
        mismatch_cnt = int((comp_df["검산판정"] == "❌ 수식불일치").sum())
        total_score_diff = float(comp_df["득점오차(수기-시스템)"].sum())

        # 검산 요약 배너
        if mismatch_cnt > 0:
            st.error(
                f"🚨 **[자동 검산 결과: 불일치 {mismatch_cnt}건 적발]** "
                f"수기 엑셀 합산 평점({comp_df['수기_산출득점'].sum():.3f}점)과 시스템 표준 엔진 평점({comp_df['시스템_산출득점'].sum():.3f}점) 간에 "
                f"총 **{total_score_diff:+.3f}점**의 왜곡이 발견되었습니다. 아래 하이라이트된 빨간색 셀을 확인하십시오."
            )
        else:
            st.success(
                f"🎉 **[자동 검산 완료: 전 지표 완벽 일치]** 업로드된 엑셀의 모든 목표치·달성률·선형보간 평점이 시스템 표준 계산값과 일치합니다."
            )

        # 차이가 나는 셀을 정밀 하이라이트하는 Styler 함수
        def highlight_discrepancy_cells(df_target: pd.DataFrame) -> pd.DataFrame:
            styles = pd.DataFrame("", index=df_target.index, columns=df_target.columns)
            for r_idx, cols in discrepancy_cells:
                if cols:
                    styles.loc[r_idx, "검산판정"] = (
                        "background-color: #fee2e2; color: #991b1b; font-weight: 700;"
                    )
                    styles.loc[r_idx, "자동검산 진단사유"] = (
                        "background-color: #fff1f2; color: #be123c; font-weight: 600;"
                    )
                    for c_name in cols:
                        if c_name in styles.columns:
                            styles.loc[r_idx, c_name] = (
                                "background-color: #fecaca; color: #7f1d1d; font-weight: 800; border: 1.5px solid #dc2626;"
                            )
                else:
                    styles.loc[r_idx, "검산판정"] = (
                        "background-color: #dcfce7; color: #166534; font-weight: 700;"
                    )
            return styles

        styled_comp = comp_df.style.apply(highlight_discrepancy_cells, axis=None).format(
            {
                "배점": "{:.1f}",
                "당해실적": "{:,.0f}",
                "수기_최저목표": "{:,.0f}",
                "시스템_최저목표": "{:,.0f}",
                "수기_최고목표": "{:,.0f}",
                "시스템_최고목표": "{:,.0f}",
                "수기_달성률(%)": "{:.2f}%",
                "시스템_달성률(%)": "{:.2f}%",
                "수기_산출득점": "{:.4f}",
                "시스템_산출득점": "{:.4f}",
                "득점오차(수기-시스템)": "{:+.4f}",
            }
        )

        st.dataframe(styled_comp, use_container_width=True, hide_index=True)
