#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""SST gel DVNext QC analyzer + Lot 종합 QC — HB-analysis-06-5.

읽은 파일: HB-analysis-06-4-1_integrated_reviewed.py
원본 VERSION: HB-analysis-06-4-1 / LotQC-2.6
원본 SHA-256: e195e4bad020e2b6a7e81f76a08a3cb4f1b111d03ecfecc12407fc4a390b27c2
파일명은 그대로 사용한다. 파일명에서 VERSION을 읽지 않는다.


Method A: step-flow Up/Down H-B profiles, common-range comparison, hysteresis.
Method B: 3ITT baseline/breakdown/recovery profile.
QC mode : Method A fingerprint + Method B 회복 + 밀도(별도 입력)를 Lot 단위로 통합.

06 설계 원칙(수정 지시서 §3~§5)
- 모든 입력 CSV에 대해 분석은 계속 수행한다. 계산에서 제외하는 것은 파싱 실패·필수 컬럼 결손·수치 계산 불능만이다.
  Hold time, Torque 범위, 유효 point 수, H-B 조건, start-up peak 유무는 제외 사유가 아니라 Step·지표·Lot 결과에
  붙는 quality metadata(상태·사유·라벨)다.
- 두 채널: STRICT(60 s hold + Torque 10–95 % 점만) / INCLUSIVE(정책이 허용한 hold + 범위 밖 점 포함).
  Lot 통계·H-B 피팅·QC 지표를 채널별로 따로 계산하고 섞지 않는다. 기존 04_2 시트(Step_Results·Flow_Profile·
  Overview·Model_Validation·StartUp)는 STRICT 채널 대표값으로 채운다. --policy standard에서
  Step 대표 응력·점도·Torque, 대표창, 파싱은 06-2와 같다.
  대표 고전단은 실제 측정 최고 RPM이다. 그 Step이 채널에서 부적격이면 더 낮은 RPM으로 대체하지 않고 NO VALUE다.
  공식 SSI는 두 개다. A_SSI_FIELD_0P2_2 = eta_up(0.2)/eta_up(2.0), A_SSI_QC_0P5_5 = eta_up(0.5)/eta(5.0 Up endpoint).
  모호한 A_SSI는 Summary·Lot 통계·long CSV 행을 만들지 않는다. 실측 10배 RPM 쌍은 pair_id별로 따로 집계한다.
  0 rpm은 REST_STABILIZATION이며 유동·H-B·SSI에 넣지 않는다. 5 rpm은 하강 sequence가 있을 때만 shared turnaround다.
  0.35 rpm은 Full H-B 저전단 확장점이고 SSI 분자가 아니다. 4 rpm은 중간 측정점이다.
  측정 최고 RPM이 채널에서 부적격하면 4 rpm으로 대체하지 않고 대표 고전단 지표는 NO VALUE다.
  대표 H-B는 Up/Down Full이므로 06-2의 common 기반 H-B와 달라질 수 있다.
- 정책은 전역 상수가 아니라 AnalysisPolicy 객체(STANDARD/INCLUSIVE/CUSTOM preset, CLI로 구성)로 주입하고,
  QCPolicy 시트·Traceability·long CSV에 정책 JSON과 해시를 남긴다. 예외 래퍼(HB-one-time-exception)의
  monkey patch 방식은 폐기하며 30/45 s hold는 --policy inclusive(accepted_hold_s)로 본 분석기가 처리한다.
- 상태 체계: OK < NOTICE < REVIEW < CENSORED < EXCEPTION < NO VALUE < INVALID.
  NOTICE = 값이 있거나 없어도 분석은 유효하며 해석만 주의(합격성 판정 제외). A_OVS_L 미산출, H-B R², GRANGE가 대상.
  _QCBuffer는 NaN+NOTICE를 NO VALUE로 덮어쓰지 않는다.
- 지표 사전에 metric_role(CORE/PROCESS_SENSITIVE/DIAGNOSTIC/DATA_QUALITY/CENSORED_METRIC), lot_acceptance,
  precision_evaluation 필드를 추가해 Lot 합격성·정밀도 판정 대상을 코드로 고정한다.
- 05-3에서 도입한 RSD N/A(0 기준·|Mean|<SD), Range·|Mean|·부호 일관성·[SIGN] 경고, Lot_QC_Compare는 유지한다.

지시서와 다르게 결정한 점(근거는 CHANGES-06.md)
- 30/45 s accepted hold의 대표창은 지시서 §F 함수(min(10 s, 25 %))가 아니라 후반 10 s 고정 — 회귀시험 2
  (예외 래퍼와 동일값) 요구를 우선. 미승인 짧은 hold(10–30 s)에만 min(10 s, 25 %) 창을 적용.
- STRICT 채널 Lot 통계는 'status=OK만'이 아니라 strict 유효 데이터의 수치(OK·NOTICE·REVIEW)를 사용 —
  05-3과 동일값(회귀시험 1)을 유지하기 위함. STRICT 평균에는 EXCEPTION·CENSORED·Torque 범위 밖 참고값을 넣지 않는다.
- 06-3 INCLUSIVE 채널은 정책이 허용하면 accepted hold 예외와 Torque 하한/상한 밖 참고값을 계산·Lot 통계에 포함한다.
  이 값은 strict 정량값으로 승격하지 않으며, value_class와 합격성(EXCLUDED/REVIEW)을 별도로 남긴다.

Dependencies: numpy, pandas, scipy, openpyxl
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import json
import math
import os
import re
import sys
from dataclasses import asdict, dataclass, field, replace
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries


STRESS_DYNE_TO_PA = 0.1
DEFAULT_TORQUE_MIN = 10.0
DEFAULT_TORQUE_MAX = 95.0
DEFAULT_WINDOW_SECONDS = 10.0
DEFAULT_WINDOW_FRACTION = 0.25
METHOD_A_STEP_SECONDS = 60.0
METHOD_A_TAIL_SECONDS = 10.0
METHOD_A_HOLD_TOLERANCE_SECONDS = 0.5
METHOD_A_MIN_HB_POINTS = 4
METHOD_A_RECOMMENDED_HB_POINTS = 5
METHOD_A_VERSION = "A-1.0"
RPM_ZERO_TOL = 0.001
SHEAR_RATE_ZERO_TOL = 0.002          # CP-52에서 RPM_ZERO_TOL×2에 해당하는 전단률 수치 오차
STABILIZATION_TARGET_S = 600.0
STABILIZATION_TEMP_TARGET_C = 25.0
FIELD_SSI_LOW_RPM = 0.2
FIELD_SSI_HIGH_RPM = 2.0
QC_SSI_LOW_RPM = 0.5
QC_SSI_HIGH_RPM = 5.0
A_TEMPLATE = "SST_Gel_StepFlow_Template.xlsx"
B_TEMPLATE = "SST_Gel_3ITT_Template.xlsx"
RECOVERY_TIMES = (10, 30, 60, 180, 300, 600, 900)


# ============================================================================ 06: AnalysisPolicy
@dataclass(frozen=True)
class AnalysisPolicy:
    """분석 정책. 전역 상수 대신 실행 시 구성해 모든 계산·시트·추적성에 주입한다(지시서 §B)."""
    name: str = "STANDARD"
    hold_target_s: float = METHOD_A_STEP_SECONDS
    hold_tolerance_s: float = METHOD_A_HOLD_TOLERANCE_SECONDS
    accepted_hold_s: tuple = (60.0,)
    allow_short_hold_analysis: bool = False
    torque_min_pct: float = DEFAULT_TORQUE_MIN
    torque_max_pct: float = DEFAULT_TORQUE_MAX
    include_out_of_torque_analysis: bool = True
    tail_seconds: float = METHOD_A_TAIL_SECONDS
    short_hold_window_fraction: float = 0.25      # 미승인 짧은 hold(10–30 s)의 대표창 비율
    min_window_s: float = 3.0
    min_hold_s: float = 10.0                       # 이보다 짧으면 대표값 계산 보류(INVALID)
    min_hb_points: int = METHOD_A_MIN_HB_POINTS
    recommended_hb_points: int = METHOD_A_RECOMMENDED_HB_POINTS
    generate_strict_values: bool = True
    generate_inclusive_values: bool = True
    censored_in_inclusive_fit: bool = False        # 전 점이 Torque 범위 밖인 Step을 inclusive H-B에 넣을지
    # 범위 밖 point를 inclusive 대표값·파생지표에 넣을지. censored_in_inclusive_fit과 함께 켜면 ALL_POINTS도 포함.
    include_out_of_range_in_inclusive: bool = False
    # inclusive Lot 평균에 CENSORED(Torque 참고값) 수치를 넣을지. strict 평균에는 절대 넣지 않는다.
    include_censored_in_lot_stats: bool = False
    lot_stats_mode: str = "both"                   # strict | inclusive | both
    # lot_stats_mode(감사용 strict/inclusive 통계 보존 범위)와 별개로 Summary·콘솔·비교표 대표 통계의 출처를 정한다.
    # primary_channel=inclusive이면 Method A의 기존 Mean~Range 열에 inclusive 통계를 표시한다.
    primary_channel: str = "strict"                 # strict | inclusive
    exception_stat_policy: str = "separate"        # exclude | separate | include
    report_values: str = "both"                    # strict | inclusive | both
    include_exception_in_lot_stats: bool = False
    include_notice_in_lot_acceptance: bool = False
    audit_level: str = "full"                      # basic | full
    require_stabilization: bool = True
    stabilization_target_s: float = STABILIZATION_TARGET_S
    stabilization_min_s: float | None = None
    stabilization_temp_target_C: float = STABILIZATION_TEMP_TARGET_C
    stabilization_temp_tolerance_C: float | None = None
    sampling_gap_review_s: float | None = None
    ssi_decade_ratio: float = 10.0
    ssi_ratio_tolerance_pct: float = 1.0
    ssi_calculate_all: bool = True
    ssi_primary_selection: str = "none"
    ssi_allow_interpolation: bool = False
    ssi_allow_model_estimation: bool = False
    ssi_include_down: bool = False
    ssi_pool_different_pairs: bool = False
    ssi_preferred_pairs: tuple = ((0.2, 2.0), (0.5, 5.0))
    ssi_explicit_pair: tuple | None = None

    def to_dict(self):
        d = asdict(self)
        d["accepted_hold_s"] = list(self.accepted_hold_s)
        return d

    def to_json(self):
        return json.dumps(self.to_dict(), ensure_ascii=False, sort_keys=True)

    def digest(self):
        return hashlib.sha256(self.to_json().encode("utf-8")).hexdigest()[:16]

    def hold_accepted(self, hold):
        return any(math.isclose(float(hold), float(x), abs_tol=self.hold_tolerance_s) for x in self.accepted_hold_s)

    def is_strict_hold(self, hold):
        return math.isclose(float(hold), self.hold_target_s, abs_tol=self.hold_tolerance_s)

    @property
    def is_standard(self):
        return self.name == "STANDARD"


POLICY_STANDARD = AnalysisPolicy(name="STANDARD", accepted_hold_s=(60.0,), allow_short_hold_analysis=False,
                                 include_out_of_torque_analysis=True, censored_in_inclusive_fit=False,
                                 include_out_of_range_in_inclusive=False, include_censored_in_lot_stats=False,
                                 primary_channel="strict", include_exception_in_lot_stats=False)
POLICY_INCLUSIVE = AnalysisPolicy(name="INCLUSIVE", accepted_hold_s=(30.0, 45.0, 60.0), allow_short_hold_analysis=True,
                                  include_out_of_torque_analysis=True, censored_in_inclusive_fit=True,
                                  include_out_of_range_in_inclusive=True, include_censored_in_lot_stats=True,
                                  primary_channel="inclusive", include_exception_in_lot_stats=True)
POLICY_CUSTOM = AnalysisPolicy(name="CUSTOM", accepted_hold_s=(30.0, 45.0, 60.0), allow_short_hold_analysis=True,
                               include_out_of_torque_analysis=True, censored_in_inclusive_fit=True,
                               include_out_of_range_in_inclusive=True, include_censored_in_lot_stats=True,
                               primary_channel="inclusive", include_exception_in_lot_stats=True)
POLICY_PRESETS = {"standard": POLICY_STANDARD, "inclusive": POLICY_INCLUSIVE, "custom": POLICY_CUSTOM}
_ACTIVE_POLICY = POLICY_STANDARD


def set_active_policy(policy):
    """process_enhanced 등 기존 호출 서명을 유지하는 경로에 정책을 전달하기 위한 실행 컨텍스트."""
    global _ACTIVE_POLICY
    _ACTIVE_POLICY = policy
    return policy


def active_policy():
    return _ACTIVE_POLICY


# Step 라벨(지시서 §N)
LBL_STRICT = "STRICT_60_VALID"
LBL_SHORT_INCL = "SHORT_HOLD_{h:g}S_INCLUDED"
LBL_SHORT_UNACC = "SHORT_HOLD_{h:g}S_NOT_ACCEPTED"
LBL_LONG = "LONG_HOLD_{h:g}S_INCLUDED"
LBL_LOW_TQ = "LOW_TORQUE_INCLUDED_REFERENCE_ONLY"
LBL_HIGH_TQ = "HIGH_TORQUE_INCLUDED_REFERENCE_ONLY"
LBL_MIXED_TQ = "OUT_OF_RANGE_MIXED_INCLUDED"
LBL_NO_POINTS = "NO_POINTS"
LBL_TOO_SHORT = "HOLD_TOO_SHORT_NOT_EVALUATED"
LBL_REST = "REST"

# 06-3 값 구분. 최종 상태 문자열을 파싱하지 않고 Step → 지표 → Lot까지 이 값을 전달한다.
VC_STRICT = "strict 정량값"
VC_EXC = "inclusive exception"
VC_LOW = "torque 하한 밖 참고값"
VC_HIGH = "torque 상한 밖 참고값"
VC_BOTH = "torque 범위 밖 참고값"
VC_NO = "NO VALUE"
VC_INVALID = "INVALID"
VC_NA = "NOT_APPLICABLE"
VC_ESTIMATED = "ESTIMATED"
VC_LEGACY = "LEGACY_UNKNOWN"
TORQUE_REFERENCE_CLASSES = {VC_LOW, VC_HIGH, VC_BOTH}
NON_NUMERIC_CLASSES = {VC_NO, VC_INVALID, VC_NA}


def inclusive_out_of_range_enabled(policy):
    """ALL_POINTS·범위 밖 point를 inclusive 계산에 넣는 정책인지."""
    if policy is None or not policy.include_out_of_torque_analysis:
        return False
    return bool(policy.censored_in_inclusive_fit or policy.include_out_of_range_in_inclusive)


def _vc_from_torque_use(used_low, used_high, n_low=0, n_high=0):
    """하한과 상한이 함께 쓰이면 한쪽을 숨기지 않고 torque 범위 밖 참고값으로 둔다."""
    if used_low and used_high:
        return VC_BOTH
    if used_low:
        return VC_LOW
    if used_high:
        return VC_HIGH
    return ""


def _pick_value_class(classes, used_low=False, used_high=False, n_low=0, n_high=0):
    """INVALID 다음, 계산된 범위 밖 참고값, inclusive exception, strict 정량값, NO VALUE."""
    classes = [c for c in classes if c]
    if VC_INVALID in classes:
        return VC_INVALID
    used_low = bool(used_low or VC_LOW in classes or VC_BOTH in classes)
    used_high = bool(used_high or VC_HIGH in classes or VC_BOTH in classes)
    if used_low or used_high:
        return _vc_from_torque_use(used_low, used_high, n_low if used_low else 0, n_high if used_high else 0)
    if VC_ESTIMATED in classes:
        return VC_ESTIMATED
    if VC_EXC in classes:
        return VC_EXC
    if VC_STRICT in classes:
        return VC_STRICT
    if VC_NA in classes:
        return VC_NA
    if VC_NO in classes:
        return VC_NO
    return VC_NO


def _default_value_class(status):
    """구조화 필드가 없는 채널(Method B 등)의 표시용 기본 구분. Method A는 Step 필드를 우선한다."""
    status = status or ""
    if status == "INVALID":
        return VC_INVALID
    if status in ("NO VALUE", "N/A", "NO DATA"):
        return VC_NO
    if status == "EXCEPTION":
        return VC_EXC
    if status == "CENSORED":
        return VC_LOW
    if status in ("OK", "NOTICE", "REVIEW", "VALID", "N/A"):
        return VC_STRICT
    return VC_NO


def _uniq_keep(items):
    out = []
    for x in items:
        if x and x not in out:
            out.append(x)
    return out


def read_lines(path: Path) -> list[str]:
    for enc in ("utf-8-sig", "cp949", "latin-1"):
        try:
            return path.read_text(encoding=enc).splitlines()
        except UnicodeDecodeError:
            pass
    raise UnicodeError(f"CSV 인코딩을 해석할 수 없습니다: {path}")


def clock_seconds(text: str):
    m = re.search(r"(\d{1,3}):(\d{2}):(\d{2})", str(text or ""))
    if not m:
        return None
    h, minute, sec = map(int, m.groups())
    return float(h * 3600 + minute * 60 + sec)


def hms_seconds(text: str):
    m = re.search(r"Time\s*=\s*(\d{1,3}):(\d{2}):(\d{2})", text, re.I)
    if not m:
        return None
    h, minute, sec = map(int, m.groups())
    return float(h * 3600 + minute * 60 + sec)


def parse_metadata(lines: list[str]) -> dict:
    meta = {}
    for line in lines[:15]:
        fields = next(csv.reader([line]))
        for i in range(0, len(fields) - 1, 2):
            key = fields[i].strip()
            value = fields[i + 1].strip()
            if key and value:
                meta[key] = value
    return meta


def parse_method_holds(lines: list[str], raw_header_idx: int) -> dict[int, float]:
    start = None
    for i, line in enumerate(lines[:raw_header_idx]):
        if line.replace(" ", "").lower().startswith("step,speed,temperature,datacollection"):
            start = i + 2
            break
    if start is None:
        return {}
    holds = {}
    modes = {}
    for line in lines[start:raw_header_idx]:
        fields = next(csv.reader([line]))
        if not fields or not fields[0].strip().isdigit():
            continue
        step = int(fields[0].strip())
        sec = hms_seconds(" ".join(fields))
        if sec is not None:
            holds[step] = sec
        collection = fields[3].strip() if len(fields) > 3 else ""
        avg = clock_seconds(fields[5]) if len(fields) > 5 else None
        modes[step] = {"collection": collection, "avg_s": avg, "single_point": "single point" in collection.lower()}
    return holds, modes


def parse_dvnext_csv(path):
    path = Path(path)
    lines = read_lines(path)
    header_idx = next((i for i, x in enumerate(lines) if x.replace(" ", "").startswith("Step,Point,Time,")), None)
    if header_idx is None:
        raise ValueError("point-by-point DATA 섹션을 찾지 못했습니다.")
    holds, collection_modes = parse_method_holds(lines, header_idx)
    expected = ["Step", "Point", "Time", "Viscosity", "Torque", "Speed", "Shear Stress", "Shear Rate"]
    header = [x.strip() for x in next(csv.reader([lines[header_idx]]))]
    if header[:8] != expected:
        raise ValueError(f"지원하지 않는 DATA 헤더: {header}")
    rows = []
    for fields in csv.reader(lines[header_idx + 2:]):
        if len(fields) < 8 or not fields[0].strip().isdigit():
            continue
        rows.append((fields[:11] + [""] * 11)[:11])
    if not rows:
        raise ValueError("측정 포인트가 없습니다.")
    cols = ["Step", "Point", "Time", "Viscosity_cP", "Torque_pct", "Speed_RPM",
            "ShearStress_dyne", "ShearRate", "Temp_C", "Density", "Accuracy_cP"]
    df = pd.DataFrame(rows, columns=cols)
    for c in cols:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df = df.dropna(subset=["Step", "Point", "Time", "Speed_RPM"]).copy()
    if df.empty:
        raise ValueError("숫자로 해석 가능한 측정 포인트가 없습니다.")
    df["Step"] = df["Step"].astype(int)
    df["Point"] = df["Point"].astype(int)
    df = df.sort_values(["Point", "Time"], kind="stable").reset_index(drop=True)
    df["ShearStress_Pa"] = df["ShearStress_dyne"] * STRESS_DYNE_TO_PA
    df["t_elapsed"] = df.groupby("Step")["Time"].transform(lambda s: s - s.min() + 1.0)
    observed = df.groupby("Step")["t_elapsed"].max().to_dict()
    for step, sec in observed.items():
        holds.setdefault(int(step), float(sec))
    df.attrs["collection_modes"] = collection_modes
    return df, holds, parse_metadata(lines)


def _plan_fields(direction, up, down, shared, step_type):
    return {"sequence_direction": direction, "used_in_up_branch": up, "used_in_down_branch": down,
            "turnaround_shared": shared, "step_type": step_type, "moving": step_type != "REST_STABILIZATION"}


def _zero_flow_step(rpm_value, gamma_value):
    gamma_ok = (not _finite_num(gamma_value)) or abs(float(gamma_value)) <= SHEAR_RATE_ZERO_TOL
    return abs(float(rpm_value)) <= RPM_ZERO_TOL and gamma_ok


def step_sequence_plan(df):
    """0 rpm 안정화를 뺀 moving sequence로 방향을 정한다.

    최고 RPM 뒤에 하강 Step이 없으면 그 점은 Up endpoint이고 shared turnaround가 아니다.
    하강 sequence가 있고 최고 RPM이 한 Step이면 그 점만 shared turnaround다.
    Step 번호로 Up/Down을 하드코딩하지 않는다.
    """
    rpm = df.groupby("Step")["Speed_RPM"].median().sort_index()
    gamma = df.groupby("Step")["ShearRate"].median() if "ShearRate" in df.columns else pd.Series(dtype=float)
    plan = {}
    moving_idx = []
    for step, value in rpm.items():
        step = int(step)
        g = float(gamma.loc[step]) if step in gamma.index and _finite_num(gamma.loc[step]) else np.nan
        if _zero_flow_step(value, g):
            plan[step] = _plan_fields("Rest", False, False, False, "REST_STABILIZATION")
        elif abs(float(value)) <= RPM_ZERO_TOL:
            plan[step] = _plan_fields("Rest", False, False, False, "UNKNOWN")
        else:
            moving_idx.append(step)
    if not moving_idx:
        return plan
    values = {s: float(rpm.loc[s]) for s in moving_idx}
    peak_rpm = max(values.values())
    peak_steps = [s for s in moving_idx if abs(values[s] - peak_rpm) <= RPM_MATCH_TOL]
    pos = {s: i for i, s in enumerate(moving_idx)}
    has_down_after = len(peak_steps) == 1 and pos[peak_steps[0]] < len(moving_idx) - 1
    if has_down_after and pos[peak_steps[0]] > 0:
        turn = peak_steps[0]
        ti = pos[turn]
        for s in moving_idx:
            i = pos[s]
            if i < ti:
                plan[s] = _plan_fields("Up", True, False, False, "MOVING_UP")
            elif i == ti:
                plan[s] = _plan_fields("Turnaround", True, True, True, "MOVING_TURNAROUND")
            else:
                plan[s] = _plan_fields("Down", False, True, False, "MOVING_DOWN")
    elif len(peak_steps) >= 2:
        end_up = pos[peak_steps[0]]
        for s in moving_idx:
            if pos[s] <= end_up:
                plan[s] = _plan_fields("Up", True, False, False, "MOVING_UP")
            else:
                plan[s] = _plan_fields("Down", False, True, False, "MOVING_DOWN")
    else:
        order = moving_idx
        peak = max(order, key=lambda s: (values[s], -order.index(s)))
        for s in order:
            if order.index(s) <= order.index(peak):
                plan[s] = _plan_fields("Up", True, False, False, "MOVING_UP")
            else:
                plan[s] = _plan_fields("Down", False, True, False, "MOVING_DOWN")
    return plan


def step_directions(df):
    """호환용 방향 라벨. 분석 membership은 step_sequence_plan의 branch 필드를 쓴다."""
    return {s: p["sequence_direction"] for s, p in step_sequence_plan(df).items()}


def select_tail_window(group, observed_s, policy, accepted):
    """대표창 선택(지시서 §F 변형). 승인 hold(60/accepted/long)는 후반 tail_seconds 고정,
    미승인 짧은 hold는 min(tail_seconds, observed×fraction), 단 min_window_s 미만이면 전체 구간."""
    if observed_s <= 0:
        return group.iloc[0:0], 0.0
    if accepted:
        cand = min(float(policy.tail_seconds), float(observed_s))
    else:
        cand = min(float(policy.tail_seconds), float(observed_s) * float(policy.short_hold_window_fraction))
        if cand < policy.min_window_s:
            cand = float(observed_s)
    start = max(0.0, float(observed_s) - cand)
    return group[group["t_elapsed"] > start], cand


def classify_hold(hold, policy):
    if not _finite_num(hold) or hold <= 0:
        return "UNKNOWN"
    if policy.is_strict_hold(hold):
        return "STRICT_60"
    if hold < policy.hold_target_s:
        return "SHORT_ACCEPTED" if policy.hold_accepted(hold) else "SHORT_UNACCEPTED"
    return "LONG"


def _finite_num(x):
    try:
        return x is not None and np.isfinite(float(x))
    except (TypeError, ValueError):
        return False


def step_representatives(df, holds, win_s=10.0, window_fraction=0.25, tq_lo=10.0, tq_hi=95.0, policy=None):
    """Method A Step 대표값 — STRICT/INCLUSIVE 2채널(지시서 §D·§E·§F).

    레거시 열(tau_Pa, eta_cP, torque_valid, n_valid, status, hold_compliant …)은 STRICT 채널 값이며
    05-3의 step_representatives와 동일하게 계산되므로 04_2 시트·H-B·hysteresis 수치는 변하지 않는다.
    win_s/window_fraction 인자는 호출 호환용이며 Method A에서는 사용하지 않는다(정책 객체가 대신함).
    """
    policy = policy or active_policy()
    if policy is None:
        policy = POLICY_STANDARD
    if not (math.isclose(policy.torque_min_pct, tq_lo) and math.isclose(policy.torque_max_pct, tq_hi)):
        policy = replace(policy, torque_min_pct=float(tq_lo), torque_max_pct=float(tq_hi))
    seq_plan = step_sequence_plan(df)
    direction = {s: p["sequence_direction"] for s, p in seq_plan.items()}
    records = []
    for step, g in df.groupby("Step", sort=True):
        step = int(step)
        hold = float(holds.get(step, g["t_elapsed"].max()))
        observed = float(g["t_elapsed"].max())
        moving = direction[step] != "Rest"
        mode = (df.attrs.get("collection_modes") or {}).get(step, {})
        # 대표창은 저장된 point의 경과시간을 쓴다. Single Point Averaging의 실제 안정화 시간은 method end condition이다.
        duration_observed = observed
        if (not moving) and mode.get("single_point") and hold > observed:
            duration_observed = hold
        hold_class = classify_hold(hold, policy) if moving else "REST"
        strict_hold = moving and hold_class == "STRICT_60"
        accepted = moving and (strict_hold or hold_class == "LONG"
                               or (hold_class == "SHORT_ACCEPTED" and policy.allow_short_hold_analysis))
        too_short = moving and observed < policy.min_hold_s
        # --- 대표창: STRICT 채널은 항상 후반 tail_seconds(05와 동일), INCLUSIVE는 select_tail_window
        strict_dur = min(observed, policy.tail_seconds)
        strict_start = max(0.0, observed - strict_dur)
        window = g[g["t_elapsed"] > strict_start]
        if accepted or not moving:
            window_incl, incl_dur = window, strict_dur
        else:
            window_incl, incl_dur = select_tail_window(g, observed, policy, accepted=False)
        incl_start = max(0.0, observed - incl_dur)
        # --- Torque 분류(STRICT 창 기준; 05와 동일)
        in_range = window[(window["Torque_pct"] >= policy.torque_min_pct) & (window["Torque_pct"] <= policy.torque_max_pct)]
        torque = window["Torque_pct"].dropna()
        n_all = int(len(window))
        n_in = int(len(in_range))
        n_low = int((torque < policy.torque_min_pct).sum())
        n_high = int((torque > policy.torque_max_pct).sum())
        first5 = g[g["t_elapsed"] <= min(5.0, observed)]["Torque_pct"].mean()
        if n_all == 0 or torque.empty:
            tq_mode = "NO_POINTS"
        elif n_in == n_all:
            tq_mode = "STRICT_VALID_ONLY"
        elif n_in > 0:
            tq_mode = "MIXED"
        else:
            tq_mode = "ALL_POINTS"                  # 전 점이 범위 밖 → inclusive만 전 점으로 계산(참고값)
        # --- 레거시 status(05와 동일 문구)
        status = "REST" if not moving else "OK"
        if moving and not strict_hold:
            status = f"HOLD NOT 60 s ({hold:g} s)"
        elif moving and in_range.empty:
            if torque.empty:
                status = "NO DATA"
            elif torque.max() < policy.torque_min_pct:
                status = "LOW TORQUE"
            elif torque.min() > policy.torque_max_pct:
                status = "OVER TORQUE"
            else:
                status = "NO VALID TORQUE"
        # --- STRICT 채널
        use = in_range if strict_hold else in_range.iloc[0:0]
        # --- INCLUSIVE 채널: 승인 hold(+정책상 범위 밖 점 포함). 미승인 짧은 hold도 참고값은 계산하되 포함 표시는 False.
        if moving and not too_short and n_all_incl(window_incl):
            src = window_incl if policy.include_out_of_torque_analysis else \
                window_incl[(window_incl["Torque_pct"] >= policy.torque_min_pct) & (window_incl["Torque_pct"] <= policy.torque_max_pct)]
        else:
            src = g.iloc[0:0]
        tau_incl = float(src["ShearStress_Pa"].mean()) if len(src) else np.nan
        eta_incl = float(src["Viscosity_cP"].mean()) if len(src) else np.nan
        censored = moving and tq_mode == "ALL_POINTS"
        included_strict = bool(moving and strict_hold and n_in > 0)
        # ALL_POINTS도 정책이 범위 밖 참고값을 inclusive에 넣으면 포함. strict는 in_range만 사용한다.
        oor_in_incl = inclusive_out_of_range_enabled(policy)
        included_incl = bool(moving and accepted and not too_short and len(src) > 0 and (not censored or oor_in_incl))
        if len(src):
            tq_src = src["Torque_pct"]
            used_low = bool((tq_src < policy.torque_min_pct).fillna(False).any())
            used_high = bool((tq_src > policy.torque_max_pct).fillna(False).any())
            n_low_src = int((tq_src < policy.torque_min_pct).fillna(False).sum())
            n_high_src = int((tq_src > policy.torque_max_pct).fillna(False).sum())
        else:
            used_low = used_high = False
            n_low_src = n_high_src = 0
        # 실제로 inclusive 평균에 넣은 범위 밖 점만 참고값으로 표시한다.
        used_low_eff = bool(included_incl and used_low)
        used_high_eff = bool(included_incl and used_high)
        if not moving:
            vc_strict, vc_incl = VC_NA, VC_NA
        elif too_short:
            vc_strict, vc_incl = VC_INVALID, VC_INVALID
        else:
            vc_strict = VC_STRICT if included_strict else VC_NO
            if len(src) == 0 or tq_mode == "NO_POINTS":
                vc_incl = VC_NO
            elif not included_incl:
                vc_incl = VC_NO
            elif used_low_eff or used_high_eff:
                vc_incl = _vc_from_torque_use(used_low_eff, used_high_eff, n_low_src, n_high_src)
            elif not strict_hold:
                vc_incl = VC_EXC
            else:
                vc_incl = VC_STRICT
        # --- 06 상태·라벨(Step 수준, 지시서 §E·§F·§N)
        if not moving:
            d_status, reason, label = "OK", "0 rpm stabilization; excluded from flow analysis", LBL_REST
        elif too_short:
            d_status, reason, label = "INVALID", f"hold {observed:g} s < {policy.min_hold_s:g} s — 대표값 계산 보류", LBL_TOO_SHORT
        elif tq_mode == "NO_POINTS":
            d_status, reason, label = "NO VALUE", "대표창에 point 없음", LBL_NO_POINTS
        else:
            parts, labels = [], []
            if hold_class == "STRICT_60":
                d_status = "OK"
            elif hold_class == "SHORT_ACCEPTED" and accepted:
                d_status = "EXCEPTION"
                parts.append(f"Observed hold {hold:g} s; accepted by {policy.name} policy, not strict {policy.hold_target_s:g} s")
                labels.append(LBL_SHORT_INCL.format(h=hold))
            elif hold_class == "LONG":
                d_status = "REVIEW"
                parts.append(f"hold {hold:g} s > {policy.hold_target_s:g} s — 후반 {policy.tail_seconds:g} s 사용")
                labels.append(LBL_LONG.format(h=hold))
            else:
                d_status = "REVIEW"
                parts.append(f"hold {hold:g} s 미승인(accepted={','.join(f'{x:g}' for x in policy.accepted_hold_s)}) — 참고값만 계산, 통계 제외")
                labels.append(LBL_SHORT_UNACC.format(h=hold))
            if tq_mode == "MIXED":
                d_status = _rank_max(d_status, "REVIEW")
                parts.append(f"Torque 범위 내 {n_in}/{n_all} point (low {n_low}, high {n_high}) — strict는 범위 내 점만, inclusive는 전 점")
                labels.append(LBL_MIXED_TQ)
            elif tq_mode == "ALL_POINTS":
                d_status = _rank_max(d_status, "CENSORED")
                if n_low > 0 and n_high > 0:
                    parts.append(
                        "범위 내 point 없음 또는 범위 밖 point 동시 사용; "
                        f"Torque 하한 미만 {n_low}개 및 상한 초과 {n_high}개; "
                        "inclusive 참고값; 합격성 판정 제외"
                    )
                    labels.append(LBL_MIXED_TQ)
                elif n_low > 0:
                    parts.append(f"전 point Torque < {policy.torque_min_pct:g}% (max {torque.max():.1f}%) — CENSORED_LOW_TORQUE, inclusive 참고값만")
                    labels.append(LBL_LOW_TQ)
                else:
                    parts.append(f"전 point Torque > {policy.torque_max_pct:g}% (min {torque.min():.1f}%) — CENSORED_HIGH_TORQUE, inclusive 참고값만")
                    labels.append(LBL_HIGH_TQ)
            if not labels:
                labels.append(LBL_STRICT)
            reason = "; ".join(parts) if parts else f"{policy.hold_target_s:g} s hold, Torque 전 point 범위 내"
            label = "|".join(labels)
        elapsed = sorted(float(x) for x in window["t_elapsed"] if _finite_num(x)) if len(window) else []
        gaps = np.diff(elapsed) if len(elapsed) > 1 else np.array([])
        median_dt = float(np.median(gaps)) if len(gaps) else np.nan
        if not _finite_num(median_dt) and _finite_num(mode.get("avg_s")):
            median_dt = float(mode["avg_s"])
        max_gap = float(np.max(gaps)) if len(gaps) else np.nan
        gap_status = "OK"
        if _finite_num(max_gap) and _finite_num(median_dt) and max_gap > max(median_dt * 1.5, median_dt + 0.2):
            gap_status = "NOTICE"
        if policy.sampling_gap_review_s is not None and _finite_num(max_gap) and max_gap > float(policy.sampling_gap_review_s):
            gap_status = "REVIEW"
        raw_visc = float(g["Viscosity_cP"].mean()) if g["Viscosity_cP"].notna().any() else np.nan
        records.append({
            # ---- 레거시(05 동일) ----
            "Step": step, "dir": direction[step],
            "step_type": seq_plan[step]["step_type"],
            "moving_flag": bool(moving),
            "is_rest": not bool(moving),
            "is_stabilization": seq_plan[step]["step_type"] == "REST_STABILIZATION",
            "sequence_direction": seq_plan[step]["sequence_direction"],
            "used_in_up_branch": bool(seq_plan[step]["used_in_up_branch"]),
            "used_in_down_branch": bool(seq_plan[step]["used_in_down_branch"]),
            "turnaround_shared": bool(seq_plan[step]["turnaround_shared"]),
            "RPM": float(g["Speed_RPM"].median()),
            "gamma": float(g["ShearRate"].median()), "hold_s": hold,
            "observed_s": observed, "hold_compliant": bool(strict_hold or not moving),
            "window_start_s": strict_start, "window_end_s": float(observed), "window_duration_s": strict_dur,
            "median_sampling_interval_s": median_dt, "max_sampling_gap_s": max_gap, "sampling_gap_status": gap_status,
            "raw_viscosity_cP": raw_visc,
            "tau_Pa": float(use["ShearStress_Pa"].mean()) if len(use) else np.nan,
            "eta_cP": float(use["Viscosity_cP"].mean()) if len(use) else np.nan,
            "torque_valid": float(use["Torque_pct"].mean()) if len(use) else np.nan,
            "torque_window": float(torque.mean()) if len(torque) else np.nan,
            "torque_first5": float(first5) if pd.notna(first5) else np.nan,
            "temp": float(window["Temp_C"].mean()) if len(window) and window["Temp_C"].notna().any() else np.nan,
            "temp_min": float(window["Temp_C"].min()) if len(window) and window["Temp_C"].notna().any() else np.nan,
            "temp_max": float(window["Temp_C"].max()) if len(window) and window["Temp_C"].notna().any() else np.nan,
            "temp_sd": float(window["Temp_C"].std(ddof=1)) if window["Temp_C"].notna().sum() > 1 else np.nan,
            "accuracy_cP": float(use["Accuracy_cP"].mean()) if len(use) else np.nan,
            "n_window": n_all, "n_valid": int(len(use)),
            "n_valid_observed": n_in,
            "valid_pct": float(len(use) / n_all * 100.0) if n_all else np.nan,
            "status": status,
            # ---- 06 채널 필드(지시서 §D) ----
            "hold_target_s": policy.hold_target_s, "hold_observed_s": duration_observed, "hold_class": hold_class,
            "strict_hold_valid": bool(strict_hold), "analysis_hold_accepted": bool(accepted),
            "torque_valid_n": n_in, "torque_all_n": n_all, "n_low_torque": n_low, "n_high_torque": n_high,
            "strict_torque_valid": bool(n_in > 0), "analysis_torque_mode": tq_mode,
            "included_in_strict": included_strict, "included_in_inclusive": included_incl,
            "exception_accepted": bool(moving and accepted and hold_class == "SHORT_ACCEPTED" and len(src) > 0),
            "used_low_torque_in_inclusive": used_low_eff,
            "used_high_torque_in_inclusive": used_high_eff,
            "used_out_of_torque_in_inclusive": bool(used_low_eff or used_high_eff),
            "value_class_strict": vc_strict, "value_class_inclusive": vc_incl, "value_class": vc_incl,
            "data_status": d_status, "status_reason": reason, "label": label,
            "tau_Pa_strict": float(use["ShearStress_Pa"].mean()) if len(use) else np.nan,
            "eta_cP_strict": float(use["Viscosity_cP"].mean()) if len(use) else np.nan,
            "tau_Pa_inclusive": tau_incl, "eta_cP_inclusive": eta_incl,
            "torque_inclusive": float(src["Torque_pct"].mean()) if len(src) else np.nan,
            "accuracy_cP_inclusive": float(src["Accuracy_cP"].mean()) if len(src) else np.nan,
            "torque_mean_all": float(torque.mean()) if len(torque) else np.nan,
            "torque_min_all": float(torque.min()) if len(torque) else np.nan,
            "torque_max_all": float(torque.max()) if len(torque) else np.nan,
            "tail_window_s": incl_dur, "tail_start_s": incl_start, "tail_end_s": observed,
            "n_inclusive": int(len(src)),
        })
    return pd.DataFrame(records)


def n_all_incl(w):
    return int(len(w)) > 0


_STEP_STATUS_ORDER = {"OK": 0, "NOTICE": 1, "REVIEW": 2, "CENSORED": 3, "EXCEPTION": 4, "NO VALUE": 5, "INVALID": 6}


def _rank_max(a, b):
    return a if _STEP_STATUS_ORDER.get(a, 0) >= _STEP_STATUS_ORDER.get(b, 0) else b


def channel_view(reps, channel):
    """analyze_step_flow / compute_a_qc가 그대로 소비할 수 있는 채널별 reps 뷰를 만든다.

    strict    : 레거시 열 그대로(05 동일).
    inclusive : tau_Pa/eta_cP ← inclusive 값, n_valid ← n_inclusive(포함 Step만), status ← 06 data_status 기반.
    """
    v = reps.copy()
    if channel == "strict":
        v["exception_accepted"] = False
        v["value_class"] = v["value_class_strict"] if "value_class_strict" in v.columns else VC_STRICT
        v["used_low_torque"] = False
        v["used_high_torque"] = False
        if "included_in_strict" in v.columns and "n_valid" in v.columns:
            v["n_valid"] = np.where(v["included_in_strict"].fillna(False).astype(bool), v["n_valid"], 0).astype(int)
        return v
    v["tau_Pa"] = v["tau_Pa_inclusive"]
    v["eta_cP"] = v["eta_cP_inclusive"]
    v["torque_valid"] = v["torque_inclusive"]
    v["accuracy_cP"] = v["accuracy_cP_inclusive"]
    v["n_valid"] = np.where(v["included_in_inclusive"], v["n_inclusive"], 0).astype(int)
    v["hold_compliant"] = v["analysis_hold_accepted"] | (v["dir"] == "Rest")
    v["window_start_s"] = v["tail_start_s"]
    v["window_duration_s"] = v["tail_window_s"]
    v["value_class"] = v["value_class_inclusive"] if "value_class_inclusive" in v.columns else VC_NO
    v["used_low_torque"] = v["used_low_torque_in_inclusive"] if "used_low_torque_in_inclusive" in v.columns else False
    v["used_high_torque"] = v["used_high_torque_in_inclusive"] if "used_high_torque_in_inclusive" in v.columns else False
    def _st(r):
        if r["dir"] == "Rest":
            return "REST"
        if r["included_in_inclusive"]:
            return "OK" if r["data_status"] == "OK" else f"{r['data_status']}: {r['label']}"
        return f"EXCLUDED: {r['label']}"
    v["status"] = v.apply(_st, axis=1)
    return v


def hb_model(x, tau_y, consistency, exponent):
    return tau_y + consistency * np.power(x, exponent)


def _hb_contributor_meta(points):
    """H-B에 실제 들어간 Step의 구조화 필드만 집계한다. 상태 문자열은 파싱하지 않는다."""
    empty = {"contributing_steps": [], "contributing_rpms": [], "contributing_gammas": [],
             "n_points": 0, "gamma_min": np.nan, "gamma_max": np.nan, "steps": [],
             "n_strict_steps": 0, "n_exception_steps": 0, "n_low_torque_steps": 0,
             "n_high_torque_steps": 0, "n_mixed_torque_steps": 0}
    if points is None or len(points) == 0:
        return empty
    steps, rpms, gammas = [], [], []
    n_strict = n_exc = n_low = n_high = n_mixed = 0
    for rec in points.to_dict("records"):
        if _finite_num(rec.get("Step")):
            steps.append(int(rec["Step"]))
        if _finite(rec.get("RPM")):
            rpms.append(float(rec["RPM"]))
        if _finite(rec.get("gamma")):
            gammas.append(float(rec["gamma"]))
        vc = str(rec.get("value_class") or "")
        low = bool(rec.get("used_low_torque")) or vc in (VC_LOW, VC_BOTH)
        high = bool(rec.get("used_high_torque")) or vc in (VC_HIGH, VC_BOTH)
        if low:
            n_low += 1
        if high:
            n_high += 1
        if str(rec.get("analysis_torque_mode") or "") == "MIXED":
            n_mixed += 1
        hold = str(rec.get("hold_class") or "")
        if bool(rec.get("exception_accepted")) or hold in ("SHORT_ACCEPTED", "LONG") or vc == VC_EXC:
            n_exc += 1
        elif vc == VC_STRICT or (hold == "STRICT_60" and not low and not high):
            n_strict += 1
    empty.update({
        "contributing_steps": steps, "contributing_rpms": rpms, "contributing_gammas": gammas,
        "n_points": int(len(points)), "steps": steps,
        "gamma_min": float(min(gammas)) if gammas else np.nan,
        "gamma_max": float(max(gammas)) if gammas else np.nan,
        "n_strict_steps": n_strict, "n_exception_steps": n_exc,
        "n_low_torque_steps": n_low, "n_high_torque_steps": n_high, "n_mixed_torque_steps": n_mixed,
    })
    return empty


def fit_hb_points(points, label):
    from scipy.optimize import curve_fit
    points = points.dropna(subset=["gamma", "tau_Pa"]).copy()
    points = points[(points["gamma"] > 0) & (points["n_valid"] > 0)]
    meta = _hb_contributor_meta(points)
    if len(points) < METHOD_A_MIN_HB_POINTS:
        return {**meta, "label": label, "status": "INVALID — 유효점 4개 미만"}
    gamma = points["gamma"].to_numpy(float)
    tau = points["tau_Pa"].to_numpy(float)
    sigma = points["accuracy_cP"].to_numpy(float) * gamma / 1000.0
    weighted = bool(np.all(np.isfinite(sigma) & (sigma > 0)))
    try:
        p0 = [max(float(tau.min()) * 0.25, 0.1), max(float(np.ptp(tau)), 1.0), 0.8]
        popt, pcov = curve_fit(hb_model, gamma, tau, p0=p0,
                               sigma=sigma if weighted else None, absolute_sigma=weighted,
                               bounds=([0, 0, 0.05], [np.inf, np.inf, 2.0]), maxfev=200000)
    except Exception as exc:
        return {**meta, "label": label, "status": f"피팅 실패: {exc}"}
    pred = hb_model(gamma, *popt)
    ss_res = float(np.sum((tau - pred) ** 2))
    ss_tot = float(np.sum((tau - tau.mean()) ** 2))
    stderr = np.sqrt(np.diag(pcov)) if np.shape(pcov) == (3, 3) else np.full(3, np.nan)
    ratio = float(gamma.max() / gamma.min())
    warnings = []
    if len(points) == METHOD_A_MIN_HB_POINTS:
        warnings.append("REVIEW — 유효점 4개(최소조건); 5개 이상 권장")
    if ratio < 10:
        warnings.append(f"전단률 범위 1 decade 미만({ratio:.2f}배)")
    if not weighted:
        warnings.append("비가중 피팅")
    if not np.all(np.isfinite(stderr)):
        warnings.append("표준오차 불안정")
    return {
        **meta,
        "label": label, "status": "OK" if not warnings else "; ".join(warnings),
        "tau_y": float(popt[0]), "K": float(popt[1]), "n": float(popt[2]),
        "R2": float(1 - ss_res / ss_tot) if ss_tot > 0 else np.nan,
        "se_tau_y": float(stderr[0]), "se_K": float(stderr[1]), "se_n": float(stderr[2]),
        "weighted": weighted,
    }


def _branch_frame(reps, branch):
    """방향별 fitting view. Turnaround Step은 행을 복제해 저장하지 않고 양쪽 frame이 같은 Step을 참조한다."""
    col = "used_in_up_branch" if branch == "up" else "used_in_down_branch"
    if col in reps.columns:
        sub = reps[reps[col].fillna(False).astype(bool) & (reps["n_valid"] > 0) & (reps["RPM"].abs() > RPM_ZERO_TOL)]
    else:
        sub = reps[(reps["dir"] == ("Up" if branch == "up" else "Down")) & (reps["n_valid"] > 0)]
    return sub.sort_values("Step")


def _moving_mask(reps):
    return reps["dir"].isin(["Up", "Down", "Turnaround"])


def _shear_match_tol():
    return RPM_MATCH_TOL * float(CP52_SHEAR_RATE_CONSTANT)


def _pair_up_down(up, down):
    """공유 turnaround는 같은 Step으로, 나머지는 RPM·shear-rate 허용오차 안의 1:1 pair. 보간하지 않는다."""
    up_rows = up.sort_values("Step").to_dict("records")
    down_rows = down.sort_values("Step").to_dict("records")
    used_down = set()
    pairs = []

    def _close(u, d):
        rpm_ok = abs(float(u["RPM"]) - float(d["RPM"])) <= RPM_MATCH_TOL
        gamma_ok = abs(float(u["gamma"]) - float(d["gamma"])) <= _shear_match_tol()
        return rpm_ok and gamma_ok

    for u in up_rows:
        if not bool(u.get("turnaround_shared")):
            continue
        for i, d in enumerate(down_rows):
            if i in used_down:
                continue
            if int(d["Step"]) == int(u["Step"]):
                pairs.append((u, d, "shared turnaround"))
                used_down.add(i)
                break
    paired_up = {int(u["Step"]) for u, _d, _rule in pairs}
    for u in up_rows:
        if int(u["Step"]) in paired_up:
            continue
        best_i, best_key = None, None
        for i, d in enumerate(down_rows):
            if i in used_down or not _close(u, d):
                continue
            key = (abs(float(u["RPM"]) - float(d["RPM"])), abs(float(u["gamma"]) - float(d["gamma"])), int(d["Step"]))
            if best_key is None or key < best_key:
                best_i, best_key = i, key
        if best_i is not None:
            pairs.append((u, down_rows[best_i], "rpm tolerance"))
            used_down.add(best_i)
            paired_up.add(int(u["Step"]))
    pairs.sort(key=lambda item: (float(item[0]["gamma"]), int(item[0]["Step"])))
    return pairs


def analyze_step_flow(reps):
    up = _branch_frame(reps, "up")
    down = _branch_frame(reps, "down")
    paired = _pair_up_down(up, down)
    pair_steps = {int(u["Step"]) for u, _d, _rule in paired} | {int(d["Step"]) for _u, d, _rule in paired}
    up_common = up[up["Step"].isin(pair_steps)] if len(up) else up
    down_common = down[down["Step"].isin(pair_steps)] if len(down) else down
    common = [float(u["gamma"]) for u, _d, _rule in paired]
    fits = {
        "up_full": fit_hb_points(up, "Up Full"),
        "down_full": fit_hb_points(down, "Down Full"),
        "up_common": fit_hb_points(up_common, "Up Common"),
        "down_common": fit_hb_points(down_common, "Down Common"),
    }
    rows = []
    signed_area = abs_area = up_area = 0.0
    prev = None
    for u, d, rule in paired:
        gamma = float(u["gamma"])
        signed = float(u["tau_Pa"] - d["tau_Pa"])
        absolute = abs(signed)
        same_step = int(u["Step"]) == int(d["Step"])
        row = {"gamma": gamma, "gamma_up": gamma, "gamma_down": float(d["gamma"]),
               "rpm_up": float(u["RPM"]), "rpm_down": float(d["RPM"]),
               "step_up": int(u["Step"]), "step_down": int(d["Step"]),
               "tau_up": float(u["tau_Pa"]), "tau_down": float(d["tau_Pa"]),
               "pair_rule": rule,
               "delta_signed": signed, "delta_abs": absolute,
               "segment_signed": np.nan, "segment_abs": np.nan,
               "segment_up": np.nan, "recovery_pct": float(d["tau_Pa"] / u["tau_Pa"] * 100) if u["tau_Pa"] else np.nan,
               "turnaround_shared": bool(same_step and u.get("turnaround_shared"))}
        if prev is not None:
            width = row["gamma"] - prev["gamma"]
            row["segment_signed"] = width * (row["delta_signed"] + prev["delta_signed"]) / 2
            row["segment_abs"] = width * (row["delta_abs"] + prev["delta_abs"]) / 2
            row["segment_up"] = width * (row["tau_up"] + prev["tau_up"]) / 2
            signed_area += row["segment_signed"]
            abs_area += row["segment_abs"]
            up_area += row["segment_up"]
        rows.append(row)
        prev = row
    relative = abs_area / up_area * 100 if up_area > 0 else np.nan
    valid_recovery = [r["recovery_pct"] for r in rows if np.isfinite(r["recovery_pct"])]
    low_recovery = valid_recovery[0] if valid_recovery else np.nan
    mean_recovery = float(np.mean(valid_recovery)) if valid_recovery else np.nan
    comp = {}
    uc, dc = fits["up_common"], fits["down_common"]
    if uc.get("tau_y") is not None and dc.get("tau_y") is not None:
        comp = {
            "tau_y_ratio_pct": dc["tau_y"] / uc["tau_y"] * 100 if uc["tau_y"] else np.nan,
            "K_ratio_pct": dc["K"] / uc["K"] * 100 if uc["K"] else np.nan,
            "n_difference": dc["n"] - uc["n"],
        }
    comp_full = {}
    uf, dfull = fits["up_full"], fits["down_full"]
    if uf.get("tau_y") is not None and dfull.get("tau_y") is not None:
        comp_full = {
            "tau_y_ratio_pct": dfull["tau_y"] / uf["tau_y"] * 100 if uf["tau_y"] else np.nan,
            "K_ratio_pct": dfull["K"] / uf["K"] * 100 if uf["K"] else np.nan,
            "n_difference": dfull["n"] - uf["n"],
        }
    shared_rows = [r for r in rows if r.get("turnaround_shared")]
    turn = {}
    if shared_rows:
        t = shared_rows[-1]
        turn = {"shared": True, "step": int(t["step_up"]), "rpm": float(t["rpm_up"]), "gamma": float(t["gamma"]),
                "delta_tau": float(t["delta_signed"])}
    elif "turnaround_shared" in reps.columns and bool(reps["turnaround_shared"].fillna(False).astype(bool).any()):
        src = reps[reps["turnaround_shared"].fillna(False).astype(bool)].iloc[0]
        turn = {"shared": False, "step": int(src["Step"]), "rpm": float(src["RPM"]), "gamma": float(src["gamma"]),
                "delta_tau": np.nan, "excluded_from_channel": True}
    rpms = [float(r["rpm_up"]) for r in rows]
    return {"fits": fits, "points": rows, "common_gammas": common, "pair_steps": sorted(pair_steps),
            "signed_area": float(signed_area), "absolute_area": float(abs_area),
            "up_area": float(up_area), "relative_area_pct": float(relative),
            "low_shear_recovery_pct": float(low_recovery), "mean_recovery_pct": float(mean_recovery),
            "comparison": comp, "comparison_full": comp_full,
            "hysteresis_gamma_min": float(rows[0]["gamma"]) if rows else np.nan,
            "hysteresis_gamma_max": float(rows[-1]["gamma"]) if rows else np.nan,
            "hysteresis_rpm_min": float(min(rpms)) if rpms else np.nan,
            "hysteresis_rpm_max": float(max(rpms)) if rpms else np.nan,
            "hysteresis_steps_up": [int(r["step_up"]) for r in rows],
            "hysteresis_steps_down": [int(r["step_down"]) for r in rows],
            "turnaround": turn}


def map_3itt(df):
    rpm = df.groupby("Step")["Speed_RPM"].median().sort_index()
    moving = rpm[rpm > 0]
    if len(moving) != 3:
        raise ValueError(f"3ITT는 Rest(선택)+Baseline+Breakdown+Recovery의 회전 Step 3개가 필요합니다: 현재 {len(moving)}개")
    steps = list(moving.index.astype(int))
    if not (moving.iloc[1] > moving.iloc[0] and math.isclose(moving.iloc[0], moving.iloc[2], rel_tol=0.02, abs_tol=0.01)):
        raise ValueError("3ITT Step은 저전단 → 고전단 → 동일 저전단 순서여야 합니다.")
    mapping = {int(s): 1 for s, v in rpm.items() if v <= 0}
    mapping.update({steps[0]: 2, steps[1]: 3, steps[2]: 4})
    return mapping


def final_window(group, hold, win_s, fraction, tq_lo, tq_hi):
    end = float(group["t_elapsed"].max())
    duration = min(end, max(win_s, hold * fraction))
    w = group[group["t_elapsed"] > max(0, end - duration)]
    return w[(w["Torque_pct"] >= tq_lo) & (w["Torque_pct"] <= tq_hi)], w


def analyze_3itt(df, holds, win_s=10, fraction=0.25, tq_lo=10, tq_hi=95):
    mapping = map_3itt(df)
    groups = {interval: df[[mapping[int(s)] == interval for s in df["Step"]]].copy() for interval in (1, 2, 3, 4)}
    summary = {}
    for interval, name in ((2, "Baseline"), (3, "Breakdown")):
        g = groups[interval]
        step = int(g["Step"].iloc[0])
        valid, window = final_window(g, holds.get(step, g["t_elapsed"].max()), win_s, fraction, tq_lo, tq_hi)
        summary[name] = {
            "step": step, "rpm": float(g["Speed_RPM"].median()), "gamma": float(g["ShearRate"].median()),
            "hold_s": float(holds.get(step, g["t_elapsed"].max())),
            "tau": float(valid["ShearStress_Pa"].mean()) if len(valid) else np.nan,
            "eta": float(valid["Viscosity_cP"].mean()) if len(valid) else np.nan,
            "torque": float(window["Torque_pct"].mean()) if len(window) else np.nan,
            "n_valid": int(len(valid)), "status": "OK" if len(valid) else "NO VALID TORQUE",
        }
    base = summary["Baseline"]
    rec = groups[4].copy()
    rec_valid = rec[(rec["Torque_pct"] >= tq_lo) & (rec["Torque_pct"] <= tq_hi)].copy()
    rows = []
    for t in RECOVERY_TIMES:
        w = rec_valid[(rec_valid["t_elapsed"] >= max(1, t - 5)) & (rec_valid["t_elapsed"] <= t + 5)]
        tau = float(w["ShearStress_Pa"].mean()) if len(w) else np.nan
        eta = float(w["Viscosity_cP"].mean()) if len(w) else np.nan
        rows.append({"time_s": t, "tau": tau, "eta": eta,
                     "recovery_tau_pct": tau / base["tau"] * 100 if np.isfinite(tau) and np.isfinite(base["tau"]) and base["tau"] else np.nan,
                     "recovery_eta_pct": eta / base["eta"] * 100 if np.isfinite(eta) and np.isfinite(base["eta"]) and base["eta"] else np.nan,
                     "torque": float(w["Torque_pct"].mean()) if len(w) else np.nan, "n_valid": int(len(w))})
    target_times = {50: np.nan, 80: np.nan, 90: np.nan}
    if np.isfinite(base["eta"]) and base["eta"] > 0 and len(rec_valid):
        rolling = rec_valid.sort_values("t_elapsed").copy()
        rolling["eta_smooth"] = rolling["Viscosity_cP"].rolling(11, center=True, min_periods=3).median()
        rolling["recovery"] = rolling["eta_smooth"] / base["eta"] * 100
        for target in target_times:
            hit = rolling[rolling["recovery"] >= target]
            if len(hit):
                target_times[target] = float(hit["t_elapsed"].iloc[0])
    recovery_step = int(rec["Step"].iloc[0])
    return {"mapping": mapping, "summary": summary, "recovery": rows, "target_times": target_times,
            "recovery_step": recovery_step, "recovery_hold_s": float(holds.get(recovery_step, rec["t_elapsed"].max())),
            "recovery_valid_pct": float(len(rec_valid) / len(rec) * 100) if len(rec) else np.nan}


def py_value(value, integer=False):
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return None
    if isinstance(value, np.generic):
        value = value.item()
    if isinstance(value, str):
        return value
    return int(round(float(value))) if integer else float(value)


def put(ws, row, col, value, fmt=None, integer=False):
    c = ws.cell(row=row, column=col)
    c.value = py_value(value, integer)
    if fmt:
        c.number_format = fmt
    return c


def clear_area(ws, r1, r2, c1, c2):
    for row in ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2):
        for c in row:
            c.value = None


def write_metadata(ws, meta, source):
    ws["C4"] = Path(source).name
    ws["C5"] = meta.get("Test Start", "")
    ws["C6"] = meta.get("Instrument", "")
    ws["C7"] = meta.get("Spindle", "")


def write_raw(ws, df, key_map, capacity=5000):
    if len(df) > capacity:
        raise ValueError(f"Raw_Data 용량 {capacity}행 초과: {len(df)}")
    clear_area(ws, 6, 5 + capacity, 2, 15)
    for i, r in enumerate(df.itertuples(), 6):
        key = key_map[int(r.Step)]
        vals = [key, int(r.Step), int(r.Point), r.t_elapsed, r.Speed_RPM, r.ShearRate,
                r.ShearStress_Pa, r.Viscosity_cP, r.Torque_pct, r.Temp_C, r.Accuracy_cP,
                r.ShearRate, r.ShearStress_Pa / r.ShearRate * 1000 if r.ShearRate else None,
                "OK" if DEFAULT_TORQUE_MIN <= r.Torque_pct <= DEFAULT_TORQUE_MAX else "FLAG"]
        fmts = ["0", "0", "0", "0.0", "0.000", "0.000", "0.000", "#,##0.0", "0.0", "0.00", "#,##0.0", "0.000", "#,##0.0", None]
        for j, (v, fmt) in enumerate(zip(vals, fmts), 2):
            put(ws, i, j, v, fmt, integer=j in (2, 3, 4))


def write_step_flow(wb, df, holds, meta, reps, analysis, source):
    write_metadata(wb["QC_Summary"], meta, source)
    setup = wb["Step_Setup"]
    clear_area(setup, 6, 45, 2, 8)
    for i, r in enumerate(reps.itertuples(), 6):
        vals = [r.Step, r.dir, r.RPM, r.gamma, r.hold_s, "Yes" if r.n_valid > 0 and r.dir != "Rest" else "No", r.status]
        for j, v in enumerate(vals, 2):
            put(setup, i, j, v, "0.000" if j in (4, 5) else ("0" if j in (2, 6) else None), integer=j in (2, 6))
    write_raw(wb["Raw_Data"], df, {int(s): int(s) for s in df["Step"].unique()})
    sr = wb["Step_Results"]
    clear_area(sr, 6, 45, 2, 21)
    fields = ["Step", "dir", "RPM", "gamma", "hold_s", "window_start_s", "tau_Pa", "eta_cP",
              "torque_valid", "torque_window", "torque_first5", "temp", "n_window", "n_valid", "valid_pct", "status",
              "observed_s", "hold_compliant", "window_duration_s"]
    formats = ["0", None, "0.000", "0.000", "0", "0.0", "0.000", "#,##0.0", "0.0", "0.0", "0.0", "0.00", "0", "0", "0.0", None,
               "0.0", None, "0.0"]
    for i, (_, r) in enumerate(reps.iterrows(), 6):
        for j, (field, fmt) in enumerate(zip(fields, formats), 2):
            value = "Yes" if field == "hold_compliant" and bool(r[field]) else ("No" if field == "hold_compliant" else r[field])
            put(sr, i, j, value, fmt, integer=field in ("Step", "hold_s", "n_window", "n_valid"))
    fp = wb["Flow_Profile"]
    fit_keys = ("up_full", "down_full", "up_common", "down_common")
    for row, key in enumerate(fit_keys, 6):
        fit = analysis["fits"][key]
        vals = [fit.get("label"), fit.get("gamma_min"), fit.get("gamma_max"), fit.get("n_points"), fit.get("tau_y"), fit.get("K"), fit.get("n"), fit.get("R2"), fit.get("se_tau_y"), fit.get("se_K"), fit.get("se_n"), "Accuracy" if fit.get("weighted") else "None", fit.get("status")]
        for col, value in enumerate(vals, 2):
            put(fp, row, col, value, "0.0000" if 3 <= col <= 12 else None, integer=col == 5)
    metrics = [
        ("Down/Up τy (Common)", analysis["comparison"].get("tau_y_ratio_pct"), "%"),
        ("Down/Up K (Common)", analysis["comparison"].get("K_ratio_pct"), "%"),
        ("n_down - n_up (Common)", analysis["comparison"].get("n_difference"), "-"),
        ("최저 공통 전단률 응력비", analysis["low_shear_recovery_pct"], "%"),
        ("공통점 평균 응력비", analysis["mean_recovery_pct"], "%"),
        ("Signed hysteresis area", analysis["signed_area"], "Pa/s"),
        ("Absolute hysteresis area", analysis["absolute_area"], "Pa/s"),
        ("Relative hysteresis area", analysis["relative_area_pct"], "%"),
    ]
    for row, (name, value, unit) in enumerate(metrics, 14):
        fp.cell(row, 2).value = name
        put(fp, row, 3, value, "0.000")
        fp.cell(row, 4).value = unit
    clear_area(fp, 27, 60, 2, 9)
    for row, p in enumerate(analysis["points"], 27):
        vals = [p[k] for k in ("gamma", "tau_up", "tau_down", "delta_signed", "delta_abs", "segment_signed", "segment_abs", "recovery_pct")]
        for col, value in enumerate(vals, 2):
            put(fp, row, col, value, "0.000")


def write_3itt(wb, df, holds, meta, result, source):
    write_metadata(wb["QC_Summary"], meta, source)
    setup = wb["Step_Setup"]
    clear_area(setup, 6, 12, 2, 8)
    reverse = {v: k for k, v in result["mapping"].items()}
    names = {1: "Rest", 2: "Baseline", 3: "Breakdown", 4: "Recovery"}
    for interval in (1, 2, 3, 4):
        step = reverse.get(interval)
        row = 5 + interval
        if step is None:
            continue
        g = df[df["Step"] == step]
        vals = [interval, names[interval], float(g["Speed_RPM"].median()), float(g["ShearRate"].median()), holds.get(step, g["t_elapsed"].max()), int(step), "실측 CSV"]
        for col, value in enumerate(vals, 2):
            put(setup, row, col, value, "0.000" if col in (4, 5) else ("0" if col in (2, 6, 7) else None), integer=col in (2, 6, 7))
    write_raw(wb["Raw_Data"], df, result["mapping"])
    rec = wb["Recovery"]
    for row, name in ((6, "Baseline"), (7, "Breakdown")):
        s = result["summary"][name]
        vals = [name, s["rpm"], s["gamma"], s["hold_s"], s["tau"], s["eta"], s["torque"], s["n_valid"], s["status"]]
        for col, value in enumerate(vals, 2):
            put(rec, row, col, value, "0.000" if col not in (2, 9, 10) else None, integer=col in (5, 9))
    clear_area(rec, 13, 25, 2, 8)
    for row, p in enumerate(result["recovery"], 13):
        vals = [p["time_s"], p["tau"], p["eta"], p["recovery_tau_pct"], p["recovery_eta_pct"], p["torque"], p["n_valid"]]
        for col, value in enumerate(vals, 2):
            put(rec, row, col, value, "0.000" if col in (3, 4) else "0.0", integer=col in (2, 8))
    for row, target in enumerate((50, 80, 90), 29):
        rec.cell(row, 2).value = f"t{target}"
        put(rec, row, 3, result["target_times"][target], "0.0")
    put(rec, 33, 3, result["recovery_valid_pct"], "0.0")


class TemplateSchemaError(RuntimeError):
    pass


STEPFLOW_TEMPLATE_SCHEMA = {
    "required_sheets": ["Overview", "README", "QC_Summary", "Step_Setup", "Raw_Data", "Step_Results",
                        "Flow_Profile", "Interpretation", "Model_Validation", "StartUp"],
    "headers": {
        "Overview": {"B2": "SST gel 종합 분석 Overview"},
        "README": {"B2": "SST gel Step Flow QC Template"},
        "QC_Summary": {"B2": "SST gel Step Flow — QC Summary", "B4": "원본 파일", "C4": "내용"},
        "Step_Setup": {"B5": "Step", "C5": "방향", "D5": "RPM", "E5": "전단률(1/s)", "F5": "유지시간(s)", "G5": "정량사용", "H5": "상태"},
        "Raw_Data": {
            "B5": "구분/Interval", "C5": "원본 Step", "D5": "Point", "E5": "경과시간(s)", "F5": "RPM",
            "G5": "전단률 γ(1/s)", "H5": "전단응력 τ(Pa)", "I5": "점도(cP)", "J5": "Torque(%)",
            "K5": "온도(°C)", "L5": "Accuracy(cP)", "M5": "γ 확인", "N5": "η 확인(cP)", "O5": "유효성",
        },
        "Step_Results": {
            "B5": "Step", "C5": "방향", "D5": "RPM", "E5": "γ(1/s)", "F5": "유지(s)", "G5": "창 시작(s)",
            "H5": "τ(Pa)", "I5": "η(cP)", "J5": "유효 Torque", "K5": "창 Torque", "L5": "초기5초 Torque",
            "M5": "온도", "N5": "창 n", "O5": "유효 n", "P5": "유효비(%)", "Q5": "상태",
        },
        "Flow_Profile": {"B5": "프로파일", "C5": "γ min", "N5": "상태"},
        "Model_Validation": {"B2": "SST gel 모델 검증"},
        "StartUp": {"B2": "SST gel 저전단 Start-up"},
    },
    "data_start": {"Raw_Data": "B6", "Step_Setup": "B6", "Step_Results": "B6"},
    "raw_capacity": 5000,
}


THREITT_TEMPLATE_SCHEMA = {
    "required_sheets": ["Overview", "README", "QC_Summary", "Step_Setup", "Raw_Data", "Recovery", "Interpretation"],
    "headers": {
        "Overview": {"B2": "SST gel 종합 분석 Overview"},
        "README": {"B2": "SST gel 3ITT QC Template"},
        "QC_Summary": {"B2": "SST gel 3ITT — QC Summary", "B4": "원본 파일", "C4": "내용"},
        "Step_Setup": {"B5": "Interval", "C5": "구간", "D5": "RPM", "E5": "전단률(1/s)", "F5": "유지시간(s)", "G5": "원본 Step", "H5": "출처"},
        "Raw_Data": {
            "B5": "구분/Interval", "C5": "원본 Step", "D5": "Point", "E5": "경과시간(s)", "F5": "RPM",
            "G5": "전단률 γ(1/s)", "H5": "전단응력 τ(Pa)", "I5": "점도(cP)", "J5": "Torque(%)",
            "K5": "온도(°C)", "L5": "Accuracy(cP)", "M5": "γ 확인", "N5": "η 확인(cP)", "O5": "유효성",
        },
        "Recovery": {"B5": "구간", "C5": "RPM", "D5": "γ(1/s)", "E5": "유지(s)", "F5": "τ(Pa)", "G5": "η(cP)",
                     "H5": "Torque(%)", "I5": "유효 n", "J5": "상태"},
    },
    "data_start": {"Raw_Data": "B6", "Step_Setup": "B6", "Recovery": "B6"},
    "raw_capacity": 5000,
}


def _merged_blocks(ws, coord):
    cell = ws[coord]
    for rng in ws.merged_cells.ranges:
        if cell.coordinate in rng and cell.coordinate != rng.start_cell.coordinate:
            return str(rng)
    return ""


def validate_template_schema(workbook, method, template_path):
    """고정 좌표에 쓰기 전에 첨부 템플릿의 시트와 header를 확인한다."""
    schema = STEPFLOW_TEMPLATE_SCHEMA if method == "A" else THREITT_TEMPLATE_SCHEMA if method == "B" else None
    if schema is None:
        raise TemplateSchemaError(f"unknown method {method}; template={template_path}")
    missing = [name for name in schema["required_sheets"] if name not in workbook.sheetnames]
    if missing:
        raise TemplateSchemaError(
            f"template={template_path}; method={method}; missing sheets={missing}; "
            f"actual sheets={workbook.sheetnames}"
        )
    for sheet, expected in schema["headers"].items():
        ws = workbook[sheet]
        for coord, text in expected.items():
            actual = ws[coord].value
            if actual != text:
                raise TemplateSchemaError(
                    f"template={template_path}; sheet={sheet}; cell={coord}; expected={text!r}; actual={actual!r}"
                )
    for sheet, coord in schema["data_start"].items():
        blocked = _merged_blocks(workbook[sheet], coord)
        if blocked:
            raise TemplateSchemaError(
                f"template={template_path}; sheet={sheet}; cell={coord} is blocked by merge {blocked}"
            )
    return schema


def prepare_workbook(template, output, method=None):
    wb = load_workbook(template)
    if method in ("A", "B"):
        validate_template_schema(wb, method, template)
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except AttributeError:
        pass
    Path(output).parent.mkdir(parents=True, exist_ok=True)
    return wb


def process(path, method, template, output, win_s, fraction, tq_lo, tq_hi):
    df, holds, meta = parse_dvnext_csv(path)
    wb = prepare_workbook(template, output, method)
    if method == "A":
        reps = step_representatives(df, holds, win_s, fraction, tq_lo, tq_hi)
        result = analyze_step_flow(reps)
        write_step_flow(wb, df, holds, meta, reps, result, path)
    else:
        result = analyze_3itt(df, holds, win_s, fraction, tq_lo, tq_hi)
        write_3itt(wb, df, holds, meta, result, path)
    wb.save(output)
    return len(df), int(df["Step"].nunique()), result


def collect_inputs(items, input_dir, recursive):
    paths = [Path(x).expanduser() for x in items]
    if input_dir:
        paths.append(Path(input_dir).expanduser())
    found = []
    for p in paths:
        if p.is_file() and p.suffix.lower() == ".csv":
            found.append(p.resolve())
        elif p.is_dir():
            found.extend(x.resolve() for x in p.glob("**/*.csv" if recursive else "*.csv"))
    unique = list(dict.fromkeys(found))
    if not unique:
        raise ValueError("처리할 CSV 파일이 없습니다.")
    return unique


def resolve_template(arg, method):
    name = A_TEMPLATE if method == "A" else B_TEMPLATE
    candidates = [Path(arg).expanduser()] if arg else [Path.cwd() / name, Path(__file__).resolve().parent / name]
    for p in candidates:
        if p.is_file():
            return p.resolve()
    if not arg:
        # 배포 파일명에 -5, -6 같은 접미사가 붙는 경우: 유일한 후보만 자동 채택한다.
        stem = Path(name).stem
        found = []
        for folder in dict.fromkeys([Path.cwd(), Path(__file__).resolve().parent]):
            found.extend(sorted(folder.glob(f"{stem}*.xlsx")))
        found = list(dict.fromkeys(p.resolve() for p in found))
        if len(found) == 1:
            return found[0]
        if len(found) > 1:
            raise FileNotFoundError(
                f"템플릿 후보가 여러 개입니다. --template(-a/-b)로 지정하세요: " + ", ".join(p.name for p in found))
    raise FileNotFoundError(f"템플릿을 찾을 수 없습니다: {name}")



from collections import defaultdict
from copy import copy
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import SeriesLabel
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.table import Table, TableStyleInfo

VERSION = "HB-analysis-06-6"
SAMPLE_RE = re.compile(r"^HT-(?P<lot>.+)-(?P<repeat>\d{2})(?:-\d+)?$", re.I)

NAVY = "1F4E78"
BLUE = "D9EAF7"
LIGHT_BLUE = "EAF3F8"
GREEN = "C6EFCE"
YELLOW = "FFEB9C"
RED = "FFC7CE"
GRAY = "E7E6E6"
WHITE = "FFFFFF"
THIN = Side(style="thin", color="B7B7B7")


def parse_sample_id(path):
    stem = Path(path).stem
    m = SAMPLE_RE.match(stem)
    if m:
        return m.group("lot"), m.group("repeat")
    return stem, "00"


def _sheet_base(ws, title, subtitle, max_col=8):
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "B5"
    ws.column_dimensions["A"].width = 3
    widths = [24, 18, 18, 18, 18, 18, 18, 28, 18, 18, 18, 18]
    for i, width in enumerate(widths, 2):
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.merge_cells(start_row=2, start_column=2, end_row=2, end_column=max_col)
    ws.cell(2, 2, title)
    ws.cell(2, 2).font = Font(name="Calibri", size=16, bold=True, color=WHITE)
    ws.cell(2, 2).fill = PatternFill("solid", fgColor=NAVY)
    ws.cell(2, 2).alignment = Alignment(vertical="center")
    ws.row_dimensions[2].height = 28
    ws.merge_cells(start_row=3, start_column=2, end_row=3, end_column=max_col)
    ws.cell(3, 2, subtitle)
    ws.cell(3, 2).font = Font(name="Calibri", size=10, italic=True, color="666666")
    ws.cell(3, 2).alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[3].height = 30


def _section(ws, row, text, c1=2, c2=8):
    ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
    cell = ws.cell(row, c1, text)
    cell.font = Font(name="Calibri", bold=True, color=WHITE)
    cell.fill = PatternFill("solid", fgColor=NAVY)
    cell.alignment = Alignment(vertical="center")
    ws.row_dimensions[row].height = 21


def _headers(ws, row, headers, start_col=2):
    for j, text in enumerate(headers, start_col):
        cell = ws.cell(row, j, text)
        cell.font = Font(name="Calibri", bold=True, color=WHITE)
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=THIN)
    ws.row_dimensions[row].height = 28


def _nav(ws, names, row=1):
    for col, name in enumerate(names, 2):
        cell = ws.cell(row, col, name)
        cell.hyperlink = Hyperlink(ref=cell.coordinate, location=f"'{name}'!B2")
        cell.font = Font(name="Calibri", size=9, color="0563C1", underline="single")


def ensure_extended_sheets(wb, method):
    if "Overview" not in wb.sheetnames:
        ws = wb.create_sheet("Overview", 0)
    else:
        ws = wb["Overview"]
    _sheet_base(
        ws,
        "SST gel 종합 분석 Overview",
        "원자료 유효성, 모델 적합성, 구조 이력 및 반복 분석용 핵심 지표를 한 화면에 표시합니다.",
        8,
    )
    if method == "A":
        if "Model_Validation" not in wb.sheetnames:
            mv = wb.create_sheet("Model_Validation")
        else:
            mv = wb["Model_Validation"]
        _sheet_base(mv, "SST gel 모델 검증", "Bingham 교차확인, H-B 잔차 및 절단 민감도를 표시합니다.", 13)
        if "StartUp" not in wb.sheetnames:
            su = wb.create_sheet("StartUp")
        else:
            su = wb["StartUp"]
        _sheet_base(su, "SST gel 저전단 Start-up", "최초 저전단과 고전단 이력 후 동일 저전단의 시간 프로파일을 비교합니다.", 12)
    names = wb.sheetnames
    for sheet in wb.worksheets:
        _nav(sheet, names[:8])


def fit_bingham_points(points, label):
    p = points.dropna(subset=["gamma", "tau_Pa"]).copy()
    p = p[(p["gamma"] > 0) & (p["n_valid"] > 0)]
    if len(p) < 3:
        return {"label": label, "status": "유효점 3개 미만", "n_points": int(len(p))}
    gamma = p["gamma"].to_numpy(float)
    tau = p["tau_Pa"].to_numpy(float)
    sigma = p["accuracy_cP"].to_numpy(float) * gamma / 1000.0
    weighted = bool(np.all(np.isfinite(sigma) & (sigma > 0)))
    X = np.column_stack([np.ones(len(gamma)), gamma])
    if weighted:
        Xw = X / sigma[:, None]
        yw = tau / sigma
        coef, *_ = np.linalg.lstsq(Xw, yw, rcond=None)
    else:
        coef, *_ = np.linalg.lstsq(X, tau, rcond=None)
    pred = X @ coef
    residual = tau - pred
    ss_res = float(np.sum(residual ** 2))
    ss_tot = float(np.sum((tau - tau.mean()) ** 2))
    return {
        "label": label, "status": "OK" if coef[0] >= 0 else "음의 절편",
        "n_points": int(len(p)), "gamma_min": float(gamma.min()), "gamma_max": float(gamma.max()),
        "tau_y": float(coef[0]), "plastic_viscosity": float(coef[1]),
        "R2": float(1 - ss_res / ss_tot) if ss_tot > 0 else np.nan,
        "RMSE": float(np.sqrt(np.mean(residual ** 2))),
        "MAE": float(np.mean(np.abs(residual))), "weighted": weighted,
    }


def hb_validation(fit, points):
    if not all(k in fit for k in ("tau_y", "K", "n")):
        return [], {}
    p = points.dropna(subset=["gamma", "tau_Pa"]).copy()
    p = p[(p["gamma"] > 0) & (p["n_valid"] > 0)].sort_values("gamma")
    rows = []
    for r in p.itertuples():
        pred = float(hb_model(r.gamma, fit["tau_y"], fit["K"], fit["n"]))
        residual = float(r.tau_Pa - pred)
        rows.append({
            "Step": int(r.Step), "gamma": float(r.gamma), "tau_meas": float(r.tau_Pa),
            "tau_pred": pred, "residual": residual, "abs_residual": abs(residual),
            "relative_error_pct": residual / r.tau_Pa * 100 if r.tau_Pa else np.nan,
            "torque": float(r.torque_valid), "status": str(r.status),
        })
    residuals = np.array([r["residual"] for r in rows], float)
    measured = np.array([r["tau_meas"] for r in rows], float)
    predicted = np.array([r["tau_pred"] for r in rows], float)
    stats = {
        "N": len(rows), "SSE": float(np.sum(residuals ** 2)),
        "RMSE": float(np.sqrt(np.mean(residuals ** 2))),
        "MAE": float(np.mean(np.abs(residuals))),
        "max_abs_residual": float(np.max(np.abs(residuals))),
        "max_abs_relative_error_pct": float(np.nanmax(np.abs([r["relative_error_pct"] for r in rows]))),
        "corr": float(np.corrcoef(measured, predicted)[0, 1]) if len(rows) > 1 else np.nan,
    }
    return rows, stats


def hb_sensitivity(points, label):
    p = points.dropna(subset=["gamma", "tau_Pa"]).copy()
    p = p[(p["gamma"] > 0) & (p["n_valid"] > 0)].sort_values("gamma")
    scenarios = [("Baseline", p)]
    if len(p) >= 5:
        scenarios.extend([("최저 γ 제외", p.iloc[1:]), ("최고 γ 제외", p.iloc[:-1])])
    out = []
    for scenario, data in scenarios:
        result = fit_hb_points(data, f"{label} {scenario}")
        result["scenario"] = scenario
        out.append(result)
    if out and "tau_y" in out[0]:
        base = out[0]
        for result in out:
            for key in ("tau_y", "K", "n"):
                value, ref = result.get(key), base.get(key)
                result[f"delta_{key}_pct"] = ((value - ref) / abs(ref) * 100
                    if value is not None and ref not in (None, 0) else np.nan)
    return out


def analyze_startup_profile(df, reps):
    if "used_in_up_branch" in reps.columns:
        up = reps[reps["used_in_up_branch"].fillna(False).astype(bool) & (reps["RPM"] > 0)].sort_values("Step")
        down = reps[reps["used_in_down_branch"].fillna(False).astype(bool) & (reps["RPM"] > 0)].sort_values("Step")
    else:
        up = reps[(reps["dir"] == "Up") & (reps["RPM"] > 0)].sort_values("Step")
        down = reps[(reps["dir"] == "Down") & (reps["RPM"] > 0)].sort_values("Step")
    if up.empty or down.empty:
        return {}, []
    up_step = int(up.iloc[0]["Step"])
    low_gamma = float(up.iloc[0]["gamma"])
    down_match = down.iloc[(down["gamma"] - low_gamma).abs().argsort()[:1]]
    down_step = int(down_match.iloc[0]["Step"])
    profiles = []
    summary = {"up_step": up_step, "down_step": down_step, "gamma": low_gamma}
    for direction, step in (("Initial Up", up_step), ("Post-shear Down", down_step)):
        g = df[df["Step"] == step].sort_values("t_elapsed").copy()
        tail = g[g["t_elapsed"] > max(0.0, float(g["t_elapsed"].max()) - 10.0)]
        plateau_tau = float(tail["ShearStress_Pa"].mean()) if len(tail) else np.nan
        plateau_eta = float(tail["Viscosity_cP"].mean()) if len(tail) else np.nan
        peak_tau = float(g["ShearStress_Pa"].max()) if len(g) else np.nan
        target = 0.9 * plateau_tau if np.isfinite(plateau_tau) else np.nan
        hit = g[g["ShearStress_Pa"] >= target] if np.isfinite(target) else g.iloc[0:0]
        t90 = float(hit["t_elapsed"].iloc[0]) if len(hit) else np.nan
        prefix = "initial" if direction.startswith("Initial") else "post"
        summary.update({f"{prefix}_plateau_tau": plateau_tau, f"{prefix}_plateau_eta": plateau_eta,
                        f"{prefix}_peak_tau": peak_tau, f"{prefix}_t90_s": t90,
                        f"{prefix}_torque_mean": float(tail["Torque_pct"].mean()) if len(tail) else np.nan})
        for r in g.itertuples():
            profiles.append({"direction": direction, "step": step, "time_s": float(r.t_elapsed),
                             "tau_Pa": float(r.ShearStress_Pa), "eta_cP": float(r.Viscosity_cP),
                             "torque_pct": float(r.Torque_pct), "temp_C": float(r.Temp_C)})
    a = summary.get("initial_plateau_tau")
    b = summary.get("post_plateau_tau")
    summary["plateau_retention_pct"] = b / a * 100 if a and np.isfinite(a) and np.isfinite(b) else np.nan
    a = summary.get("initial_peak_tau")
    b = summary.get("post_peak_tau")
    summary["peak_retention_pct"] = b / a * 100 if a and np.isfinite(a) and np.isfinite(b) else np.nan
    return summary, profiles


def build_qc_checks(reps, analysis, validations, startup):
    moving = reps[_moving_mask(reps)].copy()
    invalid_steps = moving[moving["status"] != "OK"]
    checks = []
    def add(name, value, criterion, status, note=""):
        checks.append({"check": name, "value": value, "criterion": criterion, "status": status, "note": note})
    hold_bad = moving[~moving["hold_compliant"]]
    add("60초 회전 Step", f"{len(moving)-len(hold_bad)}/{len(moving)}",
        "모든 Up/Down Step 60초", "OK" if hold_bad.empty else "REVIEW",
        ", ".join(f"S{int(r.Step)}={r.hold_s:g}s" for r in hold_bad.itertuples()))
    add("정량사용 Step", f"{len(moving)-len(invalid_steps)}/{len(moving)}",
        "60초 및 Torque 10–95%", "OK" if invalid_steps.empty else "REVIEW",
        ", ".join(f"S{int(r.Step)} {r.status}" for r in invalid_steps.itertuples()))
    common_n = len(analysis.get("common_gammas", []))
    point_status = "INVALID" if common_n < METHOD_A_MIN_HB_POINTS else ("REVIEW" if common_n == METHOD_A_MIN_HB_POINTS else "OK")
    add("공통 전단률 pair", common_n, "최소 4개; 5개 이상 권장", point_status)
    for key, name in (("up_common", "Up Common"), ("down_common", "Down Common")):
        fit = analysis["fits"].get(key, {})
        r2 = fit.get("R2")
        add(f"{name} H-B R²", r2, "0.98 이상", "OK" if r2 is not None and r2 >= 0.98 else "REVIEW")
        gmin, gmax = fit.get("gamma_min"), fit.get("gamma_max")
        ratio = gmax / gmin if gmin and gmax else np.nan
        add(f"{name} 전단률 범위", ratio, "10배 이상 권장", "OK" if np.isfinite(ratio) and ratio >= 10 else "REVIEW")
        stat = validations.get(key, {}).get("stats", {})
        err = stat.get("max_abs_relative_error_pct")
        add(f"{name} 최대 상대잔차", err, "5% 이하 참고", "OK" if err is not None and err <= 5 else "REVIEW")
    temp_span = float(moving["temp"].max() - moving["temp"].min()) if len(moving) else np.nan
    add("대표구간 온도 변동폭", temp_span, "1.0 °C 이하", "OK" if np.isfinite(temp_span) and temp_span <= 1 else "REVIEW")
    tq = startup.get("initial_torque_mean")
    add("저전단 Start-up Torque", tq, "10–95%", "OK" if tq is not None and 10 <= tq <= 95 else "REVIEW",
        "범위 밖이면 Start-up 지표는 탐색값")
    statuses = [r["status"] for r in checks]
    overall = "INVALID" if "INVALID" in statuses else ("REVIEW" if "REVIEW" in statuses else "VALID")
    return checks, overall


def _write_status_cf(ws, cell_range):
    first = cell_range.split(":")[0]
    ws.conditional_formatting.add(cell_range, FormulaRule(formula=[f'OR({first}="OK",{first}="VALID",{first}="GOOD")'], fill=PatternFill("solid", fgColor=GREEN), stopIfTrue=True))
    ws.conditional_formatting.add(cell_range, FormulaRule(formula=[f'OR({first}="REVIEW",{first}="WARNING")'], fill=PatternFill("solid", fgColor=YELLOW), stopIfTrue=True))
    ws.conditional_formatting.add(cell_range, FormulaRule(formula=[f'OR({first}="INVALID",{first}="HIGH VAR",{first}="ACTION")'], fill=PatternFill("solid", fgColor=RED), stopIfTrue=True))


def write_overview_a(ws, source, lot, repeat, meta, analysis, bingham, checks, overall, startup, run_info):
    clear_area(ws, 5, 60, 2, 8)
    _section(ws, 5, "시료 및 시험 정보")
    info = [("원본 파일", Path(source).name), ("Lot", lot), ("반복", repeat),
            ("시험 시작", meta.get("Test Start", "")), ("장비", meta.get("Instrument", "")),
            ("Spindle", meta.get("Spindle", "")), ("데이터 QC", overall)]
    for i, (name, value) in enumerate(info, 6):
        ws.cell(i, 2, name).font = Font(bold=True)
        ws.cell(i, 3, value)
    _section(ws, 14, "핵심 H-B 및 구조 이력 지표 — 대표값은 Full H-B")
    _headers(ws, 15, ["지표", "Up Full", "Down Full", "Down/Up", "단위", "해석"])
    uf, df_ = analysis["fits"]["up_full"], analysis["fits"]["down_full"]
    uc, dc = analysis["fits"]["up_common"], analysis["fits"]["down_common"]
    full_cmp = analysis.get("comparison_full", {})
    common_cmp = analysis.get("comparison", {})
    core = [
        ("H-B τy Full", uf.get("tau_y"), df_.get("tau_y"), full_cmp.get("tau_y_ratio_pct"), "Pa", "분석 가능 Step 전체"),
        ("H-B K Full", uf.get("K"), df_.get("K"), full_cmp.get("K_ratio_pct"), "Pa·sⁿ", "Consistency index"),
        ("H-B n Full", uf.get("n"), df_.get("n"), full_cmp.get("n_difference"), "-", "1 미만이면 전단박화"),
        ("H-B R² Full", uf.get("R2"), df_.get("R2"), None, "-", "Full 모델 적합도"),
        ("H-B τy Common", uc.get("tau_y"), dc.get("tau_y"), common_cmp.get("tau_y_ratio_pct"), "Pa", "공통 전단률 진단"),
        ("H-B K Common", uc.get("K"), dc.get("K"), common_cmp.get("K_ratio_pct"), "Pa·sⁿ", "공통 전단률 진단"),
        ("H-B n Common", uc.get("n"), dc.get("n"), None, "-", "공통 전단률 진단"),
        ("H-B R² Common", uc.get("R2"), dc.get("R2"), None, "-", "공통 전단률 진단"),
        ("Bingham τy", bingham["up_common"].get("tau_y"), bingham["down_common"].get("tau_y"), None, "Pa", "선형 모델 교차확인(공통점)"),
        ("평균 응력 회복률", None, None, analysis.get("mean_recovery_pct"), "%", "공통 γ의 Down/Up 평균"),
        ("Relative hysteresis", analysis.get("relative_area_pct"), None, None, "%", "공통 pair만. Up 면적 대비 절대 이력면적"),
        ("저전단 plateau 유지율", startup.get("initial_plateau_tau"), startup.get("post_plateau_tau"), startup.get("plateau_retention_pct"), "Pa / %", "Torque 유효성 동시 확인"),
    ]
    for i, row in enumerate(core, 16):
        for j, value in enumerate(row, 2):
            put(ws, i, j, value, "0.000" if isinstance(value, (int, float, np.number)) else None)
    valid_row = 16 + len(core) + 1
    _section(ws, valid_row, "데이터 및 모델 유효성")
    _headers(ws, valid_row + 1, ["확인 항목", "값", "권장/판정 기준", "상태", "비고"])
    for i, item in enumerate(checks, valid_row + 2):
        vals = [item[k] for k in ("check", "value", "criterion", "status", "note")]
        for j, value in enumerate(vals, 2):
            put(ws, i, j, value, "0.000" if isinstance(value, (int, float, np.number)) else None)
        ws.cell(i, 6).alignment = Alignment(horizontal="center")
        ws.cell(i, 7).alignment = Alignment(wrap_text=True)
    end = valid_row + 1 + len(checks)
    _write_status_cf(ws, f"F{valid_row + 2}:F{end}")
    ws.cell(end + 2, 2, "주의: 제품 규격이 제공되지 않았으므로 VALID/REVIEW/INVALID는 데이터·모델 적합성 판정이며 제품 합격 판정이 아닙니다.")
    ws.merge_cells(start_row=end + 2, start_column=2, end_row=end + 2, end_column=8)
    ws.cell(end + 2, 2).alignment = Alignment(wrap_text=True)
    ws.cell(end + 2, 2).font = Font(italic=True, color="9C6500")
    ws.row_dimensions[end + 2].height = 32
    trace_row = end + 4
    _section(ws, trace_row, "분석 실행 및 추적성")
    trace = [
        ("분석기", run_info.get("analyzer")), ("Method", run_info.get("method")),
        ("대표값 규칙", run_info.get("representative_rule")), ("H-B 점수 규칙", run_info.get("hb_rule")),
        ("Torque 범위", run_info.get("torque_rule")), ("템플릿", run_info.get("template")),
        ("원본 SHA-256", run_info.get("raw_sha256")), ("실행 시각", run_info.get("timestamp")),
        ("실행 명령", run_info.get("command")),
    ]
    for i, (name, value) in enumerate(trace, trace_row + 1):
        ws.cell(i, 2, name).font = Font(bold=True)
        ws.merge_cells(start_row=i, start_column=3, end_row=i, end_column=8)
        ws.cell(i, 3, value)
        ws.cell(i, 3).alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[trace_row + len(trace)].height = 42


def write_model_validation(ws, bingham, validations, sensitivities):
    clear_area(ws, 5, 90, 2, 13)
    _section(ws, 5, "Bingham 교차확인", 2, 13)
    _headers(ws, 6, ["프로파일", "γ min", "γ max", "N", "τy(Pa)", "Plastic viscosity(Pa·s)", "R²", "RMSE(Pa)", "MAE(Pa)", "가중", "상태"], 2)
    for i, key in enumerate(("up_common", "down_common"), 7):
        f = bingham[key]
        vals = [f.get("label"), f.get("gamma_min"), f.get("gamma_max"), f.get("n_points"), f.get("tau_y"),
                f.get("plastic_viscosity"), f.get("R2"), f.get("RMSE"), f.get("MAE"),
                "Accuracy" if f.get("weighted") else "None", f.get("status")]
        for j, value in enumerate(vals, 2):
            put(ws, i, j, value, "0.0000" if isinstance(value, (int, float, np.number)) else None)
    _section(ws, 11, "H-B 잔차 통계 — Full이 대표, Common은 진단", 2, 13)
    _headers(ws, 12, ["프로파일", "N", "SSE", "RMSE(Pa)", "MAE(Pa)", "Max |resid|(Pa)", "Max |상대오차|(%)", "Corr"], 2)
    labels = {"up_full": "Up Full", "down_full": "Down Full", "up_common": "Up Common", "down_common": "Down Common"}
    row = 13
    for key in ("up_full", "down_full", "up_common", "down_common"):
        if key not in validations:
            continue
        s = validations[key]["stats"] or {}
        vals = [labels[key], s.get("N"), s.get("SSE"), s.get("RMSE"), s.get("MAE"),
                s.get("max_abs_residual"), s.get("max_abs_relative_error_pct"), s.get("corr")]
        for j, value in enumerate(vals, 2):
            put(ws, row, j, value, "0.0000" if isinstance(value, (int, float, np.number)) else None)
        row += 1
    row += 1
    _section(ws, row, "H-B 절단 민감도", 2, 13)
    _headers(ws, row + 1, ["방향", "시나리오", "N", "τy", "K", "n", "R²", "Δτy(%)", "ΔK(%)", "Δn(%)", "상태"], 2)
    row += 2
    for key, direction in (("up_full", "Up Full"), ("down_full", "Down Full"), ("up_common", "Up Common"), ("down_common", "Down Common")):
        for f in sensitivities.get(key, []):
            vals = [direction, f.get("scenario"), f.get("n_points"), f.get("tau_y"), f.get("K"), f.get("n"), f.get("R2"),
                    f.get("delta_tau_y_pct"), f.get("delta_K_pct"), f.get("delta_n_pct"), f.get("status")]
            for j, value in enumerate(vals, 2):
                put(ws, row, j, value, "0.0000" if isinstance(value, (int, float, np.number)) else None)
            row += 1
    _section(ws, row + 1, "전단률별 H-B 잔차 (Full 및 Common)", 2, 13)
    _headers(ws, row + 2, ["방향", "Step", "γ(1/s)", "측정 τ", "예측 τ", "잔차", "|잔차|", "상대오차(%)", "Torque(%)", "상태"], 2)
    row += 3
    for key, direction in (("up_full", "Up Full"), ("down_full", "Down Full"), ("up_common", "Up Common"), ("down_common", "Down Common")):
        for r in validations.get(key, {}).get("rows", []):
            vals = [direction, r["Step"], r["gamma"], r["tau_meas"], r["tau_pred"], r["residual"], r["abs_residual"],
                    r["relative_error_pct"], r["torque"], r["status"]]
            for j, value in enumerate(vals, 2):
                put(ws, row, j, value, "0.0000" if isinstance(value, (int, float, np.number)) else None)
            row += 1
    ws.freeze_panes = "B13"


def write_startup(ws, summary, profiles):
    clear_area(ws, 5, 500, 2, 12)
    _section(ws, 5, "저전단 Start-up 요약", 2, 12)
    _headers(ws, 6, ["지표", "Initial Up", "Post-shear Down", "Down/Up", "단위", "주의"], 2)
    rows = [
        ("원본 Step", summary.get("up_step"), summary.get("down_step"), None, "#", "동일/근접 저전단률"),
        ("전단률", summary.get("gamma"), summary.get("gamma"), None, "1/s", ""),
        ("후반 10초 τ", summary.get("initial_plateau_tau"), summary.get("post_plateau_tau"), summary.get("plateau_retention_pct"), "Pa / %", ""),
        ("Peak τ", summary.get("initial_peak_tau"), summary.get("post_peak_tau"), summary.get("peak_retention_pct"), "Pa / %", ""),
        ("후반 10초 η", summary.get("initial_plateau_eta"), summary.get("post_plateau_eta"), None, "cP", ""),
        ("t90", summary.get("initial_t90_s"), summary.get("post_t90_s"), None, "s", "단순 90% 최초 도달"),
        ("후반 10초 Torque", summary.get("initial_torque_mean"), summary.get("post_torque_mean"), None, "%", "10% 미만은 탐색값"),
    ]
    for i, values in enumerate(rows, 7):
        for j, value in enumerate(values, 2):
            put(ws, i, j, value, "0.000" if isinstance(value, (int, float, np.number)) else None)
    _section(ws, 16, "시간 프로파일", 2, 12)
    _headers(ws, 17, ["방향", "Step", "시간(s)", "τ(Pa)", "η(cP)", "Torque(%)", "온도(°C)"], 2)
    for i, r in enumerate(profiles, 18):
        vals = [r[k] for k in ("direction", "step", "time_s", "tau_Pa", "eta_cP", "torque_pct", "temp_C")]
        for j, value in enumerate(vals, 2):
            put(ws, i, j, value, "0.000" if isinstance(value, (int, float, np.number)) else None)
    if profiles:
        last = 17 + len(profiles)
        chart = LineChart()
        chart.title = "Low-shear stress profile"
        chart.y_axis.title = "Shear stress (Pa)"
        chart.x_axis.title = "Point sequence"
        data = Reference(ws, min_col=5, min_row=17, max_row=last)
        chart.add_data(data, titles_from_data=True)
        chart.height = 7
        chart.width = 14
        ws.add_chart(chart, "J6")
        ws.auto_filter.ref = f"B17:H{last}"
        ws.freeze_panes = "B18"


def write_overview_b(ws, source, lot, repeat, meta, result):
    clear_area(ws, 5, 45, 2, 8)
    _section(ws, 5, "시료 및 시험 정보")
    info = [("원본 파일", Path(source).name), ("Lot", lot), ("반복", repeat), ("시험 시작", meta.get("Test Start", "")),
            ("장비", meta.get("Instrument", "")), ("Spindle", meta.get("Spindle", ""))]
    for i, (name, value) in enumerate(info, 6):
        ws.cell(i, 2, name).font = Font(bold=True)
        ws.cell(i, 3, value)
    _section(ws, 14, "3ITT 핵심 지표")
    _headers(ws, 15, ["지표", "값", "단위", "상태/해석"])
    base = result.get("summary", {}).get("Baseline", {})
    breakdown = result.get("summary", {}).get("Breakdown", {})
    rows = [("Baseline τ", base.get("tau"), "Pa", base.get("status")), ("Baseline η", base.get("eta"), "cP", base.get("status")),
            ("Breakdown τ", breakdown.get("tau"), "Pa", breakdown.get("status")), ("Breakdown η", breakdown.get("eta"), "cP", breakdown.get("status")),
            ("Recovery 유효 데이터", result.get("recovery_valid_pct"), "%", "10–95% Torque"),
            ("t50", result.get("target_times", {}).get(50), "s", "미도달 시 공란"),
            ("t80", result.get("target_times", {}).get(80), "s", "미도달 시 공란"),
            ("t90", result.get("target_times", {}).get(90), "s", "미도달 시 공란")]
    for i, values in enumerate(rows, 16):
        for j, value in enumerate(values, 2):
            put(ws, i, j, value, "0.000" if isinstance(value, (int, float, np.number)) else None)



def sha256_file(path):
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def make_run_info(source, template, tq_lo, tq_hi):
    return {
        "analyzer": VERSION,
        "method": f"Method A {METHOD_A_VERSION}",
        "representative_rule": "Up/Down 각 60초 고정; 50초 초과~60초(후반 10초) 유효 Torque 평균",
        "hb_rule": "4점 최소(REVIEW), 5점 이상 권장; 3점 이하 INVALID/피팅 금지",
        "torque_rule": f"{tq_lo:g}–{tq_hi:g}%",
        "template": Path(template).name,
        "raw_sha256": sha256_file(source),
        "timestamp": datetime.now().astimezone().isoformat(timespec="seconds"),
        "command": " ".join([Path(sys.executable).name, *sys.argv]),
    }


def _cli_num(value, digits=4):
    try:
        value = float(value)
        return "N/A" if not np.isfinite(value) else f"{value:.{digits}f}"
    except (TypeError, ValueError):
        return "N/A"


def print_method_a_cli(source, output, point_count, step_count, result, summary):
    reps = summary["reps"]
    print(f"  [식별] Lot={summary['lot']} | 반복={summary['repeat']} | points={point_count} | steps={step_count}")
    print(f"  [규칙] 회전 Step 60초 고정 | 대표구간 후반 10초 | Torque {summary['run_info']['torque_rule']}")
    moving = reps[_moving_mask(reps)]
    bad = moving[~moving["hold_compliant"]]
    print(f"  [Step QC] 60초 적합 {len(moving)-len(bad)}/{len(moving)} | 정량사용 {(moving['status']=='OK').sum()}/{len(moving)}")
    if len(bad):
        print("  [제외] " + ", ".join(f"S{int(r.Step)}={r.hold_s:g}s" for r in bad.itertuples()) + " (정량계산 제외)")
    for key in ("up_full", "down_full", "up_common", "down_common"):
        fit = result["fits"].get(key, {})
        print(f"  [H-B {fit.get('label', key)}] N={fit.get('n_points', 0)} | γ={_cli_num(fit.get('gamma_min'),3)}–{_cli_num(fit.get('gamma_max'),3)} s^-1 | τy={_cli_num(fit.get('tau_y'))} Pa | K={_cli_num(fit.get('K'))} | n={_cli_num(fit.get('n'))} | R²={_cli_num(fit.get('R2'),5)} | {fit.get('status','')}")
    print(f"  [구조 이력] 평균 응력회복={_cli_num(result.get('mean_recovery_pct'),3)}% | 저전단 회복={_cli_num(result.get('low_shear_recovery_pct'),3)}% | Relative hysteresis={_cli_num(result.get('relative_area_pct'),3)}%")
    print(f"  [판정] 데이터/모델={summary['overall']} (제품 적합성 판정 아님)")
    review = [c for c in summary["checks"] if c["status"] != "OK"]
    for c in review:
        print(f"    - {c['status']}: {c['check']} | 값={c['value']} | 기준={c['criterion']} | {c['note']}")
    print(f"  [출력] {output}")
    print(f"  [원본 SHA-256] {summary['run_info']['raw_sha256']}")


def print_method_b_cli(source, output, point_count, step_count, result, summary):
    base = result.get("summary", {}).get("Baseline", {})
    breakdown = result.get("summary", {}).get("Breakdown", {})
    print(f"  [식별] Lot={summary['lot']} | 반복={summary['repeat']} | points={point_count} | steps={step_count}")
    print(f"  [3ITT] Baseline η={_cli_num(base.get('eta'),2)} cP, Torque={_cli_num(base.get('torque'),2)}% | Breakdown η={_cli_num(breakdown.get('eta'),2)} cP, Torque={_cli_num(breakdown.get('torque'),2)}%")
    print(f"  [Recovery] 유효={_cli_num(result.get('recovery_valid_pct'),2)}% | t50={_cli_num(result.get('target_times',{}).get(50),1)}s | t80={_cli_num(result.get('target_times',{}).get(80),1)}s | t90={_cli_num(result.get('target_times',{}).get(90),1)}s")
    print(f"  [출력] {output}")

def process_enhanced(path, method, template, output, win_s, fraction, tq_lo, tq_hi):
    df, holds, meta = parse_dvnext_csv(path)
    lot, repeat = parse_sample_id(path)
    wb = prepare_workbook(template, output, method)
    ensure_extended_sheets(wb, method)
    if method == "A":
        reps = step_representatives(df, holds, win_s, fraction, tq_lo, tq_hi)
        result = analyze_step_flow(reps)
        write_step_flow(wb, df, holds, meta, reps, result, path)
        up = _branch_frame(reps, "up")
        down = _branch_frame(reps, "down")
        pair_steps = set(int(s) for s in result.get("pair_steps") or [])
        up_common = up[up["Step"].isin(pair_steps)]
        down_common = down[down["Step"].isin(pair_steps)]
        bingham = {"up_full": fit_bingham_points(up, "Up Full"),
                   "down_full": fit_bingham_points(down, "Down Full"),
                   "up_common": fit_bingham_points(up_common, "Up Common"),
                   "down_common": fit_bingham_points(down_common, "Down Common")}
        validations = {}
        for key, points in (("up_full", up), ("down_full", down), ("up_common", up_common), ("down_common", down_common)):
            rows, stats = hb_validation(result["fits"][key], points)
            validations[key] = {"rows": rows, "stats": stats}
        sensitivities = {"up_full": hb_sensitivity(up, "Up Full"),
                         "down_full": hb_sensitivity(down, "Down Full"),
                         "up_common": hb_sensitivity(up_common, "Up Common"),
                         "down_common": hb_sensitivity(down_common, "Down Common")}
        startup, profiles = analyze_startup_profile(df, reps)
        checks, overall = build_qc_checks(reps, result, validations, startup)
        run_info = make_run_info(path, template, tq_lo, tq_hi)
        write_overview_a(wb["Overview"], path, lot, repeat, meta, result, bingham, checks, overall, startup, run_info)
        write_model_validation(wb["Model_Validation"], bingham, validations, sensitivities)
        write_startup(wb["StartUp"], startup, profiles)
        uf, dfit = result["fits"]["up_full"], result["fits"]["down_full"]
        record = {
            "lot": lot, "repeat": repeat, "file": Path(path).name, "test_start": meta.get("Test Start", ""), "qc_status": overall,
            "up_tau_y_Pa": uf.get("tau_y"), "up_K_Pa_s_n": uf.get("K"), "up_n": uf.get("n"), "up_R2": uf.get("R2"),
            "down_tau_y_Pa": dfit.get("tau_y"), "down_K_Pa_s_n": dfit.get("K"), "down_n": dfit.get("n"), "down_R2": dfit.get("R2"),
            "tau_y_retention_pct": result.get("comparison_full", {}).get("tau_y_ratio_pct"),
            "K_retention_pct": result.get("comparison_full", {}).get("K_ratio_pct"),
            "mean_stress_recovery_pct": result.get("mean_recovery_pct"), "low_shear_recovery_pct": result.get("low_shear_recovery_pct"),
            "relative_hysteresis_pct": result.get("relative_area_pct"), "absolute_hysteresis": result.get("absolute_area"),
            "initial_plateau_Pa": startup.get("initial_plateau_tau"), "post_plateau_Pa": startup.get("post_plateau_tau"),
            "plateau_retention_pct": startup.get("plateau_retention_pct"), "temp_mean_C": float(reps["temp"].mean()),
            "valid_moving_steps": int(((_moving_mask(reps)) & (reps["status"] == "OK")).sum()),
            "moving_steps": int(_moving_mask(reps).sum()),
            "method_version": METHOD_A_VERSION, "analyzer_version": VERSION,
            "raw_sha256": run_info["raw_sha256"],
        }
        step_records = reps[_moving_mask(reps)].copy()
        step_records.insert(0, "file", Path(path).name)
        step_records.insert(0, "repeat", repeat)
        step_records.insert(0, "lot", lot)
        result["_lot_record"] = record
        result["_step_records"] = step_records.to_dict("records")
        result["_cli_summary"] = {"lot": lot, "repeat": repeat, "reps": reps, "checks": checks,
                                  "overall": overall, "bingham": bingham, "validations": validations,
                                  "sensitivities": sensitivities, "startup": startup, "run_info": run_info}
        # [05] QC layer(STRICT 채널): 04_2 결과(_lot_record)는 그대로 두고 별도 레코드를 만든다.
        policy = active_policy()
        qc = compute_a_qc(df, holds, meta, channel_view(reps, "strict"), result, startup, checks, overall, tq_lo, tq_hi,
                          channel="strict", policy=policy)
        # [06] INCLUSIVE 채널: 정책이 허용한 hold + 범위 밖 점 포함 대표값으로 H-B·hysteresis·QC 지표를 다시 계산
        reps_incl = channel_view(reps, "inclusive")
        result_incl = analyze_step_flow(reps_incl)
        qc_incl = compute_a_qc(df, holds, meta, reps_incl, result_incl, startup, checks, overall, tq_lo, tq_hi,
                               channel="inclusive", policy=policy)
        audit_rows = [{k: (v.item() if isinstance(v, np.generic) else v) for k, v in r.items()}
                      for r in reps.to_dict("records")]
        strict_limits = qc.get("shear_limits") or {}
        incl_limits = qc_incl.get("shear_limits") or {}
        qc.update({"lot": lot, "repeat": repeat, "file": Path(path).name, "method": "A",
                   "test_start": meta.get("Test Start", ""), "data_qc": overall,
                   "raw_sha256": run_info["raw_sha256"], "template": Path(template).name,
                   "output": Path(output).name,
                   "strict_shear": strict_limits, "inclusive_shear": incl_limits,
                   "strict_low_rpm": strict_limits.get("fixed_low_rpm", strict_limits.get("low_rpm")),
                   "strict_high_rpm": strict_limits.get("high_rpm"),
                   "inclusive_low_rpm": incl_limits.get("fixed_low_rpm", incl_limits.get("low_rpm")),
                   "inclusive_high_rpm": incl_limits.get("high_rpm"),
                   "strict_measured_max_rpm": strict_limits.get("measured_max_rpm"),
                   "inclusive_measured_max_rpm": incl_limits.get("measured_max_rpm"),
                   "values_incl": qc_incl["values"], "status_incl": qc_incl["status"],
                   "value_class_strict": {c: (s or {}).get("value_class", "") for c, s in qc["status"].items()},
                   "value_class_incl": {c: (s or {}).get("value_class", "") for c, s in qc_incl["status"].items()},
                   "exception_steps": qc_incl["exception_steps"], "route": policy.name,
                   "policy_name": policy.name, "policy_hash": policy.digest(),
                   "audit_rows": audit_rows, "ovs_present": qc.get("ovs_present"),
                   "decade_pairs_strict": list(qc.get("decade_pairs") or []),
                   "decade_pairs_incl": list(qc_incl.get("decade_pairs") or []),
                   "stabilization": qc.get("stabilization") or {},
                   "analysis_incl": {"fits": {k: {kk: vv for kk, vv in f.items() if kk != "steps"} for k, f in result_incl["fits"].items()},
                                     "relative_area_pct": result_incl.get("relative_area_pct")}})
        write_qc_indices_sheet(wb, "A", qc, policy=policy)
        write_step_audit_sheet(wb, qc, policy)
        write_policy_sheet(wb, policy, run_info)
        result["_qc_record"] = qc
        result["_inclusive"] = result_incl
    else:
        result = analyze_3itt(df, holds, win_s, fraction, tq_lo, tq_hi)
        write_3itt(wb, df, holds, meta, result, path)
        write_overview_b(wb["Overview"], path, lot, repeat, meta, result)
        base = result.get("summary", {}).get("Baseline", {})
        breakdown = result.get("summary", {}).get("Breakdown", {})
        rec_map = {int(r["time_s"]): r for r in result.get("recovery", [])}
        record = {"lot": lot, "repeat": repeat, "file": Path(path).name, "test_start": meta.get("Test Start", ""),
                  "qc_status": "VALID" if base.get("status") == "OK" and breakdown.get("status") == "OK" else "REVIEW",
                  "baseline_tau_Pa": base.get("tau"), "baseline_eta_cP": base.get("eta"),
                  "breakdown_tau_Pa": breakdown.get("tau"), "breakdown_eta_cP": breakdown.get("eta"),
                  "recovery_valid_pct": result.get("recovery_valid_pct"), "t50_s": result.get("target_times", {}).get(50),
                  "t80_s": result.get("target_times", {}).get(80), "t90_s": result.get("target_times", {}).get(90)}
        for t in RECOVERY_TIMES:
            record[f"R{t}_eta_pct"] = rec_map.get(t, {}).get("recovery_eta_pct")
        result["_lot_record"] = record
        result["_step_records"] = []
        result["_cli_summary"] = {"lot": lot, "repeat": repeat, "result": result}
        # [05] QC layer
        qc = compute_b_qc(df, holds, meta, result, tq_lo, tq_hi)
        policy = active_policy()
        qc.update({"lot": lot, "repeat": repeat, "file": Path(path).name, "method": "B",
                   "test_start": meta.get("Test Start", ""), "data_qc": record["qc_status"],
                   "raw_sha256": sha256_file(path), "template": Path(template).name,
                   "output": Path(output).name, "route": policy.name,
                   "values_incl": dict(qc["values"]), "status_incl": dict(qc["status"]),
                   "policy_name": policy.name, "policy_hash": policy.digest(), "audit_rows": []})
        write_qc_indices_sheet(wb, "B", qc, policy=policy)
        write_policy_sheet(wb, policy, make_run_info(path, template, tq_lo, tq_hi))
        result["_qc_record"] = qc
    wb.save(output)
    return len(df), int(df["Step"].nunique()), result


def excel_ranges_overlap(ref_a, ref_b):
    min_col_a, min_row_a, max_col_a, max_row_a = range_boundaries(ref_a)
    min_col_b, min_row_b, max_col_b, max_row_b = range_boundaries(ref_b)

    return not (
        max_col_a < min_col_b
        or max_col_b < min_col_a
        or max_row_a < min_row_b
        or max_row_b < min_row_a
    )
    
def _safe_table(ws, ref, name):
    _add_checked_table(
        ws=ws,
        ref=ref,
        name=name,
        style_name="TableStyleMedium2",
    )


def write_lot_workbook(lot, records, step_records, method, outdir):
    records = sorted(records, key=lambda x: (str(x.get("repeat", "")), x.get("file", "")))
    wb = Workbook()
    ov = wb.active
    ov.title = "Lot_Overview"
    _sheet_base(ov, f"HT-{lot} Lot 종합 분석", "개별 반복의 모델 결과와 반복 재현성을 통합합니다. 제품 규격 판정이 아닌 분석·데이터 QC 결과입니다.", 11)
    rep = wb.create_sheet("Replicates")
    stat = wb.create_sheet("Metric_Statistics")
    step = wb.create_sheet("Step_Statistics") if method == "A" else None
    names = wb.sheetnames
    for ws in wb.worksheets:
        _nav(ws, names)
        ws.sheet_view.showGridLines = False
        ws.column_dimensions["A"].width = 3

    columns = list(records[0].keys())
    for j, name in enumerate(columns, 1):
        rep.cell(1, j, name)
    for i, record in enumerate(records, 2):
        for j, name in enumerate(columns, 1):
            value = record.get(name)
            rep.cell(i, j, py_value(value))
    rep.freeze_panes = "A2"
    
    _safe_table(rep, f"A1:{get_column_letter(len(columns))}{len(records)+1}", "ReplicateData")
    for j, name in enumerate(columns, 1):
        rep.column_dimensions[get_column_letter(j)].width = min(max(12, len(name) + 2), 26)
    metric_cols = [c for c in columns if c not in ("lot", "repeat", "file", "test_start", "qc_status")]
    _sheet_base(stat, "반복 측정 통계", "평균, 표본표준편차, RSD, 최소·최대값을 산출합니다.", 9)
    _headers(stat, 5, ["Metric", "n", "Mean", "SD", "RSD(%)", "Min", "Max", "재현성 상태"], 2)
    first_data_row = 2
    last_data_row = len(records) + 1

    col_map = {
        name: get_column_letter(columns.index(name) + 1)
        for name in columns
    }

    for row_idx, metric in enumerate(metric_cols, start=6):
        source_col = col_map[metric]
        source_range = (
            f"Replicates!{source_col}{first_data_row}:"
            f"{source_col}{last_data_row}"
        )

        stat.cell(row_idx, 2).value = metric

        # Raw/Replicates 데이터를 Excel에서 편집하면 아래 모든 셀이 자동 갱신된다.
        stat.cell(row_idx, 3).value = f"=COUNT({source_range})"
        stat.cell(row_idx, 4).value = f'=IFERROR(AVERAGE({source_range}),"")'

        # STDEV.S 대신 호환성이 넓은 표본 표준편차 함수 STDEV를 사용.
        stat.cell(row_idx, 5).value = (
            f'=IF(C{row_idx}<2,"",STDEV({source_range}))'
        )

        stat.cell(row_idx, 6).value = (
            f'=IFERROR(E{row_idx}/ABS(D{row_idx})*100,"")'
        )

        stat.cell(row_idx, 7).value = f'=IFERROR(MIN({source_range}),"")'
        stat.cell(row_idx, 8).value = f'=IFERROR(MAX({source_range}),"")'

        stat.cell(row_idx, 9).value = (
            f'=IF(C{row_idx}<2,"N/A",'
            f'IF(F{row_idx}<=10,"GOOD",'
            f'IF(F{row_idx}<=20,"REVIEW","HIGH VAR")))'
        )

        for col_idx in range(4, 9):
            stat.cell(row_idx, col_idx).number_format = "0.000"
    _write_status_cf(stat, f"I6:I{5+len(metric_cols)}")
    stat.freeze_panes = "B6"

    _section(ov, 5, "Lot 정보", 2, 11)
    info = [("Lot", lot), ("반복 수", len(records)), ("반복", ", ".join(str(r["repeat"]) for r in records)),
            ("개별 데이터 QC", ", ".join(f"{r['repeat']}:{r['qc_status']}" for r in records))]
    for i, (name, value) in enumerate(info, 6):
        ov.cell(i, 2, name).font = Font(bold=True)
        ov.cell(i, 3, value)
    _section(ov, 11, "핵심 반복 통계", 2, 11)
    _headers(ov, 12, ["Metric", "Mean", "SD", "RSD(%)", "Min", "Max", "재현성", "단위"], 2)
    preferred = (["up_tau_y_Pa", "up_K_Pa_s_n", "up_n", "down_tau_y_Pa", "down_K_Pa_s_n", "down_n",
                  "mean_stress_recovery_pct", "relative_hysteresis_pct", "plateau_retention_pct"] if method == "A" else
                 ["baseline_tau_Pa", "baseline_eta_cP", "breakdown_tau_Pa", "breakdown_eta_cP", "recovery_valid_pct", "t50_s", "t80_s", "t90_s"])
    units = {"up_tau_y_Pa": "Pa", "up_K_Pa_s_n": "Pa·sⁿ", "up_n": "-", "down_tau_y_Pa": "Pa",
             "down_K_Pa_s_n": "Pa·sⁿ", "down_n": "-", "mean_stress_recovery_pct": "%", "relative_hysteresis_pct": "%",
             "plateau_retention_pct": "%", "baseline_tau_Pa": "Pa", "baseline_eta_cP": "cP", "breakdown_tau_Pa": "Pa",
             "breakdown_eta_cP": "cP", "recovery_valid_pct": "%", "t50_s": "s", "t80_s": "s", "t90_s": "s"}
    stat_rows = {m: 6 + metric_cols.index(m) for m in metric_cols}
    out_row = 13
    for metric in preferred:
        if metric not in stat_rows:
            continue
        sr = stat_rows[metric]
        ov.cell(out_row, 2, metric)
        for dest, src in zip(range(3, 9), range(4, 10)):
            ov.cell(out_row, dest, f"='Metric_Statistics'!{get_column_letter(src)}{sr}")
            if dest < 8:
                ov.cell(out_row, dest).number_format = "0.000"
        ov.cell(out_row, 9, units.get(metric, ""))
        out_row += 1
    _write_status_cf(ov, f"H13:H{out_row-1}")
    ov.merge_cells(start_row=out_row + 2, start_column=2, end_row=out_row + 2, end_column=11)
    ov.cell(out_row + 2, 2, "RSD 상태는 임시 데이터 품질 기준입니다: ≤10% GOOD, 10–20% REVIEW, >20% HIGH VAR. 제품별 관리한계는 충분한 Lot 축적 후 별도로 설정하십시오.")
    ov.cell(out_row + 2, 2).alignment = Alignment(wrap_text=True)
    ov.cell(out_row + 2, 2).font = Font(italic=True, color="9C6500")
    ov.row_dimensions[out_row + 2].height = 34

    if method == "A" and step_records:
        sdf = pd.DataFrame(step_records)
        sdf = sdf.sort_values(["dir", "gamma", "repeat"])
        metrics = ["tau_Pa", "eta_cP", "torque_valid", "temp"]
        rows = []
        for (direction, gamma), g in sdf.groupby(["dir", "gamma"], sort=True):
            row = {"Direction": direction, "gamma_1_s": float(gamma), "n_repeats": int(g["repeat"].nunique())}
            for m in metrics:
                values = pd.to_numeric(g[m], errors="coerce").dropna()
                row[f"{m}_mean"] = float(values.mean()) if len(values) else np.nan
                row[f"{m}_SD"] = float(values.std(ddof=1)) if len(values) > 1 else np.nan
                row[f"{m}_RSD_pct"] = float(values.std(ddof=1) / abs(values.mean()) * 100) if len(values) > 1 and values.mean() else np.nan
            rows.append(row)
        headers = list(rows[0].keys()) if rows else []
        for j, h in enumerate(headers, 1):
            step.cell(1, j, h)
        for i, row in enumerate(rows, 2):
            for j, h in enumerate(headers, 1):
                step.cell(i, j, py_value(row[h]))
                if isinstance(row[h], (int, float, np.number)):
                    step.cell(i, j).number_format = "0.000"
        step.freeze_panes = "A2"
        
        _safe_table(
            step,
            f"A1:{get_column_letter(len(headers))}{len(rows)+1}",
            "StepStats"
        )

        for j, h in enumerate(headers, 1):
            step.column_dimensions[get_column_letter(j)].width = min(max(12, len(h) + 2), 24)
        if rows:
            chart = LineChart()
            chart.title = "Lot mean Up / Down flow profile"
            chart.y_axis.title = "Shear stress (Pa)"
            chart.x_axis.title = "Shear rate (1/s)"
            gamma_col = headers.index("gamma_1_s") + 1
            tau_col = headers.index("tau_Pa_mean") + 1
            for direction in ("Down", "Up"):
                idx = [i + 2 for i, r in enumerate(rows) if r["Direction"] == direction]
                if idx:
                    data = Reference(step, min_col=tau_col, min_row=min(idx)-1, max_row=max(idx))
                    cats = Reference(step, min_col=gamma_col, min_row=min(idx), max_row=max(idx))
                    chart.add_data(data, titles_from_data=True)
                    chart.series[-1].tx = SeriesLabel(v=direction)
                    chart.set_categories(cats)
            chart.height = 7
            chart.width = 14
            ov.add_chart(chart, "K5")
    output = Path(outdir) / f"HT-{lot}_Lot_Analysis.xlsx"
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except AttributeError:
        pass
    wb.save(output)
    return output


# =============================================================================
# [05] Lot QC layer
#   - Method A: 지정 RPM fingerprint(τ·SSI), 구조 이력(D/U·hysteresis·Step 내 붕괴),
#               모델 진단(H-B·Start-up), 데이터 품질(Torque·분해능·드리프트·온도)
#   - Method B: 기준 구조, 고전단 붕괴, 시점 회복률, t90(미도달 censoring)
#   - 밀도   : 별도 입력 CSV(밀도계/중량법) → ρ25, 온도 적합성, 직독–질량식 교차확인
#   - 출력   : 개별 보고서 QC_Indices 시트, Lot_QC 워크북, long-format CSV(+DB 누적)
#   규칙: 유효하지 않은 값은 대체하지 않고 NaN + 사유를 남긴다. 상태는 데이터 품질이며
#         제품 적합성 판정이 아니다.
# =============================================================================
from openpyxl.chart import ScatterChart, Series

QC_LAYER_VERSION = "LotQC-2.8"
INCLUDE_LEGACY_4RPM_DIAGNOSTICS = False
# L·M·H_REF는 legacy 비교점이다. H_REF=4 rpm은 5 rpm이 있으면 고전단 최고점이 아니다.
# 동적 최저·최고 RPM은 이 딕셔너리에서 읽지 않고 채널의 분석 가능 Step에서 고른다.
A_LEGACY_REF_RPM = {"L": 0.5, "M": 2.0, "H_REF": 4.0}
FIXED_LOW_RPM = 0.5
FIXED_MID_RPM = 2.0
LEGACY_HIGH_RPM = 4.0
FIELD_SSI_SELECTION_RULE = "exact measured Up 0.2 rpm / exact measured Up 2.0 rpm"
QC_SSI_SELECTION_RULE = "exact measured Up 0.5 rpm / exact measured 5.0 rpm Up endpoint or shared turnaround"
DECADE_SSI_SELECTION_RULE = "measured RPM pair with configured target ratio"
HIGH_SELECTION_RULE = "actual measured maximum RPM"
SSI_SELECTION_RULE = QC_SSI_SELECTION_RULE
PRIMARY_METRIC_ORDER = ["A_TAU_L_UP", "A_TAU_M_UP", "A_TAU_H_UP", "A_ETA_MIN_UP", "A_ETA_HIGH",
                        "A_SSI_FIELD_0P2_2", "A_SSI_QC_0P5_5"]
PROVENANCE_COLUMNS = [
    "A_SSI_FIELD_0P2_2_value", "A_SSI_FIELD_0P2_2_num_rpm", "A_SSI_FIELD_0P2_2_den_rpm",
    "A_SSI_FIELD_0P2_2_num_step", "A_SSI_FIELD_0P2_2_den_step", "A_SSI_FIELD_0P2_2_value_class",
    "A_SSI_QC_0P5_5_value", "A_SSI_QC_0P5_5_num_rpm", "A_SSI_QC_0P5_5_den_rpm",
    "A_SSI_QC_0P5_5_num_step", "A_SSI_QC_0P5_5_den_step", "A_SSI_QC_0P5_5_value_class",
    "A_TAU_H_UP_rpm", "A_TAU_H_UP_step", "A_ETA_HIGH_rpm", "A_ETA_HIGH_step",
]
CP52_SHEAR_RATE_CONSTANT = 2.0                  # γ̇ = 2.0 × RPM (CSV의 SRC 2.000)
RPM_MATCH_TOL = 0.011
PEAK_SEARCH_SECONDS = 30.0
EQ_DRIFT_REVIEW_PCT = 2.0                        # 후반 10초 준평형 확인(문서 §8-4) — 임시 기준
TORQUE_RESOLUTION_PCT = 0.1                      # DVNext Torque 표시 단위
TEST_TEMP_TARGET_C = 25.0
TEST_TEMP_TOL_C = 0.5
TEMP_SPAN_MAX_C = 1.0
DENSITY_CROSSCHECK_TOL = 0.0005                  # g/mL, 직독값 vs 질량/부피 계산값
RSD_NA_ZERO_REF = "N/A(0 기준)"                   # 0 기준 차이형 지표: RSD 정의상 부적합
RSD_NA_NEAR_ZERO = "N/A(|Mean|<SD)"               # |Mean|<SD → RSD>100 %, 비교지표로 무의미
SIGN_FLIP_STATUS = "REVIEW"                       # zero_ref 지표의 양·음 혼재 시 Lot 데이터 상태
# 06 상태 체계(지시서 §4): OK < NOTICE < REVIEW < CENSORED < EXCEPTION < NO VALUE < INVALID
STATUS_RANK = {"OK": 0, "NOTICE": 1, "N/A": 1, "REVIEW": 2, "CENSORED": 3, "EXCEPTION": 4, "NO VALUE": 5, "INVALID": 6}
NOTICE_BLUE = "DDEBF7"
LOT_STATUS_FILL = {"VALID": GREEN, "GOOD": GREEN, "OK": GREEN, "ELIGIBLE": GREEN, "NOTICE": NOTICE_BLUE,
                   "REVIEW": YELLOW, "WARNING": YELLOW, "EXCEPTION": "F8CBAD", "EXCLUDED": "F8CBAD",
                   "INVALID": RED, "HIGH VAR": RED, "NO DATA": GRAY, "N/A": GRAY, "NO VALUE": GRAY, "CENSORED": "E4DFEC",
                   "—": WHITE}
VALUE_CLASS_FILL = {VC_STRICT: GREEN, VC_EXC: "F4B183", VC_LOW: "E4DFEC", VC_HIGH: "E4DFEC", VC_BOTH: "E4DFEC",
                    VC_NO: GRAY, VC_INVALID: RED}
METRIC_ROLES = ("CORE", "PROCESS_SENSITIVE", "DIAGNOSTIC", "DATA_QUALITY", "CENSORED_METRIC")
ROLE_DEFAULT_BY_TIER = {"T1": "CORE", "T2": "PROCESS_SENSITIVE", "T3": "CORE", "DQ": "DATA_QUALITY"}
ROLE_OVERRIDE = {"A_HB_R2_UP": "DIAGNOSTIC", "A_HB_R2_DN": "DIAGNOSTIC", "A_HB_GRANGE": "DIAGNOSTIC", "A_OVS_L": "DIAGNOSTIC",
                 "A_BRK_PEAK": "PROCESS_SENSITIVE", "A_HYS_REL": "PROCESS_SENSITIVE", "A_HDI": "PROCESS_SENSITIVE",
                 "A_HB_TAUY_RATIO": "PROCESS_SENSITIVE", "A_HYS_AREA_SIGNED": "PROCESS_SENSITIVE",
                 "A_HYS_AREA_ABS": "PROCESS_SENSITIVE", "A_ETA_RECOVERY_LOW": "PROCESS_SENSITIVE",
                 "A_POST_SHEAR_STRESS_RETENTION": "PROCESS_SENSITIVE",
                 "A_HB_TAUY_UP_COMMON": "DIAGNOSTIC", "A_HB_K_UP_COMMON": "DIAGNOSTIC", "A_HB_N_UP_COMMON": "DIAGNOSTIC",
                 "A_HB_R2_UP_COMMON": "DIAGNOSTIC", "A_HB_TAUY_DN_COMMON": "DIAGNOSTIC", "A_HB_K_DN_COMMON": "DIAGNOSTIC",
                 "A_HB_N_DN_COMMON": "DIAGNOSTIC", "A_HB_R2_DN_COMMON": "DIAGNOSTIC", "A_HB_TAUY_RATIO_COMMON": "DIAGNOSTIC",
                 "A_HB_K_RATIO": "DIAGNOSTIC", "A_HB_N_DIFF": "DIAGNOSTIC",
                 "A_ETA_HIGH_UP": "DIAGNOSTIC", "A_ETA_HIGH_DN": "DIAGNOSTIC",
                 "A_STARTUP_PLATEAU_UP": "DIAGNOSTIC", "A_POST_SHEAR_PLATEAU": "DIAGNOSTIC", "A_PLATEAU_RETENTION": "DIAGNOSTIC",
                 "A_DUL_FLOOR": "CENSORED_METRIC", "B_REC_FLOOR": "CENSORED_METRIC", "B_T90": "CENSORED_METRIC"}
TIER_ORDER = ["T1", "T2", "T3", "DQ"]
TIER_LABELS = {
    "T1": "T1 Lot 동일성 핵심",
    "T2": "T2 구조 이력·회복",
    "T3": "T3 모델·진단",
    "DQ": "DQ 데이터 품질",
}
TIER_NOTES = {
    "T1": "매 Lot 비교의 1차 fingerprint. 조성·원료·혼합·탈포 편차를 가장 직접적으로 반영하며, 측정 정밀도가 확보된 직접 관측값 위주로 구성.",
    "T2": "전단 이력에 대한 구조 붕괴와 재형성(요변성의 시간 측면). 원심 중 유동화와 원심 후 장벽 재겔화의 간접 지표.",
    "T3": "모델 외삽·파라미터 상관이 있는 파생값과 진단값. 개발·변경관리·이상조사용이며 기능 상관 검증 전 규격 사용 금지.",
    "DQ": "측정 자체의 유효성(신호 크기·분해능·준평형·온도·정량 Step 수). 물성이 아니라 해당 Lot 값의 신뢰도를 말함.",
}


def _gamma(rpm):
    return CP52_SHEAR_RATE_CONSTANT * float(rpm)


def metric_definitions(ref=None):
    """QC 지표 사전. 각 지표가 가리키는 물성과 해석 한계를 한 곳에서 정의한다.

    prec = (규칙, GOOD 한계, REVIEW 한계): 반복 재현성의 '임시' 데이터 품질 기준.
    규격(USL/LSL)이 아니며 GR&R·Lot 분포 확보 후 교체한다.
    zero_ref=True: 0을 기준으로 부호가 바뀔 수 있는 차이형 지표. RSD는 정의상 부적합하므로
    계산하지 않고(N/A) SD·Range·부호 일관성으로 반복성을 평가한다. 양·음 반복이 섞이면 [SIGN] 경고.
    """
    L, M = FIXED_LOW_RPM, FIXED_MID_RPM
    gL, gM = _gamma(L), _gamma(M)
    step_rule = "60초 회전 Step·후반 10초·Torque 10–95% 유효점만 사용. 미충족 시 공란(대체값 없음)."
    d = [
        # ------------------------------------------------------------------ T1
        dict(code="D_RHO25", method="D", tier="T1", unit="g/mL",
             name="25 °C 탈포 젤 겉보기 밀도 ρ25",
             formula="독립 충전별 ρ = (m충전 − m공용기) / V교정(시험온도 물 교정 부피). 밀도계 직독값이 있으면 직독값을 쓰고 질량식과 교차확인.",
             property="젤 단위 부피당 질량. 수지·가소제 대 실리카 등 충전재의 배합비와 잔류 기포량이 합쳐진 값.",
             function="원심 중 젤이 혈청층과 혈구·혈병층 사이 어느 높이에 정착하는지를 정하는 1차 인자(부력 평형 위치).",
             up="충전재 과량·가소제 부족·계량 오차 가능. 장벽이 혈구층 쪽으로 낮게 형성되거나 상승 이동이 늦어질 수 있음.",
             down="기포 혼입·충전재 부족 가능. 장벽이 혈청층 쪽으로 뜨거나 혈청 중 젤 파편 위험.",
             limits="밀도는 '어디에 정착하는가'만 말하며, '제때 이동·정착·재겔화하는가'는 유변 지표가 담당. DVNext CSV의 Density 열은 장비 입력 설정값이므로 사용하지 않음. 온도 보정 없음.",
             validity=f"시험온도 {TEST_TEMP_TARGET_C:g}±{TEST_TEMP_TOL_C:g} °C, 탈포·충전·상면 정리 조건 고정, 독립 재충전 ≥3회. 소량 중량법은 기준법 비교로 적격화된 경우만.",
             prec=("SD", 0.002, 0.005)),
        dict(code="A_TAU_L_UP", method="A", tier="T1", unit="Pa",
             name=f"저전단 구조 저항 τ↑({L:g} rpm)",
             formula=f"Method A Up 경로 {L:g} rpm(γ̇={gL:g} s⁻¹) Step 대표 전단응력.",
             property="휴지 중 형성된 실리카·겔화제 네트워크가 최저 정량 전단에서 흐름에 저항하는 크기(저전단 구조강도의 상대값).",
             function="보관·운송 중 튜브 바닥 형태 유지, 원심 초기 유동 개시 저항.",
             up="네트워크 강화(겔화제 과다, 냉각 이력에 따른 결정성 구조 등) 가능. 원심 중 이동 개시 지연 여부 확인.",
             down="구조 약화 가능. 보관·운송 중 흘러내림·경사 유동 여부 확인.",
             limits="이미 흐르고 있는 상태의 값이며 정적 항복응력이 아님. 고정 저전단 평가점은 0.5 rpm이다. 0.35 rpm은 Full H-B 확장점이며 이 지표를 대체하지 않는다. H-B fitting 입력 RPM을 제한하지 않는다. Torque가 10% 부근이면 0.1% 표시 단위가 약 1% 상대오차에 해당(A_RES_L 참조).",
             validity=step_rule, prec=("RSD", 10, 20)),
        dict(code="A_TAU_M_UP", method="A", tier="T1", unit="Pa",
             name=f"중간 전단 유동 저항 τ↑({M:g} rpm)",
             formula=f"Method A Up 경로 {M:g} rpm(γ̇={gM:g} s⁻¹) Step 대표 전단응력.",
             property="구조 일부가 전단으로 해체된 상태의 유동 저항. 측정창 중 Torque 여유가 커 반복 정밀도가 가장 좋은 기준점.",
             function="충전·이송 및 원심 중 젤 이동 과정의 대표 유동성(Lot 간 상대 비교).",
             up="배합 점도 상승, 충전재 증가 또는 분산 불량(응집) 가능.",
             down="가소제 과다, 네트워크 형성 부족 가능.",
             limits="실제 원심분리 전단률보다 훨씬 낮은 조건. 절대 이동속도의 예측값이 아님.",
             validity=step_rule, prec=("RSD", 10, 20)),
        dict(code="A_TAU_H_UP", method="A", tier="T1", unit="Pa",
             name="고전단 유동 저항 (실제 최고 RPM)",
             formula="Up 또는 공유 turnaround 중 해당 채널에서 유효한 측정 최고 RPM Step의 대표 전단응력. 더 높은 RPM이 부적격하면 낮은 RPM으로 대체하지 않고 NO VALUE.",
             property="그 측정에서 실제로 도달한 최고 전단의 유동 저항.",
             function="원심 중 이동성 방향의 대리지표.",
             up="매트릭스 점도 상승 또는 충전재 증가.",
             down="매트릭스 점도 저하.",
             limits="5 rpm이 측정되면 4 rpm 응력을 이 지표로 보고하지 않는다. 4 rpm은 중간 측정점이다.",
             validity=step_rule, prec=("RSD", 10, 20)),
        dict(code="A_SSI_FIELD_0P2_2", method="A", tier="T1", unit="-",
             name="현장 전단박화 지수 SSI (0.2/2 rpm)", summary_name="현장 SSI (0.2/2 rpm)",
             formula="eta_up(0.2 rpm) / eta_up(2.0 rpm)",
             property="현장 전단박화 지수. 정확한 0.2 rpm과 2.0 rpm Up 대표 점도의 비.",
             function="저전단 구조 대비 2 rpm 유동화.",
             up="0.2 rpm 대비 2 rpm 점도가 낮음.", down="뉴턴 거동에 가까워짐.",
             limits="0.2 대신 0.35 또는 0.5를 쓰지 않는다. 2.0 대신 1.5 또는 3.0을 쓰지 않는다. 보간·모델값 금지.",
             validity="같은 채널에서 두 실측 Up Step representative가 존재해야 한다. 다른 RPM, 보간값 또는 모델값으로 대체하지 않는다.",
             prec=("RSD", 10, 20)),
        dict(code="A_SSI_QC_0P5_5", method="A", tier="T1", unit="-",
             name="Step-flow QC 전단박화 지수 SSI (0.5/5 rpm)", summary_name="Step-flow QC SSI (0.5/5 rpm)",
             formula="eta_up(0.5 rpm) / eta(5.0 rpm Up endpoint 또는 실제 shared turnaround)",
             property="Step-flow QC 전단박화 지수. 정확한 0.5 rpm Up과 5.0 rpm endpoint의 비.",
             function="0.5 rpm 구조 대비 5 rpm 유동화.",
             up="0.5 rpm 대비 5 rpm 점도가 낮음.", down="뉴턴 거동에 가까워짐.",
             limits="0.5 대신 0.35를 쓰지 않는다. 5.0 대신 4.0을 쓰지 않는다. 실제 최저/최고 RPM 비율로 정의를 바꾸지 않는다.",
             validity="같은 채널에서 두 실측 Step representative가 존재해야 한다. 0.35 또는 4 rpm으로 대체하지 않는다.",
             prec=("RSD", 10, 20)),
        dict(code="REST_STABILIZATION", method="A", tier="T1", unit="-", output_policy="dictionary_only",
             name="무전단 구조 안정화",
             formula="0 rpm, 25 °C, 목표 600초",
             property="0 rpm, 25 °C, 목표 600초. 시료 이송 중 구조 교란 회복을 위한 전처리 구간. 유동·H-B·SSI 계산에서는 제외한다.",
             function="측정 위치로 옮기는 동안 생긴 구조 교란을 유동 측정 전에 회복한다.",
             up="—", down="—",
             limits="moving Step, SSI 후보, H-B 후보, hysteresis 후보가 아니다.",
             validity="유동 지표에서는 제외하고 시험 전처리 QC로만 평가한다.",
             prec=("-", None, None)),
        dict(code="A_SSI_DECADE", method="A", tier="T1", unit="-", output_policy="decade_sheet",
             name="실측 10배 RPM 점도비 SSI",
             formula="eta(low RPM) / eta(high RPM), high RPM / low RPM = 10 within configured tolerance",
             property="프로토콜에서 실제로 측정된 10배 RPM 쌍의 점도비. pair_id별 독립 지표.",
             function="공식 0.2/2, 0.5/5 이외의 실측 decade 쌍을 프로토콜 변경에 맞춰 보존한다.",
             up="저점 대비 고점 점도가 낮음.", down="뉴턴 거동에 가까워짐.",
             limits="서로 다른 pair를 같은 Lot 통계에 혼합하지 않는다. 0 rpm pair는 만들지 않는다. 0.35/4는 decade가 아니다.",
             validity="pair_id별 독립 지표로 관리한다. 서로 다른 pair를 같은 Lot 통계에 혼합하지 않는다.",
             prec=("RSD", 10, 20)),
        # ------------------------------------------------------------------ T2
        dict(code="A_DU_L", method="A", tier="T2", unit="%",
             name=f"저전단 잔류 구조율 D/U({L:g} rpm)",
             formula=f"τ_down({L:g} rpm) / τ_up({L:g} rpm) × 100 (동일 전단률, 각 Step 대표값).",
             property="최고 전단 이력 후 하강 경로에서 같은 저전단으로 돌아왔을 때 남아 있거나 재형성된 구조의 비율(Method A 시간척도).",
             function="원심으로 유동화된 젤이 전단 감소 시 구조를 다시 세우는 경향. 장벽 재겔화의 간접 지표.",
             up="100%에 가까울수록 전단 이력 영향이 작거나 빠르게 재구축됨.",
             down="전단에 의한 구조 손실이 크고 수십 초 내 재구축이 부족. Method B 회복 확인 우선.",
             limits="100% 초과는 하강 중 구조 형성(rheopexy), 휴지 구조 미평형, 벽면 미끄럼 변화 등을 조사. 무전단 휴지 회복률이 아님. 하강 Torque가 10% 미만이면 값 대신 CENSORED(검출하한 A_DUL_FLOOR 미만)로 표시.",
             validity=step_rule + " Up·Down 모두 유효해야 계산.", prec=("SD", 2, 5)),
        dict(code="A_DU_M", method="A", tier="T2", unit="%",
             name=f"중간 전단 잔류 구조율 D/U({M:g} rpm)",
             formula=f"τ_down({M:g} rpm) / τ_up({M:g} rpm) × 100.",
             property="중간 전단에서 고전단 이력 전후 유동 저항의 비. 저전단 D/U보다 Torque 여유가 커 정밀도가 좋음.",
             function="원심 감속 구간에서의 구조 재형성 경향.",
             up="전단 이력 영향이 작음.",
             down="고전단 이력으로 구조가 더 많이 손상됨.",
             limits="A_DU_L과 함께 해석. 단독 합부 판단 금지.",
             validity=step_rule + " Up·Down 모두 유효해야 계산.", prec=("SD", 2, 5)),
        dict(code="A_HYS_REL", method="A", tier="T2", unit="%",
             name="상대 응력 히스테리시스",
             formula="공통 전단률에서 |τ_up − τ_down| 사다리꼴 적분 / τ_up 사다리꼴 적분 × 100 (04_2와 동일 계산).",
             property="상승·하강 경로 전체에서 전단 이력으로 생긴 곡선 차이의 정규화 크기(구조 붕괴와 재구축의 불균형 총량).",
             function="원심 전단 이력에 대한 구조 취약성 비교.",
             up="전단 이력에 의한 구조 변화가 큼.",
             down="상승·하강 곡선이 유사. 구조가 안정하거나 매우 빠르게 재구축.",
             limits="γ̇ 간격이 넓은 고전단 구간에 가중됨. Step 구성·유지시간·rest를 고정한 경우만 Lot 간 비교. 에너지 값이 아님.",
             validity="공통 전단률 5쌍 이상 권장(4쌍 REVIEW).", prec=("SD", 1, 2)),
        dict(code="A_BRK_PEAK", method="A", tier="T2", unit="%",
             name="최고 전단 Step 내 구조 붕괴율",
             formula="Up 최고 RPM Step에서 (τ_peak − τ_대표) / τ_peak × 100. τ_peak는 Step 초기 30초(또는 절반) 내 Torque 유효점의 3점 이동중앙값 최대.",
             property="일정한 고전단 유지 중 시간에 따라 구조가 계속 해체되는 정도(요변성의 붕괴 측면). 0 % 기준의 차이형 지표.",
             function="원심 중 지속 전단에서 젤이 얼마나 더 묽어지는가.",
             up="지속 전단 시 구조 약화가 큼.",
             down="0에 가까우면 구조가 전단에 안정하거나 Step 초기에 이미 평형. 음수는 대표창 응력이 초기 peak보다 높다는 뜻(peak 미검출 또는 전단 중 구조 형성) — 원시곡선 확인.",
             limits="Step 초기 수 초의 장비 응답·탄성 부하 영향이 포함됨. 최고 RPM이 다른 프로그램끼리는 비교 불가. 평균이 0 부근이므로 RSD는 산정하지 않고 SD·Range·부호 일관성으로 반복성을 봄. 부호 전환이 있으면 반복성 GOOD이라도 [SIGN] REVIEW.",
             validity=step_rule, prec=("SD", 1, 2), zero_ref=True),
        dict(code="B_ETA_BASE", method="B", tier="T2", unit="cP",
             name="3ITT 기준 저전단 점도 η_base",
             formula="Interval 1(저전단) 후반 창의 Torque 유효 평균 점도.",
             property="시험 직전 휴지 구조의 저전단 저항. 회복률 계산의 기준 구조(분모).",
             function="휴지·보관 상태 구조의 기준선.",
             up="휴지 구조가 강함.",
             down="휴지 구조가 약함.",
             limits="Method A τ↑(L)과 RPM이 같아도 rest 시간·선행 전단이 다르면 값이 다를 수 있음.",
             validity="기준 구간 Torque 10–95%.", prec=("RSD", 10, 20)),
        dict(code="B_BRK_DECAY", method="B", tier="T2", unit="%",
             name="3ITT 고전단 구간 붕괴율",
             formula="Interval 2에서 (τ_peak − τ_후반창) / τ_peak × 100.",
             property="정해진 고전단 유지 중 구조 해체의 진행 정도(원심 전단 모사 구간의 붕괴량). 0 % 기준의 차이형 지표.",
             function="원심 중 구조 해체의 크기.",
             up="고전단 지속 시 구조 손실이 큼.",
             down="고전단에 대한 구조 안정성이 큼. 음수는 후반 창 응력이 peak보다 높음 — 원시곡선 확인.",
             limits="고전단 크기·시간이 실제 원심 조건과 다름. 동일 프로그램 내 상대 비교. 평균이 0 부근이면 RSD 산정 안 함.",
             validity="고전단 구간 Torque 10–95%.", prec=("SD", 1, 2), zero_ref=True),
        dict(code="B_R60", method="B", tier="T2", unit="%",
             name="60초 구조 회복률 R60",
             formula="Interval 3 시작 후 55–65초 η 평균 / η_base × 100.",
             property="고전단 해제 직후 1분 안에 저전단에서 기준 대비 재형성된 구조의 비율.",
             function="원심 정지 직후 장벽이 다시 굳어 이송·기울임을 견디기 시작하는 초기 속도.",
             up="빠른 재겔화.",
             down="재겔화 지연. 원심 직후 이송·진동 시 장벽 흐트러짐 위험 검토.",
             limits="저전단 회전 중 회복(무전단 휴지 회복 아님). 100% 초과는 기준 구간 미평형 또는 기준 이상의 재구축. 회복 중 Torque<10%이면 CENSORED(검출하한 B_REC_FLOOR 미만).",
             validity="기준·회복 구간 Torque 10–95%.", prec=("SD", 3, 6)),
        dict(code="B_R300", method="B", tier="T2", unit="%",
             name="300초 구조 회복률 R300",
             formula="Interval 3 시작 후 295–305초 η 평균 / η_base × 100.",
             property="5분 시간척도에서의 구조 재형성 정도(느린 재구축 성분 포함).",
             function="원심 후 검체 대기·운반 동안의 장벽 안정성.",
             up="중기 재구축이 큼.",
             down="느린 재구축 또는 비가역적 구조 손실.",
             limits="회복 구간이 300초 미만이면 공란.",
             validity="기준·회복 구간 Torque 10–95%.", prec=("SD", 3, 6)),
        dict(code="B_T90", method="B", tier="T2", unit="s",
             name="90% 회복 도달시간 t90",
             formula="Interval 3에서 11점 이동중앙값 η가 η_base의 90%에 처음 도달한 시간. 미도달은 't90 > 구간길이'(right-censored), 교차가 Torque<10% 구간에서 일어나 첫 유효점에서 이미 90% 이상이면 't90 ≤ 첫 유효점'(상한값)으로 표시하고 공란 처리.",
             property="구조 재형성의 시간척도.",
             function="원심 정지 후 장벽이 기능 수준으로 굳는 데 걸리는 시간의 대리값.",
             up="재형성이 느림.",
             down="재형성이 빠름.",
             limits="미도달 반복이 있으면 평균은 도달 반복만의 값이므로 편향됨(사유 열의 미도달 수 확인).",
             validity="기준·회복 구간 Torque 10–95%.", prec=("RSD", 20, 35)),
        # ------------------------------------------------------------------ T3
        dict(code="A_HB_TAUY_UP", method="A", tier="T3", unit="Pa",
             name="동적 항복응력 — 상승 Full H-B",
             formula="τ = τy + K·γ̇ⁿ. 분석 가능한 모든 Up Step 대표값(Up Full). legacy reference RPM이나 공통 전단률로 제한하지 않음.",
             property="상승 유동곡선 전체를 γ̇→0으로 외삽한 모델상 흐름 개시 응력.",
             function="초기 구조강도의 모델 기반 요약값(개발·변경관리용).",
             up="외삽상 흐름 개시 응력 증가.",
             down="외삽상 흐름 개시 응력 감소.",
             limits="외삽 추정값이며 정적 항복응력이 아님. 0.5/2 rpm과 4 rpm reference는 H-B 입력을 제한하지 않는다. Common H-B는 A_HB_*_COMMON.",
             validity="유효점 4개 최소(REVIEW), 5개 이상 권장.", prec=("RSD", 10, 20)),
        dict(code="A_HB_K_UP", method="A", tier="T3", unit="Pa·sⁿ",
             name="Consistency K↑ (H-B, Up Full)",
             formula="Up Full H-B 피팅의 K.",
             property="항복 후 유동 저항의 크기. n과 결합되어 의미를 가짐.",
             function="원심 중 유동 저항의 모델 요약값.",
             up="같은 n에서 유동 저항 증가.", down="같은 n에서 유동 저항 감소.",
             limits="n이 달라지면 K 단위 자체가 달라지므로 n이 다른 Lot 간 직접 비교 금지.",
             validity="유효점 4개 최소(REVIEW), 5개 이상 권장.", prec=("RSD", 10, 20)),
        dict(code="A_HB_N_UP", method="A", tier="T3", unit="-",
             name="Flow index n↑ (H-B, Up Full)",
             formula="Up Full H-B 피팅의 n.",
             property="전단률 의존 형상. n<1 전단박화, n=1 Bingham형.",
             function="전단 증가에 따른 유동화 형상.",
             up="1에 가까워짐(전단박화 약화).", down="전단박화 강화.",
             limits="K와 상호 보상되는 파라미터. SSI(T1)가 모델 없는 대응 지표.",
             validity="유효점 4개 최소(REVIEW), 5개 이상 권장.", prec=("RSD", 10, 20)),
        dict(code="A_HB_R2_UP", method="A", tier="T3", unit="-",
             name="H-B 적합도 R²↑ Full",
             formula="1 − SSE/SST (Up Full).",
             property="측정점과 H-B 모델의 일치도. 품질 우수성 지표가 아니라 계산 신뢰성 확인값(진단).",
             function="—", up="모델 적합 양호.", down="기포·불균일·시간의존성·모델 부적합 조사.",
             limits="R² 단독으로 Lot 합격/불합격 또는 반복 제외 금지. 상태는 항상 NOTICE.",
             validity="—", prec=("-", None, None)),
        dict(code="A_HB_R2_DN", method="A", tier="T3", unit="-",
             name="H-B 적합도 R²↓ Full",
             formula="1 − SSE/SST (Down Full).",
             property="하강 경로 측정점과 H-B 모델의 일치도. 계산 신뢰성 확인값(진단).",
             function="—", up="모델 적합 양호.", down="기포·불균일·시간의존성·모델 부적합 조사.",
             limits="R² 단독으로 Lot 합격/불합격 또는 반복 제외 금지. 상태는 항상 NOTICE.",
             validity="—", prec=("-", None, None)),
        dict(code="A_HB_TAUY_DN", method="A", tier="T3", unit="Pa",
             name="동적 항복응력 — 하강 Full H-B",
             formula="분석 가능한 모든 Down Step 대표값(Down Full) H-B 피팅의 τy.",
             property="고전단 이력 후 하강 곡선 전체의 외삽 흐름 개시 응력.",
             function="전단 이력 후 구조의 모델 요약값.",
             up="하강 중 재구축이 큼.", down="고전단 후 구조 약화가 큼.",
             limits="완전 휴지 회복값이 아님. Common fit으로 대체하지 않는다.",
             validity="유효점 4개 최소(REVIEW), 5개 이상 권장.", prec=("RSD", 10, 20)),
        dict(code="A_HB_K_DN", method="A", tier="T3", unit="Pa·sⁿ",
             name="Consistency K↓ (H-B, Down Full)", formula="Down Full H-B 피팅의 K.",
             property="고전단 이력 후 하강 상태의 유동 저항 크기.", function="—",
             up="하강 중 유동 저항 큼.", down="하강 중 유동 저항 작음.",
             limits="n↓과 함께 해석.", validity="유효점 4개 최소(REVIEW).", prec=("RSD", 10, 20)),
        dict(code="A_HB_N_DN", method="A", tier="T3", unit="-",
             name="Flow index n↓ (H-B, Down Full)", formula="Down Full H-B 피팅의 n.",
             property="하강 프로파일 형상.", function="—",
             up="상승 대비 전단률 의존성 약화.", down="상승 대비 전단박화 강화.",
             limits="단독 좋음/나쁨 판정 금지.", validity="유효점 4개 최소(REVIEW).", prec=("RSD", 10, 20)),
        dict(code="A_HB_TAUY_RATIO", method="A", tier="T3", unit="%",
             name="H-B 항복응력 유지율 — Down/Up Full",
             formula="Down Full τy / Up Full τy × 100. Up τy가 0이면 NO VALUE.",
             property="전단 이력 후 외삽 항복 구조의 유지 비율. Full fit끼리의 비.",
             function="A_DU_L의 모델 기반 대응값.",
             up="상승 상태에 근접.", down="전단 후 구조 손실 큼.",
             limits="두 Full fit의 γ̇ 범위가 다르면 REVIEW. Common 비율은 A_HB_TAUY_RATIO_COMMON.",
             validity="두 Full 피팅 모두 계산된 경우.", prec=("SD", 5, 10)),
        dict(code="A_HDI", method="A", tier="T3", unit="%",
             name="등가중 이력 차이 지수 HDI",
             formula="Σ|τ_up − τ_down| / Στ_up × 100 (공통 전단률, 점별 등가중).",
             property="γ̇ 간격 가중이 없는 상승·하강 차이의 평균적 크기. 저전단 쪽 차이도 동등하게 반영.",
             function="A_HYS_REL의 가중 편향을 점검하는 보조 지표.",
             up="전단 이력 차이가 큼.", down="차이가 작음.",
             limits="Step 구성 고정 시에만 비교.", validity="공통 전단률 4쌍 이상.", prec=("SD", 1, 2)),
        dict(code="A_OVS_L", method="A", tier="T3", unit="%",
             name=f"휴지 후 Start-up 응력 초과율({L:g} rpm)",
             formula=f"휴지 직후 첫 회전 Step에서 (τ_peak / τ_대표 − 1) × 100. 기준 RPM {L:g} rpm 여부와 무관.",
             property="휴지 중 형성된 구조가 흐름 개시 순간 깨지며 생기는 응력 초과. 첫 회전 Step이 기준 RPM보다 낮아도 실제 start-up 전이 응답을 사용.",
             function="보관 중 형성되는 휴지 구조의 크기(정적 항복응력의 정성적 대리).",
             up="휴지 구조 형성이 큼.", down="휴지 구조 형성이 작음.",
             limits=f"rest 직후 첫 회전 Step의 실제 RPM을 사용한다. {L:g} rpm이 아니어도 계산하며 결과 사유에 실제 RPM을 기록한다. 저토크 분해능과 장비 가속 응답 포함. Vane 정적 항복응력을 대체하지 않음. 0 기준 차이형 지표라 RSD 산정 안 함.",
             validity="rest 직후 첫 회전 Step, Torque 유효. 미산출은 NOTICE(No quantifiable start-up overshoot)로 표시하며 Run·Lot 분석은 유지. Lot 합격성·정밀도 판정 대상 아님.",
             prec=("-", None, None), zero_ref=True),
        dict(code="A_HB_GRANGE", method="A", tier="T3", unit="-",
             name="H-B 피팅 전단률 범위 비 γ̇max/γ̇min",
             formula="Up Full 피팅의 γ̇max / γ̇min. legacy reference RPM은 이 범위를 제한하지 않는다.",
             property="외삽 거리의 역지표. 작을수록 τy·K·n 추정 불확실성이 큼.",
             function="—", up="모델 파라미터 신뢰도 증가.", down="10 미만(1 decade 미만)이면 T3 모델값을 경향 참고로만 사용.",
             limits="데이터 유효성 지표(진단). 10 미만은 NOTICE이며 Lot 유효성에 전파하지 않음.", validity="—", prec=("-", None, None)),
        dict(code="B_ETA_BRK", method="B", tier="T3", unit="cP",
             name="3ITT 고전단 점도 η_brk", formula="Interval 2 후반 창 Torque 유효 평균 점도.",
             property="구조가 해체된 고전단 상태의 유동 저항.", function="원심 중 유동성 대리.",
             up="고전단에서도 저항이 큼.", down="고전단 저항이 작음.",
             limits="전단박화와 구조 붕괴가 합쳐진 값.", validity="고전단 구간 Torque 10–95%.", prec=("RSD", 10, 20)),
        dict(code="B_R10", method="B", tier="T3", unit="%",
             name="10초 즉시 회복률 R10", formula="Interval 3 시작 후 5–15초 η 평균 / η_base × 100.",
             property="고전단 해제 직후 즉시 회복(탄성·빠른 접점 재형성).", function="원심 정지 순간의 장벽 즉시 안정성.",
             up="즉시 재구축이 빠름.", down="초기 재구축이 느림.",
             limits="속도 전환 지연·장비 응답의 영향이 가장 큼.", validity="Torque 10–95%.", prec=("SD", 3, 6)),
        dict(code="B_R900", method="B", tier="T3", unit="%",
             name="900초 장기 회복률 R900", formula="Interval 3 시작 후 895–905초 η 평균 / η_base × 100.",
             property="장기 재구축 또는 비가역 손실.", function="기준 구축 단계용.",
             up="15분 내 구조 복귀가 큼.", down="느린 재구축·비가역 손실.",
             limits="회복 구간이 900초 미만이면 공란.", validity="Torque 10–95%.", prec=("SD", 3, 6)),
        # ------------------------------------------------------------------ 동적 RPM 물성·Full/Common 진단
        dict(code="A_ETA_MIN_UP", method="A", tier="T1", unit="cP",
             name="최저 측정 RPM 겉보기 점도 — 상승",
             formula="분석 가능한 Up Step 중 실제 최저 RPM의 대표 겉보기 점도. 실제 최저 측정 RPM 사용; SSI의 고정 0.5 rpm 분자와 다름.",
             property="상승 경로에서 가장 낮은 전단의 겉보기 점도.",
             function="휴지 구조가 흐름에 저항하는 점성 크기.",
             up="저전단 저항 증가.", down="저전단 저항 감소.",
             limits="실제 최저 측정 RPM 사용; SSI의 고정 0.5 rpm 분자와 다름. 0.35 rpm이 있으면 그 점도를 쓴다. 최저 RPM이 반복마다 다르면 Lot 통계는 REVIEW.",
             validity=step_rule, prec=("RSD", 10, 20)),
        dict(code="A_ETA_HIGH", method="A", tier="T1", unit="cP",
             name="고전단 겉보기 점도 (실제 최고 RPM)",
             formula="해당 채널에서 사용 가능한 실제 최고 RPM Step의 겉보기 점도. 한 번만 측정되면 공유 turnaround 점도를 A_ETA_HIGH에만 기록하고, 더 높은 RPM이 부적격하면 낮은 RPM으로 대체하지 않는다.",
             property="프로그램 최고 전단에서의 겉보기 점도.",
             function="원심에 가까운 고전단 유동 저항.",
             up="고전단 저항 큼.", down="고전단 저항 작음.",
             limits="최고 RPM이 Up과 Down에 각각 따로 측정되면 A_ETA_HIGH_UP/DN을 따로 본다. 단일 turnaround 값은 복제하지 않는다. 4 rpm reference로 대체하지 않는다.",
             validity=step_rule, prec=("RSD", 10, 20)),
        dict(code="A_ETA_HIGH_UP", method="A", tier="T1", unit="cP", output_policy="when_finite",
             name="고전단 겉보기 점도 — 상승 최고 RPM",
             formula="최고 RPM이 Up과 Down에 각각 있을 때만 Up 쪽 대표 점도.",
             property="상승 최고속의 겉보기 점도.", function="—",
             up="—", down="—",
             limits="단일 peak-speed Step이면 값을 복제하지 않고 비어 있다. A_ETA_HIGH를 본다.",
             validity="—", prec=("-", None, None)),
        dict(code="A_ETA_HIGH_DN", method="A", tier="T1", unit="cP", output_policy="when_finite",
             name="고전단 겉보기 점도 — 하강 최고 RPM",
             formula="최고 RPM이 Up과 Down에 각각 있을 때만 Down 쪽 대표 점도.",
             property="하강 최고속의 겉보기 점도.", function="—",
             up="—", down="—",
             limits="단일 peak-speed Step이면 값을 복제하지 않는다.",
             validity="—", prec=("-", None, None)),
        dict(code="A_ETA_LOW_DN", method="A", tier="T1", unit="cP",
             name="저전단 겉보기 점도 — 하강",
             formula="분석 가능한 Down Step 중 실제 최저 RPM의 대표 겉보기 점도.",
             property="고전단 이력 후 가장 낮은 하강 전단의 겉보기 점도.",
             function="전단 이력 후 저전단 저항.",
             up="잔류 저항 큼.", down="잔류 저항 작음.",
             limits="Up 최저 RPM과 다르면 그 RPM을 그대로 쓰고 비고에 불일치를 남긴다. 점도 회복률의 분모로 직접 나누지 않는다.",
             validity=step_rule, prec=("RSD", 10, 20)),
        dict(code="A_STATIC_YIELD_APPROX_UP", method="A", tier="T2", unit="Pa",
             name="근사 정적 항복응력 — 상승 Start-up",
             formula="분석 가능한 Up 최저 RPM Step의 start-up 검색창에서 3점 이동중앙값 응력 최대의 절대값.",
             property="초기 저전단에서 흐름이 시작될 때의 응력 피크. A_OVS_L(초과율 %)과 다른 절대 응력이다.",
             function="휴지 구조의 흐름 개시 응력에 대한 근사.",
             up="개시 응력 큼.", down="개시 응력 작음.",
             limits="Vane 정적 항복응력을 대체하지 않음. STRICT는 Torque 10–95% 점만, INCLUSIVE는 정책이 허용하면 범위 밖 점도 peak에 포함한다.",
             validity="유효 peak가 있을 것.", prec=("RSD", 15, 30)),
        dict(code="A_POST_SHEAR_STARTUP_STRESS", method="A", tier="T2", unit="Pa",
             name="Post-shear residual start-up stress",
             formula="Down 방향 실제 최저 RPM Step의 start-up 검색창 응력 피크 절대값.",
             property="고전단 이력 후 저전단으로 돌아왔을 때의 잔류 개시 응력.",
             function="전단 후 남은 구조의 개시 저항.",
             up="잔류 개시 응력 큼.", down="잔류 개시 응력 작음.",
             limits="Up 최저 RPM과 다르면 REVIEW. 동일 RPM 전후 비는 A_POST_SHEAR_STRESS_RETENTION.",
             validity="유효 peak가 있을 것.", prec=("RSD", 15, 30)),
        dict(code="A_POST_SHEAR_STRESS_RETENTION", method="A", tier="T2", unit="%",
             name="Post-shear stress retention",
             formula="동일 최저 공통 RPM에서 Down peak stress / Up peak stress × 100. 공통 RPM이 없으면 NO VALUE.",
             property="같은 저전단에서 고전단 이력 전후 개시 응력의 유지 비율.",
             function="구조 유지의 직접 비교.",
             up="개시 응력이 유지됨.", down="개시 응력이 줄어듦.",
             limits="plateau retention과 다르다. 보간하지 않는다.",
             validity="Up·Down 공통 RPM과 두 peak가 있을 것.", prec=("SD", 5, 10)),
        dict(code="A_ETA_RECOVERY_LOW", method="A", tier="T2", unit="%",
             name="점도 회복률 — 최저 공통 RPM",
             formula="Up·Down 모두 분석 가능한 RPM 중 가장 낮은 RPM에서 η_down / η_up × 100. 점도 열에서 계산.",
             property="최저 공통 전단에서의 겉보기 점도 회복률.",
             function="고전단 이력 후 저전단 점도가 얼마나 돌아오는가.",
             up="점도 회복이 큼.", down="점도 회복이 작음.",
             limits="low_shear_recovery_pct(응력비)와 정의를 바꾸지 않는다. 공통 RPM이 없으면 보간하지 않고 NO VALUE. 유일한 공통점이 단일 turnaround이면 같은 Step을 한 번만 세고 회복률은 100%다.",
             validity="공통 RPM이 있을 것.", prec=("SD", 3, 6)),
        dict(code="A_HYS_AREA_SIGNED", method="A", tier="T2", unit="Pa/s",
             name="응력 히스테리시스 signed area",
             formula="공통 shear-rate pair를 γ̇ 오름차순으로 (τ_up − τ_down) 사다리꼴 적분.",
             property="상승·하강 응력 차이의 부호 있는 면적.",
             function="이력의 방향(하강이 더 낮은지).",
             up="하강 응력이 더 낮음.", down="하강 응력이 더 높거나 차이가 작음.",
             limits="공유 turnaround는 적분 상한에 포함되고 그 점의 Δτ는 0이다. 한쪽에만 있는 그 외 RPM은 이 적분에서 빠진다. 공통 pair 2개 미만이면 NO VALUE.",
             validity="공통 pair ≥2.", prec=("SD", 5, 10), zero_ref=True),
        dict(code="A_HYS_AREA_ABS", method="A", tier="T2", unit="Pa/s",
             name="응력 히스테리시스 absolute area",
             formula="공통 pair에서 |τ_up − τ_down| 사다리꼴 적분.",
             property="부호를 접은 이력 면적.",
             function="경로 차이의 총량.",
             up="경로 차이 큼.", down="경로 차이 작음.",
             limits="unpaired Step은 제외. A_HYS_REL의 분자.",
             validity="공통 pair ≥2.", prec=("SD", 5, 10)),
        dict(code="A_HB_K_RATIO", method="A", tier="T3", unit="%",
             name="K 유지비 Down/Up Full",
             formula="Down Full K / Up Full K × 100.",
             property="Full fit consistency의 유지 비율.", function="—",
             up="—", down="—",
             limits="n이 다르면 K 단위가 달라 해석에 주의. 진단값.",
             validity="두 Full 피팅.", prec=("-", None, None)),
        dict(code="A_HB_N_DIFF", method="A", tier="T3", unit="-",
             name="n 차이 Down−Up Full",
             formula="Down Full n − Up Full n.",
             property="Full fit flow index 차이.", function="—",
             up="—", down="—",
             limits="진단값.", validity="두 Full 피팅.", prec=("-", None, None)),
        dict(code="A_HB_TAUY_UP_COMMON", method="A", tier="T3", unit="Pa",
             name="τy↑ Common H-B 진단", formula="Up·Down 공통 전단률만으로 피팅한 τy.",
             property="직접 비교가 가능한 구간의 외삽 항복응력.", function="—",
             up="—", down="—", limits="대표 H-B가 아니다. unpaired RPM은 제외.", validity="공통점 4개 이상.", prec=("-", None, None)),
        dict(code="A_HB_K_UP_COMMON", method="A", tier="T3", unit="Pa·sⁿ",
             name="K↑ Common H-B 진단", formula="Up Common H-B의 K.",
             property="—", function="—", up="—", down="—", limits="진단값.", validity="—", prec=("-", None, None)),
        dict(code="A_HB_N_UP_COMMON", method="A", tier="T3", unit="-",
             name="n↑ Common H-B 진단", formula="Up Common H-B의 n.",
             property="—", function="—", up="—", down="—", limits="진단값.", validity="—", prec=("-", None, None)),
        dict(code="A_HB_R2_UP_COMMON", method="A", tier="T3", unit="-",
             name="R²↑ Common H-B 진단", formula="Up Common H-B의 R².",
             property="—", function="—", up="—", down="—", limits="진단값. NOTICE.", validity="—", prec=("-", None, None)),
        dict(code="A_HB_TAUY_DN_COMMON", method="A", tier="T3", unit="Pa",
             name="τy↓ Common H-B 진단", formula="Down Common H-B의 τy.",
             property="—", function="—", up="—", down="—", limits="진단값.", validity="—", prec=("-", None, None)),
        dict(code="A_HB_K_DN_COMMON", method="A", tier="T3", unit="Pa·sⁿ",
             name="K↓ Common H-B 진단", formula="Down Common H-B의 K.",
             property="—", function="—", up="—", down="—", limits="진단값.", validity="—", prec=("-", None, None)),
        dict(code="A_HB_N_DN_COMMON", method="A", tier="T3", unit="-",
             name="n↓ Common H-B 진단", formula="Down Common H-B의 n.",
             property="—", function="—", up="—", down="—", limits="진단값.", validity="—", prec=("-", None, None)),
        dict(code="A_HB_R2_DN_COMMON", method="A", tier="T3", unit="-",
             name="R²↓ Common H-B 진단", formula="Down Common H-B의 R².",
             property="—", function="—", up="—", down="—", limits="진단값. NOTICE.", validity="—", prec=("-", None, None)),
        dict(code="A_HB_TAUY_RATIO_COMMON", method="A", tier="T3", unit="%",
             name="τy 유지비 Common 진단", formula="Down Common τy / Up Common τy × 100.",
             property="공통 구간만의 항복응력 유지비.", function="—",
             up="—", down="—", limits="Full 유지비와 섞지 않는다.", validity="두 Common 피팅.", prec=("-", None, None)),
        dict(code="A_STARTUP_PLATEAU_UP", method="A", tier="T3", unit="Pa",
             name="상승 저전단 plateau 응력", formula="초기 저전단 Step 후반 10초 평균 응력.",
             property="start-up 이후 안정 응력. peak와 다르다.", function="—",
             up="—", down="—", limits="보조 진단.", validity="—", prec=("-", None, None)),
        dict(code="A_POST_SHEAR_PLATEAU", method="A", tier="T3", unit="Pa",
             name="하강 저전단 plateau 응력", formula="post-shear 저전단 Step 후반 10초 평균 응력.",
             property="이력 후 안정 응력.", function="—",
             up="—", down="—", limits="보조 진단.", validity="—", prec=("-", None, None)),
        dict(code="A_PLATEAU_RETENTION", method="A", tier="T3", unit="%",
             name="plateau 유지율", formula="post plateau / initial plateau × 100.",
             property="안정 응력의 유지. peak retention과 다르다.", function="—",
             up="—", down="—", limits="보조 진단.", validity="—", prec=("-", None, None)),
        # ------------------------------------------------------------------ DQ
        dict(code="A_TQ_L", method="A", tier="DQ", unit="%",
             name=f"저전단 기준점 Torque({L:g} rpm)", formula=f"Up {L:g} rpm Step 대표창의 유효 Torque 평균.",
             property="저전단 측정 신호의 크기와 10% 하한까지의 여유.", function="—",
             up="신호 여유 증가.", down="하한 근접 — Lot에 따라 공란이 생길 수 있음.",
             limits="물성이 아닌 측정 여유.", validity="10–95%.", prec=("-", None, None)),
        dict(code="A_RES_L", method="A", tier="DQ", unit="%",
             name=f"저전단 기준점 분해능 한계({L:g} rpm)",
             formula=f"{TORQUE_RESOLUTION_PCT:g}% / Torque_L × 100.",
             property="Torque 표시 1단위가 τ↑(L)에서 차지하는 상대 크기. 이보다 작은 Lot 간 차이는 판별할 수 없음.",
             function="—", up="작은 차이 판별 불가.", down="판별력 양호.",
             limits="측정 시스템 지표.", validity="—", prec=("-", None, None)),
        dict(code="A_DUL_FLOOR", method="A", tier="DQ", unit="%",
             name=f"저전단 D/U 검출하한({L:g} rpm)",
             formula=f"Torque 하한({DEFAULT_TORQUE_MIN:g}%) / Torque_L(Up) × 100.",
             property="하강 경로 저전단 Torque가 10% 미만으로 떨어져 A_DU_L이 공란(CENSORED)이 되는 경계. A_DU_L이 이 값에 가까우면 구조 손실이 큰 Lot일수록 값이 사라지는 구조적 편향이 생김.",
             function="—", up="측정 가능한 D/U 범위가 좁아짐 — 저전단 기준 RPM 상향 검토.", down="측정 범위 넓음.",
             limits="측정 시스템 지표.", validity="—", prec=("-", None, None)),
        dict(code="A_EQ_DRIFT_MAX", method="A", tier="DQ", unit="%",
             name="대표창 최대 드리프트",
             formula="정량 Step마다 후반 10초 유효점의 선형 기울기 × 10초 / 평균 × 100, 그 절댓값의 최대.",
             property="대표창에서 응력이 아직 변하고 있는 정도(준평형 확인, QC-RHEO-A-001 §8-4).",
             function="—", up=f"{EQ_DRIFT_REVIEW_PCT:g}% 초과 시 REVIEW — 유지시간 연장 검토.", down="준평형에 가까움.",
             limits="임시 기준. Torque 분해능 계단이 드리프트처럼 보일 수 있음.", validity="—", prec=("-", None, None)),
        dict(code="A_TEMP_MEAN", method="A", tier="DQ", unit="°C",
             name="정량 Step 평균 시료 온도", formula="정량 Step 대표창 온도 평균.",
             property="측정 온도 조건. 점도는 온도에 민감하므로 Lot 간 비교의 전제.", function="—",
             up=f"{TEST_TEMP_TARGET_C:g}±{TEST_TEMP_TOL_C:g} °C 밖이면 REVIEW.", down="동일.",
             limits="CSV의 Temperature Control이 No이면 측정만 하고 제어는 하지 않은 것.", validity="—", prec=("-", None, None)),
        dict(code="A_TEMP_SPAN", method="A", tier="DQ", unit="°C",
             name="정량 Step 온도 변동폭", formula="정량 Step 대표창 온도 max − min.",
             property="시험 중 온도 안정성.", function="—", up=f"{TEMP_SPAN_MAX_C:g} °C 초과 시 REVIEW.", down="안정.",
             limits="—", validity="—", prec=("-", None, None)),
        dict(code="A_N_QUANT", method="A", tier="DQ", unit="개",
             name="정량 사용 회전 Step 수", formula="60초·Torque 유효 조건을 충족해 대표값이 계산된 Up/Down Step 수.",
             property="유동곡선 정보량. 제외 Step은 사유 열 참조.", function="—",
             up="정보량 증가.", down="지정점 공란·H-B 불확실성 증가.",
             limits="—", validity="—", prec=("-", None, None)),
        dict(code="B_TQ_BASE", method="B", tier="DQ", unit="%",
             name="3ITT 기준 구간 Torque", formula="Interval 1 후반 창 Torque 평균.",
             property="회복률 분모의 신호 크기.", function="—", up="여유 증가.", down="10% 미만이면 회복률 전부 무효.",
             limits="—", validity="10–95%.", prec=("-", None, None)),
        dict(code="B_REC_FLOOR", method="B", tier="DQ", unit="%",
             name="3ITT 회복률 검출하한",
             formula=f"Torque 하한({DEFAULT_TORQUE_MIN:g}%) / 기준 구간 Torque × 100.",
             property="이보다 낮은 회복률은 Torque 10% 미만이라 정량할 수 없음(CENSORED). 기준 Torque가 10% 근처면 회복 곡선 대부분이 측정 불가.",
             function="—", up="회복 곡선 관측 범위가 좁음 — 회복 구간 RPM 상향 필요.", down="관측 범위 넓음.",
             limits="측정 시스템 지표.", validity="—", prec=("-", None, None)),
        dict(code="B_REC_VALID", method="B", tier="DQ", unit="%",
             name="회복 구간 유효 데이터 비율", formula="Interval 3 점 중 Torque 10–95% 비율.",
             property="회복 곡선의 정량 가능 범위.", function="—", up="정량 범위 넓음.", down="저전단 조건 재설계 필요.",
             limits="—", validity="—", prec=("-", None, None)),
        dict(code="D_TEMP", method="D", tier="DQ", unit="°C",
             name="밀도 시험온도", formula="입력 파일의 temp_C.",
             property="밀도 측정 온도 조건.", function="—",
             up=f"{TEST_TEMP_TARGET_C:g}±{TEST_TEMP_TOL_C:g} °C 밖이면 REVIEW.", down="동일.",
             limits="—", validity="—", prec=("-", None, None)),
    ]
    # 06: metric_role / lot_acceptance / precision_evaluation / allow_notice / strict_value_required (지시서 §H)
    for m in d:
        role = ROLE_OVERRIDE.get(m["code"], ROLE_DEFAULT_BY_TIER[m["tier"]])
        m.setdefault("metric_role", role)
        m.setdefault("lot_acceptance", role in ("CORE", "PROCESS_SENSITIVE"))
        m.setdefault("precision_evaluation", m["prec"][0] != "-")
        m.setdefault("allow_notice", True)
        m.setdefault("strict_value_required", role in ("CORE", "PROCESS_SENSITIVE"))
        m.setdefault("zero_ref", False)
        m.setdefault("output_policy", "primary")
    return d


def _st(status, reason="", value_class=None):
    return {"status": status, "reason": reason or "", "value_class": value_class or "",
            "used_low": False, "used_high": False, "low_steps": [], "high_steps": [],
            "exc_steps": [], "exc_holds": []}


def _merge_vc_fields(items):
    """여러 기여 상태의 value_class·Step 목록을 합친다. 상태 문자열은 바꾸지 않는다."""
    items = [i for i in items if i]
    low_steps = _uniq_keep(s for i in items for s in (i.get("low_steps") or []))
    high_steps = _uniq_keep(s for i in items for s in (i.get("high_steps") or []))
    exc_steps = _uniq_keep(s for i in items for s in (i.get("exc_steps") or []))
    holds = []
    for i in items:
        for h in i.get("exc_holds") or []:
            if _finite_num(h) and not any(math.isclose(float(h), float(x), abs_tol=0.5) for x in holds):
                holds.append(float(h))
    used_low = any(bool(i.get("used_low")) for i in items) or bool(low_steps)
    used_high = any(bool(i.get("used_high")) for i in items) or bool(high_steps)
    vc = _pick_value_class([i.get("value_class") for i in items], used_low, used_high,
                           len(low_steps) or int(used_low), len(high_steps) or int(used_high))
    return {"value_class": vc, "used_low": used_low, "used_high": used_high,
            "low_steps": low_steps, "high_steps": high_steps, "exc_steps": exc_steps, "exc_holds": holds}


def _worst(*items):
    items = [i for i in items if i]
    if not items:
        return _st("NO VALUE", "근거 없음", VC_NO)
    status = max((i["status"] for i in items), key=lambda s: STATUS_RANK.get(s, 4))
    reasons = "; ".join(dict.fromkeys(i["reason"] for i in items if i.get("reason")))
    out = _st(status, reasons)
    out.update(_merge_vc_fields(items))
    if not out.get("value_class"):
        out["value_class"] = _default_value_class(status)
    return out


def _finite(x):
    try:
        return x is not None and not isinstance(x, str) and bool(np.isfinite(float(x)))
    except (TypeError, ValueError):
        return False


def _row_provenance(st, r):
    """Step의 구조화 value_class를 지표 상태에 붙인다. strict 채널은 used_low/high가 항상 False다."""
    st = dict(st)
    vc = str(r.get("value_class") or "")
    if not vc:
        vc = _default_value_class(st.get("status"))
    low = bool(r.get("used_low_torque", False))
    high = bool(r.get("used_high_torque", False))
    sid = f"S{int(r['Step'])}"
    hold_class = str(r.get("hold_class", ""))
    nonstrict = hold_class in ("SHORT_ACCEPTED", "LONG") and bool(r.get("analysis_hold_accepted", False))
    st["value_class"] = vc
    st["used_low"] = low
    st["used_high"] = high
    st["low_steps"] = [sid] if low else []
    st["high_steps"] = [sid] if high else []
    note_exc = vc == VC_EXC or (nonstrict and (vc in (VC_LOW, VC_HIGH) or bool(r.get("exception_accepted", False))))
    if note_exc and vc not in (VC_NO, VC_INVALID):
        st["exc_steps"] = [sid]
        st["exc_holds"] = [float(r["hold_s"])] if _finite_num(r.get("hold_s")) else []
    else:
        st["exc_steps"] = []
        st["exc_holds"] = []
    if vc in (VC_LOW, VC_HIGH) and nonstrict:
        bits = ["inclusive exception 포함"]
        if st["exc_holds"]:
            bits.append(f"accepted hold={st['exc_holds'][0]:g} s")
        if low:
            bits.append(f"Torque 하한 미만 Step {sid} 사용")
        if high:
            bits.append(f"Torque 상한 초과 Step {sid} 사용")
        extra = "; ".join(bits)
        if extra not in (st.get("reason") or ""):
            st["reason"] = (st["reason"] + "; " if st.get("reason") else "") + extra
    return st


def _row_status(r):
    label = f"S{int(r['Step'])} {r['dir']} {float(r['RPM']):g} rpm"
    ds = str(r.get("data_status", ""))
    if int(r["n_valid"]) <= 0:
        # 06: 전 point가 Torque 범위 밖인 Step은 NO VALUE가 아니라 CENSORED(검출범위 밖, 참고값은 StepAudit)
        if ds == "CENSORED":
            return _row_provenance(_st("CENSORED", f"{label}: {r['status']} — {r.get('label', '')}"), r)
        if ds == "INVALID":
            return _row_provenance(_st("INVALID", f"{label}: {r['status']}"), r)
        return _row_provenance(_st("NO VALUE", f"{label}: {r['status']}"), r)
    # 범위 밖 참고값을 계산에 넣어도 OK로 올리지 않는다. ALL_POINTS는 CENSORED를 유지한다.
    if ds == "CENSORED":
        return _row_provenance(_st("CENSORED", f"{label}: {r.get('status_reason') or r['status']} — {r.get('label', '')}"), r)
    if bool(r.get("exception_accepted", False)) or ds == "EXCEPTION":
        return _row_provenance(_st("EXCEPTION", f"{label}: {r['status']}"), r)
    if r["status"] != "OK":
        return _row_provenance(_st("REVIEW", f"{label}: {r['status']}"), r)
    return _row_provenance(_st("OK", label), r)


def _ref_step(reps, direction, rpm):
    sub = reps[(reps["RPM"] - float(rpm)).abs() <= RPM_MATCH_TOL]
    flag = "used_in_up_branch" if direction == "Up" else "used_in_down_branch"
    if flag in reps.columns:
        sub = sub[sub[flag].fillna(False).astype(bool)]
    else:
        sub = sub[sub["dir"] == direction]
    if sub.empty:
        return None, _st("NO VALUE", f"{direction} {float(rpm):g} rpm Step 없음")
    r = sub.sort_values("Step").iloc[0]
    return r, _row_status(r)


def _peak_in_step(df, step, tq_lo, tq_hi, search_s=PEAK_SEARCH_SECONDS, include_out=False):
    g = df[df["Step"] == int(step)].sort_values("t_elapsed")
    if g.empty:
        return np.nan, np.nan, np.nan, np.nan
    limit = min(float(search_s), float(g["t_elapsed"].max()) / 2.0)
    smooth = g["ShearStress_Pa"].rolling(3, center=True, min_periods=1).median()
    if include_out:
        mask = (g["t_elapsed"] <= limit) & g["Torque_pct"].notna() & g["ShearStress_Pa"].notna()
    else:
        mask = (g["t_elapsed"] <= limit) & g["Torque_pct"].between(tq_lo, tq_hi)
    if not mask.any():
        return np.nan, np.nan, np.nan, float(limit)
    idx = smooth[mask].idxmax()
    return float(smooth[idx]), float(g.loc[idx, "t_elapsed"]), float(g.loc[idx, "Torque_pct"]), float(limit)


def _tail_drift_pct(df, step, observed, tq_lo, tq_hi, include_out=False):
    g = df[df["Step"] == int(step)]
    w = g[g["t_elapsed"] > observed - METHOD_A_TAIL_SECONDS]
    if not include_out:
        w = w[w["Torque_pct"].between(tq_lo, tq_hi)]
    else:
        w = w[w["Torque_pct"].notna() & w["ShearStress_Pa"].notna()]
    if len(w) < 3 or w["t_elapsed"].nunique() < 2:
        return np.nan
    slope = np.polyfit(w["t_elapsed"].to_numpy(float), w["ShearStress_Pa"].to_numpy(float), 1)[0]
    mean = float(w["ShearStress_Pa"].mean())
    return float(slope * METHOD_A_TAIL_SECONDS / mean * 100.0) if mean else np.nan


class _QCBuffer:
    """지표 값·상태 버퍼(지시서 §J). NaN + NOTICE/CENSORED/INVALID/EXCEPTION/NO VALUE는 그대로 보존한다.
    값이 있어야만 성립하는 OK/REVIEW에 NaN이 오면 NO VALUE로 바꾼다."""
    _keep_without_value = ("NO VALUE", "CENSORED", "N/A", "NOTICE", "INVALID", "EXCEPTION", "REVIEW")

    def __init__(self):
        self.values, self.status = {}, {}

    def set(self, code, value, st):
        has = _finite(value)
        self.values[code] = float(value) if has else np.nan
        if has or st["status"] in self._keep_without_value:
            self.status[code] = {**st, "has_numeric_value": has}
        else:
            self.status[code] = {**_st("NO VALUE", st.get("reason") or "계산 불가", VC_NO), "has_numeric_value": False}


def _carry(new, old):
    """상태를 다시 만들 때 Step에서 온 value_class·Step 목록을 유지한다."""
    if not old:
        return new
    for key in ("value_class", "used_low", "used_high", "low_steps", "high_steps", "exc_steps", "exc_holds"):
        if not new.get(key) and old.get(key):
            new[key] = old.get(key)
    return new


def _stamp_from_steps(st, reps, step_ids):
    """기여 Step의 value_class를 파생 지표에 복사한다. OK를 범위 밖 참고값·예외로 덮어 승격되지 않게 한다."""
    base = _st(st.get("status", "NO VALUE"), st.get("reason", ""), st.get("value_class"))
    base.update(st)
    ids = []
    for s in step_ids or []:
        try:
            ids.append(int(s))
        except (TypeError, ValueError):
            continue
    if reps is None or not len(ids) or "Step" not in reps.columns:
        if not base.get("value_class"):
            base["value_class"] = _default_value_class(base.get("status"))
        return base
    sub = reps[reps["Step"].isin(ids)]
    if "n_valid" in sub.columns:
        used = sub[sub["n_valid"] > 0]
        if len(used):
            sub = used
    pieces = [_row_provenance(_st("OK", ""), rec) for rec in sub.to_dict("records")]
    if pieces:
        base.update(_merge_vc_fields(pieces))
    elif not base.get("value_class"):
        base["value_class"] = _default_value_class(base.get("status"))
    vc = base.get("value_class") or ""
    if vc in (VC_LOW, VC_HIGH, VC_BOTH) and base.get("status") == "OK":
        modes = set(str(x) for x in sub["analysis_torque_mode"]) if "analysis_torque_mode" in sub.columns else set()
        base["status"] = "CENSORED" if "ALL_POINTS" in modes else "REVIEW"
    elif vc == VC_EXC and base.get("status") == "OK":
        base["status"] = "EXCEPTION" if base.get("exc_steps") else "REVIEW"
    extra = []
    if base.get("used_low"):
        extra.append("Torque 하한 미만 Step " + ",".join(base.get("low_steps") or []) + " 사용")
    if base.get("used_high"):
        extra.append("Torque 상한 초과 Step " + ",".join(base.get("high_steps") or []) + " 사용")
    if base.get("exc_steps") and (vc == VC_EXC or vc in (VC_LOW, VC_HIGH)):
        holds = ", ".join(f"{float(h):g} s" for h in (base.get("exc_holds") or []))
        extra.append(("inclusive exception 포함; " if vc in (VC_LOW, VC_HIGH) else "")
                     + (f"accepted hold={holds}; " if holds else "")
                     + "Step " + ",".join(base.get("exc_steps") or []))
    add = "; ".join(x for x in extra if x and x not in (base.get("reason") or ""))
    if add:
        base["reason"] = (base["reason"] + "; " if base.get("reason") else "") + add
    return base


def hb_r2_status(fit, policy=None):
    """H-B R²는 항상 NOTICE(진단값). 4점·1 decade 미만은 사유에 명시(지시서 §I)."""
    policy = policy or active_policy()
    r2 = fit.get("R2", np.nan)
    npts = int(fit.get("n_points", 0) or 0)
    if "tau_y" not in fit or not _finite(r2):
        return _st("NO VALUE", f"{fit.get('label', '')}: {fit.get('status', 'H-B R² 계산 불가')} — 피팅 진단값, Lot 판정 제외")
    base = f"{fit.get('label')} R²={float(r2):.5f}, N={npts}"
    gmin, gmax = fit.get("gamma_min"), fit.get("gamma_max")
    ratio = gmax / gmin if gmin and gmax else np.nan
    notes = []
    if npts < policy.recommended_hb_points:
        notes.append(f"only {npts} fitting points")
    if _finite(ratio) and ratio < 10:
        notes.append(f"shear-rate span {ratio:.2f}× (<1 decade)")
    tail = "; ".join(notes)
    return _st("NOTICE", base + (f"; {tail}" if tail else "") + " — reference-only fit diagnostic, not an acceptance criterion")


def _rpm_close(a, b):
    try:
        return math.isclose(float(a), float(b), abs_tol=RPM_MATCH_TOL)
    except (TypeError, ValueError):
        return False


def _analyzable_steps(reps):
    """채널 뷰에서 Full H-B·동적 RPM 지표에 넣을 수 있는 moving Step."""
    if reps is None or len(reps) == 0:
        return reps
    m = reps[_moving_mask(reps)].copy()
    m = m[(m["gamma"] > 0) & (m["n_valid"] > 0)]
    m = m[m["tau_Pa"].map(_finite)]
    if "data_status" in m.columns:
        m = m[m["data_status"] != "INVALID"]
    return m


def _rpm_label(rpm):
    """정수 RPM은 5.0처럼 한 자리로, 0.35처럼 소수가 있으면 그대로 적는다."""
    v = float(rpm)
    if abs(v - round(v)) <= 1e-9:
        return f"{v:.1f}"
    return f"{v:g}"


def _legacy_ref(ref=None):
    """L/M/H_REF legacy 비교점. 예전 키 H가 있으면 H_REF로 읽는다. 동적 최고 RPM은 넣지 않는다."""
    src = dict(A_LEGACY_REF_RPM if ref is None else ref)
    if "H_REF" not in src and "H" in src:
        src["H_REF"] = src["H"]
    return src


def _higher_excluded_note(reps, selected_rpm, channel):
    """채널에서 빠진 더 높은 RPM이 있으면, 그 점을 측정 최고점이라고 남기고 선택 RPM을 최고점이라고 부르지 않는다."""
    if reps is None or len(reps) == 0 or not _finite(selected_rpm):
        return ""
    moving = reps[_moving_mask(reps)]
    if len(moving) == 0:
        return ""
    higher = moving[moving["RPM"] > float(selected_rpm) + RPM_MATCH_TOL]
    if higher.empty:
        return ""
    top = higher.sort_values(["RPM", "Step"]).iloc[-1]
    why = top.get("status_reason") or top.get("status") or "대표값 없음"
    return (f"측정 최고 RPM {_rpm_label(top['RPM'])} rpm Step S{int(top['Step'])}은 {channel}에서 제외 ({why}); "
            f"이 채널에서 사용 가능한 최고 RPM은 {_rpm_label(selected_rpm)} rpm이며 측정 최고점이 아님")


def _tag_selection(st, meta):
    if not meta:
        return st
    st["apply_rpm"] = f"{float(meta['selected_rpm']):g}"
    st["apply_gamma"] = f"{float(meta['selected_gamma']):g}"
    st["contrib_steps"] = f"S{int(meta['selected_step'])}"
    st["selected_rpm"] = float(meta["selected_rpm"])
    st["selected_gamma"] = float(meta["selected_gamma"])
    st["selected_step"] = int(meta["selected_step"])
    return st


def select_directional_extreme_step(reps, direction, extreme, channel=None):
    """direction의 분석 가능 Step 중 최저 또는 최고 RPM. 동일 RPM이 여러 개면 마지막 Step."""
    sub = _analyzable_steps(reps)
    flag = "used_in_up_branch" if direction == "Up" else "used_in_down_branch"
    if len(sub) and flag in sub.columns:
        sub = sub[sub[flag].fillna(False).astype(bool)]
    elif len(sub):
        sub = sub[sub["dir"] == direction]
    if sub is None or len(sub) == 0:
        return None, _st("NO VALUE", f"{direction} 분석 가능 Step 없음", VC_NO), {}
    target = float(sub["RPM"].min() if extreme == "min" else sub["RPM"].max())
    cand = sub[(sub["RPM"] - target).abs() <= RPM_MATCH_TOL].sort_values("Step")
    row = cand.iloc[-1]
    st = _row_status(row)
    if len(cand) > 1:
        st = _worst(st, _st("REVIEW", f"동일 RPM {float(row['RPM']):g} 후보 {len(cand)}개 — 마지막 Step S{int(row['Step'])} 사용"))
    meta = {"selected_step": int(row["Step"]), "selected_rpm": float(row["RPM"]),
            "selected_gamma": float(row["gamma"]), "n_candidates": int(len(cand)),
            "candidate_steps": [int(s) for s in cand["Step"]]}
    return row, _tag_selection(st, meta), meta


def select_low_shear_pair(reps, channel=None):
    """Up·Down 모두 분석 가능한 RPM 중 가장 낮은 공통 RPM의 마지막 Step 쌍."""
    scope = _analyzable_steps(reps)
    if scope is None or len(scope) == 0:
        return None
    if "used_in_up_branch" in scope.columns:
        up = scope[scope["used_in_up_branch"].fillna(False).astype(bool)]
        down = scope[scope["used_in_down_branch"].fillna(False).astype(bool)]
    else:
        up = scope[scope["dir"] == "Up"]
        down = scope[scope["dir"] == "Down"]
    if up.empty or down.empty:
        return None
    commons = []
    for rpm in up["RPM"]:
        if any(_rpm_close(rpm, d) for d in down["RPM"]) and not any(_rpm_close(rpm, c) for c in commons):
            commons.append(float(rpm))
    if not commons:
        return None
    rpm = min(commons)
    u = up[(up["RPM"] - rpm).abs() <= RPM_MATCH_TOL].sort_values("Step").iloc[-1]
    d = down[(down["RPM"] - rpm).abs() <= RPM_MATCH_TOL].sort_values("Step").iloc[-1]
    return {"rpm": float(u["RPM"]), "gamma": float(u["gamma"]), "up": u, "down": d}


def aggregate_value_class(contributing_rows):
    """기여 Step 행의 구조화 value_class를 가장 보수적인 값으로 합친다."""
    pieces = [_row_provenance(_st("OK", ""), rec) for rec in contributing_rows]
    return _merge_vc_fields(pieces) if pieces else {"value_class": VC_NO}


def _decorate_fit_status(st, fit, model):
    steps = fit.get("contributing_steps") or fit.get("steps") or []
    rpms = [float(x) for x in (fit.get("contributing_rpms") or [])]
    st["fit_model"] = model
    st["contrib_steps"] = ",".join(f"S{int(s)}" for s in steps)
    st["apply_rpm"] = ",".join(f"{x:g}" for x in rpms)
    if _finite(fit.get("gamma_min")) and _finite(fit.get("gamma_max")):
        st["apply_gamma"] = f"{float(fit['gamma_min']):g}–{float(fit['gamma_max']):g}"
    st["fit_n"] = int(fit.get("n_points") or 0)
    note = f"{model} H-B | Steps {st['contrib_steps']} | RPM {st['apply_rpm']} | n={st['fit_n']} | γ̇ {st.get('apply_gamma', '')}"
    if note not in (st.get("reason") or ""):
        st["reason"] = (st["reason"] + " | " if st.get("reason") else "") + note
    return st


def _mark_peak_torque(st, torque, tq_lo, tq_hi):
    if not _finite(torque):
        return st
    low = bool(st.get("used_low")) or float(torque) < float(tq_lo)
    high = bool(st.get("used_high")) or float(torque) > float(tq_hi)
    if float(torque) < float(tq_lo) or float(torque) > float(tq_hi):
        st["used_low"] = low
        st["used_high"] = high
        st["value_class"] = _pick_value_class([st.get("value_class")], used_low=low, used_high=high)
    return st


def _set_eta_from_row(q, code, row, st):
    if row is None or not _finite(row.get("eta_cP")) or st.get("status") == "NO VALUE":
        q.set(code, np.nan, st if st.get("value_class") else _st("NO VALUE", st.get("reason", "Step 없음"), VC_NO))
        return
    q.set(code, float(row["eta_cP"]), st)


def _set_peak_metric(q, df, code, row, st, tq_lo, tq_hi, include_out):
    if row is None:
        q.set(code, np.nan, st)
        return st
    peak, t_peak, ptq, limit = _peak_in_step(df, int(row["Step"]), tq_lo, tq_hi, include_out=include_out)
    note = f"peak search window {float(limit):g} s (PEAK_SEARCH_SECONDS={PEAK_SEARCH_SECONDS:g} s)" if _finite(limit) else "peak search 없음"
    if _finite(t_peak):
        note += f"; peak @ {float(t_peak):g} s"
    st["reason"] = (st.get("reason") + "; " if st.get("reason") else "") + note
    if include_out:
        st = _mark_peak_torque(st, ptq, tq_lo, tq_hi)
    if not _finite(peak):
        q.set(code, np.nan, _carry(_st("NO VALUE", (st.get("reason") or "") + " — 유효 peak 없음", VC_NO), st))
        return st
    q.set(code, float(peak), st)
    return st


def _export_defs(defs, records=None):
    """대표 출력용 정의. legacy는 빼고, 값이 하나도 없는 분할 고전단 점도는 행을 만들지 않는다."""
    records = list(records or [])

    def _seen(code):
        for rec in records:
            for key in ("values", "values_incl"):
                if _finite((rec.get(key) or {}).get(code, np.nan)):
                    return True
        return False

    out = []
    for d in defs:
        mode = d.get("output_policy", "primary")
        if mode == "legacy":
            continue
        if mode == "when_finite" and not _seen(d["code"]):
            continue
        out.append(d)
    return out


def _ordered_defs(defs):
    def key(d):
        if d["code"] in PRIMARY_METRIC_ORDER:
            return (0, PRIMARY_METRIC_ORDER.index(d["code"]), 0)
        return (1, TIER_ORDER.index(d["tier"]), "ADB".index(d["method"]))
    return sorted(defs, key=key)


def _step_label(value):
    if value in (None, ""):
        return None
    if isinstance(value, str):
        text = value.strip()
        return text or None
    try:
        return f"S{int(value)}"
    except (TypeError, ValueError):
        return None


def _rpm_cell(text):
    if text in (None, ""):
        return ""
    try:
        return float(text)
    except (TypeError, ValueError):
        return text


def _gamma_cell(text):
    return _rpm_cell(text)


def _lot_measured_max_text(recs, policy):
    if not recs:
        return "—"
    key = "inclusive_shear" if getattr(policy, "primary_channel", "strict") == "inclusive" else "strict_shear"
    parts, nums = [], []
    for rec in recs:
        lim = rec.get(key) or rec.get("shear_limits") or {}
        mx = lim.get("measured_max_rpm")
        label = _rpm_label(mx) if _finite(mx) else "—"
        parts.append(f"repeat {rec.get('repeat')}={label}")
        if _finite(mx):
            nums.append(round(float(mx), 5))
    uniq = list(dict.fromkeys(nums))
    if len(uniq) == 1:
        return f"{_rpm_label(uniq[0])} rpm"
    if uniq:
        return "mixed (" + ", ".join(parts) + ") — REVIEW"
    return "—"


def provenance_from_status(sts):
    """반복 시트의 공식 SSI·고전단 provenance. A_SSI alias 열은 만들지 않는다."""
    sts = sts or {}
    field = sts.get("A_SSI_FIELD_0P2_2") or {}
    qc = sts.get("A_SSI_QC_0P5_5") or {}
    tau = sts.get("A_TAU_H_UP") or {}
    eta = sts.get("A_ETA_HIGH") or {}

    def rpm(st, key):
        value = st.get(key)
        return float(value) if _finite(value) else None

    def one(prefix, st):
        return {
            f"{prefix}_value": rpm(st, "value") if _finite(st.get("value")) else None,
            f"{prefix}_num_rpm": rpm(st, "numerator_rpm"),
            f"{prefix}_den_rpm": rpm(st, "denominator_rpm"),
            f"{prefix}_num_step": _step_label(st.get("numerator_step")),
            f"{prefix}_den_step": _step_label(st.get("denominator_step")),
            f"{prefix}_value_class": st.get("value_class") or "",
        }

    out = {}
    out.update(one("A_SSI_FIELD_0P2_2", field))
    out.update(one("A_SSI_QC_0P5_5", qc))
    out.update({
        "A_TAU_H_UP_rpm": rpm(tau, "selected_high_rpm"),
        "A_TAU_H_UP_step": _step_label(tau.get("selected_high_step")),
        "A_ETA_HIGH_rpm": rpm(eta, "selected_high_rpm"),
        "A_ETA_HIGH_step": _step_label(eta.get("selected_high_step")),
    })
    return out


def select_reference_up_step(reps, target_rpm=FIXED_LOW_RPM, tolerance=None, channel=None):
    """Up 가지의 고정 RPM. 허용오차 안 Step이 없으면 다른 RPM으로 대체하지 않는다."""
    tolerance = RPM_MATCH_TOL if tolerance is None else tolerance
    sub = _analyzable_steps(reps)
    if sub is None or len(sub) == 0:
        return None
    if "used_in_up_branch" in sub.columns:
        sub = sub[sub["used_in_up_branch"].fillna(False).astype(bool)]
    else:
        sub = sub[sub["dir"].isin(["Up", "Turnaround"])]
    if len(sub) == 0:
        return None
    hit = sub[(sub["RPM"] - float(target_rpm)).abs() <= float(tolerance)].sort_values("Step")
    if hit.empty:
        return None
    return hit.iloc[0]


def rpm_token(value):
    """0.2 → 0P2, 1.0 → 1, 2.5 → 2P5. pair_id에 공통으로 쓴다."""
    number = float(value)
    if abs(number - round(number)) < 1e-8:
        return str(int(round(number)))
    text = f"{number:.6f}".rstrip("0").rstrip(".")
    return text.replace(".", "P").replace("-", "M")


def _row_eta(row):
    value = row.get("eta_cP", np.nan)
    return float(value) if _finite(value) else np.nan


def _row_gamma(row):
    value = row.get("gamma", np.nan)
    return float(value) if _finite(value) else np.nan


def _ssi_acceptance(vc):
    if vc in (VC_LOW, VC_HIGH, VC_BOTH, VC_ESTIMATED, VC_NO, VC_INVALID, VC_NA):
        return "EXCLUDED"
    if vc == VC_EXC:
        return "REVIEW"
    return "ELIGIBLE"


def _ssi_pair_class(low, high):
    classes = [low.get("value_class"), high.get("value_class")]
    used_low = bool(low.get("used_low_torque_in_inclusive")) or low.get("value_class") in (VC_LOW, VC_BOTH)
    used_high = bool(high.get("used_high_torque_in_inclusive")) or high.get("value_class") in (VC_HIGH, VC_BOTH)
    n_low = int(low.get("n_low_torque") or 0) + int(high.get("n_low_torque") or 0)
    n_high = int(low.get("n_high_torque") or 0) + int(high.get("n_high_torque") or 0)
    vc = _pick_value_class(classes, used_low=used_low, used_high=used_high, n_low=n_low, n_high=n_high)
    return vc, _ssi_acceptance(vc)


def _row_moving(row):
    if "moving" in row and row.get("moving") is not None and not (isinstance(row.get("moving"), float) and not _finite(row.get("moving"))):
        return bool(row.get("moving"))
    if "moving_flag" in row:
        return bool(row.get("moving_flag"))
    return row.get("dir") in ("Up", "Down", "Turnaround") or row.get("sequence_direction") in ("Up", "Down", "Turnaround")


def channel_row_is_eligible(row, channel):
    """strict/inclusive 대표값이 이 Step에서 계산에 쓰일 수 있는지. 채널마다 따로 복사하지 않는다."""
    if not _row_moving(row):
        return False
    if row.get("step_type") == "REST_STABILIZATION":
        return False
    if not _finite(row.get("eta_cP")) or not _finite(row.get("gamma")):
        return False
    if float(row["gamma"]) <= SHEAR_RATE_ZERO_TOL:
        return False
    if int(row.get("n_valid") or 0) <= 0:
        return False
    if channel == "strict":
        if "included_in_strict" in row and not bool(row.get("included_in_strict")):
            return False
    elif channel == "inclusive":
        if "included_in_inclusive" in row and not bool(row.get("included_in_inclusive")):
            return False
    else:
        raise ValueError(f"Unknown channel: {channel}")
    if row.get("value_class") in NON_NUMERIC_CLASSES:
        return False
    return True


def _empty_rpm_selection(target_rpm, channel, branch):
    return {
        "row": None, "found": False, "target_rpm": float(target_rpm), "measured_rpm": None,
        "selected_step": None, "selected_gamma": None, "selected_direction": None,
        "turnaround_shared": False, "data_status": "NO VALUE", "value_class": VC_NO,
        "acceptance_status": "EXCLUDED", "exclusion_reason": "분석 가능한 Step이 없음",
        "selection_rule": "exact nominal RPM", "candidate_steps": [],
        "measured_candidate_steps": [], "branch_candidate_steps": [], "eligible_candidate_steps": [],
        "duplicate_count": 0, "ambiguous": False, "channel": channel, "branch": branch,
    }


def _branch_mask(frame, branch, allow_turnaround):
    if branch == "any":
        return pd.Series(True, index=frame.index)
    if branch == "up":
        if "used_in_up_branch" in frame.columns:
            mask = frame["used_in_up_branch"].fillna(False).astype(bool)
        else:
            mask = frame["sequence_direction"].eq("Up") if "sequence_direction" in frame.columns else frame["dir"].eq("Up")
            if allow_turnaround and "sequence_direction" in frame.columns:
                mask = mask | frame["sequence_direction"].eq("Turnaround")
    elif branch == "down":
        if "used_in_down_branch" in frame.columns:
            mask = frame["used_in_down_branch"].fillna(False).astype(bool)
        else:
            mask = frame["sequence_direction"].eq("Down") if "sequence_direction" in frame.columns else frame["dir"].eq("Down")
    else:
        raise ValueError(f"Unknown branch: {branch}")
    if "turnaround_shared" in frame.columns and not allow_turnaround:
        mask = mask & ~frame["turnaround_shared"].fillna(False).astype(bool)
    return mask


def _prefer_unique_turnaround(eligible):
    if "turnaround_shared" not in eligible.columns:
        return eligible
    shared = eligible[eligible["turnaround_shared"].fillna(False).astype(bool)]
    if len(shared) == 1:
        return shared
    return eligible


def select_exact_rpm_step(reps, target_rpm, channel, branch="up", tolerance=RPM_MATCH_TOL,
                          allow_turnaround=False, duplicate_policy="require_unique_eligible"):
    """정확한 nominal RPM. branch는 강제 조건이고, 적격 후보가 둘이면 고르지 않는다."""
    del duplicate_policy
    frame = channel_view(reps, channel)
    result = _empty_rpm_selection(target_rpm, channel, branch)
    if frame.empty or "RPM" not in frame.columns:
        return result
    moving = frame["moving_flag"].fillna(False).astype(bool) if "moving_flag" in frame.columns else frame.apply(_row_moving, axis=1)
    if "step_type" in frame.columns:
        moving = moving & frame["step_type"].ne("REST_STABILIZATION")
    measured = frame[moving].copy()
    measured = measured[measured["RPM"].apply(
        lambda x: _finite(x) and abs(float(x)) > RPM_ZERO_TOL and abs(float(x) - float(target_rpm)) <= float(tolerance)
    )]
    result["measured_candidate_steps"] = sorted(int(x) for x in measured["Step"].tolist()) if len(measured) else []
    result["candidate_steps"] = list(result["measured_candidate_steps"])
    if measured.empty:
        result["exclusion_reason"] = f"{float(target_rpm):g} rpm moving Step 없음; 다른 RPM으로 대체하지 않음"
        return result
    result["measured_rpm"] = float(measured.iloc[0]["RPM"])
    branched = measured[_branch_mask(measured, branch, allow_turnaround)].copy()
    result["branch_candidate_steps"] = sorted(int(x) for x in branched["Step"].tolist()) if len(branched) else []
    if branched.empty:
        only_turn = False
        if "turnaround_shared" in measured.columns and not allow_turnaround:
            only_turn = bool(measured["turnaround_shared"].fillna(False).astype(bool).all())
        if only_turn:
            result["exclusion_reason"] = (
                f"{float(target_rpm):g} rpm shared turnaround는 allow_turnaround=False에서 사용하지 않음"
            )
        elif branch == "up":
            result["exclusion_reason"] = (
                f"{float(target_rpm):g} rpm 측정 Step은 있으나 Up 후보 없음; 다른 branch로 대체하지 않음"
            )
        else:
            result["exclusion_reason"] = (
                f"{float(target_rpm):g} rpm 측정 Step은 있으나 요구 branch={branch} 후보가 없음; 다른 branch로 대체하지 않음"
            )
        return result
    eligible = branched[branched.apply(lambda row: channel_row_is_eligible(row, channel), axis=1)].copy()
    result["eligible_candidate_steps"] = sorted(int(x) for x in eligible["Step"].tolist()) if len(eligible) else []
    if eligible.empty:
        steps = ", ".join(f"S{s}" for s in result["branch_candidate_steps"])
        reason = (
            f"{float(target_rpm):g} rpm {branch} Step은 있으나 {channel} 채널에서 부적격 ({steps}); "
            "다른 RPM 또는 다른 branch로 대체하지 않음"
        )
        if abs(float(target_rpm) - QC_SSI_HIGH_RPM) <= RPM_MATCH_TOL:
            reason += "; 4.0 rpm으로 대체하지 않음"
        result["exclusion_reason"] = reason
        result["duplicate_count"] = len(result["measured_candidate_steps"])
        return result
    if allow_turnaround:
        eligible = _prefer_unique_turnaround(eligible)
        result["eligible_candidate_steps"] = sorted(int(x) for x in eligible["Step"].tolist())
    if len(eligible) != 1:
        result["ambiguous"] = True
        result["duplicate_count"] = len(eligible)
        result["data_status"] = "REVIEW"
        result["exclusion_reason"] = f"{float(target_rpm):g} rpm 적격 후보가 {len(eligible)}개이므로 자동 선택하지 않음"
        result["eligible_candidate_steps"] = sorted(int(x) for x in eligible["Step"].tolist())
        return result
    row = eligible.iloc[0]
    dropped = [s for s in result["measured_candidate_steps"] if int(s) != int(row["Step"])]
    review = bool(dropped)
    result.update({
        "row": row, "found": True, "measured_rpm": float(row["RPM"]),
        "selected_step": int(row["Step"]), "selected_gamma": float(row["gamma"]),
        "selected_direction": row.get("sequence_direction"),
        "turnaround_shared": bool(row.get("turnaround_shared")),
        "data_status": "REVIEW" if review else (row.get("data_status") or "OK"),
        "value_class": row.get("value_class"),
        "acceptance_status": row.get("acceptance_status") or _ssi_acceptance(row.get("value_class")),
        "duplicate_count": len(result["measured_candidate_steps"]) if review else 1,
        "selection_rule": (
            f"exact measured {branch} {float(target_rpm):g} rpm"
            + ("; shared turnaround preferred" if allow_turnaround and bool(row.get("turnaround_shared")) else "")
        ),
        "exclusion_reason": (
            f"측정 후보 {result['measured_candidate_steps']} 중 유일한 적격 Step S{int(row['Step'])}; "
            f"제외 Step {dropped}"
        ) if review else "",
    })
    return result


def _decade_direction_token(branch):
    return {"up": "UP", "down": "DN"}.get(branch, str(branch).upper())


def discover_decade_rpm_pairs(reps, channel, branch="up", target_ratio=10.0, ratio_tolerance_pct=1.0,
                              include_turnaround_as_high=True):
    """실측 moving Step에서 high/low가 target_ratio인 쌍만 찾는다. 0 rpm은 후보가 아니다."""
    frame = channel_view(reps, channel)
    if frame.empty:
        return []
    rows = []
    for rec in frame.to_dict("records"):
        if not channel_row_is_eligible(rec, channel):
            continue
        direction = rec.get("sequence_direction")
        if branch == "up" and direction == "Down":
            continue
        if branch == "down" and direction == "Up":
            continue
        if direction == "Turnaround" and not include_turnaround_as_high:
            continue
        if branch == "up" and direction not in ("Up", "Turnaround"):
            continue
        if branch == "down" and direction not in ("Down", "Turnaround"):
            continue
        rows.append(rec)
    found = []
    token = _decade_direction_token(branch)
    for low in rows:
        for high in rows:
            if int(high["Step"]) == int(low["Step"]):
                continue
            low_rpm, high_rpm = float(low["RPM"]), float(high["RPM"])
            if high_rpm <= low_rpm or low_rpm <= RPM_ZERO_TOL:
                continue
            if branch == "up" and low.get("sequence_direction") != "Up":
                continue
            if branch == "down" and low.get("sequence_direction") != "Down":
                continue
            if high.get("sequence_direction") == "Turnaround" and not include_turnaround_as_high:
                continue
            if branch == "up" and high.get("sequence_direction") == "Down":
                continue
            if branch == "down" and high.get("sequence_direction") == "Up":
                continue
            actual = high_rpm / low_rpm
            error = abs(actual / float(target_ratio) - 1.0)
            if error > float(ratio_tolerance_pct) / 100.0:
                continue
            vc, acceptance = _ssi_pair_class(low, high)
            eta_low, eta_high = _row_eta(low), _row_eta(high)
            ratio = eta_low / eta_high if _finite(eta_high) and eta_high != 0 else np.nan
            official = ""
            if branch == "up" and abs(low_rpm - FIELD_SSI_LOW_RPM) <= RPM_MATCH_TOL and abs(high_rpm - FIELD_SSI_HIGH_RPM) <= RPM_MATCH_TOL:
                official = "A_SSI_FIELD_0P2_2"
            elif branch == "up" and abs(low_rpm - QC_SSI_LOW_RPM) <= RPM_MATCH_TOL and abs(high_rpm - QC_SSI_HIGH_RPM) <= RPM_MATCH_TOL:
                official = "A_SSI_QC_0P5_5"
            decision = lot_inclusion_decision(
                metric_code=official or f"SSI_DECADE_{token}_{rpm_token(low_rpm)}_{rpm_token(high_rpm)}",
                value=ratio, channel=channel, data_status="OK" if _finite(ratio) else "NO VALUE",
                value_class=vc, acceptance_status=acceptance, source_type="measured",
                pair_id=f"SSI_DECADE_{token}_{rpm_token(low_rpm)}_{rpm_token(high_rpm)}", policy=active_policy(),
            )
            found.append({
                "pair_id": f"SSI_DECADE_{token}_{rpm_token(low_rpm)}_{rpm_token(high_rpm)}",
                "target_ratio": float(target_ratio), "actual_ratio": actual, "ratio_error_pct": error * 100.0,
                "low_rpm": low_rpm, "high_rpm": high_rpm, "low_step": int(low["Step"]), "high_step": int(high["Step"]),
                "low_gamma": float(low["gamma"]), "high_gamma": float(high["gamma"]),
                "low_eta_cP": eta_low, "high_eta_cP": eta_high, "viscosity_ratio": ratio, "ssi": ratio,
                "direction": branch, "channel": channel,
                "low_value_class": low.get("value_class"), "high_value_class": high.get("value_class"),
                "value_class": vc, "data_status": "OK" if _finite(ratio) else "INVALID",
                "acceptance_status": acceptance,
                "selection_rule": DECADE_SSI_SELECTION_RULE, "reason": "", "source_type": "measured",
                "official_metric_link": official,
                "included_in_lot_stats": bool(decision["included"]) and not official,
                "lot_stat_exclusion_reason": "official metric link; dynamic 통계에 중복 집계하지 않음" if official else decision["reason"],
                "primary_pair_selected": False, "selection_reason": "",
            })
    return found


def _official_selection_rule(metric_code):
    if metric_code == "A_SSI_FIELD_0P2_2":
        return FIELD_SSI_SELECTION_RULE
    if metric_code == "A_SSI_QC_0P5_5":
        return QC_SSI_SELECTION_RULE
    return DECADE_SSI_SELECTION_RULE


def compute_exact_ssi_metric(reps, channel, metric_code, low_rpm, high_rpm, low_branch="up", high_branch="up",
                             allow_high_turnaround=False, policy=None):
    """지정 RPM 두 Step의 대표 점도비. 하나라도 없으면 NO VALUE이고 측정 후보는 남긴다."""
    policy = policy or active_policy()
    low_sel = select_exact_rpm_step(reps, low_rpm, channel, branch=low_branch, allow_turnaround=False)
    high_sel = select_exact_rpm_step(reps, high_rpm, channel, branch=high_branch, allow_turnaround=allow_high_turnaround)
    rule = _official_selection_rule(metric_code)
    base = {
        "metric_code": metric_code, "channel": channel, "source_type": "measured", "selection_rule": rule,
        "target_ratio": 10.0, "pair_id": f"SSI_DECADE_{_decade_direction_token(high_branch)}_{rpm_token(low_rpm)}_{rpm_token(high_rpm)}",
        "numerator_candidate_steps": list(low_sel.get("measured_candidate_steps") or []),
        "denominator_candidate_steps": list(high_sel.get("measured_candidate_steps") or []),
        "numerator_exclusion_reason": low_sel.get("exclusion_reason") or "",
        "denominator_exclusion_reason": high_sel.get("exclusion_reason") or "",
        "numerator_status": low_sel, "denominator_status": high_sel,
        "required_low_rpm": float(low_rpm), "required_high_rpm": float(high_rpm),
        "acceptance_status": "EXCLUDED", "included_in_lot_stats": False, "lot_stat_exclusion_reason": "",
        "ambiguous": bool(low_sel.get("ambiguous") or high_sel.get("ambiguous")),
    }
    if not low_sel.get("found") or not high_sel.get("found"):
        reason = "; ".join(x for x in (low_sel.get("exclusion_reason"), high_sel.get("exclusion_reason")) if x)
        status = "REVIEW" if base["ambiguous"] else "NO VALUE"
        base.update({
            "status": status, "data_status": status, "value_class": VC_NO, "value": np.nan,
            "reason": reason, "numerator_rpm": float(low_rpm), "denominator_rpm": float(high_rpm),
            "numerator_step": None, "denominator_step": None,
            "numerator_gamma": None, "denominator_gamma": None,
            "selected_low_rpm": None, "selected_high_rpm": None,
            "selected_low_step": None, "selected_high_step": None,
            "measured_high_rpm": high_sel.get("measured_rpm"),
            "measured_high_step": (high_sel.get("measured_candidate_steps") or [None])[0],
            "contributing_steps": "", "stat_value": np.nan,
        })
        base["lot_stat_exclusion_reason"] = reason or "NO VALUE"
        return base
    low, high = low_sel["row"], high_sel["row"]
    eta_low, eta_high = _row_eta(low), _row_eta(high)
    if not _finite(eta_high) or eta_high == 0 or not _finite(eta_low):
        base.update({"status": "INVALID", "data_status": "INVALID", "value_class": VC_INVALID, "value": np.nan,
                     "reason": "분모 점도가 0이거나 유한하지 않음", "numerator_step": int(low["Step"]),
                     "denominator_step": int(high["Step"]), "numerator_rpm": float(low["RPM"]),
                     "denominator_rpm": float(high["RPM"])})
        return base
    value = float(eta_low) / float(eta_high)
    vc, acceptance = _ssi_pair_class(low, high)
    data = "REVIEW" if low_sel.get("data_status") == "REVIEW" or high_sel.get("data_status") == "REVIEW" else (
        "CENSORED" if vc in TORQUE_REFERENCE_CLASSES else ("EXCEPTION" if vc == VC_EXC else "OK"))
    where = "shared turnaround" if bool(high.get("turnaround_shared")) else "Up endpoint"
    reason = f"eta_up({float(low['RPM']):g} rpm, S{int(low['Step'])}) / eta({float(high['RPM']):g} rpm, {where} S{int(high['Step'])})"
    extra = "; ".join(x for x in (low_sel.get("exclusion_reason"), high_sel.get("exclusion_reason")) if x)
    if extra:
        reason = reason + "; " + extra
    if vc in TORQUE_REFERENCE_CLASSES and VC_EXC in (low.get("value_class"), high.get("value_class")):
        reason += "; hold exception도 함께 기록"
    decision = lot_inclusion_decision(
        metric_code=metric_code, value=value, channel=channel, data_status=data, value_class=vc,
        acceptance_status=acceptance, source_type="measured", pair_id=base["pair_id"], policy=policy,
    )
    base.update({
        "status": data, "data_status": data, "value_class": vc, "value": value, "reason": reason,
        "numerator_rpm": float(low["RPM"]), "denominator_rpm": float(high["RPM"]),
        "numerator_step": int(low["Step"]), "denominator_step": int(high["Step"]),
        "numerator_gamma": float(low["gamma"]), "denominator_gamma": float(high["gamma"]),
        "numerator_eta_cP": eta_low, "denominator_eta_cP": eta_high,
        "numerator_value_class": low.get("value_class"), "denominator_value_class": high.get("value_class"),
        "numerator_hold_s": float(low.get("hold_observed_s") or np.nan),
        "denominator_hold_s": float(high.get("hold_observed_s") or np.nan),
        "actual_ratio": float(high["RPM"]) / float(low["RPM"]),
        "ratio_error_pct": abs((float(high["RPM"]) / float(low["RPM"])) / 10.0 - 1.0) * 100.0,
        "selected_low_rpm": float(low["RPM"]), "selected_high_rpm": float(high["RPM"]),
        "selected_low_step": int(low["Step"]), "selected_high_step": int(high["Step"]),
        "measured_high_rpm": float(high["RPM"]), "measured_high_step": int(high["Step"]),
        "acceptance_status": acceptance,
        "contributing_steps": f"S{int(low['Step'])},S{int(high['Step'])}",
        "used_low": vc in (VC_LOW, VC_BOTH), "used_high": vc in (VC_HIGH, VC_BOTH),
        "apply_rpm": f"{float(low['RPM']):g},{float(high['RPM']):g}",
        "apply_gamma": f"{float(low['gamma']):g},{float(high['gamma']):g}",
        "contrib_steps": f"S{int(low['Step'])},S{int(high['Step'])}",
        "included_in_lot_stats": bool(decision["included"]),
        "lot_stat_exclusion_reason": decision["reason"],
        "stat_value": decision["stat_value"],
        "turnaround_shared": bool(high.get("turnaround_shared")),
    })
    return base


def select_primary_decade_pair(pairs, mode="none", preferred_pairs=None, explicit_pair=None):
    """dynamic decade의 진단용 대표 pair. 공식 SSI 값은 바꾸지 않는다."""
    mode = mode or "none"
    groups = {}
    for pair in pairs or []:
        if pair.get("source_type") != "measured":
            continue
        groups.setdefault((pair.get("direction"), pair.get("channel")), []).append(pair)

    def _match(pair, lo, hi):
        return abs(float(pair["low_rpm"]) - float(lo)) <= RPM_MATCH_TOL and abs(float(pair["high_rpm"]) - float(hi)) <= RPM_MATCH_TOL

    def _one(items):
        if mode == "none" or not items:
            return None, "none" if mode == "none" else "실측 pair 없음"
        if mode == "preferred_first":
            for lo, hi in preferred_pairs or []:
                hits = [p for p in items if _match(p, lo, hi)]
                if hits:
                    return hits[0], "preferred_first"
            return None, "preferred pair 없음"
        if mode == "lowest_pair":
            return sorted(items, key=lambda p: (float(p["low_rpm"]), float(p["high_rpm"])))[0], "lowest_pair"
        if mode == "highest_pair":
            return sorted(items, key=lambda p: (-float(p["high_rpm"]), -float(p["low_rpm"])))[0], "highest_pair"
        if mode == "explicit":
            if not explicit_pair:
                return None, "explicit pair 없음"
            hits = [p for p in items if _match(p, explicit_pair[0], explicit_pair[1])]
            return (hits[0], "explicit") if hits else (None, "explicit pair 실측 없음")
        return None, f"unknown mode {mode}"

    selections = []
    for key, items in groups.items():
        chosen, why = _one(items)
        selections.append({"direction": key[0], "channel": key[1], "pair": chosen, "selection_reason": why})
    chosen_rows = [s for s in selections if s["pair"] is not None]
    info = {
        "primary_selection_mode": mode,
        "selections": selections,
        "selected_pair": chosen_rows[0]["pair"] if len(chosen_rows) == 1 else None,
        "selected_pair_id": chosen_rows[0]["pair"].get("pair_id") if len(chosen_rows) == 1 else None,
        "selected_low_rpm": chosen_rows[0]["pair"].get("low_rpm") if len(chosen_rows) == 1 else None,
        "selected_high_rpm": chosen_rows[0]["pair"].get("high_rpm") if len(chosen_rows) == 1 else None,
        "selection_reason": chosen_rows[0]["selection_reason"] if len(chosen_rows) == 1 else (
            "none" if mode == "none" else "direction/channel별 독립 선택"),
    }
    chosen_ids = {(s["pair"].get("pair_id"), s["direction"], s["channel"]) for s in chosen_rows}
    for pair in pairs or []:
        pair["primary_pair_selected"] = (pair.get("pair_id"), pair.get("direction"), pair.get("channel")) in chosen_ids
        pair["selection_reason"] = info["selection_reason"] if pair["primary_pair_selected"] else ""
    return info


def _apply_official_ssi(q, reps, channel):
    policy = active_policy()
    if policy.ssi_pool_different_pairs:
        raise ValueError("Pooling different SSI RPM pairs is not supported because pair identity must be preserved.")
    if policy.ssi_allow_interpolation or policy.ssi_allow_model_estimation:
        raise NotImplementedError("Estimated decade SSI is not implemented in HB-analysis-06-6. Use measured RPM pairs only.")
    if policy.ssi_primary_selection == "explicit" and not policy.ssi_explicit_pair:
        raise ValueError("--ssi-explicit-pair LOW:HIGH is required when primary-selection=explicit")
    for code, low_rpm, high_rpm, allow_turn in (
        ("A_SSI_FIELD_0P2_2", FIELD_SSI_LOW_RPM, FIELD_SSI_HIGH_RPM, False),
        ("A_SSI_QC_0P5_5", QC_SSI_LOW_RPM, QC_SSI_HIGH_RPM, True),
    ):
        st = compute_exact_ssi_metric(
            reps, channel, metric_code=code, low_rpm=low_rpm, high_rpm=high_rpm,
            low_branch="up", high_branch="up", allow_high_turnaround=allow_turn, policy=policy,
        )
        for end in ("numerator_status", "denominator_status"):
            sel = dict(st.get(end) or {})
            sel.pop("row", None)
            st[end] = sel
        q.set(code, st.get("value", np.nan), st)
    pairs = discover_decade_rpm_pairs(
        reps, channel, branch="up", target_ratio=float(policy.ssi_decade_ratio),
        ratio_tolerance_pct=float(policy.ssi_ratio_tolerance_pct), include_turnaround_as_high=True,
    )
    if policy.ssi_include_down:
        pairs.extend(discover_decade_rpm_pairs(
            reps, channel, branch="down", target_ratio=float(policy.ssi_decade_ratio),
            ratio_tolerance_pct=float(policy.ssi_ratio_tolerance_pct), include_turnaround_as_high=True,
        ))
    q.primary_decade = select_primary_decade_pair(
        pairs, mode=policy.ssi_primary_selection, preferred_pairs=policy.ssi_preferred_pairs,
        explicit_pair=policy.ssi_explicit_pair,
    )
    q.decade_pairs = pairs


def _rpm_close_pair(record, low, high):
    def _num(value):
        try:
            if value is None or value == "":
                return np.nan
            return float(value)
        except (TypeError, ValueError):
            return np.nan
    num, den = _num(record.get("numerator_rpm")), _num(record.get("denominator_rpm"))
    if not np.isfinite(num) or not np.isfinite(den):
        return False
    return abs(num - float(low)) <= RPM_MATCH_TOL and abs(den - float(high)) <= RPM_MATCH_TOL


def canonicalize_metric_record(record):
    """구형 A_SSI는 RPM·방향·source가 맞을 때만 공식 코드로 옮긴다."""
    record = dict(record or {})
    original = record.get("metric_code") or record.get("code") or ""
    base = {
        "original_metric_code": original,
        "canonical_metric_code": original,
        "migration_status": "unchanged",
        "migration_reason": "",
        "original_numerator_rpm": record.get("numerator_rpm"),
        "original_denominator_rpm": record.get("denominator_rpm"),
        "original_direction": record.get("direction"),
        "source_file": record.get("source_file") or record.get("file") or "",
        "included_in_lot_stats": True,
        "value_class": record.get("value_class"),
        "data_status": record.get("data_status") or record.get("status"),
        "acceptance_status": record.get("acceptance_status"),
    }
    if original != "A_SSI":
        return base
    direction = record.get("direction") or record.get("sequence_direction")
    source = record.get("source_type")
    high_shared = bool(record.get("turnaround_shared"))
    if source == "measured" and direction == "Up" and _rpm_close_pair(record, 0.2, 2.0):
        base.update({"canonical_metric_code": "A_SSI_FIELD_0P2_2", "migration_status": "mapped",
                     "migration_reason": "measured Up 0.2/2.0"})
        return base
    if source == "measured" and _rpm_close_pair(record, 0.5, 5.0) and (direction == "Up" or high_shared):
        base.update({"canonical_metric_code": "A_SSI_QC_0P5_5", "migration_status": "mapped",
                     "migration_reason": "measured Up 0.5/5.0"})
        return base
    base.update({
        "canonical_metric_code": "A_SSI_LEGACY_UNKNOWN",
        "migration_status": "legacy_unknown",
        "migration_reason": "A_SSI provenance가 공식 0.2/2 또는 0.5/5 measured Up pair와 일치하지 않음",
        "value_class": VC_LEGACY, "data_status": "REVIEW", "acceptance_status": "EXCLUDED",
        "included_in_lot_stats": False,
    })
    return base


def evaluate_stabilization(reps, policy):
    """0 rpm 안정화의 전처리 QC. 유동 값은 바꾸지 않는다."""
    frame = reps if isinstance(reps, pd.DataFrame) else pd.DataFrame(reps)
    rest = frame.iloc[0:0]
    if len(frame) and "step_type" in frame.columns:
        rest = frame[frame["step_type"] == "REST_STABILIZATION"]
    missing = "require_stabilization인데 안정화 Step이 없음" if policy.require_stabilization else "안정화 Step 없음"
    info = {"stabilization_present": bool(len(rest)), "stabilization_step": None,
            "stabilization_target_s": float(policy.stabilization_target_s), "stabilization_observed_s": np.nan,
            "stabilization_complete": False,
            "stabilization_temperature_target_C": float(policy.stabilization_temp_target_C),
            "stabilization_temperature_mean_C": np.nan, "stabilization_temperature_min_C": np.nan,
            "stabilization_temperature_max_C": np.nan, "stabilization_temperature_sd_C": np.nan,
            "stabilization_temperature_in_range_pct": np.nan,
            "stabilization_sampling_interval_median_s": np.nan, "stabilization_n_points": 0,
            "stabilization_status": "INVALID" if policy.require_stabilization else "NOTICE",
            "stabilization_reason": missing}
    if not len(rest):
        return info
    row = rest.iloc[0]
    observed = float(row.get("hold_observed_s") or 0)
    target = float(policy.stabilization_target_s)
    minimum = float(policy.stabilization_min_s) if policy.stabilization_min_s is not None else target
    info.update({"stabilization_step": int(row["Step"]), "stabilization_observed_s": observed,
                 "stabilization_complete": bool(observed + 1e-6 >= minimum),
                 "stabilization_n_points": int(row.get("n_window") or row.get("n_points") or 0),
                 "stabilization_sampling_interval_median_s": row.get("median_sampling_interval_s"),
                 "stabilization_temperature_mean_C": row.get("temp"),
                 "stabilization_temperature_min_C": row.get("temp_min"),
                 "stabilization_temperature_max_C": row.get("temp_max"),
                 "stabilization_temperature_sd_C": row.get("temp_sd")})
    status = "VALID"
    notes = ["시료 이송 중 발생한 구조 교란 회복을 위한 무전단 안정화; 유동·H-B·SSI 계산에서는 제외"]
    if not info["stabilization_complete"]:
        status = "REVIEW"
        notes.append(f"안정화 시간 부족: 목표 {target:g} s, 실제 {observed:g} s")
    tol = policy.stabilization_temp_tolerance_C
    if tol is not None and _finite(row.get("temp_min")) and _finite(row.get("temp_max")):
        target_t = float(policy.stabilization_temp_target_C)
        inside_lo = abs(float(row["temp_min"]) - target_t) <= float(tol)
        inside_hi = abs(float(row["temp_max"]) - target_t) <= float(tol)
        info["stabilization_temperature_in_range_pct"] = 100.0 if inside_lo and inside_hi else 0.0
        if not (inside_lo and inside_hi):
            status = "REVIEW"
            notes.append(f"온도 tolerance {float(tol):g} °C 초과")
    gap = row.get("sampling_gap_status")
    if gap == "REVIEW":
        status = "REVIEW"
    elif gap == "NOTICE" and status == "VALID":
        status = "NOTICE"
    info["stabilization_status"] = status
    info["stabilization_reason"] = "; ".join(notes)
    return info


def aggregate_decade_pair_stats(records, channel="inclusive"):
    """pair_id와 direction이 다른 SSI는 한 평균으로 합치지 않는다."""
    groups = {}
    key_name = "decade_pairs_incl" if channel == "inclusive" else "decade_pairs_strict"
    for rec in records:
        for pair in rec.get(key_name) or []:
            if pair.get("source_type") != "measured":
                continue
            if pair.get("official_metric_link"):
                continue
            key = (pair.get("pair_id"), round(float(pair.get("low_rpm")), 6), round(float(pair.get("high_rpm")), 6),
                   pair.get("direction"), channel, "measured")
            groups.setdefault(key, []).append(pair)
    out = []
    for key, rows in groups.items():
        vals = [float(r["ssi"]) for r in rows if _finite(r.get("ssi")) and r.get("included_in_lot_stats")]
        out.append({"pair_id": key[0], "numerator_rpm": key[1], "denominator_rpm": key[2],
                    "direction": key[3], "channel": key[4], "source_type": key[5], "n": len(vals),
                    "low_step": rows[0].get("low_step"), "high_step": rows[0].get("high_step"),
                    "value_class": rows[0].get("value_class"), "selection_rule": rows[0].get("selection_rule")})
    return out


def metric_alias(code):
    """코드 문자열만으로 A_SSI를 공식 지표에 넣지 않는다."""
    if code == "A_SSI":
        raise ValueError("A_SSI cannot be canonicalized from the code alone; use canonicalize_metric_record")
    return code



def select_actual_highest_step(reps, channel=None, direction_scope="up_or_turnaround", prefer_shared_turnaround=True):
    """측정 최고 RPM과 채널에서 유효한 최고 RPM을 구분한다.

    측정 최고점이 이 채널에서 빠지면 selected_high_rpm은 None이다.
    eligible_max_rpm만 감사값으로 남기고, 더 낮은 RPM을 대표 고전단으로 쓰지 않는다.
    """
    empty = {
        "row": None, "measured_max_rpm": np.nan, "measured_max_step": None, "measured_max_gamma": np.nan,
        "eligible_max_rpm": np.nan, "eligible_max_step": None, "eligible_max_gamma": np.nan,
        "representative_high_rpm": None, "selected_high_rpm": None,
        "selected_high_step": None, "selected_high_gamma": None, "selected_high_direction": None,
        "turnaround_shared": False, "higher_rpm_excluded": False, "higher_rpm_exclusion_reason": "",
        "selection_rule": HIGH_SELECTION_RULE, "representative_high_used": False,
    }
    if reps is None or len(reps) == 0:
        empty["higher_rpm_exclusion_reason"] = "분석 가능 Step 없음"
        return empty
    moving = reps[_moving_mask(reps)].copy()
    moving = moving[pd.to_numeric(moving["RPM"], errors="coerce").notna()]
    moving = moving[(moving["gamma"] > SHEAR_RATE_ZERO_TOL) & (moving["RPM"].abs() > RPM_ZERO_TOL)]
    if len(moving) == 0:
        empty["higher_rpm_exclusion_reason"] = "moving Step 없음"
        return empty
    top_measured = moving.sort_values(["RPM", "Step"]).iloc[-1]
    measured_max = float(top_measured["RPM"])
    empty["measured_max_rpm"] = measured_max
    empty["measured_max_step"] = int(top_measured["Step"])
    empty["measured_max_gamma"] = float(top_measured["gamma"])
    pool = _analyzable_steps(reps)
    if pool is None or len(pool) == 0:
        empty["higher_rpm_excluded"] = True
        empty["higher_rpm_exclusion_reason"] = _exclusion_sentence(top_measured, measured_max, np.nan, channel)
        return empty
    if direction_scope == "up_or_turnaround" and "used_in_up_branch" in pool.columns:
        pool = pool[pool["used_in_up_branch"].fillna(False).astype(bool)]
    elif direction_scope == "up_or_turnaround":
        pool = pool[pool["dir"].isin(["Up", "Turnaround"])]
    if len(pool) == 0:
        empty["higher_rpm_excluded"] = True
        empty["higher_rpm_exclusion_reason"] = "Up/turnaround eligible step 없음"
        return empty
    eligible_max = float(pool["RPM"].max())
    empty["eligible_max_rpm"] = eligible_max
    cand = pool[(pool["RPM"] - eligible_max).abs() <= RPM_MATCH_TOL].copy()
    cand["_shared"] = cand["turnaround_shared"].fillna(False).astype(bool) if "turnaround_shared" in cand.columns else False
    if not prefer_shared_turnaround:
        cand["_shared"] = False
    cand = cand.sort_values(["_shared", "Step"], ascending=[False, True])
    row = cand.iloc[0]
    higher = measured_max > float(row["RPM"]) + RPM_MATCH_TOL
    reason = ""
    if higher:
        top = moving[moving["RPM"] > float(row["RPM"]) + RPM_MATCH_TOL].sort_values(["RPM", "Step"]).iloc[-1]
        reason = _exclusion_sentence(top, measured_max, float(row["RPM"]), channel)
    return {
        "row": row,
        "measured_max_rpm": measured_max,
        "measured_max_step": int(top_measured["Step"]),
        "measured_max_gamma": float(top_measured["gamma"]),
        "eligible_max_rpm": eligible_max,
        "eligible_max_step": int(row["Step"]),
        "eligible_max_gamma": float(row["gamma"]),
        "representative_high_rpm": None if higher else float(row["RPM"]),
        "selected_high_rpm": None if higher else float(row["RPM"]),
        "selected_high_step": None if higher else int(row["Step"]),
        "selected_high_gamma": None if higher else float(row["gamma"]),
        "selected_high_direction": None if higher else str(row.get("dir") or ""),
        "turnaround_shared": bool(row.get("turnaround_shared")),
        "higher_rpm_excluded": higher,
        "higher_rpm_exclusion_reason": reason,
        "selection_rule": HIGH_SELECTION_RULE,
        "representative_high_used": not higher,
    }


def _exclusion_cause(top):
    """제외 원인은 hold 시간이 아니라 Step의 구조화 상태로 정한다."""
    ds = str(top.get("data_status") or "")
    mode = str(top.get("analysis_torque_mode") or "")
    n_low = float(top.get("n_low_torque") or 0)
    n_high = float(top.get("n_high_torque") or 0)
    if ds == "INVALID":
        return str(top.get("status_reason") or top.get("status") or "INVALID")
    if mode == "ALL_POINTS":
        if n_low > 0 and n_high > 0:
            return f"all points Torque <{DEFAULT_TORQUE_MIN:g}% and >{DEFAULT_TORQUE_MAX:g}%"
        if n_low > 0:
            return f"all points Torque <{DEFAULT_TORQUE_MIN:g}%"
        if n_high > 0:
            return f"all points Torque >{DEFAULT_TORQUE_MAX:g}%"
        return "all points outside Torque range"
    n_valid = top.get("n_valid")
    if _finite(n_valid) and int(n_valid) == 0:
        return "Torque 유효 point 없음 또는 대표값 부재"
    hold_ok = top.get("strict_hold_valid")
    if hold_ok is False or str(hold_ok) == "No":
        hold = top.get("hold_s")
        if _finite(hold):
            return f"hold {float(hold):g} s is not the strict {METHOD_A_STEP_SECONDS:g} s window"
        return "hold 시간 사유"
    return str(top.get("status_reason") or top.get("status") or "대표값 없음")


def _exclusion_sentence(top, measured_max, eligible_max, channel):
    who = channel or "channel"
    return (f"measured maximum {_rpm_label(measured_max)} rpm; "
            f"{who} representative high endpoint unavailable; "
            f"Step S{int(top['Step'])} excluded because {_exclusion_cause(top)}")


def _separate_down_peak(reps, high_row):
    """같은 최고 RPM이 다른 Step으로 Down에 있을 때만 그 행을 돌려준다."""
    if high_row is None:
        return None
    scope = _analyzable_steps(reps)
    if scope is None or len(scope) == 0:
        return None
    if "used_in_down_branch" in scope.columns:
        down = scope[scope["used_in_down_branch"].fillna(False).astype(bool)]
    else:
        down = scope[scope["dir"] == "Down"]
    same = down[(down["RPM"] - float(high_row["RPM"])).abs() <= RPM_MATCH_TOL]
    same = same[same["Step"] != int(high_row["Step"])]
    if same.empty or bool(high_row.get("turnaround_shared")):
        return None
    return same.sort_values("Step").iloc[0]


def _apply_high_fields(st, row, rule=None):
    st["selection_rule"] = rule or HIGH_SELECTION_RULE
    st["selected_high_rpm"] = float(row["RPM"])
    st["selected_high_step"] = int(row["Step"])
    st["selected_high_gamma"] = float(row["gamma"])
    st["selected_high_direction"] = str(row.get("dir") or "")
    st["turnaround_shared"] = bool(row.get("turnaround_shared"))
    st["contrib_steps"] = f"S{int(row['Step'])}"
    st["apply_rpm"] = _rpm_label(row["RPM"])
    st["apply_gamma"] = f"{float(row['gamma']):g}"
    st["value_class"] = st.get("value_class") or _default_value_class(st.get("status"))
    return st


def _hold_clause(row):
    hold = row.get("hold_s") if hasattr(row, "get") else None
    if not _finite(hold) or abs(float(hold) - METHOD_A_STEP_SECONDS) <= 0.1:
        return ""
    return f"; {_rpm_label(row['RPM'])} rpm Step S{int(row['Step'])}, {float(hold):g} s"


def _high_reason(row):
    shared = bool(row.get("turnaround_shared"))
    where = f"turnaround Step S{int(row['Step'])}" if shared else f"Step S{int(row['Step'])}"
    text = f"실제 최고 RPM {_rpm_label(row['RPM'])} rpm, {where}, shear rate {float(row['gamma']):g} s⁻¹"
    if shared:
        text += "; 단일 turnaround point로 측정됨; Up/Down Full H-B 양쪽 및 Common/Hysteresis 경계점으로 공유; 원시값과 Lot 반복값은 중복 집계하지 않음"
    return text + _hold_clause(row)


def _set_dynamic_rheology(q, df, reps, tq_lo, tq_hi, include_out, startup, channel="strict"):
    up_row, up_st, up_meta = select_directional_extreme_step(reps, "Up", "min")
    dn_row, dn_st, dn_meta = select_directional_extreme_step(reps, "Down", "min")
    if up_meta and dn_meta and not _rpm_close(up_meta["selected_rpm"], dn_meta["selected_rpm"]):
        note = f"Up 최저 {up_meta['selected_rpm']:g} rpm ≠ Down 최저 {dn_meta['selected_rpm']:g} rpm"
        up_st = _worst(up_st, _st("REVIEW", note))
        dn_st = _worst(dn_st, _st("REVIEW", note))
        _tag_selection(up_st, up_meta)
        _tag_selection(dn_st, dn_meta)
    if up_meta:
        up_st["selection_rule"] = "actual minimum measured eligible Up RPM; fixed low reference 0.5 rpm이 아님"
        up_st["selected_low_rpm"] = float(up_meta["selected_rpm"])
        up_st["selected_low_step"] = int(up_meta["selected_step"])
        up_st["selected_low_gamma"] = float(up_meta["selected_gamma"])
    _set_eta_from_row(q, "A_ETA_MIN_UP", up_row, up_st)
    _set_eta_from_row(q, "A_ETA_LOW_DN", dn_row, dn_st)
    chosen = select_actual_highest_step(reps, channel=channel, prefer_shared_turnaround=True)
    low_row = select_reference_up_step(reps, target_rpm=FIXED_LOW_RPM, tolerance=RPM_MATCH_TOL, channel=channel)
    q.shear_limits = {
        "channel": channel,
        "low_rpm": float(up_meta["selected_rpm"]) if up_meta else np.nan,
        "low_step": int(up_meta["selected_step"]) if up_meta else None,
        "measured_max_rpm": chosen["measured_max_rpm"],
        "measured_max_step": chosen.get("measured_max_step"),
        "measured_max_gamma": chosen.get("measured_max_gamma", np.nan),
        "eligible_max_rpm": chosen.get("eligible_max_rpm", np.nan),
        "eligible_max_step": chosen.get("eligible_max_step"),
        "eligible_max_gamma": chosen.get("eligible_max_gamma", np.nan),
        "representative_high_rpm": chosen.get("representative_high_rpm"),
        "selected_high_rpm": chosen["selected_high_rpm"],
        "selected_high_step": chosen["selected_high_step"],
        "selected_high_gamma": chosen["selected_high_gamma"],
        "selected_high_direction": chosen["selected_high_direction"],
        "turnaround_shared": chosen["turnaround_shared"],
        "higher_rpm_excluded": chosen["higher_rpm_excluded"],
        "higher_rpm_exclusion_reason": chosen["higher_rpm_exclusion_reason"],
        "selection_rule": chosen["selection_rule"],
        "representative_high_used": chosen["representative_high_used"],
        "high_rpm": float(chosen["selected_high_rpm"]) if chosen["representative_high_used"] else np.nan,
        "high_step": chosen["selected_high_step"] if chosen["representative_high_used"] else None,
        "fixed_low_rpm": float(low_row["RPM"]) if low_row is not None else np.nan,
        "fixed_low_step": int(low_row["Step"]) if low_row is not None else None,
    }
    high_row = chosen["row"] if chosen["representative_high_used"] else None
    if high_row is None:
        why = chosen["higher_rpm_exclusion_reason"] or "측정 최고 RPM을 대표 고전단으로 쓸 수 없음"
        nov = _st("NO VALUE", why + "; 더 낮은 RPM으로 대체하지 않음", VC_NO)
        nov["measured_max_rpm"] = chosen["measured_max_rpm"]
        nov["eligible_high_rpm"] = chosen["selected_high_rpm"]
        nov["higher_rpm_excluded"] = bool(chosen["higher_rpm_excluded"])
        nov["selection_rule"] = HIGH_SELECTION_RULE
        for code in ("A_ETA_HIGH", "A_TAU_H_UP"):
            q.set(code, np.nan, dict(nov))
    else:
        tau_st = _apply_high_fields(_row_status(high_row), high_row)
        where = (f"turnaround Step S{int(high_row['Step'])}" if bool(high_row.get("turnaround_shared"))
                 else f"Step S{int(high_row['Step'])}")
        tau_st["reason"] = _high_reason(high_row) + f"; { _rpm_label(high_row['RPM']) } rpm {where} 사용"
        q.set("A_TAU_H_UP", float(high_row["tau_Pa"]) if _finite(high_row.get("tau_Pa")) else np.nan, tau_st)
        eta_st = _apply_high_fields(_row_status(high_row), high_row)
        eta_st["reason"] = _high_reason(high_row)
        eta_st["direction_role"] = "Up+Down boundary" if bool(high_row.get("turnaround_shared")) else str(high_row.get("dir") or "")
        _set_eta_from_row(q, "A_ETA_HIGH", high_row, eta_st)
        down_peak = _separate_down_peak(reps, high_row)
        if down_peak is not None:
            up_only = _apply_high_fields(_row_status(high_row), high_row)
            up_only["reason"] = _high_reason(high_row) + "; 최고 RPM이 Up과 Down에 각각 측정됨"
            dn_only = _apply_high_fields(_row_status(down_peak), down_peak)
            _set_eta_from_row(q, "A_ETA_HIGH_UP", high_row, up_only)
            _set_eta_from_row(q, "A_ETA_HIGH_DN", down_peak, dn_only)
    _apply_official_ssi(q, reps, channel)
    _set_peak_metric(q, df, "A_STATIC_YIELD_APPROX_UP", up_row, dict(up_st), tq_lo, tq_hi, include_out)
    _set_peak_metric(q, df, "A_POST_SHEAR_STARTUP_STRESS", dn_row, dict(dn_st), tq_lo, tq_hi, include_out)
    pair = select_low_shear_pair(reps)
    if pair is None:
        nov = _st("NO VALUE", "공통 RPM 없음 — 보간하지 않음", VC_NO)
        q.set("A_ETA_RECOVERY_LOW", np.nan, nov)
        q.set("A_POST_SHEAR_STRESS_RETENTION", np.nan, nov)
    else:
        st = _worst(_row_status(pair["up"]), _row_status(pair["down"]))
        st["apply_rpm"] = f"{pair['rpm']:g}"
        st["apply_gamma"] = f"{pair['gamma']:g}"
        su, sd = int(pair["up"]["Step"]), int(pair["down"]["Step"])
        if su == sd:
            st["contrib_steps"] = f"S{su}"
            st["turnaround_shared"] = True
            pair_note = (f"paired Up/Down 동일 Step S{su} @ {pair['rpm']:g} rpm "
                         "(단일 turnaround 경계값, Up=Down, 원시값 1회)")
        else:
            st["contrib_steps"] = f"S{su},S{sd}"
            pair_note = f"paired Up S{su} / Down S{sd} @ {pair['rpm']:g} rpm"
        st["reason"] = (st.get("reason") + "; " if st.get("reason") else "") + pair_note
        eu, ed = float(pair["up"]["eta_cP"]), float(pair["down"]["eta_cP"])
        q.set("A_ETA_RECOVERY_LOW", ed / eu * 100 if eu else np.nan, st)
        pu, _, _, limu = _peak_in_step(df, int(pair["up"]["Step"]), tq_lo, tq_hi, include_out=include_out)
        pd_, _, ptq, limd = _peak_in_step(df, int(pair["down"]["Step"]), tq_lo, tq_hi, include_out=include_out)
        rst = dict(st)
        rst["reason"] = (rst.get("reason") or "") + f"; retention peaks search {limu:g}/{limd:g} s"
        if include_out:
            rst = _mark_peak_torque(rst, ptq, tq_lo, tq_hi)
        if _finite(pu) and pu and _finite(pd_):
            q.set("A_POST_SHEAR_STRESS_RETENTION", float(pd_) / float(pu) * 100, rst)
        else:
            q.set("A_POST_SHEAR_STRESS_RETENTION", np.nan, _carry(_st("NO VALUE", rst["reason"] + " — peak 없음", VC_NO), rst))
    if isinstance(startup, dict) and startup:
        q.set("A_STARTUP_PLATEAU_UP", startup.get("initial_plateau_tau", np.nan), _st("NOTICE", "plateau 보조값"))
        q.set("A_POST_SHEAR_PLATEAU", startup.get("post_plateau_tau", np.nan), _st("NOTICE", "plateau 보조값"))
        q.set("A_PLATEAU_RETENTION", startup.get("plateau_retention_pct", np.nan), _st("NOTICE", "plateau retention. peak retention과 다름"))
    else:
        for code in ("A_STARTUP_PLATEAU_UP", "A_POST_SHEAR_PLATEAU", "A_PLATEAU_RETENTION"):
            q.set(code, np.nan, _st("NO VALUE", "start-up profile 없음", VC_NO))


def compute_a_qc(df, holds, meta, reps, analysis, startup, checks, overall, tq_lo, tq_hi, ref=None, channel="strict",
                 policy=None):
    """Method A 반복 1건의 QC 지표. channel='strict'|'inclusive' 뷰(channel_view)를 받아 같은 규칙으로 계산한다."""
    ref = _legacy_ref(ref)
    policy = policy or active_policy()
    q = _QCBuffer()
    include_out = channel == "inclusive" and inclusive_out_of_range_enabled(policy)
    exc_col = "exception_accepted" in reps.columns
    exc_steps = set()
    if exc_col:
        exc_steps = set(int(s) for s in reps.loc[reps["exception_accepted"].astype(bool) & (reps["n_valid"] > 0), "Step"])
    route = policy.name if policy else "STANDARD"
    fixed_points = (("L", FIXED_LOW_RPM), ("M", FIXED_MID_RPM))
    rows = {(k, d): _ref_step(reps, d, rpm) for k, rpm in fixed_points for d in ("Up", "Down")}
    diag = _legacy_ref(ref)
    rows[("H_REF", "Up")] = _ref_step(reps, "Up", diag["H_REF"])
    rows[("H_REF", "Down")] = _ref_step(reps, "Down", diag["H_REF"])

    def val(key, direction, field):
        r, st = rows[(key, direction)]
        return (float(r[field]) if r is not None and st["status"] != "NO VALUE" else np.nan), st

    # 대표 저점·중간점은 CLI ref-rpm과 무관하게 0.5 rpm, 2.0 rpm이다.
    fixed_rpm = {"L": FIXED_LOW_RPM, "M": FIXED_MID_RPM}
    for key, which in (("L", "low"), ("M", "mid")):
        v, st = val(key, "Up", "tau_Pa")
        row = rows[(key, "Up")][0]
        out = dict(st)
        out["selection_rule"] = f"fixed {which} reference {_rpm_label(fixed_rpm[key])} rpm"
        out["apply_rpm"] = _rpm_label(fixed_rpm[key])
        if row is not None and st.get("status") != "NO VALUE":
            out["contrib_steps"] = f"S{int(row['Step'])}"
            out["apply_gamma"] = f"{float(row['gamma']):g}"
            if which == "low":
                out["selected_low_rpm"] = float(row["RPM"])
                out["selected_low_step"] = int(row["Step"])
                out["selected_low_gamma"] = float(row["gamma"])
        q.set(f"A_TAU_{key}_UP", v, out)
    v4, s4 = val("H_REF", "Up", "tau_Pa")
    e4, se4 = val("H_REF", "Up", "eta_cP")
    # legacy diagnostic의 저점만 CLI ref-rpm을 따른다. 대표 A_TAU_L_UP·A_SSI는 FIXED_LOW_RPM이다.
    r_leg, sL = _ref_step(reps, "Up", diag["L"])
    eL = float(r_leg["eta_cP"]) if r_leg is not None and sL["status"] != "NO VALUE" else np.nan
    q.legacy_4rpm = {
        "tau_up_4": v4,
        "eta_up_4": e4,
        "ssi_0p5_4": (eL / e4) if _finite(eL) and _finite(e4) and e4 else np.nan,
        "status_tau": s4.get("status"),
        "status_eta": se4.get("status"),
        "status_ssi": _worst(sL, se4).get("status"),
        "note": ("legacy diagnostic only; not a high-shear endpoint; not used for representative SSI; "
                 "not used for acceptance; not shown in primary Summary"),
    }
    # T2 구조 이력
    for key in ("L", "M"):
        tu, su = val(key, "Up", "tau_Pa")
        td, sd = val(key, "Down", "tau_Pa")
        st = _worst(su, sd)
        st["apply_rpm"] = _rpm_label(fixed_rpm[key])
        st["selection_rule"] = f"fixed {'low' if key == 'L' else 'mid'} reference {_rpm_label(fixed_rpm[key])} rpm"
        ru, rd = rows[(key, "Up")][0], rows[(key, "Down")][0]
        if _finite(tu) and not _finite(td) and rd is not None and str(rd["status"]) == "LOW TORQUE" \
                and _finite(ru["torque_valid"]) and ru["torque_valid"]:
            floor = tq_lo / float(ru["torque_valid"]) * 100
            st = _carry(_st("CENSORED", f"S{int(rd['Step'])} Down Torque<{tq_lo:g}% → D/U < 약 {floor:.1f}% (검출하한 미만)"), st)
        q.set(f"A_DU_{key}", td / tu * 100 if _finite(tu) and _finite(td) and tu else np.nan, st)
    common = [float(x) for x in analysis.get("common_gammas", [])]
    pair_steps = set(int(s) for s in (analysis.get("pair_steps") or []))
    moving = reps[_moving_mask(reps)]
    if pair_steps:
        common_steps = set(int(s) for s in moving.loc[moving["Step"].isin(pair_steps) & (moving["n_valid"] > 0), "Step"])
    else:
        common_steps = set(int(s) for s in moving.loc[moving["gamma"].round(8).isin([round(x, 8) for x in common]) & (moving["n_valid"] > 0), "Step"])
    n_common = len(analysis.get("points") or [])
    if n_common < 2:
        hy = _st("NO VALUE", f"공통 전단률 {n_common}쌍")
    else:
        rng = f"공통 γ̇ {n_common}쌍 {min(common):g}–{max(common):g} s⁻¹"
        hy = _st("OK", rng) if n_common >= METHOD_A_RECOMMENDED_HB_POINTS else _st("REVIEW", rng + " (5쌍 미만)")
        if common_steps & exc_steps:
            hy = _st("EXCEPTION", rng + f"; 예외 Step 포함 S{','.join(map(str, sorted(common_steps & exc_steps)))}")
    hy = _stamp_from_steps(hy, reps, common_steps)
    unpaired = _analyzable_steps(reps)
    if len(unpaired):
        unpaired = unpaired[~unpaired["Step"].isin(common_steps)]
    unpaired_txt = ", ".join(f"S{int(r.Step)} {float(r.RPM):g} rpm" for r in unpaired.itertuples()) if len(unpaired) else ""
    if unpaired_txt:
        hy["reason"] = (hy.get("reason") or "") + f" | Full H-B only; no paired point for hysteresis: {unpaired_txt}"
    turn = analysis.get("turnaround") or {}
    if turn.get("shared") and _finite(analysis.get("hysteresis_rpm_max")):
        rpms = sorted({round(float(p["rpm_up"]), 5) for p in analysis.get("points", []) if _finite(p.get("rpm_up"))})
        prev = rpms[-2] if len(rpms) >= 2 else analysis.get("hysteresis_rpm_min")
        hy["reason"] = (hy.get("reason") or "") + (
            f" | 공통 적분 범위 {float(analysis['hysteresis_rpm_min']):g}–{float(analysis['hysteresis_rpm_max']):g} rpm; "
            f"최고 {float(turn['rpm']):g} rpm turnaround에서 Up=Down 동일 경계값 사용; "
            f"{float(prev):g}–{float(turn['rpm']):g} rpm 구간 포함")
        hy["turnaround_shared"] = True
        hy["turnaround_step"] = int(turn["step"])
    area_st = hy if n_common >= 2 else _st("NO VALUE", f"공통 전단률 {n_common}쌍(<2) — 면적 적분 불가", VC_NO)
    if n_common >= 2:
        area_st["apply_gamma"] = f"{min(common):g}–{max(common):g}"
        area_st["contrib_steps"] = ",".join(f"S{s}" for s in sorted(common_steps))
        paired = _analyzable_steps(reps)
        if len(paired):
            paired = paired[paired["Step"].isin(common_steps)]
            area_st["apply_rpm"] = ",".join(f"{float(x):g}" for x in sorted(set(round(float(v), 5) for v in paired["RPM"])))
    q.set("A_HYS_REL", analysis.get("relative_area_pct") if n_common >= 2 else np.nan, area_st if n_common >= 2 else hy)
    q.set("A_HYS_AREA_SIGNED", analysis.get("signed_area") if n_common >= 2 else np.nan, area_st)
    q.set("A_HYS_AREA_ABS", analysis.get("absolute_area") if n_common >= 2 else np.nan, area_st)
    pts = analysis.get("points", [])
    sum_up = sum(p["tau_up"] for p in pts)
    hdi_st = hy if n_common >= METHOD_A_MIN_HB_POINTS else _st("NO VALUE", f"공통 전단률 {n_common}쌍(<4)", VC_NO)
    q.set("A_HDI", sum(p["delta_abs"] for p in pts) / sum_up * 100 if pts and sum_up else np.nan, hdi_st)
    if "used_in_up_branch" in reps.columns:
        up_rows = reps[reps["used_in_up_branch"].fillna(False).astype(bool)]
    else:
        up_rows = reps[reps["dir"] == "Up"]
    if len(up_rows):
        pr = up_rows.loc[up_rows["RPM"].idxmax()]
        pst = _row_status(pr)
        brk = np.nan
        if pst["status"] != "NO VALUE":
            peak, t_peak, _ptq, _plim = _peak_in_step(df, int(pr["Step"]), tq_lo, tq_hi, include_out=include_out)
            if _finite(peak) and peak:
                tail = float(pr["tau_Pa"])
                brk = (peak - tail) / peak * 100
                detail = f"; peak {peak:.1f} Pa @ {t_peak:g} s, 대표 {tail:.1f} Pa"
                if brk < 0:
                    # 대표창 응력이 초기 peak보다 높음: 붕괴가 아니라 peak 미검출 또는 전단 중 구조 형성.
                    # 값은 그대로 남기고(대체하지 않음) 해당 반복만 REVIEW로 표시한다.
                    pst = _worst(pst, _st("REVIEW", f"음수 {brk:.2f}% (대표창 > 초기 peak) — 원시곡선 확인"))
                pst = _carry(_st(pst["status"], pst["reason"] + detail), pst)
        q.set("A_BRK_PEAK", brk, pst)
    # T3 Start-up (지시서 §I: 미산출은 NOTICE, Run/Lot 유효성에 전파하지 않음)
    # Start-up은 rest 직후 첫 회전 Step의 실제 전이 응답으로 정의한다.
    # 따라서 그 RPM이 L 기준값(예: 0.5 rpm)보다 낮거나 높아도 Step 자체가 유효하면 계산한다.
    first = moving.sort_values("Step").iloc[0] if len(moving) else None
    ovs, ost, ovs_present = np.nan, _st("NOTICE", "Start-up Step 없음"), "NOT_EVALUABLE"
    if first is None:
        ost = _st("NOTICE", "No quantifiable start-up overshoot under current method — 첫 회전 Step 없음; 진단값, Lot 판정 제외")
    else:
        startup_row = first
        startup_status = _row_status(startup_row)
        startup_step = int(startup_row["Step"])
        startup_rpm = float(startup_row["RPM"])
        if startup_status["status"] == "NO VALUE" or not _finite(startup_row["tau_Pa"]) or not startup_row["tau_Pa"]:
            ovs_present = "NO"
            ost = _st("NOTICE", "No quantifiable start-up overshoot under current method — "
                                f"S{startup_step} {startup_rpm:g} rpm 대표 응력 없음/무효({startup_status['reason']})")
        else:
            peak, t_peak, _ptq, _plim = _peak_in_step(df, startup_step, tq_lo, tq_hi, include_out=include_out)
            if _finite(peak):
                ovs = (peak / float(startup_row["tau_Pa"]) - 1) * 100
                ovs_present = "YES" if ovs > 0 else "NO"
                rpm_note = (f"기준 {FIXED_LOW_RPM:g} rpm과 다름; "
                            if not math.isclose(startup_rpm, float(FIXED_LOW_RPM), abs_tol=RPM_MATCH_TOL) else "")
                ost = _st("NOTICE", f"Start-up overshoot stress calculated from first moving Step "
                                    f"S{startup_step} {startup_rpm:g} rpm; {rpm_note}peak @ {t_peak:g} s; "
                                    f"{startup_status['reason']} — 진단값, Lot 판정 제외")
            else:
                ovs_present = "NO"
                ost = _st("NOTICE", "No quantifiable start-up overshoot under current method — "
                                    f"S{startup_step} {startup_rpm:g} rpm peak 미검출(Torque 유효점 없음)")
        ost = _stamp_from_steps(ost, reps, [startup_step])
        if ost.get("status") != "NO VALUE":
            ost["status"] = "NOTICE"
    q.set("A_OVS_L", ovs, ost)
    # T3 H-B
    fits = analysis.get("fits", {})

    def fit_status(key, model):
        f = fits.get(key, {})
        steps = f.get("contributing_steps") or f.get("steps") or []
        if "tau_y" not in f:
            st = _stamp_from_steps(_st("NO VALUE", f"{f.get('label', key)}: {f.get('status', '피팅 없음')}", VC_NO), reps, steps)
            return _decorate_fit_status(st, f, model)
        base = f"{model} H-B {f.get('label')} N={f.get('n_points')}"
        if set(int(s) for s in steps) & exc_steps:
            st = _st("EXCEPTION", base + "; 예외 Step 포함")
        else:
            st = _st("OK", base) if f.get("status") == "OK" else _st("REVIEW", base + f"; {f.get('status')}")
        st = _stamp_from_steps(st, reps, steps)
        if f.get("n_low_torque_steps") and f.get("n_high_torque_steps"):
            extra = "torque 하한 밖 참고값 및 torque 상한 밖 참고값 포함"
            if extra not in st.get("reason", ""):
                st["reason"] = (st.get("reason") + "; " if st.get("reason") else "") + extra
        turn = analysis.get("turnaround") or {}
        if turn.get("shared") and int(turn.get("step") or -1) in set(int(s) for s in steps):
            share = (f"최고 RPM {float(turn['rpm']):g} rpm Step S{int(turn['step'])}는 단일 turnaround point로 측정됨; "
                     "Up/Down Full H-B 양쪽 및 Common/Hysteresis 경계점으로 공유; "
                     "원시값과 Lot 반복값은 중복 집계하지 않음")
            if share not in st.get("reason", ""):
                st["reason"] = (st.get("reason") + "; " if st.get("reason") else "") + share
            if key == "down_full":
                rpms = ",".join(f"{float(x):g}" for x in (f.get("contributing_rpms") or []))
                down_note = (f"Down Full H-B는 turnaround {float(turn['rpm']):g} rpm Step S{int(turn['step'])}를 "
                             f"시작점으로 포함; contributing RPM = {rpms}")
                if down_note not in st.get("reason", ""):
                    st["reason"] = st.get("reason", "") + "; " + down_note
        return _decorate_fit_status(st, f, model)

    for code, key, field in (("A_HB_TAUY_UP", "up_full", "tau_y"), ("A_HB_K_UP", "up_full", "K"),
                             ("A_HB_N_UP", "up_full", "n"),
                             ("A_HB_TAUY_DN", "down_full", "tau_y"), ("A_HB_K_DN", "down_full", "K"),
                             ("A_HB_N_DN", "down_full", "n")):
        q.set(code, fits.get(key, {}).get(field), fit_status(key, "Full"))
    for code, key, field in (("A_HB_TAUY_UP_COMMON", "up_common", "tau_y"), ("A_HB_K_UP_COMMON", "up_common", "K"),
                             ("A_HB_N_UP_COMMON", "up_common", "n"),
                             ("A_HB_TAUY_DN_COMMON", "down_common", "tau_y"), ("A_HB_K_DN_COMMON", "down_common", "K"),
                             ("A_HB_N_DN_COMMON", "down_common", "n")):
        q.set(code, fits.get(key, {}).get(field), fit_status(key, "Common"))
    for code, key, model in (("A_HB_R2_UP", "up_full", "Full"), ("A_HB_R2_DN", "down_full", "Full"),
                             ("A_HB_R2_UP_COMMON", "up_common", "Common"), ("A_HB_R2_DN_COMMON", "down_common", "Common")):
        f = fits.get(key, {})
        st = hb_r2_status(f, policy)
        if set(int(x) for x in (f.get("steps") or [])) & exc_steps and st["status"] != "NO VALUE":
            st = _carry(_st("NOTICE", st["reason"] + "; 예외 Step 포함"), st)
        st = _stamp_from_steps(st, reps, f.get("contributing_steps") or f.get("steps") or [])
        if st.get("status") != "NO VALUE":
            st["status"] = "NOTICE"
        q.set(code, f.get("R2"), _decorate_fit_status(st, f, model))
    uf, dfull = fits.get("up_full", {}), fits.get("down_full", {})
    ratio_st = _worst(fit_status("up_full", "Full"), fit_status("down_full", "Full"))
    if _finite(uf.get("gamma_min")) and _finite(dfull.get("gamma_min")) and (
            abs(float(uf["gamma_min"]) - float(dfull["gamma_min"])) > 1e-6
            or abs(float(uf.get("gamma_max", 0)) - float(dfull.get("gamma_max", 0))) > 1e-6):
        ratio_st = _worst(ratio_st, _st("REVIEW",
            "Up/Down Full fitting range differs; "
            f"Up γ̇ {float(uf['gamma_min']):.4g}–{float(uf['gamma_max']):.4g} N={uf.get('n_points')} RPM {uf.get('contributing_rpms')}; "
            f"Down γ̇ {float(dfull['gamma_min']):.4g}–{float(dfull['gamma_max']):.4g} N={dfull.get('n_points')} RPM {dfull.get('contributing_rpms')}"))
    ratio = analysis.get("comparison_full", {}).get("tau_y_ratio_pct", np.nan)
    if not _finite(uf.get("tau_y")) or not uf.get("tau_y"):
        ratio = np.nan
        ratio_st = _st("NO VALUE", "Up Full τy가 0이거나 없음", VC_NO)
    ratio_st["fit_model"] = "Full"
    ratio_st["apply_rpm"] = f"Up {uf.get('contributing_rpms')} / Down {dfull.get('contributing_rpms')}"
    ratio_st["apply_gamma"] = (
        f"Up {uf.get('gamma_min')}–{uf.get('gamma_max')} / Down {dfull.get('gamma_min')}–{dfull.get('gamma_max')}")
    ratio_st["contrib_steps"] = (
        "Up " + ",".join(f"S{int(s)}" for s in (uf.get("contributing_steps") or []))
        + " / Down " + ",".join(f"S{int(s)}" for s in (dfull.get("contributing_steps") or [])))
    q.set("A_HB_TAUY_RATIO", ratio, ratio_st)
    q.set("A_HB_K_RATIO", np.nan if not _finite(ratio) else analysis.get("comparison_full", {}).get("K_ratio_pct", np.nan), ratio_st)
    q.set("A_HB_N_DIFF", np.nan if not _finite(uf.get("n")) or not _finite(dfull.get("n")) else dfull.get("n") - uf.get("n"), ratio_st)
    q.set("A_HB_TAUY_RATIO_COMMON", analysis.get("comparison", {}).get("tau_y_ratio_pct", np.nan),
          _worst(fit_status("up_common", "Common"), fit_status("down_common", "Common")))
    uf = fits.get("up_full", {})
    grange = uf["gamma_max"] / uf["gamma_min"] if uf.get("gamma_min") else np.nan
    if _finite(grange) and grange >= 10:
        gst = _st("NOTICE", f"{grange:.2f}× ≥10 — 외삽 범위 양호(진단값)")
    elif _finite(grange):
        gst = _st("NOTICE", f"{grange:.2f}× <10 (1 decade 미만) — τy·K·n 외삽 불확실, HB_RANGE_LT_1DECADE_NOTICE")
    else:
        gst = _st("NO VALUE", "피팅 없음", VC_NO)
    gst = _decorate_fit_status(_stamp_from_steps(gst, reps, uf.get("contributing_steps") or uf.get("steps") or []), uf, "Full")
    if gst.get("status") != "NO VALUE":
        gst["status"] = "NOTICE"
    q.set("A_HB_GRANGE", grange, gst)
    _set_dynamic_rheology(q, df, reps, tq_lo, tq_hi, include_out, startup, channel=channel)
    # DQ
    tqL, stL = val("L", "Up", "torque_valid")
    q.set("A_TQ_L", tqL, stL)
    q.set("A_RES_L", TORQUE_RESOLUTION_PCT / tqL * 100 if _finite(tqL) and tqL else np.nan, stL)
    q.set("A_DUL_FLOOR", tq_lo / tqL * 100 if _finite(tqL) and tqL else np.nan, stL)
    step_rows, drifts = [], []
    for r in moving.sort_values("Step").itertuples():
        used = int(r.n_valid) > 0
        drift = _tail_drift_pct(df, r.Step, float(r.observed_s), tq_lo, tq_hi, include_out=include_out) if used else np.nan
        if _finite(drift):
            drifts.append((abs(drift), int(r.Step)))
        step_rows.append({
            "Step": int(r.Step), "dir": r.dir, "RPM": float(r.RPM), "gamma": float(r.gamma),
            "hold_s": float(r.hold_s), "used": used,
            "exception": bool(getattr(r, "exception_accepted", False)),
            "tau_Pa": float(r.tau_Pa) if used else np.nan, "eta_cP": float(r.eta_cP) if used else np.nan,
            "torque_valid": float(r.torque_valid) if used else np.nan, "torque_window": float(r.torque_window),
            "temp": float(r.temp), "tail_drift_pct": drift,
            "resolution_pct": TORQUE_RESOLUTION_PCT / float(r.torque_valid) * 100 if used and r.torque_valid else np.nan,
            "status": str(r.status),
            "label": str(getattr(r, "label", "")), "data_status": str(getattr(r, "data_status", "")),
            "channel": channel,
        })
    if drifts:
        mx, s = max(drifts)
        dst = _st("OK", f"최대 S{s}") if mx <= EQ_DRIFT_REVIEW_PCT else \
            _st("REVIEW", f"S{s} {mx:.2f}% > {EQ_DRIFT_REVIEW_PCT:g}% — 준평형 미확인")
        q.set("A_EQ_DRIFT_MAX", mx, _stamp_from_steps(dst, reps, [s]))
    else:
        q.set("A_EQ_DRIFT_MAX", np.nan, _st("NO VALUE", "정량 Step 없음"))
    quant = moving[moving["n_valid"] > 0]
    tc = str(meta.get("Temperature Control", "")).strip()
    tc_note = f"; Temperature Control={tc}" if tc else ""
    if len(quant):
        tmean = float(quant["temp"].mean())
        tspan = float(quant["temp"].max() - quant["temp"].min())
        q.set("A_TEMP_MEAN", tmean, (_st("OK", "25.0±0.5 °C" + tc_note) if abs(tmean - TEST_TEMP_TARGET_C) <= TEST_TEMP_TOL_C
                                     else _st("REVIEW", f"{tmean:.2f} °C — 목표 밖" + tc_note)))
        q.set("A_TEMP_SPAN", tspan, _st("OK", "") if tspan <= TEMP_SPAN_MAX_C else _st("REVIEW", f"{tspan:.2f} °C > {TEMP_SPAN_MAX_C:g}"))
    else:
        q.set("A_TEMP_MEAN", np.nan, _st("NO VALUE", "정량 Step 없음"))
        q.set("A_TEMP_SPAN", np.nan, _st("NO VALUE", "정량 Step 없음"))
    excluded = moving[moving["n_valid"] <= 0]
    q.set("A_N_QUANT", len(quant), _st("OK", f"{len(quant)}/{len(moving)}") if excluded.empty else
          _st("REVIEW", f"{len(quant)}/{len(moving)}; 제외 " + ", ".join(f"S{int(x.Step)} {x.status}" for x in excluded.itertuples())))
    flags = [{"code": c["check"], "status": c["status"], "reason": f"값={c['value']}; 기준={c['criterion']}; {c['note']}"}
             for c in checks if c["status"] != "OK"]
    return {"values": q.values, "status": q.status, "route": route, "channel": channel, "step_rows": step_rows,
            "check_flags": flags, "temp_control": tc, "exception_steps": sorted(exc_steps), "ovs_present": ovs_present,
            "shear_limits": dict(getattr(q, "shear_limits", {}) or {}),
            "legacy_4rpm": dict(getattr(q, "legacy_4rpm", {}) or {}),
            "decade_pairs": list(getattr(q, "decade_pairs", []) or []),
            "stabilization": evaluate_stabilization(reps, policy or active_policy())}


def compute_b_qc(df, holds, meta, result, tq_lo, tq_hi):
    q = _QCBuffer()
    summ = result.get("summary", {})
    base, brk = summ.get("Baseline", {}), summ.get("Breakdown", {})

    def interval_status(s, name):
        label = f"{name} S{s.get('step')} {float(s.get('rpm', np.nan)):g} rpm"
        return _st("OK", label) if s.get("n_valid", 0) > 0 else _st("NO VALUE", f"{label}: {s.get('status')}")

    bst, kst = interval_status(base, "Baseline"), interval_status(brk, "Breakdown")
    q.set("B_ETA_BASE", base.get("eta"), bst)
    q.set("B_ETA_BRK", brk.get("eta"), kst)
    decay = np.nan
    if kst["status"] == "OK":
        peak, _, _, _ = _peak_in_step(df, int(brk["step"]), tq_lo, tq_hi)
        if _finite(peak) and peak:
            decay = (peak - float(brk["tau"])) / peak * 100
    q.set("B_BRK_DECAY", decay, kst)
    rec_hold = float(result.get("recovery_hold_s", np.nan))
    rec_map = {int(r["time_s"]): r for r in result.get("recovery", [])}
    tq = base.get("torque")
    floor = tq_lo / tq * 100 if _finite(tq) and tq else np.nan
    rec = df[df["Step"] == int(result.get("recovery_step", -1))].sort_values("t_elapsed")
    rec_ok = rec[rec["Torque_pct"].between(tq_lo, tq_hi)]
    for t, code in ((10, "B_R10"), (60, "B_R60"), (300, "B_R300"), (900, "B_R900")):
        if _finite(rec_hold) and t > rec_hold + 0.5:
            q.set(code, np.nan, _st("NO VALUE", f"회복 구간 {rec_hold:g} s < {t} s"))
        elif bst["status"] != "OK":
            q.set(code, np.nan, _st("NO VALUE", "Baseline 무효"))
        else:
            r = rec_map.get(t, {})
            w = rec[(rec["t_elapsed"] >= max(1, t - 5)) & (rec["t_elapsed"] <= t + 5)]
            if _finite(r.get("recovery_eta_pct")):
                st = _st("OK", f"n={r.get('n_valid', 0)}")
            elif len(w) and float(w["Torque_pct"].max()) < tq_lo:
                st = _st("CENSORED", f"{t} s 창 Torque<{tq_lo:g}% → R{t} < 약 {floor:.1f}% (검출하한 미만)")
            else:
                st = _st("NO VALUE", f"{t} s 시점 Torque 유효점 없음")
            q.set(code, r.get("recovery_eta_pct"), st)
    t90 = result.get("target_times", {}).get(90)
    if bst["status"] != "OK":
        q.set("B_T90", np.nan, _st("NO VALUE", "Baseline 무효"))
    elif _finite(t90):
        first_ok = float(rec_ok["t_elapsed"].iloc[0]) if len(rec_ok) else np.nan
        low_before = rec[(rec["t_elapsed"] < t90) & (rec["Torque_pct"] < tq_lo)]
        if len(low_before) and _finite(first_ok) and t90 <= first_ok + 1.0:
            q.set("B_T90", np.nan, _st("CENSORED", f"t90 ≤ {t90:g} s — 90% 교차가 Torque<{tq_lo:g}% 구간에서 발생(상한값만 확인)"))
        else:
            q.set("B_T90", t90, _st("OK", "도달"))
    elif rec_ok.empty:
        q.set("B_T90", np.nan, _st("CENSORED", f"회복 구간 전체 Torque<{tq_lo:g}% → R < 약 {floor:.1f}% 유지, t90 > {rec_hold:g} s"))
    else:
        q.set("B_T90", np.nan, _st("CENSORED", f"미도달 — t90 > {rec_hold:g} s (right-censored)"))
    q.set("B_REC_FLOOR", floor, _st("OK", "") if _finite(floor) else _st("NO VALUE", "기준 Torque 없음"))
    q.set("B_TQ_BASE", tq, _st("OK", "") if _finite(tq) and tq_lo <= tq <= tq_hi else _st("REVIEW", "Torque 범위 밖"))
    rv = result.get("recovery_valid_pct")
    q.set("B_REC_VALID", rv, _st("OK", "") if _finite(rv) and rv >= 99.999 else _st("REVIEW", "회복 구간 일부 Torque 범위 밖"))
    return {"values": q.values, "status": q.status, "route": "STANDARD", "step_rows": [], "check_flags": [],
            "temp_control": str(meta.get("Temperature Control", "")).strip(), "exception_steps": []}


# ----------------------------------------------------------------------------- 밀도 입력
DENSITY_COLUMNS = ["lot", "fill_id", "method", "temp_C", "mass_empty_g", "mass_filled_g",
                   "volume_mL", "density_g_mL", "operator", "measured_at", "note"]


def normalize_lot(text):
    return re.sub(r"^HT-", "", str(text).strip(), flags=re.I).upper()


def _num(text):
    try:
        v = float(str(text).strip().replace(",", ""))
        return v if np.isfinite(v) else np.nan
    except (TypeError, ValueError):
        return np.nan


def load_density_file(path):
    """밀도 입력 CSV → {LOT: [record, ...]}. 한 행 = 독립 충전 1회."""
    path = Path(path)
    lines = [x for x in read_lines(path) if x.strip() and not x.lstrip().startswith("#")]
    reader = csv.DictReader(lines)
    reader.fieldnames = [f.strip().lower() for f in (reader.fieldnames or [])]
    missing = [c for c in ("lot", "fill_id") if c not in reader.fieldnames]
    if missing:
        raise ValueError(f"밀도 입력 필수 열 누락: {missing} (필요 열: {', '.join(DENSITY_COLUMNS)})")
    sha = sha256_file(path)
    out = defaultdict(list)
    for i, row in enumerate(reader, 2):
        row = {k.strip().lower(): (v or "").strip() for k, v in row.items() if k}
        if not row.get("lot"):
            continue
        q = _QCBuffer()
        me, mf, vol = _num(row.get("mass_empty_g")), _num(row.get("mass_filled_g")), _num(row.get("volume_mL".lower()))
        direct, temp = _num(row.get("density_g_ml")), _num(row.get("temp_c"))
        calc = (mf - me) / vol if _finite(me) and _finite(mf) and _finite(vol) and vol > 0 else np.nan
        if _finite(direct):
            rho, rst = direct, _st("OK", "직독값")
            if _finite(calc):
                diff = direct - calc
                rst = _st("OK", f"직독–질량식 차 {diff:+.4f}") if abs(diff) <= DENSITY_CROSSCHECK_TOL else \
                    _st("REVIEW", f"직독–질량식 차 {diff:+.4f} g/mL > {DENSITY_CROSSCHECK_TOL:g}")
        elif _finite(calc):
            rho, rst = calc, _st("OK", "질량/교정부피 계산")
        else:
            rho, rst = np.nan, _st("NO VALUE", "밀도·질량·부피 입력 없음")
        if not _finite(temp):
            tst = _st("REVIEW", "시험온도 미기록")
        elif abs(temp - TEST_TEMP_TARGET_C) > TEST_TEMP_TOL_C:
            tst = _st("REVIEW", f"시험온도 {temp:g} °C — {TEST_TEMP_TARGET_C:g}±{TEST_TEMP_TOL_C:g} 밖")
        else:
            tst = _st("OK", f"{temp:g} °C")
        q.set("D_RHO25", rho, _worst(rst, tst) if _finite(rho) else rst)
        q.set("D_TEMP", temp, tst)
        worst = _worst(*q.status.values())["status"]
        out[normalize_lot(row["lot"])].append({
            "lot": normalize_lot(row["lot"]), "repeat": row.get("fill_id") or f"row{i}", "file": path.name,
            "method": "D", "test_start": row.get("measured_at", ""), "route": "STANDARD",
            "data_qc": "VALID" if worst == "OK" else ("INVALID" if worst == "NO VALUE" else "REVIEW"),
            "raw_sha256": sha, "template": "", "output": "", "values": q.values, "status": q.status,
            "step_rows": [], "check_flags": [], "temp_control": "", "exception_steps": [],
            "density_meta": {"method": row.get("method", ""), "operator": row.get("operator", ""),
                             "mass_empty_g": me, "mass_filled_g": mf, "volume_mL": vol,
                             "density_direct": direct, "density_calc": calc, "note": row.get("note", "")},
        })
    return dict(out)


def write_density_template(path):
    rows = [
        "# SST gel 밀도 입력 양식 — 한 행 = 독립 충전 1회. '#' 행은 무시됩니다.",
        "# lot: 파일명 HT-<lot>-<반복>의 <lot>과 동일(HT- 접두 허용). method 예: REF_25mL_PYCNO / MICRO_GRAV_1mL / DENSITOMETER",
        "# density_g_mL(직독) 또는 mass_empty_g·mass_filled_g·volume_mL(시험온도 물 교정부피) 중 하나 이상 입력. 둘 다 있으면 교차확인.",
        "# 예) 1FI011,F1,REF_25mL_PYCNO,25.0,30.1234,56.2744,25.000,,작업자,2026-09-28 10:00,  ← '#'을 지우고 실제 값으로 사용",
        ",".join(DENSITY_COLUMNS),
    ]
    Path(path).write_text("\n".join(rows) + "\n", encoding="utf-8-sig")


# ----------------------------------------------------------------------------- 통계·출력 공통
def _defs_by_code(defs):
    return {d["code"]: d for d in defs}


def rsd_note(mean, sd, n, zero_ref=False):
    """RSD를 보고할 수 있는지 판정한다. 빈 문자열이면 RSD 보고 가능.

    - zero_ref 지표: 0을 기준으로 부호가 바뀌는 차이형 지표라 RSD(SD/|Mean|)가 정의상 부적합.
    - |Mean| < SD: RSD > 100 %. 평균이 0에 가까워 지표가 폭증한 상태이며 반복 산포의 비교값이 아니다.
    """
    if n is None or n < 2:
        return ""
    if zero_ref:
        return RSD_NA_ZERO_REF
    if not _finite(mean) or not _finite(sd):
        return ""
    if mean == 0 or abs(float(mean)) < float(sd):
        return RSD_NA_NEAR_ZERO
    return ""


def _rec_channel(r, channel):
    if channel == "inclusive":
        return r.get("values_incl", r["values"]), r.get("status_incl", r["status"])
    return r["values"], r["status"]


def lot_metric_stats(recs, code, d=None, channel="strict", policy=None):
    """Lot 반복 통계(채널별).

    STRICT: OK·NOTICE·REVIEW만 평균. EXCEPTION·CENSORED·Torque 참고값은 제외.
    INCLUSIVE: 정책이 허용하면 inclusive exception과 Torque 하한/상한 참고값도 평균에 넣는다.
    참고값을 평균에 넣어도 acceptance는 ELIGIBLE로 올리지 않는다.
    """
    d = d or {}
    policy = policy or active_policy()
    zero_ref = bool(d.get("zero_ref", False))
    pairs = []
    for r in recs:
        vals, sts = _rec_channel(r, channel)
        if code in vals:
            pairs.append((vals.get(code, np.nan), sts.get(code) or _st("NO VALUE", ""), str(r["repeat"])))
    n_total = len(pairs)
    statuses = [st["status"] for _, st, _ in pairs]
    counts = {k: statuses.count(k) for k in ("OK", "NOTICE", "REVIEW", "CENSORED", "EXCEPTION", "NO VALUE", "INVALID", "N/A")}
    # EXCEPTION 값: strict 채널 평균에는 절대 넣지 않음. inclusive 채널은 exception_stat_policy가
    # separate(기본)/include면 평균에 포함하고 별도 n·mean도 표시한다. exclude면 개수만.
    # CENSORED(Torque 참고값)는 include_censored_in_lot_stats일 때만 inclusive 평균에 넣는다.
    include_exc = channel == "inclusive" and (policy.exception_stat_policy in ("separate", "include")
                                              or policy.include_exception_in_lot_stats)
    include_cens = channel == "inclusive" and bool(policy.include_censored_in_lot_stats)
    core_pairs = []
    for v, st, rep in pairs:
        decision = lot_inclusion_decision(
            metric_code=code, value=v, channel=channel,
            data_status=st.get("data_status") or st.get("status"),
            value_class=st.get("value_class") or _default_value_class(st.get("status")),
            acceptance_status=st.get("acceptance_status"),
            source_type=st.get("source_type") or "measured",
            pair_id=st.get("pair_id"), policy=policy,
        )
        if decision["included"]:
            core_pairs.append((decision["stat_value"], st, rep))
    core = [float(v) for v, _, _ in core_pairs]
    exc_vals = [float(v) for v, st, _ in pairs if _finite(v) and st["status"] == "EXCEPTION"]
    cens_vals = [float(v) for v, st, _ in pairs if _finite(v) and st["status"] == "CENSORED"]
    n_numeric = sum(1 for v, _, _ in pairs if _finite(v))
    n = len(core)
    arr = np.array(core, float)
    mean = float(arr.mean()) if n else np.nan
    sd = float(arr.std(ddof=1)) if n > 1 else np.nan
    rsd = sd / abs(mean) * 100 if n > 1 and mean else np.nan
    note = rsd_note(mean, sd, n, zero_ref)
    n_pos = int((arr > 0).sum()) if n else 0
    n_neg = int((arr < 0).sum()) if n else 0
    n_zero = int((arr == 0).sum()) if n else 0
    sign_flip = bool(n_pos and n_neg)
    if n_total == 0:
        data = "NO DATA"
    elif n == 0:
        if counts["NOTICE"] + counts["N/A"] == n_total:
            data = "NOTICE"
        elif cens_vals or counts["CENSORED"]:
            data = "CENSORED"
        elif exc_vals:
            data = "EXCEPTION"
        elif counts["NO VALUE"] and not counts["INVALID"]:
            data = "NO VALUE"
        else:
            data = "INVALID"
    elif exc_vals:
        data = "EXCEPTION"                      # 예외 hold 값이 평균에 있거나(inclusive) 별도 집계됨 — 표준 통계와 혼용 금지
    elif counts["REVIEW"] or counts["CENSORED"] or counts["NO VALUE"] or counts["INVALID"] or n_numeric < n_total:
        data = "REVIEW"
    elif counts["NOTICE"] or counts["N/A"]:
        data = "NOTICE"
    else:
        data = "VALID"
    groups = defaultdict(list)
    for v, st, rep_id in pairs:
        reason = st.get("reason") or ""
        keep = st["status"] != "OK" or any(
            token in reason for token in ("Full H-B only", "fitting range differs", "turnaround",
                                          "실제 최고 RPM", "측정 최고 RPM", "legacy reference", "eta_up(",
                                          "fixed low reference", "actual maximum", "hold="))
        if keep and reason:
            groups[f"[{st['status']}] {reason}"].append(rep_id)
    reasons = " | ".join(f"{k} ({','.join(v)})" for k, v in groups.items())
    if zero_ref and sign_flip:
        if data in ("VALID", "NOTICE"):
            data = SIGN_FLIP_STATUS
        reasons = (reasons + " | " if reasons else "") + (
            f"[SIGN] 부호 전환 +{n_pos}/−{n_neg}, Range {float(arr.max() - arr.min()):.2f} — 원시곡선(peak 검출·대표창) 확인")
    if include_cens and any(st.get("status") == "CENSORED" for _, st, _ in core_pairs) and data in ("VALID", "NOTICE", "REVIEW"):
        data = "CENSORED"
    src_pairs = core_pairs if core_pairs else pairs
    merged_vc = _merge_vc_fields([{
        "value_class": (st.get("value_class") or _default_value_class(st.get("status"))),
        "used_low": bool(st.get("used_low")), "used_high": bool(st.get("used_high")),
        "low_steps": list(st.get("low_steps") or []), "high_steps": list(st.get("high_steps") or []),
        "exc_steps": list(st.get("exc_steps") or []), "exc_holds": list(st.get("exc_holds") or []),
    } for _, st, _ in src_pairs])
    vc = merged_vc.get("value_class") or (VC_NO if n == 0 else VC_STRICT)

    def _low(st):
        return bool(st.get("used_low")) or st.get("value_class") in (VC_LOW, VC_BOTH)

    def _high(st):
        return bool(st.get("used_high")) or st.get("value_class") in (VC_HIGH, VC_BOTH)

    n_low_ref = sum(1 for v, st, _ in pairs if _finite(v) and _low(st))
    n_high_ref = sum(1 for v, st, _ in pairs if _finite(v) and _high(st))
    exc_details, low_details, high_details = [], [], []
    for v, st, rep in core_pairs:
        if st.get("exc_steps") or st.get("value_class") == VC_EXC or st.get("exc_holds"):
            steps = ",".join(st.get("exc_steps") or [])
            holds = ", ".join(f"{float(h):g} s" for h in (st.get("exc_holds") or []))
            exc_details.append("repeat " + str(rep) + (f", Step {steps}" if steps else "") + (f", hold {holds}" if holds else ""))
        if _low(st):
            steps = ",".join(st.get("low_steps") or [])
            low_details.append("repeat " + str(rep) + (f" Step {steps}" if steps else ""))
        if _high(st):
            steps = ",".join(st.get("high_steps") or [])
            high_details.append("repeat " + str(rep) + (f" Step {steps}" if steps else ""))
    def _join_meta(pairs):
        vals = [x for _, x in pairs if x]
        uniq = list(dict.fromkeys(vals))
        if not uniq:
            return "", False
        if len(uniq) == 1:
            return uniq[0], False
        return "; ".join(f"repeat {a}: {b}" for a, b in pairs if b), True
    rpm_pairs, step_pairs, gamma_pairs, models, rule_pairs = [], [], [], [], []
    for _, st, rep in core_pairs:
        rpm_pairs.append((rep, st.get("apply_rpm") or ""))
        step_pairs.append((rep, st.get("contrib_steps") or ""))
        gamma_pairs.append((rep, st.get("apply_gamma") or ""))
        rule_pairs.append((rep, st.get("selection_rule") or ""))
        if st.get("fit_model"):
            models.append(st.get("fit_model"))
    apply_rpm, rpm_mix = _join_meta(rpm_pairs)
    selection_rule, _rule_mix = _join_meta(rule_pairs)
    contrib_steps, _step_mix = _join_meta(step_pairs)
    apply_gamma, _gamma_mix = _join_meta(gamma_pairs)

    def _join_field(key, as_step=False):
        packed = []
        for _, st, rep in pairs:
            raw = st.get(key)
            if as_step:
                text = _step_label(raw) or ""
            elif _finite(raw):
                text = _rpm_label(raw) if key.endswith("rpm") else f"{float(raw):g}"
            else:
                text = "" if raw in (None, "") else str(raw)
            packed.append((rep, text))
        return _join_meta(packed)

    low_rpm, low_mix = _join_field("selected_low_rpm")
    high_rpm, high_mix = _join_field("selected_high_rpm")
    low_step, _low_step_mix = _join_field("selected_low_step", as_step=True)
    high_step, _high_step_mix = _join_field("selected_high_step", as_step=True)
    low_gamma, _low_g_mix = _join_field("selected_low_gamma")
    high_gamma, _high_g_mix = _join_field("selected_high_gamma")
    if (rpm_mix or high_mix or low_mix) and data in ("VALID", "NOTICE"):
        data = "REVIEW"
        reasons = (reasons + " | " if reasons else "") + "반복별 적용 RPM이 다름"
    if not d.get("lot_acceptance", True):
        acceptance = "—"
    else:
        acceptance = {"VALID": "ELIGIBLE", "NOTICE": "ELIGIBLE", "REVIEW": "REVIEW", "EXCEPTION": "EXCLUDED",
                      "CENSORED": "EXCLUDED", "INVALID": "INVALID", "NO VALUE": "EXCLUDED", "NO DATA": "NO DATA"}.get(data, data)
        if data == "EXCEPTION" and policy.exception_stat_policy == "include":
            acceptance = "REVIEW"
        if core_pairs and vc in (VC_LOW, VC_HIGH, VC_BOTH):
            acceptance = "EXCLUDED"
        elif core_pairs and vc == VC_EXC and channel == "inclusive" and include_exc:
            acceptance = "REVIEW"
    return {"channel": channel, "n_total": n_total, "n": n, "n_numeric": n_numeric, "mean": mean, "sd": sd, "rsd": rsd,
            "value_class": vc, "n_low_ref": n_low_ref, "n_high_ref": n_high_ref,
            "apply_rpm": apply_rpm, "apply_gamma": apply_gamma, "contrib_steps": contrib_steps,
            "selection_rule": selection_rule,
            "low_rpm": low_rpm, "high_rpm": high_rpm, "low_step": low_step, "high_step": high_step,
            "low_gamma": low_gamma, "high_gamma": high_gamma,
            "numerator_rpm": _join_field("numerator_rpm")[0], "denominator_rpm": _join_field("denominator_rpm")[0],
            "numerator_step": _join_field("numerator_step", as_step=True)[0],
            "denominator_step": _join_field("denominator_step", as_step=True)[0],
            "numerator_gamma": _join_field("numerator_gamma")[0], "denominator_gamma": _join_field("denominator_gamma")[0],
            "actual_ratio": _join_field("actual_ratio")[0], "source_type": _join_field("source_type")[0],
            "measured_high_step": _join_field("measured_high_step", as_step=True)[0],
            "measured_high_rpm": _join_field("measured_high_rpm")[0],
            "rpm_mismatch": rpm_mix, "fit_model": models[0] if models else "",
            "exc_details": exc_details, "low_details": low_details, "high_details": high_details,
            "rsd_note": note, "rsd_text": note if note else (f"{rsd:.2f}" if _finite(rsd) else ""),
            "min": float(arr.min()) if n else np.nan, "max": float(arr.max()) if n else np.nan,
            "range": float(arr.max() - arr.min()) if n else np.nan,
            "abs_mean": float(np.abs(arr).mean()) if n else np.nan,
            "n_pos": n_pos, "n_neg": n_neg, "n_zero": n_zero, "sign_flip": sign_flip,
            "sign_text": (f"+{n_pos}/−{n_neg}" + (f"/0×{n_zero}" if n_zero else "")) if n else "",
            "n_notice": counts["NOTICE"] + counts["N/A"], "n_review": counts["REVIEW"], "n_exception": len(exc_vals),
            "n_censored": counts["CENSORED"] + sum(
                1 for _, st, _ in pairs
                if st.get("value_class") in TORQUE_REFERENCE_CLASSES and st.get("status") != "CENSORED"
            ), "n_no_value": counts["NO VALUE"], "n_invalid": counts["INVALID"],
            "included_in_lot_stats": bool(n),
            "lot_stat_exclusion_reason": "" if n else " | ".join(dict.fromkeys(
                (st.get("lot_stat_exclusion_reason") or st.get("reason") or "")
                for _, st, _ in pairs if (st.get("lot_stat_exclusion_reason") or st.get("reason"))
            )),
            "exc_mean": float(np.mean(exc_vals)) if exc_vals else np.nan,
            "exc_sd": float(np.std(exc_vals, ddof=1)) if len(exc_vals) > 1 else np.nan,
            "cens_n": len(cens_vals),
            "zero_ref": zero_ref, "data_status": data, "acceptance_status": acceptance, "reasons": reasons}


def lot_metric_stats_both(recs, code, d=None, policy=None):
    """strict/inclusive 두 채널 통계와 Δ(inclusive − strict)."""
    policy = policy or active_policy()
    st = lot_metric_stats(recs, code, d, "strict", policy)
    inc = lot_metric_stats(recs, code, d, "inclusive", policy)
    delta = inc["mean"] - st["mean"] if _finite(inc["mean"]) and _finite(st["mean"]) else np.nan
    # primary는 Summary·콘솔·LotCompare가 쓰는 대표 통계다.
    # Method A만 inclusive 채널을 가지므로 Method B·Density는 strict 통계를 대표값으로 유지한다.
    primary = (
        inc
        if d is not None
        and d.get("method") == "A"
        and policy.primary_channel == "inclusive"
        else st
    )
    out = dict(primary)
    out.update({"strict": st, "inclusive": inc, "delta": delta,
                "delta_pct": (delta / abs(st["mean"]) * 100) if _finite(delta) and st["mean"] else np.nan})
    return out


def prec_status(d, s):
    """Lot_QC_Summary의 Excel 수식과 같은 규칙으로 반복성 상태를 Python에서 계산한다(Compare·CLI용)."""
    rule, good, rev = d["prec"]
    if rule == "-" or s["n"] < 2:
        return "N/A"
    if rule == "RSD":
        if s["rsd_note"] or not _finite(s["rsd"]):
            return "HIGH VAR" if _finite(s["sd"]) else "N/A"
        x = s["rsd"]
    else:
        x = s["sd"]
    if not _finite(x):
        return "N/A"
    return "GOOD" if x <= good else ("REVIEW" if x <= rev else "HIGH VAR")


def _est_lines(text, width):
    if text is None:
        return 1
    s = str(text)
    total = 0
    for part in s.split("\n"):
        units = sum(2.0 if ord(ch) >= 0x1100 else 1.05 for ch in part)
        total += max(1, math.ceil(units / max(width - 2.5, 4)))
    return total


def _fit_row_height(ws, row, cols_widths, base=15.5, max_lines=12):
    lines = 1
    for col, width in cols_widths.items():
        lines = max(lines, _est_lines(ws.cell(row, col).value, width))
    ws.row_dimensions[row].height = min(max_lines, lines) * base


def _status_cf(ws, cell_range):
    first = cell_range.split(":")[0]
    col = re.match(r"[A-Z]+", first).group(0)
    row = re.search(r"\d+", first).group(0)
    ref = f"${col}{row}"
    for text, color in LOT_STATUS_FILL.items():
        ws.conditional_formatting.add(cell_range, FormulaRule(formula=[f'{ref}="{text}"'],
                                                              fill=PatternFill("solid", fgColor=color), stopIfTrue=True))


def _value_class_cf(ws, cell_range):
    """값 상태 열 전용. 데이터 상태 exact-match 서식과 셀을 섞지 않는다."""
    first = cell_range.split(":")[0]
    col = re.match(r"[A-Z]+", first).group(0)
    row = re.search(r"\d+", first).group(0)
    ref = f"${col}{row}"
    for text, color in VALUE_CLASS_FILL.items():
        ws.conditional_formatting.add(cell_range, FormulaRule(formula=[f'{ref}="{text}"'],
                                                              fill=PatternFill("solid", fgColor=color), stopIfTrue=True))


def _fmt_for(unit):
    return {"g/mL": "0.0000", "Pa": "0.00", "cP": "#,##0", "%": "0.00", "-": "0.000",
            "Pa·sⁿ": "0.000", "s": "0.0", "°C": "0.00", "개": "0", "Pa/s": "0.00"}.get(unit, "0.000")


def _banner(ws, text, c2, row=4):
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=c2)
    c = ws.cell(row, 2, text)
    c.font = Font(bold=True, color="9C0006")
    c.fill = PatternFill("solid", fgColor="FFC7CE")
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 30


def _print_setup(ws):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


def write_qc_indices_sheet(wb, method, qc, defs=None, policy=None):
    """개별 반복 보고서의 QC 지표 시트(04_2 시트는 변경하지 않음). 06: Strict/Inclusive 값·상태·Δ 병렬 표시."""
    policy = policy or active_policy()
    defs = _export_defs([d for d in (defs or metric_definitions()) if d["method"] == method], [qc])
    if "QC_Indices" in wb.sheetnames:
        del wb["QC_Indices"]
    ws = wb.create_sheet("QC_Indices", 1)
    title = "Method A" if method == "A" else "Method B 3ITT"
    show_s = policy.report_values in ("strict", "both")
    show_i = policy.report_values in ("inclusive", "both") and method == "A"
    headers = ["Tier", "Role", "Code", "지표", "단위"]
    if show_s:
        headers += ["값 Strict", "상태 Strict", "Strict 값 구분"]
    if show_i:
        headers += ["값 Inclusive", "상태 Inclusive", "Inclusive 값 구분", "Δ(Incl−Strict)"]
    headers += ["근거·사유", "가리키는 물성", "합격성 반영"]
    n_cols = len(headers) + 1
    _sheet_base(ws, f"SST gel {title} — QC 지표와 가리키는 물성",
                "Strict = 60 s hold·Torque 10–95 % 점만 / Inclusive = 정책이 허용한 hold·범위 밖 점 포함. "
                "상태는 데이터·계산 유효성이며 제품 합격 판정이 아닙니다. 값 구분은 별도 열입니다. 공란 = 산출 불가(대체값 없음).", n_cols)
    _nav(ws, wb.sheetnames[:9])
    if policy.name != "STANDARD":
        _banner(ws, f"POLICY {policy.name} — accepted holds {','.join(f'{x:g}' for x in policy.accepted_hold_s)} s; "
                    "EXCEPTION/CENSORED 값은 표준 60 s 통계·규격 설정에 혼용 금지", n_cols)
    _section(ws, 5, f"반복 {qc.get('repeat')} | {qc.get('file')} | 데이터 QC {qc.get('data_qc')} | 정책 {policy.name} "
                    f"| A_OVS_L_PRESENT={qc.get('ovs_present', '—')} "
                    f"| 고정 저점 0.5 rpm | 측정 최고 RPM strict/incl "
                    f"{qc.get('strict_measured_max_rpm', '—')}/{qc.get('inclusive_measured_max_rpm', '—')} "
                    f"| 대표 고점 strict/incl {qc.get('strict_high_rpm', '—')}/{qc.get('inclusive_high_rpm', '—')}", 2, n_cols)
    _headers(ws, 6, headers, 2)
    col_w = {"Tier": 14, "Role": 12, "Code": 16, "지표": 32, "단위": 7, "값 Strict": 12, "상태 Strict": 11,
             "Strict 값 구분": 24, "값 Inclusive": 12, "상태 Inclusive": 11, "Inclusive 값 구분": 24,
             "Δ(Incl−Strict)": 11, "근거·사유": 48, "가리키는 물성": 46, "합격성 반영": 9}
    for j, h in enumerate(headers, 2):
        ws.column_dimensions[get_column_letter(j)].width = col_w.get(h, 12)
    row = 7
    status_cols = []
    vi, si = qc.get("values_incl", {}), qc.get("status_incl", {})
    for d in sorted(defs, key=lambda x: TIER_ORDER.index(x["tier"])):
        if d.get("output_policy") in ("decade_sheet", "dictionary_only", "legacy"):
            continue
        code = d["code"]
        v = qc["values"].get(code, np.nan)
        st = qc["status"].get(code, _st("NO VALUE", "미계산"))
        v2 = vi.get(code, np.nan)
        st2 = si.get(code, _st("NO VALUE", "미계산"))
        vals = [TIER_LABELS[d["tier"]], d["metric_role"], code, d["name"], d["unit"]]
        if show_s:
            vals += [v if _finite(v) else None, st["status"], st.get("value_class") or _default_value_class(st["status"])]
        if show_i:
            vals += [v2 if _finite(v2) else None, st2["status"],
                     st2.get("value_class") or _default_value_class(st2["status"]),
                     (v2 - v) if _finite(v) and _finite(v2) else None]
        reason = st["reason"] if not show_i else (
            f"[S] {st['reason']}" + (f" ‖ [I] {st2['reason']}" if st2["reason"] != st["reason"] else ""))
        vals += [reason, d["property"], "예" if d["lot_acceptance"] else "아니오(진단)"]
        for j, value in enumerate(vals, 2):
            h = headers[j - 2]
            fmt = _fmt_for(d["unit"]) if h in ("값 Strict", "값 Inclusive", "Δ(Incl−Strict)") else None
            c = put(ws, row, j, value, fmt)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            c.border = Border(bottom=THIN)
        _fit_row_height(ws, row, {headers.index("지표") + 2: 32, headers.index("근거·사유") + 2: 48,
                                  headers.index("가리키는 물성") + 2: 46}, max_lines=8)
        row += 1
    for h in ("상태 Strict", "상태 Inclusive"):
        if h in headers:
            col = get_column_letter(headers.index(h) + 2)
            _status_cf(ws, f"{col}7:{col}{row - 1}")
    for h in ("Strict 값 구분", "Inclusive 값 구분"):
        if h in headers:
            col = get_column_letter(headers.index(h) + 2)
            _value_class_cf(ws, f"{col}7:{col}{row - 1}")
    ws.freeze_panes = "E7"
    _print_setup(ws)
    return ws


STEP_AUDIT_COLUMNS = ["lot", "repeat", "file", "step", "direction", "rpm", "shear_rate", "hold_config_s", "hold_observed_s",
                      "hold_class", "tail_window_s", "tail_start_s", "tail_end_s", "torque_min_pct", "torque_max_pct",
                      "n_all", "n_in_range", "n_low_torque", "n_high_torque", "tau_pa_strict", "tau_pa_inclusive",
                      "eta_cp_strict", "eta_cp_inclusive", "torque_mean_all", "torque_min_all", "torque_max_all",
                      "strict_hold_valid", "strict_torque_valid", "analysis_torque_mode",
                      "included_in_strict", "included_in_inclusive", "status", "label", "reason",
                      "valueclassstrict", "valueclassinclusive", "usedlowtorque", "usedhightorque",
                      "sequence_direction", "used_in_up_branch", "used_in_down_branch",
                      "used_in_strict_full_hb", "used_in_inclusive_full_hb",
                      "used_in_strict_up_full_hb", "used_in_strict_down_full_hb",
                      "used_in_inclusive_up_full_hb", "used_in_inclusive_down_full_hb",
                      "used_in_common_hb", "used_in_hysteresis", "turnaround_shared",
                      "used_for_eta_high", "used_for_tau_high", "used_for_ssi", "used_for_ssi_dynamic",
                      "step_type", "target_temperature_C", "mean_temperature_C", "min_temperature_C", "max_temperature_C",
                      "sampling_interval_s", "n_points", "stabilization_status", "stabilization_reason",
                      "used_in_ssi_field_0p2_2", "used_in_ssi_qc_0p5_5", "used_in_decade_ssi", "value_class_strict", "value_class_inclusive",
                      "full_hb_direction", "exclusion_reason_full_hb", "exclusion_reason_common_hb",
                      "exclusion_reason_hysteresis", "gamma",
                      "used_in_decade_ssi_up", "used_in_decade_ssi_down",
                      "exact_rpm_candidate_for", "exact_rpm_selected_for", "exact_rpm_exclusion_reason"]


def step_audit_rows(qc, policy=None):
    """qc['audit_rows'](reps 06 필드)를 지시서 §M 컬럼으로 정렬한다. 모든 회전 Step이 한 줄씩 남는다."""
    policy = policy or active_policy()
    raw = list(qc.get("audit_rows", []))

    def _on(r, branch):
        key = "used_in_up_branch" if branch == "up" else "used_in_down_branch"
        if key in r:
            return bool(r.get(key))
        d = r.get("sequence_direction") or r.get("dir")
        if d == "Turnaround":
            return True
        return d == ("Up" if branch == "up" else "Down")

    def _valued(r, channel):
        d = r.get("sequence_direction") or r.get("dir")
        if d not in ("Up", "Down", "Turnaround"):
            return False
        if not _finite(r.get("gamma")) or float(r.get("gamma") or 0) <= 0:
            return False
        if str(r.get("data_status")) == "INVALID":
            return False
        if channel == "strict":
            return bool(r.get("included_in_strict")) and _finite(r.get("tau_Pa_strict"))
        return bool(r.get("included_in_inclusive")) and _finite(r.get("tau_Pa_inclusive"))

    def _paired(r, channel):
        if not _valued(r, channel):
            return False
        if bool(r.get("turnaround_shared")) and _on(r, "up") and _on(r, "down"):
            return True
        def _match(x):
            return (_valued(x, channel) and _finite(x.get("RPM")) and _finite(r.get("RPM"))
                    and abs(float(x["RPM"]) - float(r["RPM"])) <= RPM_MATCH_TOL
                    and _finite(x.get("gamma")) and _finite(r.get("gamma"))
                    and abs(float(x["gamma"]) - float(r["gamma"])) <= _shear_match_tol())
        if _on(r, "up") and not _on(r, "down"):
            return any(_match(x) and _on(x, "down") for x in raw)
        if _on(r, "down") and not _on(r, "up"):
            return any(_match(x) and _on(x, "up") for x in raw)
        return False

    def _max_ids(rows_):
        if not rows_:
            return set()
        mx = max(float(x["RPM"]) for x in rows_)
        return {int(x["Step"]) for x in rows_ if abs(float(x["RPM"]) - mx) <= RPM_MATCH_TOL}

    def _step_id(store, code, key):
        value = ((store or {}).get(code) or {}).get(key)
        try:
            return {int(value)} if value is not None and value != "" else set()
        except (TypeError, ValueError):
            return set()

    field_steps, qc_steps, decade_steps = set(), set(), set()
    decade_up, decade_dn = set(), set()
    exact_candidate, exact_selected, exact_reason = {}, {}, {}
    for store in (qc.get("status") or {}, qc.get("status_incl") or {}):
        field_steps |= _step_id(store, "A_SSI_FIELD_0P2_2", "numerator_step") | _step_id(store, "A_SSI_FIELD_0P2_2", "denominator_step")
        qc_steps |= _step_id(store, "A_SSI_QC_0P5_5", "numerator_step") | _step_id(store, "A_SSI_QC_0P5_5", "denominator_step")
        for code in ("A_SSI_FIELD_0P2_2", "A_SSI_QC_0P5_5"):
            st = (store or {}).get(code) or {}
            for end in ("numerator", "denominator"):
                sel = st.get(f"{end}_status") or {}
                label = f"{code}:{end}"
                for step in sel.get("measured_candidate_steps") or []:
                    exact_candidate.setdefault(int(step), []).append(label)
                    if sel.get("exclusion_reason") and not sel.get("found"):
                        exact_reason.setdefault(int(step), []).append(sel.get("exclusion_reason"))
                if sel.get("selected_step") is not None:
                    exact_selected.setdefault(int(sel["selected_step"]), []).append(label)
    for key in ("decade_pairs", "decade_pairs_strict", "decade_pairs_incl"):
        for pair in qc.get(key) or []:
            decade_steps.add(int(pair["low_step"]))
            decade_steps.add(int(pair["high_step"]))
            bucket = decade_dn if pair.get("direction") == "down" else decade_up
            bucket.add(int(pair["low_step"]))
            bucket.add(int(pair["high_step"]))
    eta_steps, tau_steps, ssi_steps = set(), set(), set()
    measured_rows = [x for x in raw if _finite(x.get("RPM")) and _finite(x.get("gamma")) and float(x.get("gamma") or 0) > 0]
    measured_max = max((float(x["RPM"]) for x in measured_rows), default=np.nan)
    for channel in ("strict", "inclusive"):
        elig = [x for x in raw if _valued(x, channel) and _on(x, "up")]
        if not elig or not _finite(measured_max):
            continue
        eligible_max = max(float(x["RPM"]) for x in elig)
        if measured_max > eligible_max + RPM_MATCH_TOL:
            continue
        high_ids = {int(x["Step"]) for x in elig if abs(float(x["RPM"]) - eligible_max) <= RPM_MATCH_TOL}
        eta_steps |= high_ids
        tau_steps |= high_ids
        low_ids = {int(x["Step"]) for x in elig if abs(float(x["RPM"]) - FIXED_LOW_RPM) <= RPM_MATCH_TOL}
        if low_ids and high_ids:
            ssi_steps |= low_ids | high_ids

    out = []
    for r in raw:
        seq = r.get("sequence_direction") or r.get("dir")
        up_b, down_b = _on(r, "up"), _on(r, "down")
        strict_up = up_b and _valued(r, "strict")
        strict_down = down_b and _valued(r, "strict")
        incl_up = up_b and _valued(r, "inclusive")
        incl_down = down_b and _valued(r, "inclusive")
        strict_full = strict_up or strict_down
        incl_full = incl_up or incl_down
        common = _paired(r, "inclusive") or _paired(r, "strict")
        hyst = common
        shared = bool(r.get("turnaround_shared"))
        reasons_full = []
        if up_b or down_b:
            if not strict_full:
                reasons_full.append("strict Full 제외: " + (r.get("status_reason") or r.get("data_status") or "대표값 없음"))
            if not incl_full:
                reasons_full.append("inclusive Full 제외: " + (r.get("status_reason") or "정책상 분석 불가 또는 대표값 없음"))
        else:
            reasons_full = ["0 rpm stabilization Step" if r.get("step_type") == "REST_STABILIZATION" else "Rest 또는 비회전 Step"]
        if shared:
            common_reason = "" if common else "turnaround Step 대표값이 없어 Common/Hysteresis 경계점에서 제외"
        else:
            common_reason = "" if common else "no paired point — Full H-B only; Common/Hysteresis 제외"
        if up_b and down_b:
            hb_dir = "Up+Down"
        elif up_b:
            hb_dir = "Up"
        elif down_b:
            hb_dir = "Down"
        else:
            hb_dir = ""
        out.append({
            "lot": qc.get("lot"), "repeat": qc.get("repeat"), "file": qc.get("file"),
            "step": int(r["Step"]), "direction": seq, "rpm": r["RPM"], "shear_rate": r["gamma"],
            "hold_config_s": r["hold_s"], "hold_observed_s": r["hold_observed_s"], "hold_class": r["hold_class"],
            "tail_window_s": r["tail_window_s"], "tail_start_s": r["tail_start_s"], "tail_end_s": r["tail_end_s"],
            "torque_min_pct": policy.torque_min_pct, "torque_max_pct": policy.torque_max_pct,
            "n_all": r["torque_all_n"], "n_in_range": r["torque_valid_n"], "n_low_torque": r["n_low_torque"],
            "n_high_torque": r["n_high_torque"], "tau_pa_strict": r["tau_Pa_strict"], "tau_pa_inclusive": r["tau_Pa_inclusive"],
            "eta_cp_strict": r["eta_cP_strict"], "eta_cp_inclusive": r["eta_cP_inclusive"],
            "torque_mean_all": r["torque_mean_all"], "torque_min_all": r["torque_min_all"], "torque_max_all": r["torque_max_all"],
            "strict_hold_valid": "Yes" if r["strict_hold_valid"] else "No",
            "strict_torque_valid": "Yes" if r["strict_torque_valid"] else "No",
            "analysis_torque_mode": r["analysis_torque_mode"],
            "included_in_strict": "Yes" if r["included_in_strict"] else "No",
            "included_in_inclusive": "Yes" if r["included_in_inclusive"] else "No",
            "status": r["data_status"], "label": r["label"], "reason": r["status_reason"],
            "valueclassstrict": r.get("value_class_strict") or "",
            "valueclassinclusive": r.get("value_class_inclusive") or "",
            "usedlowtorque": "Yes" if r.get("used_low_torque_in_inclusive") else "No",
            "usedhightorque": "Yes" if r.get("used_high_torque_in_inclusive") else "No",
            "sequence_direction": seq,
            "used_in_up_branch": "Yes" if up_b else "No",
            "used_in_down_branch": "Yes" if down_b else "No",
            "used_in_strict_full_hb": "Yes" if strict_full else "No",
            "used_in_inclusive_full_hb": "Yes" if incl_full else "No",
            "used_in_strict_up_full_hb": "Yes" if strict_up else "No",
            "used_in_strict_down_full_hb": "Yes" if strict_down else "No",
            "used_in_inclusive_up_full_hb": "Yes" if incl_up else "No",
            "used_in_inclusive_down_full_hb": "Yes" if incl_down else "No",
            "used_in_common_hb": "Yes" if common else "No",
            "used_in_hysteresis": "Yes" if hyst else "No",
            "turnaround_shared": "Yes" if shared else "No",
            "used_for_eta_high": "Yes" if int(r["Step"]) in eta_steps else "No",
            "used_for_tau_high": "Yes" if int(r["Step"]) in tau_steps else "No",
            "used_for_ssi": "Yes" if int(r["Step"]) in (field_steps | qc_steps) else "No",
            "used_for_ssi_dynamic": "Yes" if int(r["Step"]) in decade_steps else "No",
            "step_type": r.get("step_type") or "",
            "target_temperature_C": policy.stabilization_temp_target_C if r.get("step_type") == "REST_STABILIZATION" else None,
            "mean_temperature_C": r.get("temp"), "min_temperature_C": r.get("temp_min"), "max_temperature_C": r.get("temp_max"),
            "sampling_interval_s": r.get("median_sampling_interval_s"), "n_points": r.get("n_window"),
            "stabilization_status": (qc.get("stabilization") or {}).get("stabilization_status") if r.get("step_type") == "REST_STABILIZATION" else "",
            "stabilization_reason": (qc.get("stabilization") or {}).get("stabilization_reason") if r.get("step_type") == "REST_STABILIZATION" else "",
            "used_in_ssi_field_0p2_2": "Yes" if int(r["Step"]) in field_steps else "No",
            "used_in_ssi_qc_0p5_5": "Yes" if int(r["Step"]) in qc_steps else "No",
            "used_in_decade_ssi": "Yes" if int(r["Step"]) in decade_steps else "No",
            "value_class_strict": r.get("value_class_strict") or "",
            "value_class_inclusive": r.get("value_class_inclusive") or "",
            "full_hb_direction": hb_dir,
            "exclusion_reason_full_hb": "" if (strict_full and incl_full) else "; ".join(reasons_full),
            "exclusion_reason_common_hb": common_reason,
            "exclusion_reason_hysteresis": common_reason,
            "gamma": r["gamma"],
            "used_in_decade_ssi_up": "Yes" if int(r["Step"]) in decade_up else "No",
            "used_in_decade_ssi_down": "Yes" if int(r["Step"]) in decade_dn else "No",
            "exact_rpm_candidate_for": "; ".join(dict.fromkeys(exact_candidate.get(int(r["Step"]), []))),
            "exact_rpm_selected_for": "; ".join(dict.fromkeys(exact_selected.get(int(r["Step"]), []))),
            "exact_rpm_exclusion_reason": "; ".join(dict.fromkeys(exact_reason.get(int(r["Step"]), []))),
        })
    return out


STEP_AUDIT_FMT = {"rpm": "0.00", "shear_rate": "0.00", "hold_config_s": "0", "hold_observed_s": "0.0", "tail_window_s": "0.0",
                  "tail_start_s": "0.0", "tail_end_s": "0.0", "tau_pa_strict": "0.00", "tau_pa_inclusive": "0.00",
                  "eta_cp_strict": "#,##0", "eta_cp_inclusive": "#,##0", "torque_mean_all": "0.00", "torque_min_all": "0.0",
                  "torque_max_all": "0.0"}


def write_step_audit_sheet(wb, qc, policy=None, rows=None, title_extra=""):
    """StepAudit 시트(지시서 §M). 개별 보고서와 Lot 워크북 공용."""
    policy = policy or active_policy()
    if "StepAudit" in wb.sheetnames:
        del wb["StepAudit"]
    ws = wb.create_sheet("StepAudit")
    rows = rows if rows is not None else step_audit_rows(qc, policy)
    _sheet_base(ws, f"StepAudit — 모든 회전 Step의 포함/제외·창·Torque 감사{title_extra}",
                "Step은 버려지지 않고 한 줄씩 남습니다. strict=60 s·Torque 범위 내 점만, inclusive=정책 허용 hold·전 점. "
                "status/label/reason이 왜 어느 채널에 포함·제외됐는지를 말합니다.", 16)
    _nav(ws, wb.sheetnames[:9])
    _write_table_sheet(ws, STEP_AUDIT_COLUMNS, rows, 5, "StepAudit", STEP_AUDIT_FMT)
    for j, h in enumerate(STEP_AUDIT_COLUMNS, 2):
        ws.column_dimensions[get_column_letter(j)].width = 60 if h == "reason" else (30 if h == "label" else (18 if h in ("file", "analysis_torque_mode") else 11))
    if rows:
        col = get_column_letter(STEP_AUDIT_COLUMNS.index("status") + 2)
        _status_cf(ws, f"{col}6:{col}{5 + len(rows)}")
        for key in ("valueclassstrict", "valueclassinclusive"):
            vcol = get_column_letter(STEP_AUDIT_COLUMNS.index(key) + 2)
            _value_class_cf(ws, f"{vcol}6:{vcol}{5 + len(rows)}")
        for i in range(6, 6 + len(rows)):
            ws.cell(i, STEP_AUDIT_COLUMNS.index("reason") + 2).alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "F6"
    _print_setup(ws)
    return ws


def policy_rows(policy, run_info=None):
    run_info = run_info or {}
    rows = [("정책 이름", policy.name), ("정책 해시(SHA-256 앞 16자)", policy.digest()),
            ("Hold 목표(s)", policy.hold_target_s), ("Hold 허용오차(s)", policy.hold_tolerance_s),
            ("승인 hold(s)", ", ".join(f"{x:g}" for x in policy.accepted_hold_s)),
            ("짧은 hold 분석 허용", policy.allow_short_hold_analysis),
            ("Torque 유효범위(%)", f"{policy.torque_min_pct:g}–{policy.torque_max_pct:g}"),
            ("범위 밖 Torque 점 inclusive 포함", policy.include_out_of_torque_analysis),
            ("대표창(s)", policy.tail_seconds), ("미승인 짧은 hold 창 비율", policy.short_hold_window_fraction),
            ("최소 창(s)", policy.min_window_s), ("최소 hold(s, 미만 INVALID)", policy.min_hold_s),
            ("H-B 최소/권장 점", f"{policy.min_hb_points}/{policy.recommended_hb_points}"),
            ("CENSORED Step inclusive H-B 포함", policy.censored_in_inclusive_fit),
            ("범위 밖 점 inclusive 대표값 포함", policy.include_out_of_range_in_inclusive),
            ("CENSORED 참고값 inclusive Lot 통계 포함", policy.include_censored_in_lot_stats),
            ("대표 채널", policy.primary_channel),
            ("Lot 통계 모드", policy.lot_stats_mode), ("EXCEPTION 통계 정책", policy.exception_stat_policy),
            ("보고 채널", policy.report_values), ("NOTICE 합격성 반영", policy.include_notice_in_lot_acceptance),
            ("감사 수준", policy.audit_level),
            ("임시 기준", f"드리프트 REVIEW >{EQ_DRIFT_REVIEW_PCT:g}%, 시험온도 {TEST_TEMP_TARGET_C:g}±{TEST_TEMP_TOL_C:g} °C, "
                         f"온도폭 ≤{TEMP_SPAN_MAX_C:g} °C, 밀도 교차확인 ±{DENSITY_CROSSCHECK_TOL:g} g/mL, RPM 일치 ±{RPM_MATCH_TOL:g}"),
            ("고정 평가점", f"Low={FIXED_LOW_RPM:g} rpm, Mid={FIXED_MID_RPM:g} rpm; 고전단은 실제 측정 최고 RPM"),
            ("분석기", f"{VERSION} / {QC_LAYER_VERSION}"), ("정책 JSON", policy.to_json())]
    for k in ("template", "raw_sha256", "command", "timestamp"):
        if run_info.get(k):
            rows.append((k, run_info[k]))
    return rows


def write_policy_sheet(wb, policy, run_info=None):
    """QCPolicy 시트(지시서 §L). 같은 코드라도 정책이 바뀌면 결과가 달라지므로 적용값 전부를 남긴다."""
    if "QCPolicy" in wb.sheetnames:
        del wb["QCPolicy"]
    ws = wb.create_sheet("QCPolicy")
    _sheet_base(ws, "QCPolicy — 이 결과에 적용된 분석 정책·임계값", "정책 해시가 같으면 같은 규칙으로 계산된 결과입니다.", 8)
    _nav(ws, wb.sheetnames[:9])
    ws.column_dimensions["B"].width = 34
    for c in "CDEFGH":
        ws.column_dimensions[c].width = 18
    for i, (k, v) in enumerate(policy_rows(policy, run_info), 5):
        ws.cell(i, 2, k).font = Font(bold=True)
        ws.merge_cells(start_row=i, start_column=3, end_row=i, end_column=8)
        ws.cell(i, 3, v if not isinstance(v, bool) else ("Yes" if v else "No")).alignment = Alignment(wrap_text=True, vertical="top")
        if k == "정책 JSON":
            ws.row_dimensions[i].height = 75
    _print_setup(ws)
    return ws


def _add_checked_table(ws, ref, name, style_name):
    """
    모든 Excel Table 생성 경로에서 공통으로 쓰는 검증 함수.
    """

    if ":" not in ref or ref.split(":")[0] == ref.split(":")[1]:
        return

    # worksheet AutoFilter와 Excel Table은 동시에 만들지 않는다.
    ws.auto_filter.ref = None

    # 같은 시트의 표 이름·범위 중복을 차단한다.
    for existing_name, existing_table in ws.tables.items():
        if existing_name == name:
            raise RuntimeError(
                f"중복 Excel Table 이름: "
                f"sheet={ws.title}, name={name}, ref={existing_table.ref}"
            )

        if excel_ranges_overlap(existing_table.ref, ref):
            raise RuntimeError(
                f"Excel Table 범위 중복: "
                f"sheet={ws.title}, new={name}:{ref}, "
                f"existing={existing_name}:{existing_table.ref}"
            )

    # Excel table displayName은 workbook 범위에서 고유해야 한다.
    for other_ws in ws.parent.worksheets:
        if other_ws is ws:
            continue

        for existing_name, existing_table in other_ws.tables.items():
            if existing_name == name:
                raise RuntimeError(
                    f"통합문서 전체 Table 이름 중복: "
                    f"new={ws.title}!{name}, "
                    f"existing={other_ws.title}!{existing_name}"
                )

    table = Table(displayName=name, ref=ref)
    table.tableStyleInfo = TableStyleInfo(
        name=style_name,
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )

    ws.add_table(table)

def _write_table_sheet(ws, column_headers, rows, start_row, name, fmts=None):
    _headers(ws, start_row, column_headers, 2)

    for row_idx, record in enumerate(rows, start_row + 1):
        for col_idx, header in enumerate(column_headers, 2):
            value = record.get(header)

            put(
                ws,
                row_idx,
                col_idx,
                (
                    None
                    if isinstance(value, (float, np.floating))
                    and not np.isfinite(value)
                    else value
                ),
                (fmts or {}).get(header),
            )

    if not rows:
        ws.auto_filter.ref = None
        return

    ref = (
        f"B{start_row}:"
        f"{get_column_letter(len(column_headers) + 1)}"
        f"{start_row + len(rows)}"
    )

    _add_checked_table(
        ws,
        ref,
        name,
        "TableStyleLight9",
    )

def lot_inclusion_decision(*, metric_code, value, channel, data_status, value_class, acceptance_status,
                          source_type, pair_id, policy):
    """통계 포함과 합격성 제외는 따로 결정한다. acceptance EXCLUDED여도 정책이 허용하면 평균에 넣는다."""
    policy = policy or active_policy()
    del acceptance_status, pair_id
    if not _finite(value):
        return {"included": False, "reason": "non-finite", "stat_value": np.nan}
    source = source_type or "measured"
    if source in ("interpolated", "model-derived", "estimated") or value_class == VC_ESTIMATED:
        return {"included": False, "reason": "estimated", "stat_value": np.nan}
    if value_class in NON_NUMERIC_CLASSES or data_status in ("INVALID", "NO VALUE") or value_class in (VC_LEGACY, "LEGACY_UNKNOWN"):
        return {"included": False, "reason": data_status or value_class or "non-numeric", "stat_value": np.nan}
    if metric_code == "A_SSI_LEGACY_UNKNOWN":
        return {"included": False, "reason": "legacy unknown", "stat_value": np.nan}
    if channel == "strict":
        if value_class == VC_STRICT and data_status not in ("INVALID", "NO VALUE"):
            return {"included": True, "reason": "", "stat_value": float(value)}
        return {"included": False, "reason": value_class or data_status or "strict excludes non-quantitative", "stat_value": np.nan}
    if value_class == VC_STRICT:
        return {"included": True, "reason": "", "stat_value": float(value)}
    if value_class == VC_EXC:
        if policy.exception_stat_policy in ("separate", "include") or policy.include_exception_in_lot_stats:
            return {"included": True, "reason": "", "stat_value": float(value)}
        return {"included": False, "reason": "exception excluded by policy", "stat_value": np.nan}
    if value_class in TORQUE_REFERENCE_CLASSES:
        if policy.include_censored_in_lot_stats:
            return {"included": True, "reason": "torque reference included; acceptance excluded", "stat_value": float(value)}
        return {"included": False, "reason": "censored excluded", "stat_value": np.nan}
    return {"included": False, "reason": str(value_class or data_status or "excluded"), "stat_value": np.nan}


def _mean_eligible(st, policy, channel="strict", value=None):
    """Lot 평균 포함. lot_inclusion_decision과 같은 결과만 쓴다."""
    if st is None:
        return False
    if value is None:
        value = st.get("value", st.get("stat_value"))
    return lot_inclusion_decision(
        metric_code=st.get("metric_code") or "", value=value, channel=channel,
        data_status=st.get("data_status") or st.get("status"),
        value_class=st.get("value_class") or _default_value_class(st.get("status")),
        acceptance_status=st.get("acceptance_status"),
        source_type=st.get("source_type") or "measured", pair_id=st.get("pair_id"), policy=policy,
    )["included"]


def format_value_remark(channel_label, stats, policy):
    """Summary·LotCompare 비고. 채널, 값 구분, 실제 Step/반복, Torque 기준, 합격성 제외를 남긴다."""
    vc = stats.get("value_class") or VC_NO
    tq = f"Torque {policy.torque_min_pct:g}–{policy.torque_max_pct:g}%"
    bits = [channel_label]
    if stats.get("fit_model"):
        bits.append(f"{stats['fit_model']} H-B")
    bits.append(vc)
    if vc == VC_STRICT:
        bits.append(f"{policy.hold_target_s:g} s hold, {tq} 범위 내")
    if stats.get("contrib_steps"):
        bits.append("Steps " + stats["contrib_steps"])
    if stats.get("apply_rpm"):
        bits.append("RPM " + stats["apply_rpm"])
    if stats.get("apply_gamma"):
        bits.append("γ̇ " + stats["apply_gamma"])
    if stats.get("selection_rule"):
        bits.append(stats["selection_rule"])
    if stats.get("exc_details"):
        bits.append("accepted hold 사용: " + "; ".join(stats["exc_details"]))
    if stats.get("low_details"):
        bits.append(f"Torque <{policy.torque_min_pct:g}% " + "; ".join(stats["low_details"]) + " 사용")
    if stats.get("high_details"):
        bits.append(f"Torque >{policy.torque_max_pct:g}% " + "; ".join(stats["high_details"]) + " 사용")
    if vc == VC_BOTH:
        bits.append("torque 하한 밖 참고값 및 torque 상한 밖 참고값 포함")
    if vc in (VC_LOW, VC_HIGH, VC_BOTH):
        if stats.get("exc_details"):
            bits.append("inclusive exception 포함")
        bits.append("통계에는 포함하되 합격성 판정 제외")
        if stats.get("rsd_text"):
            bits.append("RSD는 참고 통계")
    elif vc == VC_EXC:
        bits.append("합격성 자동 ELIGIBLE 아님")
    if stats.get("reasons"):
        bits.append(stats["reasons"])
    return " | ".join(bits)


def write_legacy_4rpm_sheet(wb, a_recs):
    """4 rpm 진단 시트. 기본 실행에서는 만들지 않는다."""
    if not INCLUDE_LEGACY_4RPM_DIAGNOSTICS:
        return None
    ws = wb.create_sheet("Legacy_4RPM_Diagnostics")
    _sheet_base(ws, "Legacy_4RPM_Diagnostics — 4 rpm legacy diagnostic only",
                "legacy diagnostic only. not a high-shear endpoint. not used for representative SSI. "
                "not used for acceptance. not shown in primary Summary. H_REF는 이 진단 시트에서만 쓴다.", 8)
    headers = ["repeat", "eta_up(0.5)/eta_up(4.0)", "tau_up(4.0)", "eta_up(4.0)", "note"]
    _headers(ws, 5, headers, 2)
    for i, rec in enumerate(a_recs, 6):
        leg = rec.get("legacy_4rpm") or {}
        vals = [rec.get("repeat"), leg.get("ssi_0p5_4"), leg.get("tau_up_4"), leg.get("eta_up_4"),
                leg.get("note") or "legacy diagnostic only; not a high-shear endpoint"]
        for j, value in enumerate(vals, 2):
            if isinstance(value, float) and not _finite(value):
                value = None
            ws.cell(i, j, value)
    ws.freeze_panes = "B6"
    return ws


def write_decade_ssi_sheet(wb, a_recs):
    """발견된 실측 decade pair. 0 rpm pair는 만들지 않고, 공식 code와 행을 이중으로 만들지 않는다."""
    if "DecadeSSI" in wb.sheetnames:
        del wb["DecadeSSI"]
    ws = wb.create_sheet("DecadeSSI")
    headers = ["Lot", "Repeat", "Channel", "Direction", "Pair ID", "Official metric link",
               "Low RPM", "High RPM", "RPM ratio", "Ratio error %", "Low Step", "High Step",
               "Low shear rate", "High shear rate", "Low viscosity", "High viscosity", "SSI",
               "Source type", "Value class", "Data status", "Acceptance status",
               "Included in Lot stats", "Lot-stat exclusion reason", "Primary pair selected", "Selection reason", "Reason"]
    rows = []
    for rec in a_recs or []:
        for channel, key in (("strict", "decade_pairs_strict"), ("inclusive", "decade_pairs_incl")):
            for pair in rec.get(key) or []:
                if abs(float(pair.get("low_rpm") or 0)) <= RPM_ZERO_TOL:
                    continue
                rows.append({"Lot": rec.get("lot"), "Repeat": rec.get("repeat"), "Channel": channel,
                             "Direction": pair.get("direction"), "Pair ID": pair.get("pair_id"),
                             "Official metric link": pair.get("official_metric_link") or "",
                             "Low RPM": pair.get("low_rpm"), "High RPM": pair.get("high_rpm"),
                             "RPM ratio": pair.get("actual_ratio"), "Ratio error %": pair.get("ratio_error_pct"),
                             "Low Step": pair.get("low_step"), "High Step": pair.get("high_step"),
                             "Low shear rate": pair.get("low_gamma"), "High shear rate": pair.get("high_gamma"),
                             "Low viscosity": pair.get("low_eta_cP"), "High viscosity": pair.get("high_eta_cP"),
                             "SSI": pair.get("ssi"), "Source type": pair.get("source_type"),
                             "Value class": pair.get("value_class"), "Data status": pair.get("data_status"),
                             "Acceptance status": pair.get("acceptance_status"),
                             "Included in Lot stats": bool(pair.get("included_in_lot_stats")),
                             "Lot-stat exclusion reason": pair.get("lot_stat_exclusion_reason") or "",
                             "Primary pair selected": bool(pair.get("primary_pair_selected")),
                             "Selection reason": pair.get("selection_reason") or "",
                             "Reason": pair.get("reason") or ""})
    _sheet_base(ws, "DecadeSSI — 실측 10배 RPM 점도비",
                "pair_id마다 독립 행이다. 공식 SSI와 같은 숫자를 여기 평균에 다시 넣지 않는다.", 8)
    _write_table_sheet(ws, headers, rows, 5, "DecadeSSI")
    ws.freeze_panes = "B6"
    return ws


def write_lot_qc_workbook(lot, a_recs, b_recs, d_recs, outdir, route="STANDARD", run_meta=None, defs=None,
                          suffix="", policy=None):
    """Lot 종합 QC 워크북. 통계는 Replicates(strict/INCL/EXC) → Metric_Statistics → Lot_QC_Summary 수식 연쇄.

    06: STRICT/INCLUSIVE 채널 병렬, EXCEPTION 값은 A_Replicates_EXC로 분리(exception_stat_policy=separate),
    StepAudit·QCPolicy 시트 추가, 파일명에 정책 태그.
    """
    policy = policy or active_policy()
    route = policy.name if route in ("STANDARD", None) or policy.name != "STANDARD" else route
    defs = _export_defs(defs or metric_definitions(), list(a_recs) + list(b_recs) + list(d_recs))
    run_meta = run_meta or {}
    recs_by_method = {"A": sorted(a_recs, key=lambda r: str(r["repeat"])),
                      "B": sorted(b_recs, key=lambda r: str(r["repeat"])),
                      "D": sorted(d_recs, key=lambda r: str(r["repeat"]))}
    sheet_by_method = {"A": "A_Replicates", "B": "B_Replicates", "D": "Density_Replicates"}
    wb = Workbook()
    sm = wb.active
    sm.title = "Lot_QC_Summary"
    dic = wb.create_sheet("QC_Dictionary")
    rep_ws = {m: wb.create_sheet(sheet_by_method[m]) for m in ("A", "B", "D")}
    rep_incl = wb.create_sheet("A_Replicates_INCL")
    rep_exc = wb.create_sheet("A_Replicates_EXC")
    stat = wb.create_sheet("Metric_Statistics")
    srep = wb.create_sheet("A_Step_Replicates")
    sstat = wb.create_sheet("A_Step_Statistics")
    flags_ws = wb.create_sheet("Replicate_Flags")
    trace = wb.create_sheet("Traceability")
    names = wb.sheetnames
    exc_text = (f"POLICY {policy.name} — accepted holds {','.join(f'{x:g}' for x in policy.accepted_hold_s)} s, 범위 밖 Torque 점 inclusive 포함. "
                "EXCEPTION·CENSORED 값은 표준 60 s Lot 통계·추세·규격 설정에 혼용 금지")
    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False
        ws.column_dimensions["A"].width = 3

    # ---------------------------------------------------------------- Replicates
    base_cols = ["lot", "repeat", "file", "test_start", "route", "data_qc"]
    col_letter = {}
    for m, ws in rep_ws.items():
        codes = [d["code"] for d in defs if d["method"] == m and d.get("output_policy") not in ("decade_sheet", "dictionary_only", "legacy")]
        # Method A 기본 Replicates는 STRICT 채널 원값이다. Inclusive 원값은 별도 감사 시트에 둔다.
        rep_title = (f"{sheet_by_method[m]} — STRICT 채널 반복별 QC 지표 원값" if m == "A"
                     else f"{sheet_by_method[m]} — 반복별 QC 지표 원값")
        rep_sub = ("STRICT: 60 s hold·Torque 10–95% 기준 값입니다. Inclusive 값은 A_Replicates_INCL, 예외 사용 내역은 A_Replicates_EXC·Replicate_Flags 참조."
                   if m == "A" else
                   "Python 공식 계산값(숫자형). 공란 = 유효 조건 미충족(사유는 Replicate_Flags). 이 시트 값을 수정하면 통계가 자동 재계산됩니다.")
        _sheet_base(ws, rep_title, rep_sub, min(len(base_cols) + len(codes) + 1, 16))
        _nav(ws, names)
        extra = ["temp_control"] if m != "D" else ["density_method", "operator", "mass_empty_g", "mass_filled_g",
                                                    "volume_mL", "density_direct", "density_calc", "note"]
        prov = list(PROVENANCE_COLUMNS) if m == "A" else []
        helpers = [name for c in codes for name in (
            f"{c}__include_in_lot_stats", f"{c}__stat_value", f"{c}__lot_exclusion_reason")]
        headers = base_cols + codes + helpers + prov + extra
        rows = []
        for r in recs_by_method[m]:
            row = {k: r.get(k) for k in base_cols}
            for c in codes:
                st = r["status"].get(c) or _st("NO VALUE")
                raw = r["values"].get(c, np.nan)
                decision = lot_inclusion_decision(
                    metric_code=c, value=raw, channel="strict" if m == "A" else "strict",
                    data_status=st.get("data_status") or st.get("status"),
                    value_class=st.get("value_class") or _default_value_class(st.get("status")),
                    acceptance_status=st.get("acceptance_status"),
                    source_type=st.get("source_type") or "measured", pair_id=st.get("pair_id"), policy=policy,
                )
                row[c] = raw if _finite(raw) else np.nan
                row[f"{c}__stat_value"] = decision["stat_value"]
                row[f"{c}__include_in_lot_stats"] = bool(decision["included"])
                row[f"{c}__lot_exclusion_reason"] = decision["reason"]
            if m == "A":
                row.update(provenance_from_status(r.get("status") or {}))
            if m != "D":
                row["temp_control"] = r.get("temp_control", "")
            else:
                dm = r.get("density_meta", {})
                row.update({"density_method": dm.get("method"), "operator": dm.get("operator"),
                            "mass_empty_g": dm.get("mass_empty_g"), "mass_filled_g": dm.get("mass_filled_g"),
                            "volume_mL": dm.get("volume_mL"), "density_direct": dm.get("density_direct"),
                            "density_calc": dm.get("density_calc"), "note": dm.get("note")})
            rows.append(row)
        dd = _defs_by_code(defs)
        fmts = {c: _fmt_for(dd[c]["unit"]) for c in codes}
        _write_table_sheet(ws, headers, rows, 5, f"{m}_Rep", fmts)
        for j, h in enumerate(headers, 2):
            ws.column_dimensions[get_column_letter(j)].width = max(11, min(len(h) + 3, 24))
            if h.endswith("__stat_value"):
                col_letter[h[: -len("__stat_value")]] = (sheet_by_method[m], get_column_letter(j), 6, 5 + len(rows))
            if "__" in h:
                ws.column_dimensions[get_column_letter(j)].hidden = True
        ws.freeze_panes = "E6"
        if policy.name != "STANDARD":
            _banner(ws, exc_text, min(len(headers) + 1, 16))

    # ---------------------------------------------------------------- A_Replicates_INCL / _EXC (06)
    a_codes = [d["code"] for d in defs if d["method"] == "A" and d.get("output_policy") not in ("decade_sheet", "dictionary_only", "legacy")]
    col_letter_incl, col_letter_exc, col_class_incl = {}, {}, {}
    for ws, key, title, sub, book in (
            (rep_incl, "incl", "A_Replicates_INCL — INCLUSIVE 채널 전체 반복값(감사용)",
             "정책 허용 hold·범위 밖 참고값을 포함한 inclusive 값입니다. CENSORED여도 셀을 비우지 않고, 오른쪽 값 구분 열에 출처를 남깁니다. "
             "Metric_Statistics의 inclusive 평균은 정책이 허용한 값 구분만 집계합니다.", col_letter_incl),
            (rep_exc, "exc", "A_Replicates_EXC — INCLUSIVE 결과 중 EXCEPTION 값 감사 필터",
             "30/45 s 등 승인 예외 hold Step을 사용한 inclusive 값만 표시합니다. 이 시트는 제외 목록이 아니라 사용 내역 확인용입니다.", col_letter_exc)):
        class_headers = [f"{c}_class" for c in a_codes] if key == "incl" else []
        helpers = [name for c in a_codes for name in (
            f"{c}__include_in_lot_stats", f"{c}__stat_value", f"{c}__lot_exclusion_reason")] if key == "incl" else []
        prov = list(PROVENANCE_COLUMNS) if key == "incl" else []
        headers = base_cols + a_codes + helpers + prov + class_headers
        _sheet_base(ws, title, sub, min(len(headers) + 1, 16))
        _nav(ws, names)
        rows = []
        for r in recs_by_method["A"]:
            vi, si = r.get("values_incl", r["values"]), r.get("status_incl", r["status"])
            row = {k: r.get(k) for k in base_cols}
            for c in a_codes:
                st = si.get(c)
                v = vi.get(c, np.nan)
                if key == "incl":
                    decision = lot_inclusion_decision(
                        metric_code=c, value=v, channel="inclusive",
                        data_status=(st.get("data_status") or st.get("status")) if st else "NO VALUE",
                        value_class=(st.get("value_class") if st else None) or (_default_value_class(st["status"]) if st else VC_NO),
                        acceptance_status=st.get("acceptance_status") if st else None,
                        source_type=(st.get("source_type") if st else None) or "measured",
                        pair_id=st.get("pair_id") if st else None, policy=policy,
                    )
                    row[c] = v if _finite(v) else np.nan
                    row[f"{c}__stat_value"] = decision["stat_value"]
                    row[f"{c}__include_in_lot_stats"] = bool(decision["included"])
                    row[f"{c}__lot_exclusion_reason"] = decision["reason"]
                    row[f"{c}_class"] = (st.get("value_class") if st else "") or (_default_value_class(st["status"]) if st else VC_NO)
                else:
                    row[c] = v if (st and st["status"] == "EXCEPTION" and _finite(v)) else np.nan
            if key == "incl":
                row.update(provenance_from_status(si))
            rows.append(row)
        dd = _defs_by_code(defs)
        _write_table_sheet(ws, headers, rows, 5, f"A_Rep_{key.upper()}", {c: _fmt_for(dd[c]["unit"]) for c in a_codes})
        for j, h in enumerate(headers, 2):
            ws.column_dimensions[get_column_letter(j)].width = max(11, min(len(str(h)) + 2, 26))
            if h.endswith("__stat_value"):
                book[h[: -len("__stat_value")]] = ("A_Replicates_INCL" if key == "incl" else "A_Replicates_EXC", get_column_letter(j), 6, 5 + len(rows))
            elif h in a_codes and key != "incl":
                book[h] = ("A_Replicates_EXC", get_column_letter(j), 6, 5 + len(rows))
            if "__" in str(h):
                ws.column_dimensions[get_column_letter(j)].hidden = True
            if key == "incl" and h.endswith("_class"):
                col_class_incl[h[:-6]] = ("A_Replicates_INCL", get_column_letter(j), 6, 5 + len(rows))
        if key == "incl" and rows:
            for cname, (_sh, letter, r1, r2) in col_class_incl.items():
                _value_class_cf(ws, f"{letter}{r1}:{letter}{r2}")
        ws.freeze_panes = "E6"
        if policy.name != "STANDARD":
            _banner(ws, exc_text, min(len(headers) + 1, 16))

    # ---------------------------------------------------------------- Metric_Statistics (수식)
    _sheet_base(stat, "Metric_Statistics — STRICT/INCLUSIVE 감사용 반복 통계(수식)",
                "감사·method 검토용 시트입니다. 좌측은 STRICT(60 s·Torque 10–95%) 통계, 우측은 INCLUSIVE 통계와 차이입니다. "
                "일상 판독은 Lot_QC_Summary의 대표 채널 값을 사용합니다. "
                f"RSD는 0 기준 차이형 지표(zero_ref)에서 '{RSD_NA_ZERO_REF}', |Mean|<SD이면 '{RSD_NA_NEAR_ZERO}'로 표시합니다. "
                "Inclusive 평균은 정책이 허용한 값 구분만 집계하고, n Low/High-TQ Ref는 참고값 개수입니다.", 26)
    _nav(stat, names)
    # 열 제목에 채널을 명시해 primary/strict/inclusive 값을 혼동하지 않게 한다.
    _headers(stat, 5, ["Code", "Method", "Tier", "n Strict", "Strict Mean", "Strict SD", "Strict RSD(%)", "Strict Min", "Strict Max",
                       "Strict Range", "Strict |Mean|", "Strict n(+)", "Strict n(−)", "Strict 원본 범위",
                       "n Inclusive", "Inclusive Mean", "Inclusive SD", "Δ Mean (Inclusive−Strict)", "n EXC", "EXC Mean",
                       "n Low-TQ Ref", "n High-TQ Ref", "Inclusive Min", "Inclusive Max", "Inclusive Range"], 2)
    stat_row = {}
    ordered = [d for d in _ordered_defs(defs) if d.get("output_policy") not in ("decade_sheet", "dictionary_only", "legacy")]
    for i, d in enumerate(ordered, 6):
        code = d["code"]
        stat_row[code] = i
        stat.cell(i, 2, code)
        stat.cell(i, 3, d["method"])
        stat.cell(i, 4, d["tier"])
        sheet, col, r1, r2 = col_letter.get(code, (None, None, 0, 0))
        if sheet and r2 >= r1:
            rng = f"'{sheet}'!{col}{r1}:{col}{r2}"
            stat.cell(i, 5, f"=COUNT({rng})")
            stat.cell(i, 6, f'=IF(E{i}=0,"",AVERAGE({rng}))')
            stat.cell(i, 7, f'=IF(E{i}<2,"",STDEV({rng}))')
            if d.get("zero_ref"):
                stat.cell(i, 8, f'=IF(E{i}<2,"","{RSD_NA_ZERO_REF}")')
            else:
                stat.cell(i, 8, f'=IF(E{i}<2,"",IF(G{i}=0,0,IF(ABS(F{i})<G{i},"{RSD_NA_NEAR_ZERO}",G{i}/ABS(F{i})*100)))')
            stat.cell(i, 9, f'=IF(E{i}=0,"",MIN({rng}))')
            stat.cell(i, 10, f'=IF(E{i}=0,"",MAX({rng}))')
            stat.cell(i, 11, f'=IF(E{i}=0,"",J{i}-I{i})')
            stat.cell(i, 12, f'=IF(E{i}=0,"",SUMPRODUCT(ABS({rng}))/E{i})')
            stat.cell(i, 13, f'=IF(E{i}=0,"",COUNTIF({rng},">0"))')
            stat.cell(i, 14, f'=IF(E{i}=0,"",COUNTIF({rng},"<0"))')
            stat.cell(i, 15, rng)
        else:
            stat.cell(i, 5, 0)
            for c in range(6, 15):
                stat.cell(i, c, '=""')
            stat.cell(i, 15, "반복 데이터 없음")
        # 06: INCLUSIVE / EXCEPTION 채널(Method A만)
        isheet, icol, ir1, ir2 = col_letter_incl.get(code, (None, None, 0, 0))
        esheet, ecol, er1, er2 = col_letter_exc.get(code, (None, None, 0, 0))
        if isheet and ir2 >= ir1:
            irng = f"'{isheet}'!{icol}{ir1}:{icol}{ir2}"
            cinfo = col_class_incl.get(code)
            if cinfo:
                _csh, cletter, cr1, cr2 = cinfo
                crng = f"'{_csh}'!{cletter}{cr1}:{cletter}{cr2}"
                stat.cell(i, 16, f"=COUNT({irng})")
                stat.cell(i, 17, f'=IF(P{i}=0,"",AVERAGE({irng}))')
                stat.cell(i, 18, f'=IF(P{i}<2,"",STDEV({irng}))')
                stat.cell(i, 22, f'=COUNTIF({crng},"{VC_LOW}")+COUNTIF({crng},"{VC_BOTH}")')
                stat.cell(i, 23, f'=COUNTIF({crng},"{VC_HIGH}")+COUNTIF({crng},"{VC_BOTH}")')
                stat.cell(i, 24, f'=IF(P{i}=0,"",MIN({irng}))')
                stat.cell(i, 25, f'=IF(P{i}=0,"",MAX({irng}))')
                stat.cell(i, 26, f'=IF(P{i}=0,"",Y{i}-X{i})')
            else:
                stat.cell(i, 16, f"=COUNT({irng})")
                stat.cell(i, 17, f'=IF(P{i}=0,"",AVERAGE({irng}))')
                stat.cell(i, 18, f'=IF(P{i}<2,"",STDEV({irng}))')
                stat.cell(i, 22, 0)
                stat.cell(i, 23, 0)
            stat.cell(i, 19, f'=IF(OR(P{i}=0,E{i}=0),"",Q{i}-F{i})')
        else:
            stat.cell(i, 16, 0)
            for c in (17, 18, 19, 22, 23):
                stat.cell(i, c, '=""')
        if esheet and er2 >= er1:
            erng = f"'{esheet}'!{ecol}{er1}:{ecol}{er2}"
            stat.cell(i, 20, f"=COUNT({erng})")
            stat.cell(i, 21, f'=IF(T{i}=0,"",AVERAGE({erng}))')
        else:
            stat.cell(i, 20, 0)
            stat.cell(i, 21, '=""')
        for c in range(6, 13):
            stat.cell(i, c).number_format = _fmt_for(d["unit"]) if c != 8 else "0.00"
        for c in (13, 14, 16, 20, 22, 23):
            stat.cell(i, c).number_format = "0"
        for c in (17, 18, 19, 21):
            stat.cell(i, c).number_format = _fmt_for(d["unit"])
    for col, w in {2: 17, 3: 8, 4: 6, 5: 8, 6: 13, 7: 13, 8: 15, 9: 13, 10: 13, 11: 11, 12: 11, 13: 6, 14: 6, 15: 34,
                   16: 12, 17: 14, 18: 13, 19: 16, 20: 7, 21: 13, 22: 14, 23: 14}.items():
        stat.column_dimensions[get_column_letter(col)].width = w
    stat.freeze_panes = "C6"

    # ---------------------------------------------------------------- Summary
    n_cols = 45
    _sheet_base(sm, f"HT-{lot} Lot 종합 QC — 정책 {policy.name}",
                "Method A fingerprint·구조 이력, Method B 회복, 밀도를 Lot 단위로 통합했습니다. "
                "Strict = 60 s·Torque 10–95 % / Inclusive = 정책 허용 hold·범위 밖 점 포함. "
                "상태는 데이터 품질 판정이며 제품 적합성(합격/불합격) 판정이 아닙니다.", n_cols)
    _nav(sm, names)
    if policy.name != "STANDARD":
        _banner(sm, exc_text, n_cols)
    _section(sm, 5, "Lot 정보 및 데이터 범위", 2, n_cols)
    info = [
        ("Lot", lot),
        ("분석 정책", f"{policy.name} (hash {policy.digest()}) | Lot 통계 모드 {policy.lot_stats_mode} | 대표 채널 {policy.primary_channel} | EXCEPTION {policy.exception_stat_policy} | 상세 QCPolicy 시트"),
        ("Method A 반복", f"{len(recs_by_method['A'])}건 — " + ", ".join(f"{r['repeat']}:{r['data_qc']}" for r in recs_by_method["A"]) if recs_by_method["A"] else "없음"),
        ("Method B 반복", f"{len(recs_by_method['B'])}건 — " + ", ".join(f"{r['repeat']}:{r['data_qc']}" for r in recs_by_method["B"]) if recs_by_method["B"] else "없음 — T2 회복 지표 NO DATA"),
        ("밀도 충전", f"{len(recs_by_method['D'])}건 — " + ", ".join(f"{r['repeat']}:{r['data_qc']}" for r in recs_by_method["D"]) if recs_by_method["D"] else "없음 — 밀도 입력 파일(--density-file) 미제공"),
        ("고정 평가점", f"Low={FIXED_LOW_RPM:g} rpm, Mid={FIXED_MID_RPM:g} rpm"),
        ("고전단 평가점", "각 반복의 실제 측정 최고 RPM. 그 Step이 채널에서 부적격이면 NO VALUE"),
        ("현재 Lot의 측정 최고 RPM", _lot_measured_max_text(recs_by_method["A"], policy)),
        ("4 rpm", "중간 측정점; 대표 고전단 기준 아님"),
        ("분석기", f"{VERSION} / {QC_LAYER_VERSION} | 실행 {run_meta.get('timestamp', '')}"),
    ]
    for rec in recs_by_method["A"]:
        stab = rec.get("stabilization") or {}
        if not stab:
            continue
        info.append(("무전단 안정화",
                     f"존재={stab.get('stabilization_present')} | Step S{stab.get('stabilization_step')} | 목표 RPM 0 | "
                     f"목표 온도 {stab.get('stabilization_temperature_target_C')} °C | "
                     f"목표 {stab.get('stabilization_target_s')} s | 실제 {stab.get('stabilization_observed_s')} s | "
                     f"평균 온도 {stab.get('stabilization_temperature_mean_C')} °C | "
                     f"min {stab.get('stabilization_temperature_min_C')} / max {stab.get('stabilization_temperature_max_C')} | "
                     f"sampling {stab.get('stabilization_sampling_interval_median_s')} s | "
                     f"상태 {stab.get('stabilization_status')} | repeat {rec.get('repeat')} | {stab.get('stabilization_reason')}"))
    for i, (k, v) in enumerate(info, 6):
        sm.cell(i, 2, k).font = Font(bold=True)
        sm.cell(i, 2).alignment = Alignment(vertical="top")
        sm.merge_cells(start_row=i, start_column=3, end_row=i, end_column=n_cols)
        sm.cell(i, 3, v).alignment = Alignment(wrap_text=True, vertical="top")
    r0 = 6 + len(info) + 1
    _section(sm, r0, "지표 층위 — 무엇을 말하고 무엇을 말하지 않는가", 2, n_cols)
    for i, t in enumerate(TIER_ORDER, r0 + 1):
        sm.cell(i, 2, TIER_LABELS[t]).font = Font(bold=True)
        sm.cell(i, 2).alignment = Alignment(vertical="top")
        sm.merge_cells(start_row=i, start_column=3, end_row=i, end_column=n_cols)
        sm.cell(i, 3, TIER_NOTES[t]).alignment = Alignment(wrap_text=True, vertical="top")
        sm.row_dimensions[i].height = 30
    h = r0 + len(TIER_ORDER) + 2
    _section(sm, h - 1, "지표별 Lot 통계 (숫자 열은 Metric_Statistics 수식 참조; 상태·사유는 Python 판정)", 2, n_cols)
    # ---------------------------------------------------------------- Summary
    # 화면용 Summary는 대표 채널 하나의 통계를 기존 통계 열에 표시한다.
    # Strict/Inclusive 두 채널의 원자료와 상세 통계는 Metric_Statistics, AReplicatesINCL,
    # AReplicatesEXC, StepAudit에 보존하므로 Summary에서 병렬 열을 반복하지 않는다.
    headers = [
        "층위", "Code", "Role", "지표", "가리키는 물성", "단위",
        "유효 n / 전체", "Mean", "SD", "RSD(%)", "Min", "Max", "Range", "Mean |SD|", "부호 +/−",
        "대표 채널", "n EXC", "n CENS", "n Low-TQ Ref", "n High-TQ Ref",
        "반복성 규칙(임시)", "GOOD ≤", "REVIEW ≤", "반복성 상태",
        "적용 RPM", "적용 shear rate", "기여 Step",
        "적용 저점 RPM", "적용 고점 RPM", "적용 저점 Step", "적용 고점 Step",
        "적용 저점 shear rate", "적용 고점 shear rate", "선택 규칙",
        "분자 RPM", "분모 RPM", "분자 Step", "분모 Step", "분자 shear rate", "분모 shear rate",
        "실제 RPM 배율", "source type",
        "Included in Lot stats", "Lot-stat exclusion reason",
        "값 상태", "데이터 상태", "합격성 반영", "비고·제외 근거(반복)",
    ]
    _headers(sm, h, headers, 2)
    n_cols = len(headers) + 1
    width_by_name = {
        "층위": 17, "Code": 16, "Role": 12, "지표": 28, "가리키는 물성": 40, "단위": 6,
        "유효 n / 전체": 12, "Mean": 12, "SD": 11, "RSD(%)": 12, "Min": 11, "Max": 11, "Range": 11,
        "Mean |SD|": 12, "부호 +/−": 10, "대표 채널": 12, "n EXC": 8, "n CENS": 8,
        "n Low-TQ Ref": 14, "n High-TQ Ref": 14, "반복성 규칙(임시)": 14, "GOOD ≤": 8, "REVIEW ≤": 8,
        "반복성 상태": 12, "적용 RPM": 28, "적용 shear rate": 18, "기여 Step": 28,
        "적용 저점 RPM": 14, "적용 고점 RPM": 14, "적용 저점 Step": 14, "적용 고점 Step": 16,
        "적용 저점 shear rate": 16, "적용 고점 shear rate": 16, "선택 규칙": 42,
        "값 상태": 24, "데이터 상태": 12, "합격성 반영": 12, "비고·제외 근거(반복)": 72,
    }
    widths = {i: width_by_name.get(name, 12) for i, name in enumerate(headers, 2)}
    for col, w in widths.items():
        sm.column_dimensions[get_column_letter(col)].width = w
    lot_stats = {}
    row = h + 1
    for d in ordered:
        code = d["code"]
        recs = recs_by_method[d["method"]]
        s = lot_metric_stats_both(recs, code, d, policy)
        lot_stats[code] = s
        rule, good, rev = d["prec"]

        # primary_channel=inclusive는 Method A에만 적용한다. Method B와 Density에는 inclusive 채널이 없다.
        use_inclusive_primary = d["method"] == "A" and policy.primary_channel == "inclusive"
        primary = s["inclusive"] if use_inclusive_primary else s["strict"]
        primary_label = "INCLUSIVE" if use_inclusive_primary else "STRICT"

        # 원본 lot_metric_stats()의 실제 snake_case 반환 key만 사용한다.
        # n EXC는 inclusive EXCEPTION 반복 수, n CENS는 대표 채널의 CENSORED 반복 수다.
        n_exc = s["inclusive"]["n_exception"] if d["method"] == "A" else 0
        n_cens = primary["n_censored"]
        n_low = primary.get("n_low_ref", 0) if d["method"] == "A" else 0
        n_high = primary.get("n_high_ref", 0) if d["method"] == "A" else 0
        remarks = format_value_remark(primary_label, primary, policy)
        other = s["strict"] if use_inclusive_primary else s["inclusive"]
        if d["method"] == "A" and _finite(primary.get("mean")) and _finite(other.get("mean")) and abs(float(primary["mean"]) - float(other["mean"])) > 1e-9:
            remarks += f" | strict mean {s['strict']['mean']:.4g} / inclusive mean {s['inclusive']['mean']:.4g}"

        # primary 채널의 Mean~Range를 기존 주 통계 열에 직접 기록한다.
        # rsd_text는 원본이 제공하는 표시용 문자열(N/A(0 기준), N/A(Mean≈0) 포함)이다.
        vals = [
            TIER_LABELS[d["tier"]], code, d["metric_role"], d.get("summary_name") or d["name"], d["property"], d["unit"],
            f'{primary["n"]} / {primary["n_total"]}',
            primary["mean"], primary["sd"], primary["rsd_text"],
            primary["min"], primary["max"], primary["range"], primary["abs_mean"], primary["sign_text"],
            primary_label, n_exc, n_cens, n_low, n_high, rule, good, rev, None,
            primary.get("apply_rpm") or "", primary.get("apply_gamma") or "", primary.get("contrib_steps") or "",
            _rpm_cell(primary.get("low_rpm")), _rpm_cell(primary.get("high_rpm")),
            primary.get("low_step") or "", primary.get("high_step") or "",
            _gamma_cell(primary.get("low_gamma")), _gamma_cell(primary.get("high_gamma")),
            primary.get("selection_rule") or "",
            primary.get("numerator_rpm") or "", primary.get("denominator_rpm") or "",
            primary.get("numerator_step") or "", primary.get("denominator_step") or "",
            primary.get("numerator_gamma") or "", primary.get("denominator_gamma") or "",
            primary.get("actual_ratio") or "", primary.get("source_type") or "",
            "Yes" if primary.get("included_in_lot_stats") else "No",
            primary.get("lot_stat_exclusion_reason") or "",
            primary.get("value_class") or VC_NO, primary["data_status"], primary["acceptance_status"], remarks,
        ]
        for j, v in enumerate(vals, 2):
            c = sm.cell(row, j, v)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            c.border = Border(bottom=THIN)

        # RSD는 rsd_text로 기록하므로 텍스트/숫자 모두 표시 가능하다. 다른 통계 열은 단위별 숫자 형식을 적용한다.
        for j in (9, 10, 12, 13, 14, 15):
            sm.cell(row, j).number_format = _fmt_for(d["unit"])
        sm.cell(row, 11).number_format = "0.00"
        hmap = {name: idx for idx, name in enumerate(headers, 2)}
        for name in ("적용 저점 RPM", "적용 고점 RPM", "적용 저점 shear rate", "적용 고점 shear rate"):
            cell = sm.cell(row, hmap[name])
            if isinstance(cell.value, float):
                cell.number_format = "0.00"

        # 반복성 상태는 Summary에 표시된 대표 통계의 n(H), SD(J), RSD(K)를 기준으로 계산한다.
        # 진단 지표(precision_evaluation=False)는 기존 규칙대로 N/A로 둔다.
        hmap = {name: idx for idx, name in enumerate(headers, 2)}
        n_ref = f"H{row}"
        rule_L = get_column_letter(hmap["반복성 규칙(임시)"])
        good_L = get_column_letter(hmap["GOOD ≤"])
        rev_L = get_column_letter(hmap["REVIEW ≤"])
        prec_col = hmap["반복성 상태"]
        if not d["precision_evaluation"]:
            sm.cell(row, prec_col, "N/A")
        else:
            sm.cell(row, prec_col, (f'=IF(OR({rule_L}{row}="-",{n_ref}<2),"N/A",IF({rule_L}{row}="RSD",'
                                    f'IF(ISNUMBER(K{row}),IF(K{row}<={good_L}{row},"GOOD",IF(K{row}<={rev_L}{row},"REVIEW","HIGH VAR")),"HIGH VAR"),'
                                    f'IF(J{row}<={good_L}{row},"GOOD",IF(J{row}<={rev_L}{row},"REVIEW","HIGH VAR"))))'))
        remark_col = hmap["비고·제외 근거(반복)"]
        for j in (prec_col, hmap["값 상태"], hmap["데이터 상태"], hmap["합격성 반영"], remark_col):
            sm.cell(row, j).alignment = Alignment(
                wrap_text=True, vertical="top", horizontal="left" if j == remark_col else "center"
            )
            sm.cell(row, j).border = Border(bottom=THIN)
        if s.get("zero_ref") and primary["sign_flip"]:
            sm.cell(row, 16).fill = PatternFill("solid", fgColor=YELLOW)
        _fit_row_height(sm, row, {5: 28, 6: 40, remark_col: 72}, max_lines=10)
        row += 1
    last = row - 1
    hmap = {name: idx for idx, name in enumerate(headers, 2)}
    for name in ("반복성 상태", "데이터 상태", "합격성 반영"):
        col = get_column_letter(hmap[name])
        _status_cf(sm, f"{col}{h + 1}:{col}{last}")
    vcol = get_column_letter(hmap["값 상태"])
    _value_class_cf(sm, f"{vcol}{h + 1}:{vcol}{last}")
    notes = [
        "대표 채널: Summary의 유효 n, Mean, SD, RSD, Min, Max, Range, 부호는 '대표 채널' 열의 통계입니다. primary_channel=inclusive이면 Method A는 정책 허용 hold를 반영한 INCLUSIVE 통계가 기존 통계 열에 표시됩니다. Strict/Inclusive 병렬 원자료·상세 통계는 Metric_Statistics, AReplicatesINCL, AReplicatesEXC, StepAudit에서 확인합니다.",
        "데이터 상태: VALID=모든 반복 유효 | NOTICE=값 표시·해석 주의(합격성 제외; A_OVS_L 미산출·H-B R²·GRANGE) | REVIEW=일부 반복 공란·경고 또는 [SIGN] 부호 전환 | EXCEPTION=승인 예외 hold Step 사용 | CENSORED=검출하한 밖(Torque<10 % 또는 >95 %) 참고값 | INVALID=유효 반복 없음 | NO DATA=해당 시험 미제공.",
        "합격성 반영: ELIGIBLE=Lot 합격성 판단에 쓸 수 있는 상태 | REVIEW=조건부 | EXCLUDED=EXCEPTION/CENSORED라 판정에서 제외 | —=진단·데이터 품질 지표(Role DIAGNOSTIC/DATA_QUALITY/CENSORED_METRIC)라 판정 대상 아님.",
        f"RSD(%)는 SD/|Mean|×100입니다. 0 기준 차이형 지표(A_BRK_PEAK·A_OVS_L·B_BRK_DECAY)는 '{RSD_NA_ZERO_REF}', |Mean|<SD(RSD>100%)이면 '{RSD_NA_NEAR_ZERO}'로 표시하고 SD·Range·부호(+/−)로 반복성을 봅니다.",
        "값 상태는 strict 정량값 / inclusive exception / torque 하한 밖 참고값 / torque 상한 밖 참고값 / torque 범위 밖 참고값 / NO VALUE / INVALID 중 하나입니다. Torque 참고값은 통계에 넣을 수 있지만 합격성 ELIGIBLE로 쓰지 않습니다.",
        "공식 SSI는 A_SSI_FIELD_0P2_2 = eta_up(0.2 rpm)/eta_up(2.0 rpm), A_SSI_QC_0P5_5 = eta_up(0.5 rpm)/eta(5.0 rpm Up endpoint)이다. 없는 RPM은 다른 RPM으로 대체하지 않는다. 실측 10배 쌍은 DecadeSSI 시트에 pair_id별로 따로 둔다.",
        "적용 저점 RPM과 적용 고점 RPM은 별도 열이다. 반복마다 RPM이 다르면 데이터 상태는 REVIEW이고 비고에 반복별 RPM을 남긴다. 측정 최고 RPM이 채널에서 부적격하면 더 낮은 RPM으로 대체하지 않고 NO VALUE다.",
        "n EXC는 INCLUSIVE 채널에서 EXCEPTION으로 분류된 반복 수, n CENS는 대표 채널의 CENSORED 반복 수, n Low/High-TQ Ref는 Torque 참고값이 들어간 반복 수입니다.",
        "반복성 규칙은 임시 데이터 품질 기준입니다(RSD=%, SD=지표 단위). GR&R·Lot 분포·원심 장벽 기능시험 상관 확인 전에는 규격(USL/LSL)으로 쓰지 않습니다. 반복성 규칙·GOOD·REVIEW 열을 수정하면 반복성 상태가 재계산됩니다.",
        "공란은 산출 불가 또는 평균 부적격(EXCEPTION/CENSORED/INVALID) 값입니다. 대체 입력하지 않았고 사유는 Replicate_Flags·StepAudit에 있습니다. 지표 정의는 QC_Dictionary.",
    ]
    for i, t in enumerate(notes, last + 2):
        sm.merge_cells(start_row=i, start_column=2, end_row=i, end_column=n_cols)
        c = sm.cell(i, 2, t)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        c.font = Font(italic=True, color="9C6500")
        sm.row_dimensions[i].height = 38

    # ---------------------------------------------------------------- Dictionary
    _sheet_base(dic, "QC_Dictionary — 각 QC 지표가 가리키는 물성과 해석",
                "정의·계산식, 가리키는 물성, SST 기능 단계 연결, 값 증감의 의미, 이 지표가 말하지 않는 것(한계), 유효 조건.", 13)
    _nav(dic, names)
    dh = ["층위", "Code", "Role / 합격성", "지표", "단위", "정의·계산식", "가리키는 물성", "SST 기능 연결", "값이 높아질 때", "값이 낮아질 때",
          "해석 한계(말하지 않는 것)", "유효 조건", "반복성 규칙(임시)"]
    _headers(dic, 5, dh, 2)
    dw = {2: 14, 3: 15, 4: 16, 5: 28, 6: 7, 7: 38, 8: 38, 9: 30, 10: 28, 11: 28, 12: 40, 13: 28, 14: 14}
    for col, w in dw.items():
        dic.column_dimensions[get_column_letter(col)].width = w
    for i, d in enumerate(_ordered_defs(defs), 6):
        rule, good, rev = d["prec"]
        vals = [TIER_LABELS[d["tier"]], d["code"], f"{d['metric_role']} / {'합격성 O' if d['lot_acceptance'] else '진단(제외)'}",
                d["name"], d["unit"], d["formula"], d["property"], d["function"],
                d["up"], d["down"], d["limits"], d["validity"],
                ("—" if rule == "-" else f"{rule}: GOOD ≤{good:g}, REVIEW ≤{rev:g}")
                + (" | RSD 미적용(0 기준), 부호 전환 시 [SIGN] REVIEW" if d.get("zero_ref") else "")]
        for j, v in enumerate(vals, 2):
            c = dic.cell(i, j, v)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            c.border = Border(bottom=THIN)
        _fit_row_height(dic, i, dw, max_lines=12)
    dic.freeze_panes = "F6"

    # ---------------------------------------------------------------- Step 수준
    step_rows = []
    for r in recs_by_method["A"]:
        for s in r.get("step_rows", []):
            step_rows.append({"repeat": r["repeat"], **s})
    _sheet_base(srep, "A_Step_Replicates — STRICT 채널 Step 대표값·준평형·분해능",
                "STRICT(60 s·Torque 10–95%) 기준 대표값입니다. used는 STRICT 정량 사용 여부이며, 승인 short-hold의 inclusive 사용 여부는 StepAudit 및 A_Replicates_INCL을 참조합니다. tail_drift_pct=후반 10초 선형 드리프트, resolution_pct=Torque 0.1% 표시 단위의 상대 크기.", 17)
    _nav(srep, names)
    sh = ["repeat", "Step", "dir", "RPM", "gamma", "hold_s", "used", "exception", "tau_Pa", "eta_cP",
          "torque_valid", "torque_window", "temp", "tail_drift_pct", "resolution_pct", "status"]
    sfm = {"RPM": "0.00", "gamma": "0.00", "hold_s": "0", "tau_Pa": "0.00", "eta_cP": "#,##0", "torque_valid": "0.00",
           "torque_window": "0.00", "temp": "0.00", "tail_drift_pct": "0.00", "resolution_pct": "0.00"}
    _write_table_sheet(srep, sh, [{**x, "used": "Yes" if x["used"] else "No",
                                  "exception": "Yes" if x["exception"] else "No"} for x in step_rows], 5, "AStepRep", sfm)
    for j, hname in enumerate(sh, 2):
        srep.column_dimensions[get_column_letter(j)].width = 22 if hname == "status" else 11
    srep.freeze_panes = "C6"
    _sheet_base(sstat, "A_Step_Statistics — 정량 Step의 반복 통계와 Lot 평균 유동곡선",
                "정량 사용(used) Step만 집계합니다. RSD는 반복 간 변동, 분해능은 판별 가능한 최소 상대 차이입니다.", 14)
    _nav(sstat, names)
    agg = []
    if step_rows:
        sdf = pd.DataFrame(step_rows)
        sdf = sdf[sdf["used"]]
        for (direction, rpm), g in sdf.groupby(["dir", "RPM"], sort=False):
            tau = g["tau_Pa"].astype(float)
            agg.append({"dir": direction, "RPM": float(rpm), "gamma": float(g["gamma"].iloc[0]), "n_rep": int(len(g)),
                        "tau_mean": float(tau.mean()), "tau_SD": float(tau.std(ddof=1)) if len(g) > 1 else np.nan,
                        "tau_RSD_pct": float(tau.std(ddof=1) / tau.mean() * 100) if len(g) > 1 and tau.mean() else np.nan,
                        "eta_mean": float(g["eta_cP"].mean()), "torque_mean": float(g["torque_valid"].mean()),
                        "drift_max_abs": float(g["tail_drift_pct"].abs().max()),
                        "resolution_pct": float(g["resolution_pct"].mean()),
                        "exception_n": int(g["exception"].sum())})
        agg.sort(key=lambda x: (0 if x["dir"] == "Up" else 1, x["gamma"] if x["dir"] == "Up" else -x["gamma"]))
    ah = ["dir", "RPM", "gamma", "n_rep", "tau_mean", "tau_SD", "tau_RSD_pct", "eta_mean", "torque_mean",
          "drift_max_abs", "resolution_pct", "exception_n"]
    afm = {"RPM": "0.00", "gamma": "0.00", "tau_mean": "0.00", "tau_SD": "0.00", "tau_RSD_pct": "0.00",
           "eta_mean": "#,##0", "torque_mean": "0.00", "drift_max_abs": "0.00", "resolution_pct": "0.00"}
    _write_table_sheet(sstat, ah, agg, 5, "AStepStat", afm)
    for j in range(2, len(ah) + 2):
        sstat.column_dimensions[get_column_letter(j)].width = 12
    if agg:
        chart = ScatterChart()
        chart.title = f"HT-{lot} Lot 평균 유동곡선 (정량 Step)"
        chart.style = 13
        chart.x_axis.title = "Shear rate γ̇ (1/s)"
        chart.y_axis.title = "Shear stress τ (Pa)"
        for direction, color in (("Up", "1F4E78"), ("Down", "C55A11")):
            idx = [i for i, a in enumerate(agg, 6) if a["dir"] == direction]
            if not idx:
                continue
            xs = Reference(sstat, min_col=4, min_row=min(idx), max_row=max(idx))
            ys = Reference(sstat, min_col=6, min_row=min(idx), max_row=max(idx))
            se = Series(ys, xs, title=direction)
            se.marker.symbol = "circle"
            se.marker.size = 7
            se.graphicalProperties.line.solidFill = color
            se.marker.graphicalProperties.solidFill = color
            chart.series.append(se)
        chart.x_axis.delete = False
        chart.y_axis.delete = False
        chart.height, chart.width = 9, 16
        sstat.add_chart(chart, f"B{8 + len(agg)}")

    # ---------------------------------------------------------------- Flags
    _sheet_base(flags_ws, "Replicate_Flags — 반복별 공란·경고·제외 사유",
                "지표 상태가 OK가 아닌 항목과 04_2 데이터·모델 검사의 비OK 항목을 모두 나열합니다.", 8)
    _nav(flags_ws, names)
    frows = []
    for m in ("A", "B", "D"):
        for r in recs_by_method[m]:
            for code, s in r["status"].items():
                if s["status"] != "OK":
                    frows.append({"method": m, "repeat": r["repeat"], "file": r["file"], "item": f"{code} [S]" if m == "A" else code,
                                  "status": s["status"], "reason": s["reason"]})
            if m == "A":
                for code, s in r.get("status_incl", {}).items():
                    if s["status"] != "OK" and s != r["status"].get(code):
                        frows.append({"method": m, "repeat": r["repeat"], "file": r["file"], "item": f"{code} [I]",
                                      "status": s["status"], "reason": s["reason"]})
            for f in r.get("check_flags", []):
                frows.append({"method": m, "repeat": r["repeat"], "file": r["file"], "item": f"[04_2 검사] {f['code']}",
                              "status": f["status"], "reason": f["reason"]})
    fh = ["method", "repeat", "file", "item", "status", "reason"]
    _write_table_sheet(flags_ws, fh, frows, 5, "Flags")
    for col, w in {2: 8, 3: 8, 4: 22, 5: 26, 6: 11, 7: 90}.items():
        flags_ws.column_dimensions[get_column_letter(col)].width = w
    for i in range(6, 6 + len(frows)):
        flags_ws.cell(i, 7).alignment = Alignment(wrap_text=True, vertical="top")
    if frows:
        _status_cf(flags_ws, f"F6:F{5 + len(frows)}")
    flags_ws.freeze_panes = "B6"

    # ---------------------------------------------------------------- Traceability
    _sheet_base(trace, "Traceability — 원본·분석기·규칙 추적", "원본 SHA-256, 분석기 버전·해시, 고정 규칙, 실행 명령을 기록합니다.", 10)
    _nav(trace, names)
    tinfo = [
        ("분석기", f"{VERSION} ({Path(__file__).name})"), ("분석기 SHA-256", run_meta.get("analyzer_sha256", "")),
        ("QC layer", QC_LAYER_VERSION), ("분석 경로(정책)", f"{policy.name} | hash {policy.digest()}"),
        ("정책 JSON", policy.to_json()), ("실행 시각", run_meta.get("timestamp", "")),
        ("실행 명령", run_meta.get("command", "")),
        ("Method A 규칙", f"회전 Step {METHOD_A_STEP_SECONDS:g}초, 후반 {METHOD_A_TAIL_SECONDS:g}초, H-B 최소 {METHOD_A_MIN_HB_POINTS}점/권장 {METHOD_A_RECOMMENDED_HB_POINTS}점"),
        ("Torque 유효범위", run_meta.get("torque_rule", f"{DEFAULT_TORQUE_MIN:g}–{DEFAULT_TORQUE_MAX:g}%")),
        ("고정 평가점", f"Low={FIXED_LOW_RPM:g} rpm, Mid={FIXED_MID_RPM:g} rpm; 고전단은 실제 측정 최고 RPM"),
        ("임시 기준", f"드리프트 REVIEW >{EQ_DRIFT_REVIEW_PCT:g}%, 시험온도 {TEST_TEMP_TARGET_C:g}±{TEST_TEMP_TOL_C:g} °C, 온도폭 ≤{TEMP_SPAN_MAX_C:g} °C, 밀도 교차확인 ±{DENSITY_CROSSCHECK_TOL:g} g/mL"),
    ]
    for i, (k, v) in enumerate(tinfo, 5):
        trace.cell(i, 2, k).font = Font(bold=True)
        trace.merge_cells(start_row=i, start_column=3, end_row=i, end_column=10)
        trace.cell(i, 3, v).alignment = Alignment(wrap_text=True, vertical="top")
    trows = []
    for m in ("A", "B", "D"):
        for r in recs_by_method[m]:
            trows.append({"method": m, "repeat": r["repeat"], "file": r["file"], "route": r.get("route"),
                          "data_qc": r.get("data_qc"), "raw_sha256": r.get("raw_sha256"),
                          "template": r.get("template", ""), "report": r.get("output", ""),
                          "exception_steps": ",".join(map(str, r.get("exception_steps", [])))})
    th = ["method", "repeat", "file", "route", "data_qc", "raw_sha256", "template", "report", "exception_steps"]
    _write_table_sheet(trace, th, trows, 5 + len(tinfo) + 2, "Trace")
    for col, w in {2: 16, 3: 8, 4: 24, 5: 18, 6: 10, 7: 66, 8: 30, 9: 36, 10: 14}.items():
        trace.column_dimensions[get_column_letter(col)].width = w

    # ---------------------------------------------------------------- StepAudit / QCPolicy (06)
    audit = []
    for r in recs_by_method["A"]:
        audit.extend(step_audit_rows(r, policy))
    write_step_audit_sheet(wb, {}, policy, rows=audit, title_extra=f" — HT-{lot} 전 반복")
    write_policy_sheet(wb, policy, run_meta)
    write_legacy_4rpm_sheet(wb, recs_by_method["A"])
    write_decade_ssi_sheet(wb, recs_by_method["A"])

    for ws in wb.worksheets:
        #ws.sheet_view.zoomScale = 90
        # zoomScale은 sheetView/pane XML 복구 문제를 피하기 위해 설정하지 않는다.
        # 기존 freeze_panes, selection, sheetView는 openpyxl 기본 상태로 둔다.
        _print_setup(ws)

    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except AttributeError:
        pass
    tag = "" if policy.name == "STANDARD" else f"_{policy.name}"
    output = Path(outdir) / f"HT-{lot}_Lot_QC{tag}{suffix}.xlsx"
    wb.save(output)
    return output, lot_stats


def write_lot_compare_workbook(lot_stats_by_lot, defs, outdir, route="STANDARD", run_meta=None, n_recs_by_lot=None,
                               decade_by_lot=None):
    """같은 실행의 여러 Lot을 지표별로 나열한 비교 워크북(Python 계산값, 참고용).

    각 Lot의 공식 통계는 HT-<lot>_Lot_QC.xlsx의 Replicates→Metric_Statistics 수식 연쇄이며,
    이 시트는 그 결과를 Lot 간 비교하기 위해 한 곳에 모은 정적 사본이다.
    """
    run_meta = run_meta or {}
    n_recs_by_lot = n_recs_by_lot or {}
    lots = list(lot_stats_by_lot.keys())
    ordered = [d for d in _ordered_defs(defs) if d.get("output_policy") not in ("decade_sheet", "dictionary_only", "legacy")]
    wb = Workbook()
    ws = wb.active
    ws.title = "Lot_Compare"
    n_cols = 32
    _sheet_base(ws, f"Lot 간 QC 지표 비교 — {len(lots)} Lot ({', '.join('HT-' + str(x) for x in lots)}) — 정책 {route}",
                "지표별로 Lot을 나열한 참고용 정적 사본입니다(Python 계산값). 공식 통계는 각 HT-<lot>_Lot_QC.xlsx의 수식 연쇄를 참조하십시오. "
                "상태는 데이터 품질 판정이며 제품 적합성 판정이 아닙니다.", n_cols)
    exc_text = f"POLICY {route} — 승인 예외 hold·범위 밖 Torque 점 포함 결과. EXCEPTION·CENSORED 값은 표준 60초 Lot 통계·추세·규격 설정에 혼용 금지"
    if route != "STANDARD":
        _banner(ws, exc_text, n_cols)
    headers = ["층위", "Code", "지표", "단위", "Lot", "유효 n / 전체", "Mean", "SD", "RSD(%)", "Min", "Max", "Range",
               "|Mean|", "부호 +/−", "Incl Mean", "Δ(Incl−Strict)", "반복성 규칙(임시)", "GOOD ≤", "REVIEW ≤", "반복성 상태",
               "값 상태", "데이터 상태", "합격성", "Lot 내 순위(Mean)", "Lot 내 순위(SD)",
               "Low RPM", "High RPM", "Low Step", "High Step", "Selection rule", "비고"]
    _headers(ws, 5, headers, 2)
    widths = {2: 17, 3: 16, 4: 30, 5: 7, 6: 12, 7: 9, 8: 11, 9: 10, 10: 14, 11: 11, 12: 11, 13: 10, 14: 10, 15: 8,
              16: 11, 17: 11, 18: 9, 19: 8, 20: 8, 21: 10, 22: 22, 23: 12, 24: 12, 25: 14, 26: 14,
              27: 12, 28: 12, 29: 12, 30: 12, 31: 36, 32: 48}
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w
    row = 6
    for d in ordered:
        code = d["code"]
        per_lot = {lot: st.get(code) for lot, st in lot_stats_by_lot.items()}
        if all(x is None or x["n_total"] == 0 for x in per_lot.values()):
            continue
        rule, good, rev = d["prec"]
        means = {lot: x["mean"] for lot, x in per_lot.items() if x and _finite(x["mean"])}
        sds = {lot: x["sd"] for lot, x in per_lot.items() if x and _finite(x["sd"])}
        rank_mean = {lot: i for i, lot in enumerate(sorted(means, key=lambda k: means[k]), 1)}
        rank_sd = {lot: i for i, lot in enumerate(sorted(sds, key=lambda k: sds[k]), 1)}
        highs = {str(x.get("high_rpm")) for x in per_lot.values() if x and x.get("high_rpm")}
        high_conflict = len(highs) > 1
        first = row
        for lot in lots:
            x = per_lot.get(lot)
            if x is None:
                continue
            pstat = prec_status(d, x)
            data_status = x["data_status"]
            remark = format_value_remark("INCLUSIVE" if x.get("channel") == "inclusive" else "STRICT", x, active_policy())
            if high_conflict:
                if data_status in ("VALID", "NOTICE"):
                    data_status = "REVIEW"
                remark = "REVIEW: Lot 간 적용 High RPM이 다름 | " + remark
            vals = [TIER_LABELS[d["tier"]], code, d["name"], d["unit"], f"HT-{lot}",
                    f"{x['n']} / {x['n_total']}",
                    x["mean"] if _finite(x["mean"]) else None, x["sd"] if _finite(x["sd"]) else None,
                    (x["rsd_note"] if x["rsd_note"] else (x["rsd"] if _finite(x["rsd"]) else None)),
                    x["min"] if _finite(x["min"]) else None, x["max"] if _finite(x["max"]) else None,
                    x["range"] if _finite(x["range"]) else None, x["abs_mean"] if _finite(x["abs_mean"]) else None,
                    x["sign_text"] or None,
                    x.get("inclusive", {}).get("mean") if _finite(x.get("inclusive", {}).get("mean", np.nan)) else None,
                    x.get("delta") if _finite(x.get("delta", np.nan)) else None,
                    rule, good, rev, pstat, x.get("value_class") or VC_NO, data_status, x.get("acceptance_status", "—"),
                    rank_mean.get(lot), rank_sd.get(lot),
                    _rpm_cell(x.get("low_rpm")), _rpm_cell(x.get("high_rpm")),
                    x.get("low_step") or "", x.get("high_step") or "", x.get("selection_rule") or "", remark]
            remark_col = headers.index("비고") + 2
            for j, v in enumerate(vals, 2):
                c = put(ws, row, j, v)
                c.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left" if j in (4, remark_col) else "center")
                c.border = Border(bottom=THIN)
            for j in (8, 9, 11, 12, 13, 14, 16, 17):
                ws.cell(row, j).number_format = _fmt_for(d["unit"])
            ws.cell(row, 10).number_format = "0.00"
            if x.get("zero_ref") and x.get("sign_flip"):
                ws.cell(row, 15).fill = PatternFill("solid", fgColor=YELLOW)
            _fit_row_height(ws, row, {4: 30, remark_col: 48}, max_lines=8)
            row += 1
        for j in range(2, n_cols + 1):
            ws.cell(row - 1, j).border = Border(bottom=Side(style="medium", color="1F4E78"))
    last = row - 1
    if last >= 6:
        cmp_headers = headers
        for name in ("반복성 상태", "데이터 상태", "합격성"):
            col = get_column_letter(cmp_headers.index(name) + 2)
            _status_cf(ws, f"{col}6:{col}{last}")
        _value_class_cf(ws, f"{get_column_letter(cmp_headers.index('값 상태') + 2)}6:{get_column_letter(cmp_headers.index('값 상태') + 2)}{last}")
    notes = [
        "Lot 내 순위: 같은 지표에서 Mean(작은 값=1위)·SD(작은 값=1위)의 순번입니다. 값이 '좋다/나쁘다'가 아니라 상대 위치이며, "
        "지표별 증감 의미는 QC_Dictionary(각 Lot 워크북) 또는 여기 Notes 시트를 참조하십시오.",
        f"RSD가 '{RSD_NA_ZERO_REF}' 또는 '{RSD_NA_NEAR_ZERO}'인 지표는 SD·Range·부호로 비교하십시오. RSD가 큰 것은 산포가 큰 것이 아니라 평균이 0에 가깝다는 뜻일 수 있습니다.",
        "부호 +/−가 섞인 0 기준 지표([SIGN])는 반복성 GOOD이라도 원시곡선 확인 대상입니다. 데이터 상태 N/A는 프로그램 구성상 해당 없음(측정 실패 아님)입니다.",
        "T3 모델값(τy·K·n·τy 유지비)은 외삽 파라미터라 단독 비교를 피하고, 직접 관측 지표(A_TAU_*·A_DU_*)와 실제 원심 장벽 성능을 함께 보십시오.",
        "이 시트는 정적 사본입니다. 값을 수정해도 각 Lot 워크북과 long CSV에는 반영되지 않습니다.",
    ]
    for i, t in enumerate(notes, last + 2):
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=n_cols + 1)
        c = ws.cell(i, 2, t)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        c.font = Font(italic=True, color="9C6500")
        ws.row_dimensions[i].height = 30
    ws.freeze_panes = "G6"
    if decade_by_lot:
        dec = wb.create_sheet("Decade_Compare")
        dheaders = ["Pair ID", "Lot", "Channel", "numerator RPM", "denominator RPM", "numerator Step", "denominator Step",
                    "n", "source type", "selection rule", "value class", "protocol"]
        pair_lots = {}
        drows = []
        for lot, recs in decade_by_lot.items():
            for item in aggregate_decade_pair_stats(recs, channel="inclusive"):
                pair_lots.setdefault(item["pair_id"], set()).add(lot)
                drows.append(item | {"lot": lot})
        all_ids = set(pair_lots)
        out_rows = []
        for item in drows:
            mismatch = all_ids - pair_lots.get(item["pair_id"], set())
            # 다른 Lot에만 있는 pair는 이 pair와 순위를 합치지 않는다.
            protocol = "protocol mismatch" if len(all_ids) > 1 and any(pair_lots[other] != pair_lots[item["pair_id"]] for other in all_ids) else "same pair"
            out_rows.append({"Pair ID": item["pair_id"], "Lot": item["lot"], "Channel": item["channel"],
                             "numerator RPM": item["numerator_rpm"], "denominator RPM": item["denominator_rpm"],
                             "numerator Step": item["low_step"], "denominator Step": item["high_step"],
                             "n": item["n"], "source type": item["source_type"], "selection rule": item["selection_rule"],
                             "value class": item["value_class"], "protocol": protocol if len({frozenset(v) for v in pair_lots.values()}) > 1 else "matched"})
        _sheet_base(dec, "Decade pair 비교 — pair_id별로 분리", "서로 다른 RPM 쌍은 한 평균이나 한 순위로 합치지 않습니다.", 8)
        _write_table_sheet(dec, dheaders, out_rows, 5, "DecadeCompare")

    nt = wb.create_sheet("Notes")
    _sheet_base(nt, "Notes — 실행 정보와 지표 증감 의미", "Lot_Compare 해석에 필요한 최소 정보입니다.", 8)
    info = [("분석기", f"{VERSION} / {QC_LAYER_VERSION}"), ("분석 정책", f"{route} | hash {run_meta.get('policy_hash', '')}"),
            ("실행 시각", run_meta.get("timestamp", "")),
            ("실행 명령", run_meta.get("command", "")), ("Torque 유효범위", run_meta.get("torque_rule", "")),
            ("고정 평가점", f"Low={FIXED_LOW_RPM:g} rpm, Mid={FIXED_MID_RPM:g} rpm; 고전단은 실제 측정 최고 RPM. 부적격이면 NO VALUE"),
            ("Lot·반복 수", ", ".join(f"HT-{lot}: A={n_recs_by_lot.get(lot, {}).get('A', 0)} B={n_recs_by_lot.get(lot, {}).get('B', 0)} "
                                   f"D={n_recs_by_lot.get(lot, {}).get('D', 0)}" for lot in lots))]
    for i, (k, v) in enumerate(info, 5):
        nt.cell(i, 2, k).font = Font(bold=True)
        nt.merge_cells(start_row=i, start_column=3, end_row=i, end_column=8)
        nt.cell(i, 3, v).alignment = Alignment(wrap_text=True, vertical="top")
    r0 = 5 + len(info) + 1
    _headers(nt, r0, ["층위", "Code", "지표", "값이 높아질 때", "값이 낮아질 때", "해석 한계"], 2)
    for col, w in {2: 17, 3: 16, 4: 30, 5: 34, 6: 34, 7: 48}.items():
        nt.column_dimensions[get_column_letter(col)].width = w
    for i, d in enumerate(ordered, r0 + 1):
        for j, v in enumerate([TIER_LABELS[d["tier"]], d["code"], d["name"], d["up"], d["down"], d["limits"]], 2):
            c = nt.cell(i, j, v)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            c.border = Border(bottom=THIN)
        _fit_row_height(nt, i, {4: 30, 5: 34, 6: 34, 7: 48}, max_lines=10)
    for w in wb.worksheets:
        w.sheet_view.showGridLines = False
        w.column_dimensions["A"].width = 3
        _print_setup(w)
    tag = "" if route == "STANDARD" else f"_{route}"
    output = Path(outdir) / f"Lot_QC_Compare{tag}.xlsx"
    wb.save(output)
    return output


LONG_COLUMNS = ["lot", "method", "repeat", "file", "test_start", "route", "policy_hash", "channel", "data_qc", "code", "tier",
                "role", "name", "value", "unit", "status", "value_class", "reason", "analyzer", "qc_layer", "raw_sha256", "run_timestamp",
                "pair_id", "target_ratio", "actual_ratio", "ratio_error_pct", "numerator_rpm", "denominator_rpm",
                "numerator_step", "denominator_step", "numerator_gamma", "denominator_gamma", "source_type", "selection_rule",
                "record_type", "original_metric_code", "canonical_metric_code", "migration_status", "migration_reason",
                "direction", "included_in_lot_stats", "lot_stat_exclusion_reason", "data_status", "acceptance_status"]


def long_rows(records, defs=None, timestamp=""):
    """long-format 행. Method A는 strict/inclusive 두 채널 행을 모두 남긴다(channel 열로 구분)."""
    dd = _defs_by_code(defs or metric_definitions())
    rows = []
    for r in records:
        channels = [("strict", r["values"], r["status"])]
        if r.get("method") == "A" and "values_incl" in r:
            channels.append(("inclusive", r["values_incl"], r["status_incl"]))
        for ch, vals, sts in channels:
            for code, value in vals.items():
                d = dd.get(code)
                if not d or d.get("output_policy") == "legacy":
                    continue
                if d.get("output_policy") == "when_finite" and not _finite(value):
                    continue
                s = sts.get(code, _st("NO VALUE"))
                rows.append({"lot": r["lot"], "method": r["method"], "repeat": r["repeat"], "file": r["file"],
                             "test_start": r.get("test_start", ""), "route": r.get("route", ""),
                             "policy_hash": r.get("policy_hash", ""), "channel": ch,
                             "data_qc": r.get("data_qc", ""), "code": code, "tier": d.get("tier", ""),
                             "role": d.get("metric_role", ""), "name": d.get("name", ""),
                             "value": value if _finite(value) else None,
                             "unit": d.get("unit", ""), "status": s["status"],
                             "value_class": s.get("value_class") or _default_value_class(s.get("status")),
                             "reason": s["reason"],
                             "analyzer": VERSION, "qc_layer": QC_LAYER_VERSION, "raw_sha256": r.get("raw_sha256", ""),
                             "run_timestamp": timestamp,
                             "pair_id": s.get("pair_id") or "", "target_ratio": s.get("target_ratio"),
                             "actual_ratio": s.get("actual_ratio"), "ratio_error_pct": s.get("ratio_error_pct"),
                             "numerator_rpm": s.get("numerator_rpm"), "denominator_rpm": s.get("denominator_rpm"),
                             "numerator_step": s.get("numerator_step"), "denominator_step": s.get("denominator_step"),
                             "numerator_gamma": s.get("numerator_gamma"), "denominator_gamma": s.get("denominator_gamma"),
                             "source_type": s.get("source_type") or "", "selection_rule": s.get("selection_rule") or "",
                             "record_type": "official_metric" if code in ("A_SSI_FIELD_0P2_2", "A_SSI_QC_0P5_5") else "step_metric",
                             "original_metric_code": code, "canonical_metric_code": code,
                             "migration_status": "", "migration_reason": "",
                             "direction": s.get("selected_direction") or ("Up" if code in ("A_SSI_FIELD_0P2_2", "A_SSI_QC_0P5_5") else ""),
                             "included_in_lot_stats": bool(s.get("included_in_lot_stats")),
                             "lot_stat_exclusion_reason": s.get("lot_stat_exclusion_reason") or "",
                             "data_status": s.get("data_status") or s.get("status"),
                             "acceptance_status": s.get("acceptance_status") or ""})
        if r.get("method") == "A":
            for ch_name, key in (("strict", "decade_pairs_strict"), ("inclusive", "decade_pairs_incl")):
                for pair in r.get(key) or []:
                    rows.append({"lot": r["lot"], "method": "A", "repeat": r["repeat"], "file": r["file"],
                                 "test_start": r.get("test_start", ""), "route": r.get("route", ""),
                                 "policy_hash": r.get("policy_hash", ""), "channel": ch_name,
                                 "data_qc": r.get("data_qc", ""), "code": pair.get("pair_id"), "tier": "T1",
                                 "role": "diagnostic", "name": pair.get("pair_id"),
                                 "value": pair.get("ssi") if _finite(pair.get("ssi")) else None, "unit": "-",
                                 "status": pair.get("data_status") or "", "value_class": pair.get("value_class") or "",
                                 "reason": pair.get("reason") or "", "analyzer": VERSION, "qc_layer": QC_LAYER_VERSION,
                                 "raw_sha256": r.get("raw_sha256", ""), "run_timestamp": timestamp,
                                 "pair_id": pair.get("pair_id") or "", "target_ratio": pair.get("target_ratio"),
                                 "actual_ratio": pair.get("actual_ratio"), "ratio_error_pct": pair.get("ratio_error_pct"),
                                 "numerator_rpm": pair.get("low_rpm"), "denominator_rpm": pair.get("high_rpm"),
                                 "numerator_step": pair.get("low_step"), "denominator_step": pair.get("high_step"),
                                 "numerator_gamma": pair.get("low_gamma"), "denominator_gamma": pair.get("high_gamma"),
                                 "source_type": pair.get("source_type") or "measured",
                                 "selection_rule": pair.get("selection_rule") or "",
                                 "record_type": "dynamic_decade_pair",
                                 "original_metric_code": pair.get("pair_id"), "canonical_metric_code": pair.get("pair_id"),
                                 "migration_status": "", "migration_reason": "",
                                 "direction": pair.get("direction") or "",
                                 "included_in_lot_stats": bool(pair.get("included_in_lot_stats")),
                                 "lot_stat_exclusion_reason": pair.get("lot_stat_exclusion_reason") or "",
                                 "data_status": pair.get("data_status") or "",
                                 "acceptance_status": pair.get("acceptance_status") or ""})
    return rows


def export_long_csv(records, path, defs=None, timestamp="", db_path=None):
    df = pd.DataFrame(long_rows(records, defs, timestamp), columns=LONG_COLUMNS)
    df.to_csv(path, index=False, encoding="utf-8-sig")
    if db_path:
        db_path = Path(db_path)
        if db_path.exists():
            # keep_default_na=False: 상태 문자열 'N/A'(프로그램 구성상 해당 없음)가 결측으로 바뀌지 않게 한다.
            # 예전 DB에 value_class가 없어도 빈 열을 보태 이어 붙인다.
            old = pd.read_csv(db_path, encoding="utf-8-sig", dtype=str, keep_default_na=False)
            for col in LONG_COLUMNS:
                if col not in old.columns:
                    old[col] = ""
            migrated = []
            for rec in old.to_dict("records"):
                if rec.get("code") == "A_SSI":
                    mig = canonicalize_metric_record({
                        "metric_code": "A_SSI",
                        "numerator_rpm": rec.get("numerator_rpm"),
                        "denominator_rpm": rec.get("denominator_rpm"),
                        "direction": rec.get("direction"),
                        "source_type": rec.get("source_type"),
                        "source_file": rec.get("file"),
                        "turnaround_shared": str(rec.get("turnaround_shared")).lower() in ("1", "true", "yes"),
                    })
                    rec["original_metric_code"] = mig["original_metric_code"]
                    rec["canonical_metric_code"] = mig["canonical_metric_code"]
                    rec["code"] = mig["canonical_metric_code"]
                    rec["migration_status"] = mig["migration_status"]
                    rec["migration_reason"] = mig["migration_reason"]
                    rec["record_type"] = "legacy_unknown" if mig["canonical_metric_code"] == "A_SSI_LEGACY_UNKNOWN" else "official_metric"
                    rec["included_in_lot_stats"] = mig["included_in_lot_stats"]
                    if mig["canonical_metric_code"] == "A_SSI_LEGACY_UNKNOWN":
                        rec["value_class"] = mig["value_class"]
                        rec["data_status"] = mig["data_status"]
                        rec["acceptance_status"] = mig["acceptance_status"]
                        rec["status"] = mig["data_status"]
                migrated.append(rec)
            old = pd.DataFrame(migrated)
            old = old.reindex(columns=LONG_COLUMNS)
            new = df.astype(str).replace({"None": "", "nan": ""})
            for col in LONG_COLUMNS:
                if col not in new.columns:
                    new[col] = ""
            new = new.reindex(columns=LONG_COLUMNS)
            df = pd.concat([old, new], ignore_index=True)
            df = df.drop_duplicates(subset=["lot", "raw_sha256", "method", "repeat", "code", "route", "policy_hash", "channel",
                                            "analyzer", "qc_layer"], keep="last")
        df.to_csv(db_path, index=False, encoding="utf-8-sig")
    return path


def print_lot_qc_cli(lot, route, stats, defs, output, n_a, n_b, n_d):
    print(f"[Lot QC] HT-{lot} | 경로={route} | A={n_a} B={n_b} 밀도={n_d} -> {Path(output).name}")
    for d in sorted(defs, key=lambda x: (TIER_ORDER.index(x["tier"]), "ADB".index(x["method"]))):
        if d["tier"] == "DQ" and d["code"] not in ("A_RES_L", "A_EQ_DRIFT_MAX"):
            continue
        s = stats.get(d["code"])
        if not s or s["n_total"] == 0:
            continue
        rsd_txt = s["rsd_note"] if s.get("rsd_note") else (f"{_cli_num(s['rsd'], 2):>6}%" if _finite(s["rsd"]) else "N/A")
        extra = ""
        if s.get("zero_ref") and s["n"]:
            extra = f" Range={_cli_num(s['range'], 3)} 부호={s['sign_text']}" + (" [SIGN]" if s.get("sign_flip") else "")
        inc = s.get("inclusive")
        if inc and d["method"] == "A" and _finite(inc.get("mean", np.nan)) and (not _finite(s["mean"]) or abs(inc["mean"] - s["mean"]) > 1e-12 or inc["n"] != s["n"]):
            extra += f" | Incl={_cli_num(inc['mean'], 4)} n={inc['n']} [{inc['data_status']}]"
        print(f"  {d['tier']:<2} {d['code']:<15} {_cli_num(s['mean'], 4):>12} {d['unit']:<6} "
              f"SD={_cli_num(s['sd'], 4):>9} RSD={rsd_txt:>14} n={s['n']}/{s['n_total']} [{s['data_status']}]{extra}"
              f"  {d['name']}")
    missing = [d["code"] for d in defs if d["tier"] != "DQ"
               and d.get("output_policy") not in ("decade_sheet", "dictionary_only", "legacy")
               and stats.get(d["code"], {}).get("n_total", 0) == 0]
    if missing:
        print(f"  NO DATA: {', '.join(missing)}")


def run_meta_info(tq_lo, tq_hi, policy=None):
    policy = policy or active_policy()
    return {"timestamp": datetime.now().astimezone().isoformat(timespec="seconds"),
            "command": " ".join([Path(sys.executable).name, *sys.argv]),
            "analyzer_sha256": sha256_file(Path(__file__).resolve()),
            "torque_rule": f"{tq_lo:g}–{tq_hi:g}%", "ref_rpm": dict(A_LEGACY_REF_RPM),
            "policy_name": policy.name, "policy_hash": policy.digest(), "policy_json": policy.to_json()}


def parser():
    p = argparse.ArgumentParser(
        description="SST gel Step Flow / 3ITT QC analyzer with policy-based strict/inclusive channels and integrated Lot QC (HB-analysis-06)")
    p.add_argument("inputs", nargs="*", help="DVNext CSV 또는 폴더 (--type A/B)")
    p.add_argument("--input-dir")
    p.add_argument("--recursive", action="store_true")
    p.add_argument("--type", choices=["A", "B", "QC", "a", "b", "qc"], default="A",
                   help="A=Step Flow, B=3ITT (04_2와 동일 출력 + QC_Indices), QC=A·B·밀도 Lot 종합 QC")
    p.add_argument("--template", help="--type A/B 템플릿")
    p.add_argument("--a-input", nargs="+", default=[], help="[QC] Method A CSV 파일 또는 폴더")
    p.add_argument("--b-input", nargs="+", default=[], help="[QC] Method B(3ITT) CSV 파일 또는 폴더")
    p.add_argument("--density-file", help="[QC] 밀도 입력 CSV (양식: --write-density-template)")
    p.add_argument("--template-a", help="[QC] Step Flow 템플릿")
    p.add_argument("--template-b", help="[QC] 3ITT 템플릿")
    p.add_argument("--db-csv", help="[QC] long-format 누적 DB CSV 경로(없으면 생성, 동일 원본·지표는 갱신)")
    p.add_argument("--ref-rpm-low", type=float, default=A_LEGACY_REF_RPM["L"], help="[QC] legacy diagnostic 저전단 RPM. A_TAU_L_UP·A_SSI 저점은 0.5로 고정")
    p.add_argument("--ref-rpm-mid", type=float, default=A_LEGACY_REF_RPM["M"], help="[QC] legacy diagnostic 중간 RPM. A_TAU_M_UP은 2.0으로 고정")
    p.add_argument("--ref-rpm-high", type=float, default=A_LEGACY_REF_RPM["H_REF"], help="[QC] legacy diagnostic 4 rpm. 대표 고전단이 아님")
    p.add_argument("--include-legacy-4rpm-diagnostics", action="store_true",
                   help="[QC] Legacy_4RPM_Diagnostics 시트를 추가한다. 기본값은 출력하지 않음")
    p.add_argument("--write-density-template", metavar="PATH", help="밀도 입력 양식 CSV를 만들고 종료")
    p.add_argument("--outdir", default="results")
    p.add_argument("--win-s", type=float, default=DEFAULT_WINDOW_SECONDS, help="Method B 대표구간(초); Method A는 10초로 고정")
    p.add_argument("--window-fraction", type=float, default=DEFAULT_WINDOW_FRACTION, help="Method B 대표구간 비율; Method A에서는 사용하지 않음")
    p.add_argument("--torque-min", type=float, default=DEFAULT_TORQUE_MIN, help="기본 10%%")
    p.add_argument("--torque-max", type=float, default=DEFAULT_TORQUE_MAX, help="기본 95%%")
    p.add_argument("--overwrite", action="store_true")
    p.add_argument("--no-lot-summary", action="store_true", help="Lot 통합 워크북 생성을 생략")
    g = p.add_argument_group("분석 정책(06)")
    g.add_argument("--policy", choices=["standard", "inclusive", "custom"], default="inclusive",
                   help="standard=60 s·Torque 10–95 %% strict만 통계 | inclusive(기본)=30/45/60 s 승인, 범위 밖 점 inclusive 채널 포함 | custom=아래 인자로 구성")
    g.add_argument("--accepted-holds", help="승인 hold(초), 예 30,45,60 (preset 덮어쓰기)")
    g.add_argument("--allow-short-hold-analysis", dest="allow_short", action="store_true", default=None)
    g.add_argument("--no-short-hold-analysis", dest="allow_short", action="store_false")
    g.add_argument("--include-out-of-torque-analysis", dest="incl_oot", action="store_true", default=None)
    g.add_argument("--no-out-of-torque-analysis", dest="incl_oot", action="store_false")
    g.add_argument("--lot-stats-mode", choices=["strict", "inclusive", "both"], default=None)
    # strict/inclusive 감사자료 계산은 보존하고, 사용자 화면의 대표 통계 채널만 별도로 지정한다.
    g.add_argument("--primary-channel", choices=["strict", "inclusive"], default=None,
                   help="Summary·콘솔·비교표 대표 통계 채널 (기본: policy preset의 strict)")
    g.add_argument("--exception-stat-policy", choices=["exclude", "separate", "include"], default=None)
    g.add_argument("--report-values", choices=["strict", "inclusive", "both"], default=None)
    g.add_argument("--no-strict-values", action="store_true")
    g.add_argument("--no-inclusive-values", action="store_true")
    g.add_argument("--audit-level", choices=["basic", "full"], default=None)
    g.add_argument("--min-hold-s", type=float, default=None, help="이보다 짧은 hold는 대표값 계산 보류(기본 10)")
    g.add_argument("--stabilization-target-s", type=float, default=None)
    g.add_argument("--stabilization-min-s", type=float, default=None)
    g.add_argument("--stabilization-temp-target", type=float, default=None)
    g.add_argument("--stabilization-temp-tolerance", type=float, default=None)
    g.add_argument("--require-stabilization", dest="require_stabilization", action="store_true", default=None)
    g.add_argument("--no-require-stabilization", dest="require_stabilization", action="store_false")
    g.add_argument("--ssi-decade-ratio", type=float, default=None)
    g.add_argument("--ssi-ratio-tolerance-pct", type=float, default=None)
    g.add_argument("--ssi-calculate-all", dest="ssi_calculate_all", action="store_true", default=None)
    g.add_argument("--ssi-primary-selection", default=None, choices=["none", "preferred_first", "lowest_pair", "highest_pair", "explicit"])
    g.add_argument("--ssi-preferred-pairs", default=None, help='예: "0.2:2,0.5:5"')
    g.add_argument("--ssi-explicit-pair", default=None, help='예: "0.5:5"')
    g.add_argument("--ssi-allow-interpolation", action="store_true")
    g.add_argument("--ssi-allow-model-estimation", action="store_true")
    g.add_argument("--ssi-include-down", action="store_true")
    g.add_argument("--ssi-pool-different-pairs", action="store_true")
    return p


def policy_from_args(a):
    """CLI 인자로 AnalysisPolicy를 구성한다(지시서 §C). preset 위에 명시 인자만 덮어쓴다."""
    if a.ssi_pool_different_pairs:
        raise ValueError("Pooling different SSI RPM pairs is not supported because pair identity must be preserved.")
    if a.ssi_allow_interpolation or a.ssi_allow_model_estimation:
        raise NotImplementedError(
            "Estimated decade SSI is not implemented in HB-analysis-06-6. Use measured RPM pairs only."
        )
    if a.ssi_primary_selection == "explicit" and not a.ssi_explicit_pair:
        raise ValueError("--ssi-explicit-pair LOW:HIGH is required when --ssi-primary-selection=explicit")
    pol = POLICY_PRESETS[a.policy]
    kw = {"torque_min_pct": float(a.torque_min), "torque_max_pct": float(a.torque_max)}
    if a.accepted_holds:
        holds = tuple(sorted(float(x) for x in a.accepted_holds.replace(";", ",").split(",") if x.strip()))
        if not holds:
            raise ValueError("--accepted-holds 형식: 30,45,60")
        kw["accepted_hold_s"] = holds
        if a.allow_short is None and any(h < pol.hold_target_s for h in holds):
            kw["allow_short_hold_analysis"] = True
    if a.allow_short is not None:
        kw["allow_short_hold_analysis"] = a.allow_short
    if a.incl_oot is not None:
        kw["include_out_of_torque_analysis"] = a.incl_oot
        # inclusive 옵션이 켜지면 범위 밖 참고값을 계산·통계에 넣고, 끄면 양쪽 모두 제외한다.
        kw["include_out_of_range_in_inclusive"] = bool(a.incl_oot)
        kw["censored_in_inclusive_fit"] = bool(a.incl_oot)
        kw["include_censored_in_lot_stats"] = bool(a.incl_oot)
    if a.lot_stats_mode:
        kw["lot_stats_mode"] = a.lot_stats_mode
    # argparse와 AnalysisPolicy 모두 snake_case를 사용한다. 옵션 미지정 시 preset의 기본값을 유지한다.
    if a.primary_channel:
        kw["primary_channel"] = a.primary_channel
    if a.exception_stat_policy:
        kw["exception_stat_policy"] = a.exception_stat_policy
        kw["include_exception_in_lot_stats"] = a.exception_stat_policy == "include"
    if a.report_values:
        kw["report_values"] = a.report_values
    if a.no_strict_values:
        kw["generate_strict_values"] = False
        kw["report_values"] = "inclusive"
    if a.no_inclusive_values:
        kw["generate_inclusive_values"] = False
        kw["report_values"] = "strict"
        kw["lot_stats_mode"] = "strict"
    if a.audit_level:
        kw["audit_level"] = a.audit_level
    if a.min_hold_s is not None:
        kw["min_hold_s"] = float(a.min_hold_s)
    if a.stabilization_target_s is not None:
        kw["stabilization_target_s"] = float(a.stabilization_target_s)
    if a.stabilization_min_s is not None:
        kw["stabilization_min_s"] = float(a.stabilization_min_s)
    if a.stabilization_temp_target is not None:
        kw["stabilization_temp_target_C"] = float(a.stabilization_temp_target)
    if a.stabilization_temp_tolerance is not None:
        kw["stabilization_temp_tolerance_C"] = float(a.stabilization_temp_tolerance)
    if a.require_stabilization is not None:
        kw["require_stabilization"] = bool(a.require_stabilization)
    if a.ssi_decade_ratio is not None:
        kw["ssi_decade_ratio"] = float(a.ssi_decade_ratio)
    if a.ssi_ratio_tolerance_pct is not None:
        kw["ssi_ratio_tolerance_pct"] = float(a.ssi_ratio_tolerance_pct)
    if a.ssi_calculate_all is not None:
        kw["ssi_calculate_all"] = bool(a.ssi_calculate_all)
    if a.ssi_primary_selection:
        kw["ssi_primary_selection"] = a.ssi_primary_selection
    if a.ssi_preferred_pairs:
        pairs = []
        for part in a.ssi_preferred_pairs.replace(";", ",").split(","):
            if part.strip():
                lo, hi = part.split(":")
                pairs.append((float(lo), float(hi)))
        kw["ssi_preferred_pairs"] = tuple(pairs)
    if a.ssi_explicit_pair:
        lo, hi = a.ssi_explicit_pair.split(":")
        kw["ssi_explicit_pair"] = (float(lo), float(hi))
    if a.ssi_include_down:
        kw["ssi_include_down"] = True
    pol = replace(pol, **kw)
    if pol.name != "CUSTOM" and any(k not in ("torque_min_pct", "torque_max_pct") for k in kw):
        preset_dict = POLICY_PRESETS[a.policy].to_dict()
        changed = [k for k, v in pol.to_dict().items() if preset_dict.get(k) != v and k not in ("torque_min_pct", "torque_max_pct")]
        if changed:
            pol = replace(pol, name=f"{pol.name}_CUSTOM")
    return pol


def print_policy(pol):
    print(f"정책: {pol.name} (hash {pol.digest()}) | 승인 hold {','.join(f'{x:g}' for x in pol.accepted_hold_s)} s | "
          f"짧은 hold 분석 {'허용' if pol.allow_short_hold_analysis else '불가'} | 범위 밖 Torque inclusive {'포함' if pol.include_out_of_torque_analysis else '제외'} | "
          f"Lot 통계 {pol.lot_stats_mode} | 대표 채널 {pol.primary_channel} | "
          f"EXCEPTION {pol.exception_stat_policy} | 보고 {pol.report_values}")
    print("STRICT 채널(60 s·Torque 10–95 %)은 정책과 무관하게 항상 계산되며 04_2 시트·Lot_Analysis는 STRICT 값입니다.")


def _unique_output(outdir, stem, suffix, overwrite, policy=None):
    policy = policy or active_policy()
    tag = "" if policy.name == "STANDARD" else f"_{policy.name}"
    output = outdir / f"{stem}_{suffix}_QC{tag}.xlsx"
    if output.exists() and not overwrite:
        n = 2
        while output.exists():
            output = outdir / f"{stem}_{suffix}_QC{tag}_{n}.xlsx"
            n += 1
    return output


def run_qc_mode(a, p):
    """[05] Method A·B·밀도 Lot 종합 QC."""
    if not a.a_input and not a.b_input:
        p.error("QC 모드는 --a-input 또는 --b-input 중 하나 이상이 필요합니다.")
    global INCLUDE_LEGACY_4RPM_DIAGNOSTICS
    INCLUDE_LEGACY_4RPM_DIAGNOSTICS = bool(getattr(a, "include_legacy_4rpm_diagnostics", False))
    A_LEGACY_REF_RPM.update({"L": a.ref_rpm_low, "M": a.ref_rpm_mid, "H_REF": a.ref_rpm_high})
    if not (0 < A_LEGACY_REF_RPM["L"] < A_LEGACY_REF_RPM["M"] < A_LEGACY_REF_RPM["H_REF"]):
        p.error("legacy reference RPM은 0 < low < mid < H_REF 이어야 합니다. 고전단 대표 RPM은 데이터에서 고릅니다.")
    try:
        a_files = sorted(collect_inputs(a.a_input, None, a.recursive), key=lambda x: x.name.lower()) if a.a_input else []
        b_files = sorted(collect_inputs(a.b_input, None, a.recursive), key=lambda x: x.name.lower()) if a.b_input else []
        tpl_a = resolve_template(a.template_a, "A") if a_files else None
        tpl_b = resolve_template(a.template_b, "B") if b_files else None
        density = load_density_file(a.density_file) if a.density_file else {}
    except (ValueError, FileNotFoundError) as exc:
        p.error(str(exc))
    outdir = Path(a.outdir).expanduser().resolve()
    outdir.mkdir(parents=True, exist_ok=True)
    policy = active_policy()
    meta = run_meta_info(a.torque_min, a.torque_max, policy)
    defs = metric_definitions()
    print("=" * 78)
    print(f"{VERSION} / {QC_LAYER_VERSION} | Lot 종합 QC")
    print(f"입력: A {len(a_files)}개 | B {len(b_files)}개 | 밀도 {sum(len(v) for v in density.values())}행 | 출력: {outdir}")
    print(f"고정 평가점: Low={FIXED_LOW_RPM:g} rpm, Mid={FIXED_MID_RPM:g} rpm | "
          "고전단 대표값은 실제 측정 최고 RPM | 그 Step이 부적격이면 NO VALUE | "
          "공식 SSI = eta_up(0.2)/eta_up(2.0) 및 eta_up(0.5)/eta(5.0 Up endpoint)")
    print(f"Method A STRICT 규칙: 회전 Step {policy.hold_target_s:g}초, 후반 {policy.tail_seconds:g}초, Torque {a.torque_min:g}–{a.torque_max:g}%")
    print_policy(policy)
    print("상태는 데이터 품질 판정이며 제품 적합성 판정이 아닙니다.")
    print("=" * 78)
    qc_records, failures = defaultdict(list), []
    jobs = [("A", f, tpl_a, "StepFlow") for f in a_files] + [("B", f, tpl_b, "3ITT") for f in b_files]
    for index, (method, source, tpl, suffix) in enumerate(jobs, 1):
        output = _unique_output(outdir, source.stem, suffix, a.overwrite)
        try:
            point_count, step_count, result = process_enhanced(source, method, tpl, output, a.win_s,
                                                                a.window_fraction, a.torque_min, a.torque_max)
            result.pop("_lot_record", None)
            result.pop("_step_records", None)
            cli_summary = result.pop("_cli_summary", None)
            qc = result.pop("_qc_record")
            qc_records[normalize_lot(qc["lot"])].append(qc)
            print(f"[{index}/{len(jobs)}] 완료 Method {method} {source.name} | 데이터 QC={qc['data_qc']} | 경로={qc['route']}")
            if method == "A":
                print_method_a_cli(source, output, point_count, step_count, result, cli_summary)
            else:
                print_method_b_cli(source, output, point_count, step_count, result, cli_summary)
        except Exception as exc:
            failures.append((source, exc))
            print(f"[{index}/{len(jobs)}] 실패 {source.name}: {type(exc).__name__}: {exc}", file=sys.stderr)
    all_records = []
    lots = list(dict.fromkeys(list(qc_records.keys()) + list(density.keys())))
    outputs = []
    stats_by_lot, n_by_lot, routes, decade_by_lot = {}, {}, set(), {}
    for key in lots:
        recs = qc_records.get(key, [])
        a_recs = [r for r in recs if r["method"] == "A"]
        b_recs = [r for r in recs if r["method"] == "B"]
        d_recs = density.get(key, [])
        lot_name = (recs[0]["lot"] if recs else key)
        for r in d_recs:
            r["lot"] = lot_name
        route = policy.name
        out, stats = write_lot_qc_workbook(lot_name, a_recs, b_recs, d_recs, outdir, route, meta, defs, policy=policy)
        outputs.append(out)
        all_records.extend(a_recs + b_recs + d_recs)
        stats_by_lot[lot_name] = stats
        decade_by_lot[lot_name] = a_recs
        n_by_lot[lot_name] = {"A": len(a_recs), "B": len(b_recs), "D": len(d_recs)}
        routes.add(route)
        print("-" * 78)
        print_lot_qc_cli(lot_name, route, stats, defs, out, len(a_recs), len(b_recs), len(d_recs))
    if len(stats_by_lot) >= 2:
        cmp_out = write_lot_compare_workbook(stats_by_lot, defs, outdir, policy.name, meta, n_by_lot, decade_by_lot)
        print("-" * 78)
        print(f"[Lot 비교] {len(stats_by_lot)} Lot -> {cmp_out.name} (참고용 정적 사본; 공식 통계는 각 Lot 워크북 수식)")
    if all_records:
        tag = "" if policy.name == "STANDARD" else f"_{policy.name}"
        long_path = outdir / f"Lot_QC_long{tag}.csv"
        export_long_csv(all_records, long_path, defs, meta["timestamp"], a.db_csv)
        print(f"[long CSV] {long_path.name}" + (f" | DB 누적 {a.db_csv}" if a.db_csv else ""))
    print("-" * 78)
    print(f"결과: 성공 {sum(len(v) for v in qc_records.values())} / 실패 {len(failures)} / Lot QC {len(outputs)}")
    for source, exc in failures:
        print(f"- {source.name}: {type(exc).__name__}: {exc}", file=sys.stderr)
    return 1 if failures else 0


def _console_safe():
    """cp949 콘솔에서 γ̇ 같은 문자가 출력 도중 분석을 중단하지 않게 한다."""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except Exception:
            pass


def main():
    _console_safe()
    p = parser()
    a = p.parse_args()
    if a.write_density_template:
        write_density_template(a.write_density_template)
        print(f"밀도 입력 양식 생성: {a.write_density_template}")
        return 0
    method = a.type.upper()
    if method in ("B", "QC") and (a.win_s <= 0 or not 0 < a.window_fraction <= 1):
        p.error("Method B 대표구간을 확인하세요.")
    if not 0 <= a.torque_min < a.torque_max <= 100:
        p.error("Torque 범위를 확인하세요.")
    try:
        set_active_policy(policy_from_args(a))
    except ValueError as exc:
        p.error(str(exc))
    if method == "QC":
        return run_qc_mode(a, p)
    try:
        inputs = sorted(collect_inputs(a.inputs, a.input_dir, a.recursive), key=lambda x: x.name.lower())
        template = resolve_template(a.template, method)
    except (ValueError, FileNotFoundError) as exc:
        p.error(str(exc))
    outdir = Path(a.outdir).expanduser().resolve()
    outdir.mkdir(parents=True, exist_ok=True)
    print("=" * 78)
    print(f"{VERSION} | {'Method A Step Flow' if method == 'A' else 'Method B 3ITT'}")
    print(f"입력: {len(inputs)}개 | 템플릿: {template.name} | 출력: {outdir}")
    if method == "A":
        print(f"STRICT 규칙: 회전 Step {METHOD_A_STEP_SECONDS:g}초, 후반 {METHOD_A_TAIL_SECONDS:g}초, H-B 최소 {METHOD_A_MIN_HB_POINTS}점/권장 {METHOD_A_RECOMMENDED_HB_POINTS}점, Torque {a.torque_min:g}–{a.torque_max:g}%")
        print_policy(active_policy())
        print("주의: --win-s/--window-fraction은 Method A 계산을 변경하지 않습니다.")
    else:
        print(f"3ITT 대표구간: max({a.win_s:g}초, 유지시간×{a.window_fraction:g}) | Torque {a.torque_min:g}–{a.torque_max:g}%")
    print("=" * 78)
    records, steps, failures = [], [], []
    for index, source in enumerate(inputs, 1):
        suffix = "StepFlow" if method == "A" else "3ITT"
        output = _unique_output(outdir, source.stem, suffix, a.overwrite)
        try:
            point_count, step_count, result = process_enhanced(source, method, template, output, a.win_s,
                                                                a.window_fraction, a.torque_min, a.torque_max)
            records.append(result.pop("_lot_record"))
            steps.extend(result.pop("_step_records"))
            cli_summary = result.pop("_cli_summary")
            result.pop("_qc_record", None)
            print(f"[{index}/{len(inputs)}] 완료 {source.name}")
            if method == "A":
                print_method_a_cli(source, output, point_count, step_count, result, cli_summary)
            else:
                print_method_b_cli(source, output, point_count, step_count, result, cli_summary)
        except Exception as exc:
            failures.append((source, exc))
            print(f"[{index}/{len(inputs)}] 실패 {source.name}: {type(exc).__name__}: {exc}", file=sys.stderr)
    lot_outputs = []
    if records and not a.no_lot_summary:
        grouped_records, grouped_steps = defaultdict(list), defaultdict(list)
        for record in records:
            grouped_records[record["lot"]].append(record)
        for record in steps:
            grouped_steps[record["lot"]].append(record)
        for lot, lot_records in grouped_records.items():
            output = write_lot_workbook(lot, lot_records, grouped_steps.get(lot, []), method, outdir)
            lot_outputs.append(output)
            print(f"[Lot 완료] {lot}: {len(lot_records)} repeats -> {output.name}")
            if method == "A":
                for metric in ("up_K_Pa_s_n", "up_n", "down_K_Pa_s_n", "down_n", "mean_stress_recovery_pct", "relative_hysteresis_pct"):
                    values = pd.to_numeric(pd.Series([r.get(metric) for r in lot_records]), errors="coerce").dropna()
                    if len(values):
                        rsd = values.std(ddof=1) / abs(values.mean()) * 100 if len(values) > 1 and values.mean() != 0 else np.nan
                        print(f"  [Lot 통계] {metric}: mean={_cli_num(values.mean())}, SD={_cli_num(values.std(ddof=1))}, RSD={_cli_num(rsd,2)}%")
    print("-" * 78)
    print(f"결과: 성공 {len(records)} / 실패 {len(failures)} / Lot {len(lot_outputs)}")
    print("Lot 종합 QC(A·B·밀도 통합, 지표 물성 설명 포함)는 --type QC로 실행하십시오.")
    if failures:
        for source, exc in failures:
            print(f"- {source.name}: {type(exc).__name__}: {exc}", file=sys.stderr)
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
