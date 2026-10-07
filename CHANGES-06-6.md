# HB-analysis-06-6 변경 기록

기준 파일은 `HB-analysis-06-5_integrated_reviewed.py`이다. 06-5와 그 이전 스크립트, 테스트, 템플릿 파일은 수정하지 않았다.

작업 공간에서 `CHANGES-06-5-4.md`, `SST_Gel_StepFlow_Template-3.xlsx`, `SST_Gel_3ITT_Template-2.xlsx`는 찾지 못했다. 스키마는 같은 폴더의 `SST_Gel_StepFlow_Template.xlsx`와 `SST_Gel_3ITT_Template.xlsx`를 직접 읽어 고정했다.

## 1. 해시

| 파일 | SHA-256 |
| --- | --- |
| HB-analysis-06-5_integrated_reviewed.py | `102ac1a1cabc5db4dd8cf0c0af00f1d72498516cb6916c06e6d8b0527907bbf3` |
| HB-analysis-06-6_integrated_reviewed.py | `c5e93632538302f5ed50c44b2712d989fdfdd9463d54faa58a2d51fdf4bcdf32` |

VERSION = `HB-analysis-06-6`  
QC_LAYER_VERSION = `LotQC-2.8`

## 2. 06-5 호출 관계 (수정 전 확인)

- `compute_a_qc`는 이미 `channel_view`를 거친 reps를 받는다. strict와 inclusive가 각각 한 번씩 호출된다.
- 공식 SSI는 그 reps로 `_apply_official_ssi`에 들어갔다. `eta_cP`, `tau_Pa`, `value_class`는 채널 뷰의 열이다.
- strict 뷰는 `included_in_strict`가 아니면 `n_valid`를 0으로 만든다. inclusive 뷰는 `n_inclusive`를 `n_valid`로 올린다.
- 06-5 `select_exact_rpm_step`은 `if len(branched): near = branched`라서 Up 후보가 없으면 Down 행을 그대로 남겼다. 적격 후보가 둘이면 Step 번호가 작은 행을 골랐다. 못 찾은 경우에도 `selected_step`에 측정 Step을 넣었다.
- `metric_alias`는 코드가 `A_SSI`이면 RPM과 무관하게 `A_SSI_QC_0P5_5`를 반환했다. 다른 함수는 이 함수를 호출하지 않았다.
- Lot 평균은 `_mean_eligible`의 status 문자열로 결정됐다. Excel inclusive 평균은 `value_class`를 다시 읽고 `VC_BOTH`를 제외 목록에 넣지 않았다.
- `--ssi-include-down`, `--ssi-primary-selection`은 policy에 저장됐지만 decade 탐색은 항상 Up이었다. pooling과 보간 옵션은 저장만 되고 계산을 바꾸지 않았다.
- 템플릿은 `prepare_workbook`가 시트 이름 검사 없이 고정 좌표에 썼다.

## 3. 수정한 함수

- `select_exact_rpm_step`, `channel_row_is_eligible`
- `discover_decade_rpm_pairs`, `select_primary_decade_pair`, `compute_exact_ssi_metric`, `_apply_official_ssi`
- `canonicalize_metric_record`, `metric_alias`, `aggregate_decade_pair_stats`
- `lot_inclusion_decision`, `_mean_eligible`, `lot_metric_stats`
- `channel_view` strict `n_valid` 정규화, `_QCBuffer._keep_without_value`에 `REVIEW`
- `write_lot_qc_workbook` helper 열과 Metric_Statistics 수식
- `write_decade_ssi_sheet`, `step_audit_rows`, `long_rows`, `export_long_csv`
- `policy_from_args`, `validate_template_schema`, `prepare_workbook`
- `evaluate_stabilization`은 선택기 교체 때 빠졌던 06-5 함수를 그대로 복원했다.

## 4. exact RPM branch 누수

원인: branch 필터 결과가 비면 필터 이전의 RPM 일치 행을 유지했다. Down 0.5 rpm이 Up 0.5 rpm 공식 SSI 분자가 될 수 있었다.

수정: moving, 0 rpm·REST 제외, nominal RPM, 요청 branch, turnaround 허용, 채널 적격, 중복 처리 순서를 고정했다. 요청 branch 후보가 없으면 `found=False`이고 다른 방향으로 바꾸지 않는다. `allow_turnaround=False`이면 shared turnaround를 제거한다.

## 5. 중복 RPM

적격 후보가 하나이면 그 Step을 쓴다. 측정 후보는 여러 개고 적격이 하나이면 `data_status=REVIEW`, `duplicate_count`는 측정 후보 수, 제외 사유를 남긴다. endpoint preference는 `allow_turnaround=True`이고 shared turnaround가 하나일 때만 적용한다. 그 뒤에도 적격이 둘 이상이면 `ambiguous=True`, 값은 NO VALUE이다.

## 6. legacy A_SSI

`canonicalize_metric_record`만 변환한다.

- measured Up 0.2/2.0 → `A_SSI_FIELD_0P2_2`
- measured Up 0.5/5.0, 또는 고점이 shared turnaround → `A_SSI_QC_0P5_5`
- provenance가 없거나 0.5/4처럼 다른 pair → `A_SSI_LEGACY_UNKNOWN`, Lot 통계 제외, acceptance EXCLUDED

코드 문자열만 보고 변환하는 `metric_alias("A_SSI")`는 `ValueError`이다. 구형 DB CSV를 이어 붙일 때도 같은 함수를 탄다.

## 7. VC_BOTH

`VC_BOTH = "torque 범위 밖 참고값"`. `TORQUE_REFERENCE_CLASSES`는 하한, 상한, 양쪽을 함께 담는다. `include_censored_in_lot_stats=False`이면 셋 다 평균에서 빠진다. True이면 inclusive 평균에 들어가고 acceptance는 EXCLUDED로 남는다. count는 `VC_BOTH`일 때 n CENS, n Low-TQ Ref, n High-TQ Ref가 각각 1 증가한다.

## 8. Python과 Excel 통계

포함 여부는 `lot_inclusion_decision` 한 곳이다. 반복 시트는 측정값을 유지하고, `<code>__stat_value`에 통계용 숫자만 넣는다. 포함되지 않으면 그 칸은 빈칸이다. Metric_Statistics의 COUNT, AVERAGE, STDEV, MIN, MAX, RANGE는 이 열만 참조한다. Excel 수식이 value class를 다시 판정하지 않는다.

## 9. CLI

- `--ssi-include-down`: Up decade에 더해 실제 Down branch pair를 만든다. 공식 SSI는 Up이다. Down pair에는 official metric link가 없다.
- `--ssi-primary-selection`: dynamic pair의 진단용 대표만 고른다. 공식 SSI 값은 바꾸지 않는다.
- `--ssi-pool-different-pairs`: `ValueError` — pair identity must be preserved.
- `--ssi-allow-interpolation`, `--ssi-allow-model-estimation`: `NotImplementedError` — measured RPM pairs only.

## 10. template schema

`validate_template_schema`가 Method A/B 결과 작성 전에 시트와 header를 확인한다. 불일치하면 `TemplateSchemaError`에 시트, 칸, 기대값, 실제값을 넣는다. 템플릿 파일은 덮어쓰지 않는다.

Step-flow Raw_Data 헤더는 5행 B–O, Step_Setup 헤더는 5행 B–H, 데이터는 6행이다. 3ITT Recovery 헤더는 5행 B–J이다. 좌표는 두 xlsx를 열어 확인한 값이다.

## 11. 테스트

`HB_ANALYZER_UNDER_TEST`에 06-6 절대경로를 넣고 `python3.11 -m pytest -q test_hb_analysis_06_6.py`를 실행했다. 22 passed, 2 warnings. 경고는 상수 열의 numpy corrcoef이다.

생성 파일:

- `results-06-6/HT-SYN-01_StepFlow.xlsx`
- `results-06-6/HT-3ITT-01_3ITT.xlsx`
- `results-06-6/HT-SYN_Lot_QC_INCLUSIVE.xlsx`

LibreOffice 재계산 후 `#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`, `#N/A`는 0건이다. Table displayName 중복과 범위 겹침은 0건이다.

## 12. 남은 제한

- 보간·모델 decade SSI는 이번 버전에서 구현하지 않고, 옵션을 켜면 즉시 오류이다.
- 서로 다른 RPM pair의 pooling은 거부한다.
- 공식 0.2/2와 0.5/5에 연결된 dynamic pair는 DecadeSSI에 보이되 Lot 평균에는 한 번만 들어간다.
- 지정 첨부 파일 세 개가 작업 공간에 없어, 스키마는 위 두 템플릿에서 읽었다.
