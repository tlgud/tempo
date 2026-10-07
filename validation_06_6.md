# HB-analysis-06-6 검증

## 테스트 대상

| 항목 | 값 |
| --- | --- |
| 절대경로 | `D:\Data\Users\Sihyoung Park\OneDrive\현재 사용중\바탕 화면(old)\projects\연구소\업무 요청\연구 지원\라벡스DX\DVNext-methods\HB-analysis-06-package\HB-analysis-06-6_integrated_reviewed.py` |
| 파일명 | `HB-analysis-06-6_integrated_reviewed.py` |
| VERSION | `HB-analysis-06-6` |
| QC_LAYER_VERSION | `LotQC-2.8` |
| SHA-256 | `c5e93632538302f5ed50c44b2712d989fdfdd9463d54faa58a2d51fdf4bcdf32` |

06-5 원본 SHA-256은 `102ac1a1cabc5db4dd8cf0c0af00f1d72498516cb6916c06e6d8b0527907bbf3`이다.

테스트는 `HB_ANALYZER_UNDER_TEST`에 적힌 경로만 import한다. 파일명으로 다른 스크립트를 고르지 않는다.

## 첨부 파일

다음 세 파일은 작업 공간과 DVNext-methods 아래에서 찾지 못했다.

- `CHANGES-06-5-4.md`
- `SST_Gel_StepFlow_Template-3.xlsx`
- `SST_Gel_3ITT_Template-2.xlsx`

스키마는 아래에 있는 템플릿을 열어 기록했다. 템플릿 파일은 수정하지 않았다.

- `SST_Gel_StepFlow_Template.xlsx`
- `SST_Gel_3ITT_Template.xlsx`

## Step-flow header

| 시트 | 칸 | 값 |
| --- | --- | --- |
| Overview | B2 | SST gel 종합 분석 Overview |
| README | B2 | SST gel Step Flow QC Template |
| QC_Summary | B2 | SST gel Step Flow — QC Summary |
| QC_Summary | B4 / C4 | 원본 파일 / 내용 |
| Step_Setup | B5–H5 | Step, 방향, RPM, 전단률(1/s), 유지시간(s), 정량사용, 상태 |
| Raw_Data | B5–O5 | 구분/Interval … 유효성 |
| Step_Results | B5–Q5 | Step … 상태 |
| Flow_Profile | B5, N5 | 프로파일, 상태 |
| Model_Validation | B2 | SST gel 모델 검증 |
| StartUp | B2 | SST gel 저전단 Start-up |

데이터 시작은 Raw_Data, Step_Setup, Step_Results의 B6이다. Raw_Data 용량은 5000행이다.

## 3ITT header

| 시트 | 칸 | 값 |
| --- | --- | --- |
| README | B2 | SST gel 3ITT QC Template |
| QC_Summary | B2 | SST gel 3ITT — QC Summary |
| Step_Setup | B5–H5 | Interval, 구간, RPM, 전단률(1/s), 유지시간(s), 원본 Step, 출처 |
| Recovery | B5–J5 | 구간, RPM, γ(1/s), 유지(s), τ(Pa), η(cP), Torque(%), 유효 n, 상태 |

## 실행

```
HB_ANALYZER_UNDER_TEST=<06-6 절대경로>
python3.11 -m pytest -q test_hb_analysis_06_6.py
```

결과: 22 passed, 2 warnings. 경고는 상수 계열의 numpy corrcoef이다.

`python3.11 -m py_compile HB-analysis-06-6_integrated_reviewed.py` 통과.

## 생성 workbook

LibreOffice로 재계산한 뒤 수식 오류, Table 이름 중복, Table 범위 겹침은 없었다.

- `results-06-6/HT-SYN-01_StepFlow.xlsx`
- `results-06-6/HT-3ITT-01_3ITT.xlsx`
- `results-06-6/HT-SYN_Lot_QC_INCLUSIVE.xlsx`

합성 Step-flow의 점도는 torque에 비례한다. 이 파일의 공식 SSI는 0.2/2와 0.5/5에서 각각 10이다. 점도를 240, 180, 150, 100으로 고정한 단위 테스트에서는 field SSI 1.6, QC SSI 1.8이다.
