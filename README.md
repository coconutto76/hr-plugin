# hr-plugin

HR 위험 신호 분석과 대시보드 생성 플러그인입니다. 인사 데이터 엑셀로 직원·조직의 퇴직 위험 신호를 상/중/하로 정리하고, 그 결과를 브라우저에서 보는 대시보드로 만듭니다. 스킬 두 개가 들어 있습니다.

## 기능

| 스킬 | 하는 일 |
|---|---|
| `hr-risk` | 인사데이터 엑셀을 분석해 위험 신호를 상/중/하로 정리한 결과 엑셀, 요약 3줄, 판단 근거, 분석의 한계를 만듭니다. 이름·사번은 결과에 넣지 않고 분석용 ID로 행을 구분합니다. |
| `make_dashboard` | hr-risk 결과로 `dashboard.html`(내부용)과 `dashboard_share.html`(부서별 요약만 있는 공유용)을 만듭니다. 외부 주소 없이 열리는 단일 HTML입니다. |

## 입력과 출력

| 스킬 | 입력 | 출력 |
|---|---|---|
| `hr-risk` | `인사데이터_샘플.xlsx` | `퇴직위험_결과_YYYYMM.xlsx`, `ID매핑_YYYYMM.xlsx`(분석용 ID와 사번의 대응표, HR만 보관) |
| `make_dashboard` | 필수: `퇴직위험_결과_YYYYMM.xlsx` / 선택: 같은 기준일의 원본 `인사데이터_샘플.xlsx`, `ID매핑_YYYYMM.xlsx` | `dashboard.html`, `dashboard_share.html` |

- 출력은 사용자의 작업 폴더에 저장됩니다. 설치된 플러그인 폴더 안에는 쓰지 않습니다(`make_dashboard`는 이를 거절합니다).
- 원본 없이 결과 엑셀만으로도 대시보드가 만들어지고, 원본이 없어 표시할 수 없는 화면에는 이유가 적힙니다.

## 설치

```
/plugin marketplace add <이 폴더의 경로 또는 GitHub owner/repo>
/plugin install hr-plugin@hr-team
```

터미널에서는 `claude plugin marketplace add <경로>`, `claude plugin install hr-plugin@hr-team`을 씁니다.

필요한 환경: Windows, Windows PowerShell 5.1, **Microsoft Excel**(`hr-risk`가 결과 엑셀을 쓸 때 사용). `make_dashboard`는 Excel 없이 동작합니다.

## 실행 방법

작업 폴더에서 Claude Code를 열고 말로 요청합니다.

1. 분석: "인사데이터_샘플.xlsx를 분석해서 이번 달 위험 신호를 정리해줘" → `hr-risk`가 실행되어 결과 엑셀을 작업 폴더에 저장합니다.
2. 대시보드: "퇴직위험_결과_202604.xlsx로 대시보드 만들어줘" → `make_dashboard`가 실행되어 두 HTML을 만들고 `dashboard.html`을 엽니다. 직원별 야근·평가·인상률까지 보려면 같은 기준일의 원본과 ID매핑 파일 경로를 함께 알려 주세요.
3. 공유: 부서별 요약만 필요하면 `dashboard_share.html`만 전달합니다. 공유 전에 개인 정보가 없는지 검사하는 절차가 포함되어 있습니다.

완성된 공유용 화면의 예는 `examples/dashboard_share.html`입니다(합성 샘플 데이터, 부서별 요약만 있음). 파일을 브라우저로 열면 됩니다.

스킬을 직접 부를 때는 `/hr-plugin:hr-risk`, `/hr-plugin:make_dashboard`입니다.

## 점수의 한계

- 점수는 퇴사 확률이 아니라 **먼저 확인할 대상을 정하는 지표**입니다. "N%가 퇴사한다"로 읽으면 안 됩니다.
- 규칙 기반 가중 점수이며 가중치와 등급 경계는 초기 설정값입니다(`hr-risk/config/risk_config.json`).
- 샘플의 퇴사 건수가 적고 기준일이 하나뿐이라 **예측 성능을 확인할 수 없습니다.** 퇴사 이력이 없으면 사후 점검 자체를 하지 않습니다.
- 입력 데이터에 없는 사정(개인 사유, 외부 제안 등)은 반영하지 못합니다.
- 기준일 이후에 생긴 정보는 점수에 쓰지 않습니다. 이름·사번·병가·설문 미응답·성별·나이도 개인 점수에 쓰지 않습니다.
- 유지 우선순위는 위험 점수와 합치지 않고 HR이 따로 판단합니다.

## 개인 정보 주의

- `ID매핑_YYYYMM.xlsx`는 HR만 보관하고 결과 파일과 함께 공유하지 마세요.
- `dashboard.html`에는 분석용 ID별 정보가 들어 있으니 HR 내부에서만 보세요.
- 입력·결과 엑셀, `HR전용_ID매핑/`, 생성된 대시보드, 확인 기록 JSON, 테스트 산출물(`make_dashboard/test_output/`), `.claude/skills/` 테스트 사본은 로컬에만 둡니다. `.gitignore`가 이 파일들을 저장소에서 제외합니다.
- `.gitignore`는 git 저장소에만 적용됩니다. 폴더 경로로 마켓플레이스를 추가하면 폴더 전체가 쓰이므로, 실제 데이터가 든 폴더는 경로로 배포하지 마세요.
