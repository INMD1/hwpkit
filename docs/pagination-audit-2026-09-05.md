# 페이지 분리 및 남은 데이터셋 오류 점검 — 2026-09-05

이전 들여쓰기 수정이 끝난 상태를 기준으로 추가 수정했다. 원본과 정답 파일은 보존했다.
최종 라이브러리와 연구 실행 코드 모두 **237/237 DOCX 변환 성공**이며, 생성된 474개를 python-docx로도 열었다.
상세 수치와 문서별 결과는 [JSON](pagination-audit-2026-09-05.json)에 있다.

## 수정 결과

| 검사 | 이전 | 수정 후 |
|---|---:|---:|
| 미정의 문단 스타일이 있는 문서 | 74 | 0 |
| 한 문단의 중복 스타일 참조 | 639 | 0 |
| 직접 들여쓰기 차이(18,434문단 검사) | 400 | 386 |
| 스타일 상속을 반영한 들여쓰기 차이(60,486문단 검사) | 407 | 220 |
| 라이브러리 쪽수 오차 합계(236문서) | 520 | 178 |
| 라이브러리 쪽수 일치 문서(236문서) | 130 | 140 |
| 같은 쪽에서 발견된 텍스트 앵커 비율 | 52.25% | 61.93% |
| 연구 실행 코드 쪽수 오차 합계(12문서) | 96 | 32 |

텍스트 문단 F1 평균은 양쪽 모두 0.914610으로 유지됐다. 직접 XML 비교와 스타일 상속 비교는
문단 대응 방식과 검사 범위가 달라 별도 지표로 기록했다. 스타일 상속 비교는 문서 순서대로 같은 텍스트를
맞추며 목록/글자 단위 들여쓰기는 해석하지 않는다.

라이브러리 쪽수 오차는 53문서에서 감소, 22문서에서 증가, 161문서에서 동일했다.
수정 후 PDF는 **237/237 렌더링 성공**이다. 이전 결과 중 `mxs0o2pc`는 180초 시간 초과로,
전후 합계에는 성공한 공통 236문서만 포함했다. 이 문서의 수정 후 출력은 187쪽이고 같은 환경의 정답은 188쪽이다.

## 쪽수 사례

| 문서 | 이전 라이브러리 | 수정 라이브러리 | 정답 DOCX를 같은 환경에서 렌더링 |
|---|---:|---:|---:|
| `0pogh0zg` | 59 | 18 | 16 |
| `853letq3` | 60 | 40 | 44 |
| `0d1tojvf` | 28 | 29 | 30 |
| `1feif858` | 18 | 18 | 17 |
| `dw23xg34` | 6 | 6 | 6 |
| `mxs0o2pc` | 시간 초과 | 187 | 188 |

제공된 PDF와의 비교도 별도로 보관했다. 해당 PDF가 있는 공통 232문서의 쪽수 오차 합계는
1,455→1,012쪽이다. 원본 PDF와 정답 DOCX의 재렌더링 결과도 서로 다르므로 두 기준을 섞지 않았다.
예를 들어 `0pogh0zg`는 제공 PDF 7쪽, 정답 DOCX 재렌더링 16쪽이다.

## 원인과 코드 변경

- **표 구조 붕괴:** DOCX 스타일 정의가 숫자 ID 0~33까지만 있었으나 실제 HWP에는 44 등 더 큰 ID가 있었다.
  사용된 스타일을 본문·중첩 표·머리말/꼬리말까지 수집해 정의하고, 목록 문단의 중복 `w:pStyle`을 제거했다.
  같은 표에서 스타일 정의만 보완하는 독립 실험으로 붕괴 원인을 확인했다.
- **쪽 나눔과 문단 보호:** HWP PARA_SHAPE 비트 16~19, HWPX breakSetting,
  DOCX의 pageBreakBefore/keepNext/keepLines/widowControl을 모델과 출력에 보존한다.
  상속된 설정을 명시적인 false로 해제하는 경우도 보존한다. 본문 앞의 쪽 나눔에 빈 줄을 더하지 않는다.
- **표 앞의 쪽 나눔:** HWPX에서 표를 포함한 문단의 나눔이 표 뒤 앵커에 붙던 오류를 고쳤다.
  표 앞의 최소 높이 앵커로 옮기고, 원래 앵커에서 중복 나눔을 제거한다.
- **의도적인 첫 빈 페이지:** 문서 첫머리의 명시적 HWP 페이지 나눔은 run break로 보존한다.
  Word가 첫 문단의 pageBreakBefore를 무시하는 차이를 실제 `dw23xg34`, `yqk7q62u`로 확인했다.
- **구역별 본문/용지:** HWP BodyText/SectionN과 DOCX sectPr를 개별 구역으로 보존한다.
  HWP/HWPX 인코더도 모든 구역을 각각 저장하며 섹션 수, CFB 하위 트리, HPF manifest/spine을 갱신한다.
  첫 구역만 HWPX로 내보내거나 HWP의 여러 용지를 한 설정으로 합치던 오류를 고쳤다.
- **DOCX 구역 설정:** 연속/새 쪽/홀수/짝수/다음 단 시작과 구역별 용지 크기·여백,
  머리말/꼬리말의 기본/첫 쪽/짝수 쪽 참조를 보존한다. 비활성인 첫 쪽·짝수 쪽 설정을 임의로 활성화하지 않는다.
- **음수 오른쪽 여백:** HWP 디코더의 0 제한을 제거해 원본 값을 보존했다.
- **연구 코드 줄 간격:** 임의의 0.75 배율을 제거했다. 실제 한컴 정답의 수천 문단을 대조하면
  HWP/HWPX 160%는 DOCX `w:line=384`이며, 이전 연구 코드는 288로 줄이고 있었다.
  정방향·역방향의 비율 변환을 함께 고쳤다. 라이브러리의 기존 비율은 이미 정답과 일치했다.
- **PDF 검사기:** 자동 삽입된 빈 페이지가 사라지지 않도록 IsSkipEmptyPages=false로 검사한다.
  시간 초과 시 해당 변환의 프로세스 그룹만 종료해 LibreOffice 자식 프로세스를 남기지 않는다.

DOCX 구역의 소속과 시작 방식은 [Microsoft SectionProperties](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.sectionproperties)
및 [SectionType](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.sectiontype)를 따랐다.
PDF 빈 페이지 옵션은 [LibreOffice 공식 설명](https://help.libreoffice.org/latest/en-US/text/shared/guide/pdf_params.html)을 확인했다.

## 검증과 재현

- 라이브러리 Vitest 110개, 연구 실행 코드 Vitest 89개, Python 79개 통과. 양쪽 타입 검사 및 번들 빌드 통과.
- `Pagination.test.ts`: 독립 DOCX/HWPX fixture, 명시적 false 상속, 스타일 참조, 첫 빈 페이지,
  12개 구역의 용지·표·공유 이미지 보존과 DOCX 3회 왕복을 검사한다.
- `tests/test_rendered_pagination.py`: LibreOffice에서 실제 쪽 나눔, 연속 구역, 홀수 쪽의 빈 페이지,
  가로/세로 용지 및 두 셀의 좌표를 검사한다. LibreOffice 미설치 환경에서는 이 7개를 skip한다.
- 4개 실제 문서에서 가능한 6개 변환 방향을 비교했다. 양쪽 각 16 task 성공이며
  텍스트 F1·표·이미지 개수의 감소는 없다. 필요한 빈 앵커와 구역 개수 변화는 별도로 검토했다.
- 렌더러: `LibreOffice 26.2.5.2 620(Build:2)`. Nanum/Noto CJK 설치, 같은 폰트 환경에서 전후와 정답을 렌더링했다.
- 최종 라이브러리 DOCX 237개의 모든 레이아웃 ZIP 항목은 렌더링 입력과 바이트 단위로 동일함을 확인했다.
- 실행 번들은 `ensure_runner_built()`로 빌드·테스트·동기화했다. 번들을 직접 수정하지 않았다.

연구 저장소에서 새 빈 출력 디렉터리로 실행한다.

```sh
PYTHONDONTWRITEBYTECODE=1 hwpkit-env/bin/python -m compare.audit_dataset_docx --bundle runner/index.mjs --output datasets/new_docx_audit
PYTHONDONTWRITEBYTECODE=1 hwpkit-env/bin/python -m compare.audit_pagination --candidate datasets/new_docx_audit --output datasets/new_pagination_audit
PYTHONDONTWRITEBYTECODE=1 hwpkit-env/bin/python -m pytest tests -q
```

최종 산출물(연구 저장소 기준):

- `datasets/docx_audit/final_pagination_library/summary.json`
- `datasets/docx_audit/final_runner_spacing/summary.json`
- `datasets/docx_audit/final_directions_spacing/summary.json`
- `datasets/pagination_audit/final_full_library/summary.json` 및 `candidate/*.pdf`
- `datasets/pagination_audit/runner_spacing/summary.json` 및 `candidate/*.pdf`

## 남아 있는 차이

**전체 원본과 동일한 페이지 배치가 완성된 상태는 아니다.** 26문서에 스타일 상속 기준 들여쓰기 차이
220건이 남아 있으며 문서별 목록을 JSON에 기록했다. 실제 문서 일부는 여전히 쪽수와 텍스트 위치가 다르다.
대표적으로 `oaj5rbm3`의 라이브러리 결과는 96쪽, 정답 DOCX 재렌더링은 119쪽이다.
표 구조가 복원된 결과라도 다른 줄바꿈·글꼴·행 높이·도형 배치 차이가 남을 수 있다.

모든 원본 글꼴(굴림, 휴먼명조 등)과 한컴 렌더러를 확보한 환경은 아니다.
쪽수/텍스트 앵커 지표는 픽셀 외관 일치 인증이 아니며, HWP/HWPX 출력은 한컴 프로그램에서 열어 확인하지 않았다.
연구 실행 코드의 PDF 검사는 12개 문서이고, 전수 PDF 검사는 라이브러리 237개 문서에 수행했다.
