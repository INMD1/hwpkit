# 실제 데이터셋 변환 점검 — 2026-09-05

후속 페이지 분리 수정과 최신 검증은 [페이지 검사 보고서](pagination-audit-2026-09-05.md)를 참고한다. 아래는 첫 번째 수정 단계의 기록이다.


`hwpkit_research/datasets/input`의 원본 237개(확장자 기준 HWP 206, HWPX 31)를
라이브러리와 연구 실행 코드에서 각각 수정 전·후 새로 변환했다. 모든 입력에 한컴 DOCX 정답이 있었다.
원본/정답/기존 변환 결과는 수정하지 않았다.

| 측정 | 라이브러리 수정 전 | 라이브러리 수정 후 | 연구 실행 코드 수정 전 | 연구 실행 코드 수정 후 |
|---|---:|---:|---:|---:|
| 변환 및 본문 XML 열기 성공 | 237/237 | 237/237 | 236/237 | 237/237 |
| 직접 들여쓰기 검사 문단 | 18,434 | 18,434 | 18,379 | 18,434 |
| 직접 들여쓰기 불일치 | 12,780 | 400 | 12,725 | 400 |
| 비어 있지 않은 문단 텍스트 F1 평균 | 0.914610 | 0.914610 | 0.916388 | 0.914610 |

라이브러리의 들여쓰기 불일치는 96.87% 감소했고 226문서에서 개선됐다. 문서별 불일치가 증가한 경우는 없다.
연구 코드의 이전 성공 236문서에서는 225문서가 개선됐고 회귀는 없다.
연구 코드의 평균 F1 변화는 기존에 실패하던 `jygmlmdy`가 측정 대상에 추가된 결과다.
동일한 성공 문서끼리 비교하면 텍스트 F1과 문단·표·이미지·섹션 개수는 전부 동일하다.

수정 후 DOCX 474개는 모두 `python-docx.Document`에서도 열렸다.
기존 성공 문서들의 `word/document.xml`을 전후 대조한 결과, 변경은 `w:ind/@w:left`와
left만 가진 `w:ind`의 추가에 한정됐다. 그 외 본문 텍스트·구조·서식 XML은 동일하다.

## 수정 근거

1. **HWP/HWPX 내어쓰기의 기준 위치 오류**
   - `0d1tojvf`의 첫 앵커 문단은 한컴 DOCX에서 left/hanging=`2692/2692`다.
     기존 결과는 `0/2692`였다. HWP PARA_SHAPE의 raw left=`0`, indent=`-26920`을 확인했다.
   - HWP/HWPX 저장 left와 공통 모델의 몸체 여백을 구분했다.
     저장 단위 배율을 적용한 뒤 `bodyLeft = storedLeft - min(0, storedIndent)`로 읽고,
     `storedLeft = bodyLeft + min(0, firstLineIndent)`로 쓴다.
   - DOCX 인코더의 몸체 여백 의미는 유지한다. HWPX lineseg의 첫 줄/후속 줄 위치도 수정했다.
     음수 여백을 0으로 잘라 왕복 위치가 변하던 HWP 디코더 처리도 제거했다.
2. **압축 해제 결과 누락으로 전체 문서 변환 실패**
   - 연구 코드에서 `jygmlmdy`의 BinData 처리 중 undefined를 이미지 바이트로 읽어 실패했다.
   - Pako가 예외 없이 undefined를 반환하면 raw DEFLATE를 재시도하고,
     유효한 Uint8Array를 얻지 못하면 원본 바이트를 유지하도록 수정했다.
3. **데이터셋 검증 자체의 오류**
   - `.backup` 안의 과거 소스 테스트까지 실행되던 문제를 `runner/vitest.config.ts`의 src 범위 지정으로 해결했다.
   - 실제 샘플 회귀 테스트는 문단 번호 대신 같은 텍스트를 순서대로 대응시킨다.
     한컴의 추가 빈 앵커 때문에 기존 19번 문단은 서로 다른 내용을 비교하고 있었다.
     수정 후 내어쓰기 89문단을 비교하며 최소 80문단 검사를 요구한다.

변경은 양쪽 저장소의 관련 디코더/인코더에만 적용했고, 기존 독자 구현을 파일째 덮어쓰지 않았다.
`dist`와 연구 실행 번들도 재빌드·동기화했다.

## 회귀 검증

- 라이브러리 Vitest **96개 통과**, 타입 검사와 ESM/CJS/타입 선언 빌드 통과.
- 연구 runner Vitest **75개 통과**, 타입 검사와 빌드·번들 동기화 통과.
- 연구 Python **67개 통과**, 데이터셋 샘플 스킵 없음.
- 독립 HWPX fixture, HWP raw 여백 필드, 음수/0/양수 여백의 3회 왕복, Pako fallback 회귀 추가.
- `0d1tojvf`, `jygmlmdy`, `1feif858`, `3jh0imd6`의 사용 가능한 각 4방향,
  총 16 task로 6방향을 포함해 추가 검사했다. 두 코드 모두 수정 후 16/16 성공.
  기존 성공 task의 텍스트 F1·요소 개수 회귀와 들여쓰기 불일치 증가는 없다.
  연구 코드는 `jygmlmdy`의 HWP→DOCX/HWPX 실패 2건이 해소됐다.

## 재현과 결과 파일

연구 저장소에서 실행한다. 출력 디렉토리는 새 빈 디렉토리여야 한다.

```bash
python3 compare/audit_dataset_docx.py --bundle ../hwpkit/dist/index.mjs \
  --output datasets/docx_audit/library_new
python3 compare/audit_dataset_docx.py --bundle runner/dist/index.mjs \
  --output datasets/docx_audit/runner_new
python3 compare/audit_dataset_docx.py --bundle runner/dist/index.mjs \
  --direction docx2hwp --sample-id 0d1tojvf --output datasets/docx_audit/reverse_new
```

검사기는 기존 파일을 덮어쓰지 않고, 변환 실패 시 과거 산출물을 성공으로 재사용하지 않는다.
문서별 지표·실패·실행 번들 SHA-256은 각 출력 폴더의 `summary.json`에 저장한다.

- 저장소에 보관한 전후 집계·문서별 결과: [dataset-audit-2026-09-05.json](dataset-audit-2026-09-05.json)
- 생성 DOCX/전체 원시 보고서: `../../hwpkit_research/datasets/docx_audit/`
- 기준 커밋: hwpkit `7b10c990987158a3146059b1cb41f96d747f7f69`, research `1dd9475676bf4c75db6dbd2ce8eb9a01b97ff8fd`.

## 남은 한계

37문서에 직접 들여쓰기 불일치 400건이 남아 있다. 이 지표는 같은 텍스트 문단을 중복 발생 순서대로
대응시켜 직접 지정된 left/right/hanging/firstLine을 2 dxa 허용오차로 비교한다.
스타일·목록 상속을 해석하지 않으므로 남은 차이를 모두 실제 화면 오류라고 단정할 수 없다.
텍스트 F1은 빈 문단을 제외한 문단 텍스트 다중집합 지표이며 문서 순서·외관 전체 정확도가 아니다.

LibreOffice가 없어 PDF 렌더링 비교는 수행하지 않았다. HWP/HWPX 출력의 역방향 검사는
동일 라이브러리로 DOCX를 다시 생성한 결과이므로 한컴 등 독립 프로그램의 외관 검증을 대신하지 않는다.
텍스트 누락·추가 앵커·서식·페이지 배치의 기존 차이 전체를 해결하거나 LoRA 성능을 측정한 결과가 아니다.
