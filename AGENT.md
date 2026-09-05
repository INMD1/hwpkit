# hwpkit 작업 메모

2026-09-05 재작성. 사용자 갱신 기준 커밋: 887661ce0af118d4cf419cee0d24ff2aca29e610.
작업 전 git status와 이 파일을 읽고 필요한 소스·테스트만 연다.
상세 근거: docs/conversion-quality-audit-2026-09-05.md.

- 공개 API/형식: src/pipeline/Pipeline.ts, registry.ts, src/index.ts.
- HWP(5.x)/HWPX/DOCX/MD/HTML 입력·출력 지원. DOC/PDF/.hwpkit 형식 미지원.
- 5×5 기본 내용 왕복: src/pipeline/ConversionMatrix.test.ts. 외관 보존 인증이 아님.
- 자동 감지: ZIP 항목명, OLE FileHeader 서명 확인. WordDocument는 DOC로 거부.
  문자열 입력 기본 MD 유지. 명시 형식은 대소문자/선행 점 정규화. File 없는 Blob 환경 검사.
- 들여쓰기: src/decoders/docx/DocxIndentation.test.ts, src/encoders/docx/HangingIndent.test.ts.
  indentPt는 본문 여백. w:left에 hanging을 다시 더하면 매 왕복마다 여백이 늘어난다.
  left >= hanging을 강제하지 않는다. hanging 우선, 0/음수 보존, 부모 문단 속성 상속.
- 최신 Markdown 변경은 보존했다. source 사본을 파일째 덮어써 회귀시키지 않는다.
- 연구 사본 ../hwpkit_research/runner/src는 별도 구현이며 라이브러리와 다르다.
  이번 재작업에서는 연구 저장소를 수정하지 않았다. 필요하면 검증된 diff만 동기화한다.
- 명세 ../hwpkit_research/specs: HWP5 PDF13쪽/인쇄8쪽 표3,
  ECMA Part1 §17.3.1.12 PDF231쪽/인쇄221쪽. 대형 PDF 전체를 프롬프트에 넣지 않는다.
- 검사: npm test, npm run typecheck, npm run build. dist도 추적되므로 소스 수정 뒤 재생성한다.
  npm PATH: /home/leehojun/.nvm/versions/node/v24.16.0/bin.
- 검증: Vitest 86개, 타입 검사, ESM/CJS/d.ts 빌드. 실제 전체 코퍼스/모델 성능은 미측정.
- 다음: 독립 정답 기반 표/이미지/페이지 배치 및 전체 방향 비회귀 검사.

후속 수정마다 원인·관련 파일·명세 조항·검증 결과·남은 한계를 간결하게 갱신한다.

## 2026-09-05 실제 데이터셋 점검

- 보고서: `docs/dataset-audit-2026-09-05.md`. 원본 237개/한컴 DOCX로 전후 검증.
- HWP/HWPX 내어쓰기에서 저장 left는 모델 몸체 여백과 다르다. 디코더는
  bodyLeft = storedLeft - min(0, storedIndent), 인코더는 그 역변환을 수행한다.
  HWPX lineseg 위치도 수정. DOCX 인코더의 몸체 여백 의미는 유지한다.
- `0d1tojvf`의 한컴 left/hanging=2692/2692에 대해 기존 0/2692 오류 재현 후 수정.
  Pako undefined 반환 시 raw 재시도/원본 복귀도 적용. 연구 사본은 jygmlmdy 실패 해결.
- `src/decoders/HancomIndentation.test.ts`: 독립 HWPX fixture, HWP raw 여백,
  음수/0/양수 여백의 HWP·HWPX·DOCX 3회 왕복, 압축 fallback 회귀 10개 추가.
- Vitest 96 통과, 타입/빌드 통과. dist 재생성. 연구 사본에도 해당 diff만 반영.
- 237/237 변환 성공, 직접 들여쓰기 불일치 12,780→400, 226문서 개선·회귀 0.
  텍스트 F1·문단/표/이미지/섹션 개수 변화 없음. 4샘플의 6방향(16 task)도 전후 비교.
- 한계: 37문서/400건 직접 들여쓰기 차이 남음. 스타일/목록 상속·PDF 외관·한컴 역방향 출력 미검증.

## 2026-09-05 페이지 분리 후속 수정

- 보고서: `docs/pagination-audit-2026-09-05.md`. 원본/정답은 보존. 이전 들여쓰기 수정 상태와 전후 비교했다.
- DOCX에서 사용된 숫자 문단 스타일을 모두 정의(74문서 미정의 참조→0), 중복 pStyle 639건→0.
  빈 스타일 참조 때문에 LibreOffice가 중첩 표를 풀어 쓰던 문제를 실제 PDF로 재현했다.
- pageBreakBefore/keepNext/keepLines/widowControl과 명시적 false 상속을 HWP/HWPX/DOCX에 보존.
  HWPX 표 앞 나눔을 뒤 앵커에서 분리. HWP 첫머리의 명시적 나눔은 의도적 빈 페이지로 보존.
- HWP/DOCX 구역 디코딩 및 HWP/HWPX/DOCX 구역별 출력. 용지·본문·CFB SectionN 트리·HPF spine 보존.
  DOCX sectionType/differentFirstPage, meta.evenAndOddHeaders 및 구역별 머리말·꼬리말 참조 추가.
- HWP 음수 오른쪽 여백 보존. 연구 사본의 0.75 줄 간격 배율은 독립 한컴 DOCX와 달라 제거:
  160% = 모델 1.6 = DOCX line 384. 라이브러리의 기존 비율은 유지했다.
- `Pagination.test.ts`: 14개 회귀. 12구역/표/공유 이미지, 설정 상속, DOCX 3회 왕복 포함.
  Python `test_rendered_pagination.py`의 7개 검사는 LibreOffice에서 실제 쪽수/용지/셀 좌표 확인.
- 검증: 라이브러리 110개 + runner 89개 + Python 79개 통과. 타입·빌드·실행 번들 동기화 통과.
  237/237 DOCX 변환(양쪽), 474개 python-docx 열기. 4문서/6방향의 각 16 task 내용 손실 증가 없음.
- 라이브러리 PDF 237/237 렌더링 성공. 공통 236문서 쪽수 오차 합계 520→178, 정확한 쪽수 130→140문서.
  이전 mxs0o2pc는 180초 렌더링 시간 초과; 수정 결과 187쪽/같은 엔진 정답 188쪽.
  runner PDF 12문서 오차 96→32. 자동 빈 페이지도 보존(IsSkipEmptyPages=false)해 검사한다.
- 직접 들여쓰기 차이 400→386; 별도 스타일 상속/순서 대응 지표 407→220(26문서).
  22문서는 같은 엔진 정답 대비 쪽수 오차 증가. 원본 폰트/한컴 렌더러 및 정밀 외관 일치 미완료.
- 재현: 연구 `compare/audit_dataset_docx.py`, `compare/audit_pagination.py`의 새 빈 출력 디렉터리 사용.
  최종 결과: datasets/docx_audit/final_pagination_library, final_runner_spacing, final_directions_spacing;
  datasets/pagination_audit/final_full_library, runner_spacing. 번들 빌드와 변환을 동시에 실행하지 않는다.
