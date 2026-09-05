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
