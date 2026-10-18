# hwpkit 작업 메모

2026-09-05 재작성. 사용자 갱신 기준 커밋: 887661ce0af118d4cf419cee0d24ff2aca29e610.
작업 전 git status와 이 파일을 읽고 필요한 소스·테스트만 연다.
상세 근거: docs/conversion-quality-audit-2026-09-05.md.

- 공개 API/형식: src/pipeline/Pipeline.ts, registry.ts, src/index.ts.
- HWP(5.x)/HWPX/DOCX/MD/HTML 입출력, DOC 입력 지원(아래 2026-09-08 참고). DOC 출력/PDF/.hwpkit 미지원.
- 5×5 기본 내용 왕복: src/pipeline/ConversionMatrix.test.ts. 외관 보존 인증이 아님.
- 자동 감지: ZIP 항목명, OLE FileHeader/WordDocument 서명 확인. DOC는 독립 디코더로 처리.
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

## 2026-10-18 HWP/HWPX→DOCX 불필요한 밑줄·취소선 수정

- 사용자 원본 입찰공고 HWP, SHA256 `228b8272637dd2ed89a3cad7472570ba4acf45231cc9b3d6dce9df80d7d0416d`.
  상세·원시 감사·전후 렌더: `docs/decoration-fix-2026-10-18/`.
- CHAR_SHAPE `0x3c0400f8/fa` 69개를 밑줄+취소선으로 오판. 구형 underline type2는
  line-through로 이관해 읽고 선모양15 비표시를 적용. 실제 밑줄1/3 및 취소선은 보존.
  이전 REA MCP sub_64f5f0 증거 재사용·retained-rea-evidence.json 보존, 새 세션 실행 아님.
- HWPX 인라인/헤더의 0/false/off/none 및 대소문자 정규화, 기존 strikeout3D OFF 유지.
  DOCX 인코더는 명시적 underline false를 w:u none으로 출력.
- 사용자 기존 output(2).docx의 밑줄/취소선 각239→0, XML텍스트4745자 전후동일.
  대응3701자 진하게/기울임 불일치0. ESM/CJS 실제 원본 변환 및 네이티브 HWPX도 확인.
  다운로드 `입찰공고서-밑줄취소선수정.docx`. 원본·이전파일 보존.
- 회귀28개 포함 전체222 통과/선택LO1 생략, 타입·라이브러리·playground 빌드, dist 갱신.
  전후DOCX PDF10쪽, 각쪽 공백제외추출텍스트동일, 첫쪽/끝쪽 화면 확인.
- 한계: 원본의 기존17자 미대응·생성HWPX경유 목록/필드45자추가는 별도. 이번 장식 수정은
  기존직접출력 텍스트를 바꾸지 않음. 전체배치/장식선 세부모양 일치 인증은 아님.

## 2026-10-18 Node 공개 API 출력 확인

- `tools/verify-code-output.mjs SOURCE.docx OUTPUT_DIRECTORY`: dist 공개 exports의 ESM/CJS를
  실제 실행, Buffer 명시/자동 및 async Blob 변환, 출력 자동 감지 재열기까지 확인.
- 사용자 DOCX→HWP/HWPX 4개 + 코드 직접 작성 모델4개 저장. 텍스트/글자별강조/구역/
  용지·여백/이미지크기·바이트·기준위치 모두 통과. 모델은 한글·영문 언어별폰트/11pt/
  강조OFF/그림/표 포함. 독립Python 감사도 원본7045자·모델51자 동일, 강조 불일치0.
- 원시 HWP/HWPX: 20mm 본문여백, 차트322.8×162.2pt/사진481.9×321.35pt 및
  inline/column-center/para-top, 첫줄 -21.6pt/18pt 확인. 실제 파일은 다운로드의
  `hwpkit-code-check-20261018/`. 보고서 `docs/code-output-check-2026-10-18/`.
- 새 발견/한계: Node 차트 라벨은 내장 bitmap 글꼴이어서 브라우저/원본과 모양이 다름.
  차트PNG·한글라벨 경고 제약 유지. 본문폰트와는 별개. 한컴 전체페이지 외관 미검증.
- 런타임 소스 변경 없음. 전체194 통과/선택LO1 생략, 타입 검사 통과.

## 2026-10-18 사용자 DOCX 차트·첫 줄·본문 여백 후속 수정

- 같은 사용자 원본, 상세: `docs/font-image-fidelity-2026-10-18/docx-layout-followup.md`.
- 빈 PNG+떠 있는 개체로 처리하던 DrawingML 차트를 캐시 데이터로 PNG 렌더링하고
  원래 inline/anchor를 보존. `toolkit/ChartRaster.ts`, DOCX 디코더 비동기 사전 준비 추가.
  2D 막대/열·일반 선 지원. 편집 가능한 차트는 아니며 복합/3D 등은 경고·대체 개체.
- HWP cached line horzpos=0, HWPX 음수 clamp 제거. 첫 줄/본문/오른쪽 들여쓰기와
  실제 줄 폭 반영. numbering.xml 수준 들여쓰기, 직접 0, numFmt=none/numId=0 수정.
- 그림 높이 자체에 줄 간격 비율을 곱하지 않고 글꼴의 추가 간격만 더함. 기준선 하단 반영.
  빈 문단 기호 크기를 상속하고 10pt보다 작은 명시 크기 보존.
- 원시 HWP/HWPX 독립 감사: 차트 32280×16220, 사진 48190×32135 HWPUNIT,
  JPEG 원본 동일. 상하좌우 본문 가장자리 5670(약20mm). 기존 머리말 영역 분할 의미 유지.
- Chrome에서 새 `file-sample_1MB-배치수정.hwp/.hwpx` 저장. 7045자 내용 동일·강조 불일치0.
  기존 HWPX 추가7자는 목록이 아니라 빈 차트의 `[차트: 차트]` 대체 텍스트였고 이제 제거됨.
- 새 회귀13개, 최종194 통과/선택LO1 생략. 타입·라이브러리·playground 빌드, dist 재생성.
  번들LO는 HWP5/HWPX 렌더러가 없어 전체 페이지 PDF 외관 검증은 하지 못함.
  임시 브라우저 검사 소스/원본 복사본 제거. 사진/차트/줄의 정밀 외관 인증은 아님.

## 2026-10-18 DOCX 일반 글자에 진하게/기울임이 추가되는 후속 확인

- 사용자 원본: `/home/leehojun/다운로드/file-sample_1MB.docx`, SHA-256
  `27cd24f7f6e1e86449c1efc75c103acbb717733be5a36377cceb59e77be9d97c`.
- 다운로드의 output (2).hwp / output (1).hwpx에 원본과 다른 강조가 실제로 있었다.
  원본 `w:b/w:i val="false"` 48개씩. 공백 제외 7,045자 독립 비교에서 원본에 없는
  진하게 6,461자·기울임 6,534자 추가. 최신 src/dist 직접 변환은 두 항목 모두 0.
- 실제 Chrome의 DocRoot 검사도 구버전 b=true/i=true를 표시. 새로 로드한 같은 서버의
  변환기로 원본을 읽어 새 HWP/HWPX를 브라우저에서 저장하고, 실제 레코드로 0건 확인.
  일반/진하게/기울임/다시 일반 4문구×3단계 12개도 UI에서 확인했다.
- Vite의 변환기 등록 부수효과가 남지 않도록 src 변경 시 전체 페이지 갱신을 명시했다.
  테스트 파일 변경은 제외. 기존 수정된 OFF 처리와 나머지 라이브러리 변경은 보존.
- 회귀 8개 추가(미지정/0/false/off × HWP/HWPX). 원본에 없는 강조를 금지하면서
  원래 bold-only/italic-only는 보존. 원시 CHAR_SHAPE/HWPX charPr도 검사한다.
- 독립 감사 도구 `tools/audit-text-emphasis.py SOURCE.docx OUTPUT.hwp OUTPUT.hwpx`.
  근거/한계: `docs/font-image-fidelity-2026-10-18/docx-emphasis-followup.md` 및 비교 JSON.
- 최종 181 통과/선택적 LO 1 생략, 타입·라이브러리·playground 빌드 통과.
  새 수정본은 다운로드의 `file-sample_1MB-서식수정.hwp/.hwpx`. 원본/기존 출력 보존.
  HWPX에 기존 목록 표시 문자 7개가 더 있으므로 전체 텍스트 무손실 인증은 아니다.

## 2026-10-18 REA MCP·글꼴/그림 위치 수정

- 근거·전후 화면·선택 전체 Evidence: `docs/font-image-fidelity-2026-10-18/`.
- REA 6.5.0 MCP/Hopper로 같은 libCoreEngine.so의 sub_64f5f0(CHAR_SHAPE),
  sub_64fe20(FACE_NAME), sub_6541e0(문단 내 확장 제어)를 확인. 기본 준비 제한 초과 후
  별도 서버 환경변수 REA_HOPPER_STARTUP_TIMEOUT_MS=600000으로 성공, 앱 설정 변경 없음.
- HWP/HWPX의 언어별 글꼴 강제 대체 제거. 모델 `fontFaces` 추가, 언어별 로컬 ID 조회 및
  DOCX ascii/hAnsi/eastAsia 보존. DOCX 기본 글자 속성·부모 서식 상속과 직접 OFF 값 수정.
- HWP 기호의 폰트/굵기/공백 추측 변경 제거. 표35에 맞게 밑줄 2–3, 첨자 15/16비트 수정.
  kerning 등 상위 속성 비트로 밑줄/취소선을 버리지 않음. 음영도 문자 모양 키에 반영.
- HWP 본문·셀 그림을 원래 문단의 글자 위치에 출력/복원. 셀 문단 수와 높이도 반영.
  page=PAPER, margin=PAGE 구분 및 0/음수 좌표 보존(HWP/HWPX). 불필요한 그림 문단
  제거로 demo 출력 344→340문단. 기존 빈 문단/빠진 셀 보충 수정은 보존.
- 검증: 새 회귀 14개 포함 173 통과/선택적 LO 1 생략. 타입·라이브러리·playground 빌드,
  dist 재생성. 실제 한컴 HWP/HWPX/DOCX 서식 fixture + demo HWP 열기·5초 유지;
  전후 화면에서 본문/셀 그림의 글자 사이 위치와 종이 기준 좌표 확인.
  재현 생성기: `node tools/create-font-image-fixture.mjs OUTPUT_DIRECTORY`.
- 한계: 특정 사용자 원본 미지정, 237개 전체 외관 인증 아님. 회전/자르기/복잡한 겹침,
  정밀 줄바꿈 및 없는 글꼴 대체는 미완료. 기존 문서는 재생성 필요.

## 2026-10-17 REA MCP·실제 한컴 HWP 열기 수정

- 보고서: `docs/hwp-opening-audit-2026-10-17.md`, 전체 선택 Evidence·실행 결과는 같은 날짜의 `hwp-opening-rea-*.json`.
- 제공된 hoffice-build의 libCoreEngine.so를 REA 6.5.0 MCP/Hopper로 분석하고 Xvfb에서 실제 실행.
  sub_6547a0은 nChars=1의 CR을 내부 생성하고 바로 태그 68을 읽는다. 빈 문단에 태그 67을
  추가하던 HwpEncoder 때문에 ReadRecord 예외/손상 파일 오류 발생. 빈 PARA_TEXT 생략 및 verifier 수정.
- 표 셀의 빈 문단도 CHAR_SHAPE를 보존. 본문·셀·중첩 셀의 빈 줄/24pt 및 잘못된 태그 회귀 검사.
- demo.docx 달력 첫 행이 14열 중 13열만 덮어 한컴 sub_5519f0(+0x551c53)에서 null 셀 참조로 충돌.
  셀 주소 계산에 세로 병합 점유 반영, 빠진 칸은 빈 셀 보충, 완전히 덮인 행은 셀 시작 수 0.
  verifier에 표 범위/겹침/빈 칸/행별 셀 수 검사를 추가. 입력 모델 불변. 달력 보충으로 343→344문단.
- 검증: Vitest 159 통과/선택적 LibreOffice 1 생략, 타입·라이브러리·playground 빌드 통과, dist 재생성.
  fixture 출력 11개 + 빈 셀/중첩 표/2구역·세로 병합·행 전체 병합 3개 모두 실제 한컴 열기 후 5초 유지.
  전체 demo는 화면 표시 후 1분 이상 유지. 재현 도구 `tools/check-hancom-open.py`는 제목만 보고 성공 처리하지 않는다.
- 한계: 초기 배치 확인이며 전체 외관·다른 한컴 버전·237개 코퍼스 전수 검증은 아님. 기존 생성 파일은 재생성 필요.

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

## 2026-09-05 playground 동기화

- 기존 Vite alias가 `src/index.ts`를 직접 참조하므로 들여쓰기·페이지 분리 수정은 이미 적용된다.
  `playground/main.ts`의 입출력 선택지를 같은 registry에서 생성해 HTML 입력/HWP 출력을 반영했다.
- 직접 입력의 MD 강제 처리를 제거하고 선택 형식을 존중한다. 바이너리 파일과 자동 모드는
  Pipeline의 내용 감지를 사용하며, MD/HTML 파일은 편집 가능하게 연다. 파일 읽기 오류도 표시한다.
- 트리에 props 외의 meta/dims/sectionType/differentFirstPage 및 headers/footers를 표시한다.
  셀 병합·이미지 배치·페이지 번호 형식도 표시하며 이미지 base64 본문은 길이만 보여준다.
- 탭 전환 함수의 타입 오류와 비활성 텍스트가 계속 표시되는 CSS를 수정하고 HTML iframe을 sandbox 처리했다.
- `playground/tsconfig.json`, `playground:typecheck`, `playground:build` 추가.
  검증: 기존 Vitest 110개, 라이브러리 타입·빌드 및 playground 타입·Vite 빌드 통과.
- 내장 demo.hwp/demo.hwpx 자동 감지·트리 검사·5종 출력(총 10변환), 명시 HTML 조각·자동 HTML 입력도 통과.
- 명세 변경 없음(UI/API 연결 수정). HTML은 실제 페이지 배치 렌더러가 아니며 브라우저 수동 검증은 미실시.

## 2026-09-05 한컴 원본 PDF 위치 보정

- 보고서: `../hwpkit_research/docs/pdf-layout-review-2026-09-05.{md,json}`. 사용자 목표 외관 94%는 미달.
- HWP 빈 문단의 PARA_CHAR_SHAPE 보존, DOCX 빈 문단 pPr/rPr 출력. 독립 2pt/24pt 빈 줄의 실제 y 위치 검사.
- PageDims는 DOCX와 같은 종이→본문/머리말 거리로 정규화. HWP/HWPX top+header=mt, top=headerPt.
  역변환은 영역을 분리, 명시적 0 보존. DOCX header/footer=0을 undefined로 버리던 디코더도 수정.
  기존 PageDims를 HWP 저장 값으로 직접 만들던 호출자는 본문 기준으로 조정해야 함. 모델 주석 갱신.
- 근거: HWP 명세 PDF59쪽 표131 + 독립 한컴 DOCX 268구역 전부에서 여백 합산/분리 관계 검증.
  세로 여백 4종 각각 268/268 일치. 원본 237/237 재변환, 기존 내용/들여쓰기 감사 행과 차이 없음.
- 검증: Vitest 114, 타입/ESM/CJS/d.ts 빌드, playground 타입/Vite 빌드 통과. 연구 Python 119 통과.
  PageMargins.test.ts의 독립 HWPX·raw PAGE_DEF·0값·3회 다방향 왕복 회귀 추가.
- 실제 한컴 PDF 128개: 엄격 위치 F1 평균 0.0239%→0.0442%, 인쇄 영역 15.07%→15.93%; 94% 통과 0.
  쪽수 오차 628→726 증가, 점수 하락 문서 20개. 여백 값 정확성과 외관 비회귀는 별개이며 외관 완료로 보지 않음.
  나머지 109개는 한컴 기준 PDF 미확보. 표·줄 간격·폰트·개체 배치 후속 필요, 자동 품질 승인/학습 승격 없음.
- 원본/정답 및 기존 playground 사용자 변경 보존. 연구 runner는 일부 대응이 이미 있어 파일째 동기화하지 않음.

## 2026-09-08 DOC 입력·HWP/HWPX 열기 및 변환 회귀 수정

- 보고서: `docs/opening-and-doc-audit-2026-09-08.md`. 기존 작업 중이던 줄바꿈 변경 의도를 보존했다.
- DOC CLX/piece table 디코더·registry/API/파일 선택 추가. 기본은 본문 읽기 + 손실 경고.
  `configureDocConverter` + 별도 `hwpkit-dev/node`의 `createLibreOfficeDocConverter`로 서식·표·그림 처리.
  로컬 playground는 LibreOffice POST 경로, 정적 빌드는 본문 모드. DOC 출력/암호화 DOC는 미지원.
- HWP PARA_HEADER bit31은 유효성 표시가 아닌 문단 목록 끝. 구역·셀별 마지막에만 설정하고 verifier 수정.
  탭은 8 WORD 인라인 제어로 기록해 뒤 7글자 누락 방지. 근거: HWP §4.1/표58 + 한컴 원본 문단 패턴.
- HWPX 공식 `hc:intent`, `hp:t/hp:lineBreak`, `hp:t/hp:tab` 사용. 혼합 텍스트 순서 보존.
  공식 한컴 margin.cpp/t.cpp 및 데이터셋 원본과 대조. 이전 indent/br 출력도 읽기 호환 유지.
- DOCX 탭·링크 글자/관계 보존. 독립 demo.docx 왕복 7,218→6,922 글자 누락을 7,218로 복구.
  MD 제목/표 인라인 서식 중복과 HTML 엔티티/공백 오류, CFB 순환 참조 및 v4 원점도 수정.
  MD 링크 표시 글자와 HTML table fallback 재입력을 처리해 출력 본문의 태그 노출 방지.
- File Office 입력은 명시 형식이 없으면 내용으로 감지. 확장자가 HWP인 실제 HWPX 샘플 확인.
  PDF/RTF bytes를 MD로 오인하지 않는다. 직접 지정한 형식은 계속 존중한다.
- 검증: 실제 LibreOffice DOC 통합 포함 Vitest 149개, 타입·ESM/CJS/d.ts·playground 빌드 통과.
  패키지 양쪽 진입점과 npm pack dry-run 검사. 추적 dist와 새 node 진입점 모두 생성.
- H2Orestart 0.7.14 + LibreOffice로 생성 HWP/HWPX/DOCX 각각 PDF 열기 및 한글·표·코드 줄·탭 뒤 내용 확인.
  실제 playground DOC 업로드도 성공. 추가 fixture 55/55 변환, 정규화 텍스트 일치 42→50.
- 실제 237문서 × 5형식 = 1,185건 생성·재읽기 실패 0, 공백 제외 텍스트 일치 849건.
  최종 MD 변경 뒤 해당 237건 재검사. 상세 JSON은 보고서와 같은 경로/이름.
  불일치 336건에는 목록 표시·확장 제어 자리표시자 차이 등이 포함되며 무손실 인증은 아니다.
- 한계: 한컴 한글 실앱은 없어 같은 버전의 사용자 증상 재현/전체 외관 인증은 미실시.
  단순 수정 전 샘플도 독립 엔진에서 열렸다. DOC 기본 모드와 복잡한 서식의 보존 제한은 보고서 참고.
