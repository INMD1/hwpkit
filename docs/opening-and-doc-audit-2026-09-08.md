# DOC 입력과 HWP/HWPX 파일 열기 점검 — 2026-09-08

## 수정 내용

- DOC 입력 디코더를 등록하고 `.doc`, `.DOC`, `application/msword`, WordDocument OLE 자동 감지를 연결했다.
  Word 97–2003의 CLX/Pcdt/PlcPcd, 0Table/1Table, 압축 문자와 UTF-16 문자를 읽는다.
  본문 문자 수 범위를 지켜 머리말·각주·필드 명령문이 본문에 섞이지 않게 했다.
- 브라우저 기본 모드는 본문 텍스트·문단·줄바꿈을 읽으며 서식·표 구조·그림·머리말/꼬리말 손실을 경고한다.
  `configureDocConverter`와 별도 `hwpkit-dev/node` 엔트리의 `createLibreOfficeDocConverter`로
  로컬 LibreOffice DOC→DOCX를 연결할 수 있다. 개발 playground도 이 경로를 사용한다.
  임시 파일/프로필 격리·정리, 타임아웃, 프로세스 종료, 실패 결과 처리를 포함한다.
- HWP PARA_HEADER의 문자 수 최상위 비트를 모든 문단에 설정하던 오류를 수정했다.
  구역·표 셀의 마지막 문단에만 설정한다. 구역 설정만 담긴 첫 문단으로 본문 목록을
  끝내던 출력도 수정된다. 검증기의 잘못된 “모든 문단에 비트 필수” 조건도 바로잡았다.
  실제 한컴 원본 `input_04sa6jcl.hwp`, `input_0d1tojvf.hwp`, `input_0las09e5.hwp`의
  본문·셀 문단 비트 패턴과 대조했다.
- HWP 탭을 1워드로 저장해 뒤의 7문자가 건너뛰어지는 오류를 수정했다.
  8워드 인라인 제어와 글자 모양/줄 시작 위치의 문자 오프셋을 함께 맞춘다.
- HWPX의 `hc:indent`를 한컴 공식 요소명 `hc:intent`로 수정했다.
  이전 hwpkit 파일을 읽기 위해 디코더는 두 이름을 계속 받아들인다.
- HWPX 줄바꿈을 잘못된 run 자식 `hp:br`에서 `hp:t` 내부의 `hp:lineBreak`로 수정했다.
  원본 HWPX의 텍스트·줄바꿈·탭 순서를 보존하도록 XML 혼합 콘텐츠를 처리한다.
  기존 줄바꿈 보존 변경의 의도와 DOCX 처리도 유지했다. 기존 `hp:br` 입력은 호환용으로 수용한다.
- DOCX의 `w:tab` 읽기/쓰기를 추가했다. XML 1.0에서 금지된 제어문자·고립 surrogate를
  출력에서 제거하고 정상 이모지 쌍은 보존한다.
- CFB FAT/DIFAT/mini FAT/디렉터리 순환 참조를 거부해 손상 파일에서 무한 반복하지 않게 했다.
  CFB v4의 4096바이트 섹터 원점도 수정했다.

- 추가 샘플 감사에서 DOCX 출력이 LinkNode 전체를 누락하던 오류를 발견했다.
  링크의 표시 글자와 외부 관계를 기록하며, 독립 demo.docx의 공백 제외 본문
  7,218자를 DOCX 왕복에서 모두 보존한다(이전 6,922자, 296자 누락).
- Markdown 제목·표 셀의 인라인 서식을 다시 파싱해 왕복마다 `**`가 늘어나던 오류를 수정했다.
  HTML 엔티티는 한 번만 해석하고 인라인 요소 사이 공백을 보존한다.
  Markdown 링크의 표시 글자를 읽고, 복잡한 표 출력에 쓰이는 HTML table 블록을 다시 표로 읽어
  HWP/HWPX/DOCX 본문에 HTML 태그가 그대로 나타나지 않게 했다.
- 확장자가 HWP지만 실제로 HWPX인 `input_xd402yls.hwp`를 확인했다.
  File/Blob API에서 명시 형식이 없으면 Office 파일의 컨테이너를 우선 감지한다.
  PDF/RTF 바이너리를 Markdown으로 잘못 변환하던 자동 감지도 거부하도록 수정했다.

## 근거

- 로컬 한컴 HWP 5.0 revision 1.3: §4.1 제어문자, §4.3.1/표 58 문단 헤더.
- [hwplib 독립 리더의 문단 목록 종료 비트 해석](https://github.com/neolord0/hwplib/blob/master/src/main/java/kr/dogfoot/hwplib/reader/bodytext/paragraph/ForParaHeader.java): bit31을 lastInList로 읽으며 한컴 원본 패턴과 일치한다.
- [한컴 공식 margin.cpp](https://github.com/hancom-io/hwpx-owpml-model/blob/main/OWPML/Class/Head/margin.cpp): `intent` 요소 매핑.
- [한컴 공식 t.cpp](https://github.com/hancom-io/hwpx-owpml-model/blob/main/OWPML/Class/Para/t.cpp): 텍스트 내부 `lineBreak`, `tab` 요소 매핑.
- 실제 한컴 HWPX 원본: `input_ws2chrxl.hwpx`의 텍스트 사이 lineBreak,
  `input_wdl1yn85.hwpx`의 텍스트 내부 tab, `input_9n4jwbz5.hwpx`의 intent.
- [Microsoft MS-DOC CLX](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/bad26767-b575-44d3-9da3-96378d56ce14),
  [Word 바이너리 명세](https://download.microsoft.com/download/0/b/e/0be8bdd7-e5e8-422a-abfd-4342ed7ad886/word97-2007binaryfileformat%28doc%29specification.pdf).

## 검증 범위

기본 5×5 변환 테스트와 DOC→5형식 검사를 수행한다. 별도 원본 DOC fixture는
LibreOffice MS Word 97 필터로 생성했으며 테스트 파일 생성 과정은 fixtures README에 기록했다.
압축/UTF-16 혼합·0Table·Prc는 독립 바이너리 fixture로 추가 검사한다.
HWP 문단 종료·중첩 셀·탭, HWPX 공식 요소 및 혼합 콘텐츠, ZIP CRC와 namespace XML 파싱을 검사한다.

외부 검사 엔진은 LibreOffice + H2Orestart 0.7.14이다. Markdown으로 만든 HWP/HWPX/DOCX를
각각 PDF로 열었으며 한글, 표의 Gamma/Delta, 마지막 Omega, 코드 First/Second, 탭 뒤 AFTERTAB가
세 파일에서 모두 표시됐다. 로컬 playground HTTP 업로드도 실제 DOCX 및 표 1개를 반환했다.
LibreOffice DOC→HWP/HWPX/DOCX 통합 검사에서 본문과 표 구조가 유지됐다.
누락을 수정한 DOCX demo도 PDF로 열어 calibre download page, paragraph level formatting 등의 링크 글자를 확인했다.

전체 실제 문서 검사는 `tools/audit-conversions.mjs`로 수행한다. 각 원본을 한 번 디코딩하고
복제한 모델을 5개 인코더에 전달한다. 공개 Pipeline 경로는 별도 방향별 테스트로 검사한다. 원본 파일은 읽기만 하며
변환 파일은 메모리에서 검사한다. 각 출력의 재열기·비어 있지 않은 본문·XML/ZIP 무결성을 확인한다.
텍스트 완전 일치 수치는 공백을 제거한 자체 디코더 결과 비교이며 외관 인증이 아니다.

최종 검사 결과는 [기계 판독 보고서](opening-and-doc-audit-2026-09-08.json)에 기록했다.
전체 5방향 감사 후 마지막 Markdown 디코더 수정의 영향을 받는 MD 출력 237건을 재검사해 합쳤다.

| 검사 | 결과 |
| --- | --- |
| 전체 테스트, 실제 LibreOffice DOC 통합 포함 | 149/149 통과 |
| 타입 검사·라이브러리 ESM/CJS/타입 선언·playground 빌드 | 통과 |
| 실제 HWP 205개 + HWPX 32개 → 5형식 | 1,185/1,185 생성·재읽기 성공 |
| 위 실제 문서의 공백 제외 텍스트 완전 일치 | 849/1,185 |
| 별도 MD/DOC/DOCX/HWP fixture 11개 → 5형식 | 55/55 생성·재읽기 성공, 텍스트 일치 50/55 |
| Markdown fixture 8개 → HWP/HWPX/DOCX | 24/24 텍스트 일치 |

텍스트가 일치하지 않는 336건을 모두 내용 손실로 분류하지는 않았다. 알려진 차이에는
Markdown/HTML의 목록 표시와 HWP 확장 제어 자리표시자(`__EXT_3__`) 제거가 포함된다.
복잡한 문서의 텍스트·서식 보존 문제는 여전히 남아 있으며, 이 수치는 무손실 변환을 뜻하지 않는다.

재현 명령:

```sh
HWPKIT_TEST_LIBREOFFICE=1 npm test
npm run typecheck
npm run build
npm run playground:build
node tools/audit-conversions.mjs ../hwpkit_research/datasets/input /tmp/hwpkit-corpus.json
node tools/audit-conversions.mjs tests/fixtures /tmp/hwpkit-fixtures.json
```

## 남은 한계

- 한컴 한글 프로그램 자체는 이 환경에 없어 사용자와 동일한 버전의 열기 동작은 직접 검증하지 못했다.
  수정 전의 단순 샘플도 H2Orestart에서는 열렸다. 따라서 외부 PDF 검사는 이번 수정 후의
  표시 성공 근거이며, 사용자 프로그램에서 발생한 증상을 동일하게 재현했다는 뜻은 아니다.
- DOC 기본 브라우저 모드는 서식·표 구조·그림을 보존하지 않는다. 전체 변환은 LibreOffice 연결이 필요하다.
- 암호화 DOC, Word 97 이전 파일, RTF/HTML에 `.doc` 확장자만 붙인 파일, DOC 출력은 기본 지원 범위 밖이다.
- 복잡한 표·수식·페이지 배치와 모든 한컴 버전의 호환성은 보증하지 않는다. 전체 방향의 기본 변환 성공과
  원본 외관·내용의 완전 보존은 구분한다.
