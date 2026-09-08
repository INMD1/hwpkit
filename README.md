# HWPKit

[![npm version](https://img.shields.io/npm/v/hwpkit.svg)](https://www.npmjs.com/package/hwpkit)
[![license](https://img.shields.io/npm/l/hwpkit.svg)](https://github.com/INMD1/hwpkit/blob/main/license.md)

**HWP / HWPX / DOCX / Markdown / HTML 문서 변환 및 DOC 입력 라이브러리**

한국 문서 포맷(HWP, HWPX)과 국제 표준(DOCX, Markdown)을 상호 변환하는 TypeScript 라이브러리입니다.
브라우저와 Node.js 환경 모두에서 동작하며, 데이터 무결성과 무중단 변환을 최우선으로 설계했습니다.

---

## 주요 특징

- **Pipeline 체이닝 API** - `Pipeline.open(file).to('hwpx')` 한 줄로 변환
- **공통 문서 모델** - 텍스트·표·이미지 등을 변환하며, 지원하지 않는 서식은 손실될 수 있음
- **무중단 변환** - 어떤 입력이 들어와도 크래시 없이 `Outcome<T>` 반환
- **4단계 표 폴백** - Full > Grid > Flat > Text 순서로 안전 변환
- **Result 모나드** - null/throw 대신 `Ok | Fail` 명시적 결과 처리
- **TypeScript 완전 지원** - 모든 노드 타입과 API에 대한 타입 정의

---

## 변환 지원 현황

| 입력 \ 출력 | HWP | HWPX | DOCX | Markdown | HTML | DOC | PDF |
|---|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
| HWP | O | O | O | O | O | X | X |
| HWPX | O | O | O | O | O | X | X |
| DOCX | O | O | O | O | O | X | X |
| Markdown | O | O | O | O | O | X | X |
| HTML | O | O | O | O | O | X | X |
| DOC | △ | △ | △ | △ | △ | X | X |
| PDF | X | X | X | X | X | X | X |

`O`는 디코더·인코더 구현과 제목·한글 본문·표 셀 내용의 기본 왕복 검증을 뜻합니다.
복잡한 표·수식·이미지 배치·페이지 레이아웃의 완전 보존을 보증하지 않습니다.
Markdown/HTML은 페이지 서식을 모두 표현하지 못합니다. `.doc`은 `.docx`의 별칭이 아닙니다.
`△`: Word 97–2003 바이너리 DOC 본문·문단·줄바꿈 읽기를 지원합니다. 기본 브라우저 모드는
표 셀 내용을 문단으로 읽으며, 서식·그림·머리말/꼬리말 손실을 `warns`로 알립니다.
아래 로컬 LibreOffice 연결을 사용하면 DOCX를 거쳐 서식·표·그림을 공통 문서 모델로 가져옵니다.
암호화된 DOC는 암호를 해제한 뒤 입력하세요. DOC 출력은 지원하지 않습니다.
PDF는 연구용 명세/렌더링 비교 대상이고 라이브러리 입력·출력 형식은 아닙니다.

2026-09-05: 최신 커밋 `887661c` 기준 5×5 조합 재점검.
[변환 정확도 점검 보고서](docs/conversion-quality-audit-2026-09-05.md)와 [작업 메모](AGENT.md)를 참고하세요.

---

## 설치

```bash
npm install hwpkit
```

---

## 사용법

### 로컬 playground

```bash
npm run playground
```

playground는 `src/index.ts`를 직접 불러오므로 라이브러리 수정이 즉시 반영됩니다.
입출력 형식 목록도 같은 라이브러리 레지스트리에서 가져옵니다.
HWP/HWPX/DOCX/DOC 파일은 내용으로 자동 감지하며, Markdown/HTML은 직접 입력하거나 파일로 열 수 있습니다.
개발 서버의 DOC 변환은 로컬 LibreOffice가 필요하며 파일 크기는 25 MiB까지입니다.
정적 playground 빌드는 DOC 본문 읽기 모드로 동작하며 손실 범위를 경고로 표시합니다.
HTML 조각을 직접 붙여 넣을 때는 입력 포맷을 HTML로 선택하세요.
DocRoot 검사에서 들여쓰기·페이지 분리 속성, 구역별 용지·머리말·꼬리말을 확인할 수 있습니다.
HTML 미리보기는 한컴의 실제 페이지 배치를 재현하는 렌더러가 아닙니다.

`npm run playground:build`는 playground 타입 검사 후 `playground/dist`에 빌드합니다.

### Pipeline API (권장)

```typescript
import { Pipeline } from 'hwpkit';

// 파일 변환
const result = await Pipeline.open(uint8ArrayData, 'docx').to('hwpx');

if (result.ok) {
  // result.data: Uint8Array (변환된 파일)
  // result.warns: string[] (폴백 발생 시 경고 목록)
  saveFile(result.data);
} else {
  console.error(result.error);
}

// 문서 구조만 추출 (인코딩 없이)
const inspectResult = await Pipeline.open(data, 'docx').inspect();
if (inspectResult.ok) {
  console.log(inspectResult.data); // DocRoot
}

// File/Blob 입력 (비동기)
const pipeline = await Pipeline.openAsync(file, 'hwpx');
const converted = await pipeline.to('docx');

// Markdown 문자열 직접 입력
const mdResult = await Pipeline.open('# Hello\n\nWorld').to('docx');
```

### DOC → HWP / HWPX / DOCX

```typescript
import { Pipeline, configureDocConverter } from 'hwpkit-dev';
import { createLibreOfficeDocConverter } from 'hwpkit-dev/node';

// Node.js: LibreOffice가 설치된 로컬 환경에서 한 번 설정합니다.
configureDocConverter(createLibreOfficeDocConverter({ timeoutMs: 60_000 }));
const result = await Pipeline.open(docBytes, 'doc').to('hwpx');
// .to('hwp'), .to('docx')도 같은 방법으로 사용합니다.
```

위 패키지 이름은 현재 `package.json`의 `hwpkit-dev` 기준입니다.
브라우저/서버 앱은 `configureDocConverter(async docBytes => docxBytes)`로 자체 변환기를
연결할 수 있습니다. 연결하지 않으면 외부 프로그램 없이 본문 읽기 모드로 동작합니다.
설정한 변환기가 실패하면 실패 결과를 반환하며, 서식을 버리는 모드로 자동 전환하지 않습니다.
LibreOffice 연결은 파일을 외부 서비스에 업로드하지 않고, 임시 문서와 프로필을 정리합니다.

검증 명령:

```bash
npm test
HWPKIT_TEST_LIBREOFFICE=1 npm test -- src/node/LibreOfficeDoc.test.ts
node tools/audit-conversions.mjs /path/to/documents /tmp/conversion-audit.json
```

[2026-09-08 파일 열기·DOC 입력 점검](docs/opening-and-doc-audit-2026-09-08.md)

### Decoder / Encoder 직접 사용

```typescript
import { registry } from 'hwpkit';

const decoder = registry.getDecoder('docx')!;
const encoder = registry.getEncoder('md')!;

const docResult = await decoder.decode(docxBytes);
if (!docResult.ok) throw new Error(docResult.error);

const mdResult = await encoder.encode(docResult.data);
if (mdResult.ok) {
  const mdText = new TextDecoder().decode(mdResult.data);
}
```

### 문서 모델 직접 구성

```typescript
import { buildRoot, buildSheet, buildPara, buildSpan, buildGrid, buildRow, buildCell } from 'hwpkit';

const doc = buildRoot({ title: '제목' }, [
  buildSheet([
    buildPara([buildSpan('Hello World', { b: true, pt: 14 })], { heading: 1 }),
    buildPara([buildSpan('본문 텍스트입니다.')]),
    buildGrid([
      buildRow([
        buildCell([buildPara([buildSpan('A1')])]),
        buildCell([buildPara([buildSpan('B1')])]),
      ]),
      buildRow([
        buildCell([buildPara([buildSpan('A2')])]),
        buildCell([buildPara([buildSpan('B2')])]),
      ]),
    ]),
  ]),
]);
```

### 스타일 적용

```typescript
// 텍스트 스타일
buildSpan('스타일 적용 텍스트', {
  font: 'Malgun Gothic',  // 글꼴
  pt: 14,                 // 글자 크기 (pt)
  b: true,                // 볼드
  i: true,                // 이탤릭
  u: true,                // 밑줄
  s: true,                // 취소선
  color: 'FF0000',        // 글색 (hex, BGR)
  bg: 'FFFF00'           // 형광펜 (hex, BGR)
});

// 표 정렬 및 선 스타일
buildGrid(rows, {
  align: 'center',        // 표 정렬: 'left' | 'center' | 'right'
  defaultStroke: {        // 기본 선 스타일
    kind: 'solid',        // 선 종류: 'solid' | 'double' | 'dash' | 'dot'
    pt: 1,                // 선 굵기 (pt)
    color: '000000'       // 선 색상 (hex, BGR)
  },
  colWidths: [100, 100]   // 열 너비 (pt)
});

// 개별 셀의 선 스타일
buildCell(content, {
  top: { kind: 'double', pt: 2, color: '0000FF' },  // 상단 선
  bot: { kind: 'dash', pt: 1.5, color: 'FF0000' },  // 하단 선
  left: { kind: 'dot', pt: 1, color: '00FF00' },    // 좌측 선
  right: { kind: 'solid', pt: 3, color: 'FFFF00' }, // 우측 선
  bg: 'FFFFFF'                                        // 셀 배경색
});
```

### 트리 순회

```typescript
import { TreeWalker, walkNode, countNodes, validateRoot } from 'hwpkit';

// 텍스트 추출
const walker = new TreeWalker();
const text = walker.extractText(docRoot);

// 노드 통계
const counts = countNodes(docRoot);
// { root: 1, sheet: 1, para: 5, span: 5, txt: 5, grid: 1, row: 2, cell: 4 }

// 유효성 검증
const errors = validateRoot(docRoot);
```

---

## 아키텍처

```
입력 파일 --> [ Decoder ] --> [ DocRoot ] --> [ Encoder ] --> 출력 파일
                                  |
                            Pipeline.inspect()
```

### 문서 추상 모델 (Doc Model)

모든 문서는 `DocRoot` 트리로 변환되어 포맷 간 변환의 중간 표현으로 사용됩니다.

```
DocRoot
  └─ SheetNode (섹션/페이지)
       ├─ ParaNode (문단)
       │    ├─ SpanNode (텍스트 런)
       │    │    └─ TxtNode / BrNode / PbNode
       │    ├─ ImgNode (이미지)
       │    └─ LinkNode (하이퍼링크)
       └─ GridNode (표)
            └─ RowNode
                 └─ CellNode
                      └─ ParaNode ...
```

### 안전 계층

- **ShieldedParser** - 개별 노드 파싱 실패가 전체를 중단시키지 않음
- **StyleBridge** - 포맷 간 스타일/단위 변환 (`Metric.*`)
- **Outcome<T>** - 모든 결과를 `Ok | Fail`로 감싸 null/throw 제거

### 디렉토리 구조

```
src/
├── model/          # 문서 추상 모델 (DocRoot, 속성, 빌더)
├── contract/       # Decoder/Encoder 인터페이스, Result 모나드
├── pipeline/       # Pipeline 오케스트레이터, 포맷 레지스트리
├── decoders/       # 입력 포맷 → DocRoot
│   ├── docx/       #   DocxDecoder
│   ├── hwpx/       #   HwpxDecoder
│   ├── hwp/        #   HwpScanner
│   └── md/         #   MdDecoder
├── encoders/       # DocRoot → 출력 포맷
│   ├── docx/       #   DocxEncoder
│   ├── hwpx/       #   HwpxEncoder
│   └── md/         #   MdEncoder
├── walk/           # 트리 순회 (TreeWalker, walkNode)
├── safety/         # ShieldedParser, StyleBridge
└── toolkit/        # XmlKit, ArchiveKit, BinaryKit, TextKit
```

---

## 개발

```bash
# 의존성 설치
npm install

# 타입 체크
npm run typecheck

# 테스트 실행
npm test

# 빌드 (ESM + CJS + d.ts)
npm run build

# 개발 모드 (watch)
npm run dev
```

### 의존성

| 패키지 | 용도 |
|--------|------|
| `pako` | ZIP inflate/deflate |
| `xml2js` | XML 파싱/빌드 |
| `saxes` | SAX 스트리밍 파서 (대용량 XML) |
| `tsup` | 빌드 (esbuild 기반) |
| `vitest` | 테스트 프레임워크 |

---

## 라이선스

이 프로젝트는 **LGPL-2.1** 라이선스를 따릅니다. 자세한 내용은 [`license.md`](./license.md)를 참고하세요.


### PDF 외관 검증 상태

한컴 원본 외관과 94% 이상 일치하는 상태는 아직 아닙니다.
본문·머리말 여백과 빈 문단 높이를 보정했으며, 실제 PDF의 표 높이·줄 간격·글꼴·쪽 흐름에는 차이가 남아 있습니다.
`PageDims.mt/mb`는 종이 가장자리에서 본문까지, `headerPt/footerPt`는 머리말/꼬리말까지의 거리(pt)입니다.
HWP/HWPX 저장값은 코덱에서 합산·분리하므로 직접 모델을 만들 때 HWP의 바깥 여백만 넣지 않습니다.
[PDF 비교 자료와 실제 측정 결과](../hwpkit_research/docs/pdf-layout-review-2026-09-05.md)를 참고하세요.
