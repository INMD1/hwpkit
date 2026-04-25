# HWPKit

[![npm version](https://img.shields.io/npm/v/hwpkit.svg)](https://www.npmjs.com/package/hwpkit)
[![license](https://img.shields.io/npm/l/hwpkit.svg)](https://github.com/INMD1/hwpkit/blob/main/license.md)

**HWP / HWPX / DOCX / Markdown 양방향 문서 변환 라이브러리**

한국 문서 포맷(HWP, HWPX)과 국제 표준(DOCX, Markdown)을 상호 변환하는 TypeScript 라이브러리입니다.
브라우저와 Node.js 환경 모두에서 동작하며, 데이터 무결성과 무중단 변환을 최우선으로 설계했습니다.

---

## 주요 특징

- **Pipeline 체이닝 API** - `Pipeline.open(file).to('hwpx')` 한 줄로 변환
- **데이터 무결성 100%** - 텍스트, 표, 이미지 누락 없이 변환
- **무중단 변환** - 어떤 입력이 들어와도 크래시 없이 `Outcome<T>` 반환
- **4단계 표 폴백** - Full > Grid > Flat > Text 순서로 안전 변환
- **Result 모나드** - null/throw 대신 `Ok | Fail` 명시적 결과 처리
- **TypeScript 완전 지원** - 모든 노드 타입과 API에 대한 타입 정의

---

## 변환 지원 현황

| 입력 \ 출력 | HWPX | DOCX | Markdown |
|------------|:----:|:----:|:--------:|
| **HWPX**   | -    | O    | O        |
| **HWP**    | □    | O    | O        |
| **DOCX**   | □    | -    | O        |
| **Markdown** | □  | □    | -        |

---

## 설치

```bash
npm install hwpkit
```

---

## 사용법

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

### Decoder / Encoder 직접 사용

```typescript
import { DocxDecoder } from 'hwpkit';
import { MdEncoder } from 'hwpkit';

const decoder = new DocxDecoder();
const encoder = new MdEncoder();

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
