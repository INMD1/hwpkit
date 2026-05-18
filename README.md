# HWPKit

[![npm version](https://img.shields.io/npm/v/hwpkit.svg)](https://www.npmjs.com/package/hwpkit)
[![license](https://img.shields.io/npm/l/hwpkit.svg)](https://github.com/INMD1/hwpkit/blob/main/license.md)

**HWP / HWPX / DOCX / Markdown / HTML 양방향 문서 변환 라이브러리**

한국 문서 포맷(HWP, HWPX)과 국제 표준(DOCX, Markdown, HTML)을 상호 변환하는 TypeScript 라이브러리입니다. 브라우저와 Node.js 환경에서 모두 동작하며, 데이터 무결성과 안전한 변환을 최우선으로 설계했습니다.

---

## ✨ v0.0.3 주요 업데이트 (2026.05)

- **복잡한 표 지원**: 이제 표 안에 또 다른 표가 들어있는 **중첩된 표(Nested Tables)**를 완벽하게 지원합니다.
- **이미지 처리 강화**: DOCX의 머리말/꼬리말 이미지 인식 문제를 해결하고, HWP 이미지 추출 안정성을 높였습니다.
- **문단 스타일 개선**: 텍스트 들여쓰기, 정렬, 행 높이 계산 로직이 더욱 정교해졌습니다.
- **마크다운 옵션**: 변환 시 이미지를 포함할지 여부를 선택할 수 있는 토글 옵션이 추가되었습니다.

---

## 주요 특징

- **Pipeline API**: `Pipeline.open(file).to('hwpx')` 단 한 줄로 변환이 끝납니다.
- **데이터 안전 보장**: 변환 중 오류가 발생해도 프로그램이 멈추지 않고 안전한 결과(`Outcome`)를 반환합니다.
- **표 폴백(Fallback)**: 복잡한 표를 최대한 원본에 가깝게 변환하며, 불가능할 경우 텍스트로라도 안전하게 변환합니다.
- **TypeScript 완전 지원**: 모든 API에 타입 정의가 되어 있어 오타 걱정 없이 개발할 수 있습니다.

---

## 변환 지원 현황

| 입력 \ 출력 | HWPX | DOCX | Markdown | HTML |
|------------|:----:|:----:|:--------:|:----:|
| **HWPX**   | -    | O    | O        | O    |
| **HWP**    | O    | O    | O        | O    |
| **DOCX**   | O    | -    | O        | O    |
| **Markdown** | O  | O    | -        | O    |
| **HTML**   | O    | O    | O        | -    |

> **참고**: HWP/HWPX 포맷의 경우 한글(Hangul) 소프트웨어 버전이나 특수 기능 사용 여부에 따라 일부 레이아웃이 다르게 보일 수 있습니다. (지속적으로 개선 중)

---

## 설치

```bash
npm install hwpkit
```

---

## 사용법

### 1. Pipeline API (가장 쉬운 방법)

```typescript
import { Pipeline } from 'hwpkit';

// 파일 변환 (DOCX -> HWPX)
const result = await Pipeline.open(uint8ArrayData, 'docx').to('hwpx');

if (result.ok) {
  // result.data: Uint8Array (변환된 파일)
  saveFile(result.data);
} else {
  console.error('변환 실패:', result.error);
}

// Markdown 문자열 직접 입력하여 DOCX 만들기
const mdResult = await Pipeline.open('# Hello\n\nWorld').to('docx');
```

### 2. Decoder / Encoder 직접 사용

```typescript
import { DocxDecoder, MdEncoder } from 'hwpkit';

const decoder = new DocxDecoder();
const encoder = new MdEncoder();

const docResult = await decoder.decode(docxBytes);
if (docResult.ok) {
  const mdResult = await encoder.encode(docResult.data);
}
```

### 3. 문서 모델 직접 구성 (빌더)

```typescript
import { buildRoot, buildSheet, buildPara, buildSpan, buildGrid, buildRow, buildCell } from 'hwpkit';

const doc = buildRoot({ title: '제목' }, [
  buildSheet([
    buildPara([buildSpan('Hello World', { b: true, pt: 14 })]),
    buildGrid([
      buildRow([
        buildCell([buildPara([buildSpan('A1')])]),
        buildCell([buildPara([buildSpan('B1')])]),
      ]),
    ]),
  ]),
]);
```

---

## 아키텍처

HWPKit은 모든 문서를 **중간 표현(DocRoot)**으로 변환한 뒤, 다시 원하는 포맷으로 내보내는 방식을 사용합니다.

```
입력 파일 --> [ Decoder ] --> [ DocRoot ] --> [ Encoder ] --> 출력 파일
                                  |
                            Pipeline.inspect()
```

### 문서 추상 모델 (Doc Model)

```
DocRoot
  └─ SheetNode (섹션/페이지)
       ├─ ParaNode (문단)
       │    ├─ SpanNode (텍스트) -> TxtNode, BrNode 등
       │    ├─ ImgNode (이미지)
       │    └─ LinkNode (하이퍼링크)
       └─ GridNode (표)
            └─ RowNode
                 └─ CellNode
                      └─ ParaNode ... (중첩 가능)
```

---

## 디렉토리 구조

```
src/
├── model/          # 문서 추상 모델 (DocRoot, 속성, 빌더)
├── contract/       # Decoder/Encoder 인터페이스, Result 모나드
├── pipeline/       # Pipeline 오케스트레이터
├── decoders/       # 입력 포맷 → DocRoot (docx, hwp, hwpx, md, html)
├── encoders/       # DocRoot → 출력 포맷 (docx, hwp, hwpx, md, html)
├── walk/           # 트리 순회 및 데이터 추출
├── safety/         # 파싱 안전 계층 (ShieldedParser)
└── toolkit/        # 공통 도구 (Xml, Binary, Text, Unit)
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

# 빌드
npm run build
```

---

## 라이선스

이 프로젝트는 **LGPL-2.1** 라이선스를 따릅니다. 자세한 내용은 [`license.md`](./license.md)를 참고하세요.
