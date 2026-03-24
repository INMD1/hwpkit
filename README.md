# 🌉 HWPKit

**브라우저 네이티브 환경에서 동작하는 초경량 양방향 문서 변환 라이브러리**

HWPKit(기존 명칭 DocBridge)는 Node.js 서버 의존성 없이 순수 브라우저 환경에서 다양한 문서 포맷(HWP, HWPX, DOCX, MD)을 상호 변환할 수 있는 강력한 TypeScript/React 라이브러리입니다. 브라우저 네이티브 API와 최소한의 의존성만 사용하여 매우 빠르고 가벼운 문서 처리 환경을 제공합니다.

---

## ✨ 주요 특징

- 🌐 **완전한 브라우저 네이티브**: 백엔드 서버 없이 클라이언트(브라우저) 환경에서 모든 문서 변환을 처리합니다.
- ⚡ **빠르고 가벼움**: 불필요한 의존성을 최소화하고 브라우저 API를 적극 활용하여 가볍게 동작합니다.
- 🔄 **양방향 변환 구조 지원**:
  - `DOCX` ↔ `HWPX`
  - `HWPX` ↔ `MD`
  - `HWP` → `MD`
  - `MD` → `HWP(HML)` / `HWPX`
  - `DOCX` ↔ `MD`
- ⚛️ **React 특화 Hooks 내장**: React 환경에서 상태(로딩, 결과, 에러)를 쉽게 관리할 수 있는 전용 Hooks 제공.
- 🛡️ **TypeScript 지원**: 모든 변환 API 및 내부 추상 구문 트리(IR)에 대한 완벽한 타입을 제공합니다.
- 🧩 **커스텀 변환기 작성 가능**: 중간 표현(IR)을 기반으로 한 구조(Reader → IR → Writer)를 채택하여, 사용자가 직접 변환 로직을 확장할 수 있습니다.
- 🤖 **AI 특화 파이프라인 (개발 중)**: AI 문서 처리를 위한 XML 기반 정밀 내용 수정 기능 설계 중.

---

## 🚀 변환 기능 지원 현황

### 마크다운 (MD) 중심 변환
- **MD → HWP(HML) / HWPX / DOCX**: 지원 완료 🟢
- **HWP / HWPX / DOCX → MD**: 지원 완료 🟢

### 일반 포맷 간 변환
- **DOCX ↔ HWPX**: 지원 완료 🟢
- **HWPX / HWP → DOCX**: 지원 완료 🟢
- **DOCX → HWP**: 개발 예정 🟡

*(※ 현재 라이브러리는 활발하게 개발 및 지속적인 개선 작업이 이루어지고 있어 세부 기능은 변경될 수 있습니다.)*

---

## 📦 설치

패키지 매니저를 통해 손쉽게 설치할 수 있습니다.

```bash
npm install hwpkit
```

---

## 💻 사용 가이드

### 1. React 환경에서 Hooks 활용하기

HWPKit는 상태 관리(`isPending`, `error`, `result`)가 내장된 다양한 Hooks를 제공합니다. `hwpkit/react` (또는 `src/hooks`) 내보내기를 통해 사용할 수 있습니다.

#### 제공되는 Hooks
- `useDocxToHwpx()`, `useHwpxToDocx()`
- `useHwpxToMd()`, `useHwpToMd()`, `useDocxToMd()`
- `useMdToHwpx()`, `useMdToDocx()`, `useMdToHwp()`

#### DOCX → HWPX 변환 컴포넌트 예시
```tsx
import React from 'react';
import { useDocxToHwpx, downloadBlob } from 'hwpkit'; // 경로에 맞게 Import

function DocxToHwpxConverter() {
  const { run, isPending, error } = useDocxToHwpx();

  const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    // run 함수 호출만으로 변환 수행
    const hwpxBlob = await run(file);
    
    if (hwpxBlob) {
      // 내장된 다운로드 헬퍼 함수를 통해 브라우저에서 다운로드
      downloadBlob(hwpxBlob, 'converted.hwpx');
    }
  };

  return (
    <div>
      <input type="file" accept=".docx" onChange={handleFileChange} />
      {isPending && <p>⌛ 문서를 변환하고 있습니다...</p>}
      {error && <p style={{ color: 'red' }}>❌ 변환 오류: {error.message}</p>}
    </div>
  );
}
```

#### HWPX → Markdown 텍스트 추출 예시
```tsx
import React from 'react';
import { useHwpxToMd } from 'hwpkit'; // 경로에 맞게 Import

function HwpxToMdViewer() {
  const { run, result, isPending, error } = useHwpxToMd();

  const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    // Markdown 문자열이 result 상태에 자동으로 저장됩니다.
    await run(file);
  };

  return (
    <div>
      <input type="file" accept=".hwpx" onChange={handleFileChange} />
      {isPending && <p>⌛ 변환 중...</p>}
      {error && <p style={{ color: 'red' }}>❌ 오류 발생: {error.message}</p>}
      
      {/* 변환된 마크다운 내용 표시 */}
      {result && <pre className="md-preview">{result}</pre>}
    </div>
  );
}
```

### 2. 일반 JS/TS 환경에서 직접 호출하기 (Transformer API)

React를 사용하지 않거나 백그라운드 스크립트에서 호출할 때는 코어 `Transformer` 객체를 직접 사용합니다(`src/index.ts` 기준).

```typescript
import { MdTransformer, HwpxTransformer, DocxTransformer } from 'hwpkit';

async function processDocuments(inputFile: File) {
  try {
    // 1. 문서 객체(File 또는 Blob)를 받아 Blob 변환
    const hwpxBlob = await HwpxTransformer.fromDocx(inputFile);

    // 2. 파일 객체를 받아 마크다운 문자열로 즉시 변환
    const markdownText = await MdTransformer.fromHwpx(inputFile);
    console.log('추출된 마크다운:', markdownText);
    
  } catch (error) {
    console.error('문서 처리 중 오류가 발생했습니다:', error);
  }
}
```

---

## 🏗️ 라이브러리 아키텍처

HWPKit는 확장에 용이한 **중간 표현(Intermediate Representation, IR)** 구조를 사용합니다.

```text
입력 파일 ──▶ [ Reader ] ──▶ [ IR Document ] ──▶ [ Writer ] ──▶ 출력 파일
```

- **Readers**: `HwpxReader`, `HwpReader`, `DocxReader`, `MarkdownReader` 등
- **IR (추상 구문 트리)**: `IrDocumentNode`, `IrParagraphNode`, `IrTableNode` 등 코어 모델
- **Writers**: `HwpxWriter`, `DocxWriter`, `MdWriter` 등

---

## 🛠️ 개발 및 기여 (Contributing)

문서 파싱 규칙을 추가하거나 버그를 수정하시려면 언제든 레포지토리에 기여해 주시기 바랍니다.

```bash
# 의존성 설치
npm install

# 실시간 변경 사항 감지 및 빌드 (Vite 기반)
npm run dev

# 타입스크립트 타입 체크
npm run typecheck

# 릴리즈용 프로덕션 빌드 (ESM / UMD)
npm run build
```

---

## 📄 라이선스

이 프로젝트는 **LGPG 2.1 라이선스** 정책을 따릅니다. 사용 및 배포에 관한 더 자세한 정보는 레포지토리 내의 [`license.md`](./license.md) 파일을 참고하십시오.
