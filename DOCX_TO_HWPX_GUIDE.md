# DOCX → HWP / HWPX 변환 구현 가이드
# Claude Code용 — 실제 코드 구조 + 한컴 바이너리 역공학 기반

---

## 1. 전체 파이프라인

```
DOCX (ZIP+XML)
  └─ DocxDecoder.decode()          src/decoders/docx/DocxDecoder.ts
       └─ DocRoot (IR)             src/model/doc-tree.ts
            └─ HwpxEncoder.encode() src/encoders/hwpx/HwpxEncoder.ts
                 └─ HWPX (ZIP+XML)
```

Pipeline 호출:
```typescript
const result = await Pipeline.open(docxBytes, 'docx').to('hwpx');
```

---

## 2. DocxDecoder가 읽는 것 (DOCX → DocRoot)

### DOCX 파일 구조
```
demo.docx (ZIP)
├── word/document.xml     ← 본문 (w:body > w:p, w:tbl)
├── word/styles.xml       ← 스타일 정의
├── word/numbering.xml    ← 목록/번호매기기
├── word/fontTable.xml    ← 폰트 목록
├── word/_rels/document.xml.rels ← 이미지/관계 참조
└── word/media/           ← 이미지 파일들
```

### DOCX 핵심 요소 → DocRoot 매핑

```
w:p  → ParaNode  (단락)
w:r  → SpanNode  (텍스트 런)
w:t  → TxtNode   (텍스트 내용)
w:tbl → GridNode (표)
w:tr  → RowNode
w:tc  → CellNode
w:drawing → ImgNode (이미지)
w:sectPr  → SheetNode.dims (페이지 설정)
```

### DOCX 단위 → DocRoot 단위 변환

```typescript
// DOCX 단위: twip (dxa)  1인치=1440twip
// DocRoot 단위: pt        1인치=72pt

// twip → pt
Metric.dxaToPt(twip)     // = twip / 20

// 글자크기: w:sz는 반포인트
// w:sz val="24" = 12pt
const pt = halfPt / 2;

// 줄간격: w:spacing w:line
// w:lineRule="auto" → 퍼센트
// 240 = 100%, 480 = 200%
const lineHeightRatio = line / 240;
```

### DocxDecoder가 만드는 DocRoot 구조

```typescript
DocRoot
  └─ SheetNode { dims: PageDims, kids: ContentNode[] }
       ├─ ParaNode { props: ParaProps, kids: (SpanNode|ImgNode|...)[] }
       │    └─ SpanNode { props: TextProps, kids: TxtNode[] }
       └─ GridNode { props: GridProps, kids: RowNode[] }
            └─ RowNode { kids: CellNode[] }
                 └─ CellNode { props: CellProps, kids: ParaNode[] }
```

---

## 3. HwpxEncoder가 만드는 것 (DocRoot → HWPX)

### HWPX 파일 구조
```
output.hwpx (ZIP)
├── mimetype                     "application/hwp+zip"
├── META-INF/container.xml
├── META-INF/container.rdf
├── Contents/content.hpf         패키지 메타 (섹션 목록)
├── Contents/header.xml          ← 스타일 정의 (charPr, paraPr, fontFace 등)
├── Contents/section0.xml        ← 본문 (hp:p, hp:tbl)
├── BinData/BIN0001.png          이미지
└── Preview/PrvText.txt
```

### HWPX 핵심 요소

```
hs:sec          → 섹션 (section0.xml 루트)
hp:p            → 단락
hp:run          → 텍스트 런
hp:t            → 텍스트 내용
hp:secPr        → 섹션/페이지 설정 (반드시 hp:p의 직접 자식!)
hp:linesegarray → 줄 레이아웃 정보 (각 hp:p 마지막에 필수)
hp:tbl          → 표
hp:tr           → 행
hp:tc           → 셀 (colSpan, rowSpan 속성 직접 기재)
```

### DocRoot 단위 → HWPX 단위 변환

```typescript
// DocRoot 단위: pt
// HWPX 단위: HWPUNIT  1pt=200HU

Metric.ptToHwp(pt)   // = pt * 200

// 글자크기: charPr height 속성
// 10pt = 1000 (HWPUNIT)
// 12pt = 2400

// 줄간격: paraPr lineSpacing 속성
// 160 = 160%

// 페이지 크기: A4 기준
// width=59528, height=84188 (HWPUNIT)
```

---

## 4. HWPX 핵심 XML 구조 (반드시 준수)

### 4-1. section0.xml 기본 골격

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>
<hs:sec xmlns:hp="..." xmlns:hs="...">

  <!-- 첫 번째 단락: secPr 포함 -->
  <hp:p id="10000" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">

    <!-- ✅ secPr는 반드시 hp:p의 직접 자식 (run 바깥) -->
    <hp:secPr id="" textDirection="HORIZONTAL" ...>
      <hp:pagePr landscape="WIDELY" width="59528" height="84188" gutterType="LEFT_ONLY">
        <hp:margin header="4000" footer="4000" gutter="0"
                   left="8504" right="8504" top="5670" bottom="4252"/>
      </hp:pagePr>
      <!-- 기타 secPr 자식들 -->
    </hp:secPr>

    <!-- 텍스트 런 -->
    <hp:run charPrIDRef="0">
      <hp:t>본문 텍스트</hp:t>
    </hp:run>

    <!-- ✅ linesegarray는 반드시 hp:p 마지막 자식 -->
    <hp:linesegarray>
      <hp:lineseg textpos="0" vertpos="0" vertsize="2000"
                  textheight="1000" baseline="850" spacing="600"
                  horzpos="0" horzsize="43120" flags="393216"/>
    </hp:linesegarray>

  </hp:p>

  <!-- 이후 단락들 (secPr 없음) -->
  <hp:p id="10001" paraPrIDRef="0" styleIDRef="0" ...>
    <hp:run charPrIDRef="0">
      <hp:t>두 번째 단락</hp:t>
    </hp:run>
    <hp:linesegarray>
      <hp:lineseg .../>
    </hp:linesegarray>
  </hp:p>

</hs:sec>
```

### 4-2. ❌ 절대 금지 패턴 (현재 버그 원인)

```xml
<!-- 잘못됨: secPr가 run 안에 있음 -->
<hp:p>
  <hp:run charPrIDRef="0">
    <hp:secPr>...</hp:secPr>  ← ❌ 한컴이 무시함
  </hp:run>
</hp:p>

<!-- 올바름: secPr가 p의 직접 자식 -->
<hp:p>
  <hp:secPr>...</hp:secPr>    ← ✅
  <hp:run charPrIDRef="0">
    <hp:t>텍스트</hp:t>
  </hp:run>
</hp:p>
```

### 4-3. header.xml charPr 구조

```xml
<hh:charProperties itemCnt="N">
  <hh:charPr id="0" height="2000" textColor="#000000"
             shadeColor="none" useFontSpace="0" useKerning="0"
             symMark="NONE" borderFillIDRef="1">
    <hh:fontRef hangul="0" latin="0" hanja="0"
                japanese="0" other="0" symbol="0" user="0"/>
    <hh:ratio hangul="100" latin="100" .../>
    <hh:spacing hangul="0" latin="0" .../>
    <hh:relSz hangul="100" latin="100" .../>
    <hh:offset hangul="0" latin="0" .../>
    <hh:underline type="NONE" shape="SOLID" color="#000000"/>
    <hh:strikeout shape="NONE" color="#000000"/>
    <hh:outline type="NONE"/>
    <hh:shadow type="NONE" color="#C0C0C0" offsetX="10" offsetY="10"/>
    <!-- 굵게: <hh:bold/> 추가 -->
    <!-- 기울임: <hh:italic/> 추가 -->
  </hh:charPr>
</hh:charProperties>
```

### 4-4. 표 구조

```xml
<!-- HWPX 표: colSpan/rowSpan 직접 기재, 병합 셀은 생략 -->
<hp:tbl id="" ...>
  <hp:tr>
    <hp:tc colAddr="0" rowAddr="0" colSpan="2" rowSpan="1"
           borderFillIDRef="2" hasMargin="0" protected="0">
      <hp:cellSz width="23360" height="3000"/>
      <hp:cellMargin left="510" right="510" top="141" bottom="141"/>
      <hp:subList>
        <hp:p>...</hp:p>
      </hp:subList>
    </hp:tc>
    <!-- colSpan=2이므로 이 행에서 다음 셀 없음 -->
  </hp:tr>
</hp:tbl>
```

---

## 5. 현재 HwpxEncoder 코드의 문제점과 수정 위치

### 문제: `buildSecPrXml()` 결과를 `hp:run` 안에 넣고 있음

**파일**: `src/encoders/hwpx/HwpxEncoder.ts`

**찾을 코드** (`encodeParaPositioned` 함수 안):
```typescript
// 현재 (잘못됨)
const prefix = secPr ? `<hp:run charPrIDRef="0">${secPr}</hp:run>` : "";
```

**수정**:
```typescript
// 수정: secPr는 run 바깥에, p의 직접 자식으로
const secPrDirect = secPr ? secPr : "";  // hp:p 직접 자식으로 배치

// ...

// xml 조립 시
const xml =
  `<hp:p id="${ctx.nextElementId++}" paraPrIDRef="${paraPrId}" ...>` +
  secPrDirect +          // ← 맨 앞, run 바깥
  hfRunXml +             // ← 헤더/푸터 런
  runsXml +              // ← 실제 텍스트 런들
  linesegXml +           // ← 마지막
  `</hp:p>`;
```

**빈 문서 fallback도 동일하게 수정**:
```typescript
// 현재 (잘못됨)
contentXml =
  `<hp:p ...>` +
  `<hp:run charPrIDRef="0">${secPrXml}<hp:t></hp:t></hp:run>` +  // ← 잘못됨
  ...

// 수정
contentXml =
  `<hp:p ...>` +
  secPrXml +                                    // ← p 직접 자식
  `<hp:run charPrIDRef="0"><hp:t> </hp:t></hp:run>` +
  hfRunXml +
  buildLineSeg(0, fs + sp, fs, ctx.availableWidth) +
  `</hp:p>`;
```

---

## 6. DOCX 속성 → HWPX 속성 변환 규칙

### 6-1. 단락 정렬

```typescript
// DOCX w:jc → HWPX hh:paraPr align
const ALIGN_MAP: Record<string, string> = {
  'left':       'LEFT',
  'start':      'LEFT',
  'center':     'CENTER',
  'right':      'RIGHT',
  'end':        'RIGHT',
  'both':       'JUSTIFY',
  'distribute': 'DISTRIBUTE',
};
```

### 6-2. 줄간격

```typescript
// DOCX w:spacing w:line + w:lineRule → HWPX lineSpacing
// lineRule="auto":  line/240*100 = 퍼센트 (240=100%)
// lineRule="exact": 고정값 (HWPUNIT)
// lineRule="atLeast": 최소값

// 예: w:line="480" lineRule="auto" → lineSpacing=200 (200%)
const lineSpacingPct = Math.round(Number(line) / 240 * 100);
```

### 6-3. 글자 크기

```typescript
// DOCX w:sz = 반포인트  (24 = 12pt)
// HWPX charPr height = HWPUNIT  (12pt = 2400)
const heightHwp = Math.round((Number(sz) / 2) * 200);
```

### 6-4. 색상

```typescript
// DOCX: "FF0000" (RGB hex, # 없음)
// HWPX: "#FF0000" (# 포함)
const hwpxColor = `#${docxColor}`;
```

### 6-5. 여백/크기

```typescript
// DOCX: twip (dxa)  1인치=1440
// HWPX: HWPUNIT     1인치=14400
const hwp = Math.round(twip * 10);

// 페이지 크기 예: A4
// DOCX: width="11906" height="16838" (twip)
// HWPX: width="59528" height="84188" (HWPUNIT, ×5)
// → Metric.dxaToPt() → Metric.ptToHwp() 두 단계 거침
```

### 6-6. 테이블 병합

```
DOCX                        HWPX
─────────────────────────────────────────────────
w:gridSpan val="2"    →  tc colSpan="2"
w:vMerge val="restart" →  tc rowSpan="N" (N 계산 필요)
w:vMerge (값 없음)     →  해당 tc를 XML에서 생략
```

```typescript
// rowSpan 계산: restart 셀에서 아래로 vMerge 셀 수 카운트
function calcRowSpan(rows: DocxRow[], rowIdx: number, colIdx: number): number {
  let span = 1;
  for (let r = rowIdx + 1; r < rows.length; r++) {
    const cell = findCellAtCol(rows[r], colIdx);
    if (cell?.vMerge === true || cell?.vMerge === '') span++;
    else break;
  }
  return span;
}
```

### 6-7. 테이블 너비 (절대 균등 배분 금지)

```typescript
// DOCX w:tblGrid > w:gridCol w:w 값을 그대로 보존
const colWidths = gridCols.map(col => Math.round(Number(col.w) * 10)); // twip→HWP

// 전체 너비가 본문 너비 초과 시만 비율 축소
const totalW = colWidths.reduce((a, b) => a + b, 0);
if (totalW > bodyWidthHwp) {
  const ratio = bodyWidthHwp / totalW;
  return colWidths.map(w => Math.max(500, Math.round(w * ratio)));
}
```

### 6-8. 선 타입 (한컴 바이너리 역공학 확정)

```typescript
const BORDER_LINE_MAP: Record<string, string> = {
  'single':    'SOLID',    // ← solid, dashed 아님
  'thick':     'SOLID',
  'double':    'DOUBLE',
  'dashed':    'DASH',
  'dotted':    'DOT',
  'dashDot':   'DASH_DOT',
  'none':      'NONE',
};
```

---

## 7. linesegarray 계산법

```typescript
// 각 hp:p에 반드시 포함해야 함
function buildLineSeg(
  vertPos: number,    // 현재 세로 위치 (HWPUNIT, 누적)
  fontSize: number,   // 글자크기 (HWPUNIT)
  lineSpacingPct: number,  // 줄간격 % (160 = 160%)
  availWidth: number  // 사용 가능 가로 폭 (HWPUNIT)
): string {
  const vertSize = Math.round(fontSize * lineSpacingPct / 100);
  const textHeight = fontSize;
  const baseline = Math.round(fontSize * 0.85);
  const spacing = vertSize - textHeight;

  return `<hp:linesegarray>` +
    `<hp:lineseg textpos="0" vertpos="${vertPos}" vertsize="${vertSize}" ` +
    `textheight="${textHeight}" baseline="${baseline}" spacing="${spacing}" ` +
    `horzpos="0" horzsize="${Math.max(1, availWidth)}" flags="393216"/>` +
    `</hp:linesegarray>`;
}

// vertPos는 단락마다 누적:
let vertPos = 0;
for (const para of paras) {
  const vertSize = calcVertSize(para);
  buildLineSeg(vertPos, ...);
  vertPos += vertSize;
}
```

---

## 8. 수정 작업 순서

아래 순서로 `src/encoders/hwpx/HwpxEncoder.ts`를 수정해줘:

### Step 1: secPr 위치 수정 (빈 문서 원인)

`encodeParaPositioned` 함수에서:
- `secPr`를 `hp:run` 안이 아닌 `hp:p`의 **직접 자식**으로 이동
- 빈 문서 fallback 코드도 동일하게 수정

### Step 2: 빈 run 방지

`encodeParaPositioned`에서:
- `hp:t` 내용이 없을 때 `<hp:t> </hp:t>` (공백 1개) 삽입
- 완전히 빈 `<hp:t></hp:t>`는 일부 뷰어에서 무시됨

### Step 3: 테이블 너비 보존 확인

`encodeGridPositioned` 또는 셀 인코딩 함수에서:
- `gridCols`의 원본 너비가 보존되는지 확인
- 균등 배분 (`totalWidth / colCount`) 패턴이 있으면 제거

### Step 4: 수정 후 검증

```typescript
// 검증용 코드 임시 추가
const section0 = /* encode 결과에서 section0.xml 추출 */;
const parser = new DOMParser();
const doc = parser.parseFromString(section0, 'text/xml');

// secPr 부모가 'p'인지 확인
const secPrs = doc.getElementsByTagNameNS('http://www.hancom.co.kr/hwpml/2011/paragraph', 'secPr');
for (const secPr of Array.from(secPrs)) {
  console.assert(secPr.parentElement?.localName === 'p', 'secPr는 p의 직접 자식이어야 함');
}
```

---

## 9. 절대 규칙 (변경 금지)

1. `hp:secPr`는 **반드시 `hp:p`의 직접 자식** — `hp:run` 안에 절대 넣지 말 것
2. `hp:linesegarray`는 **반드시 `hp:p`의 마지막 자식**
3. 테이블 열 너비는 **원본 `w:gridCol` 값 보존** — 균등 배분 금지
4. `w:vMerge` 점유 셀은 HWPX에서 **완전히 생략** (빈 단락도 출력 금지)
5. 선 타입 `single` → `SOLID` (DASH 아님)
6. `Outcome<T>` 반환 — throw 직접 사용 금지
7. 개별 셀/단락 실패가 전체 변환 중단시키면 안 됨 (ShieldedParser 사용)