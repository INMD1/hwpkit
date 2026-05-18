/**
 * HWP 인코더 스타일 기능 테스트
 *
 * 테스트 항목:
 * 1. 표 위치 배치 (align: left/center/right)
 * 2. 표 선 스타일 (각 변별 선 종류, 굵기, 색상)
 * 3. 텍스트 스타일 (글꼴, bold, italic, underline, 글색, 형광펜)
 */

import { buildRoot, buildSheet, buildGrid, buildRow, buildCell, buildPara, buildSpan } from './src/model/builders';
import { HwpEncoder } from './src/encoders/hwp/HwpEncoder';
import { HwpxEncoder } from './src/encoders/hwpx/HwpxEncoder';
import fs from 'fs';

// 표 정렬 및 선 스타일 테스트 문서
const doc = buildRoot(
  { title: 'HWP 인코더 스타일 테스트 문서' },
  [
    buildSheet([
      // 제목
      buildPara([
        buildSpan('HWP 인코더 스타일 기능 테스트', { b: true, pt: 18 }),
      ]),

      // 1. 표 정렬 테스트 - 왼쪽 정렬
      buildPara([buildSpan('1. 표 정렬: 왼쪽 정렬', { b: true })]),
      buildGrid(
        [
          buildRow([
            buildCell([buildPara([buildSpan('왼쪽 정렬 표', { b: true })])]),
            buildCell([buildPara([buildSpan('데이터', {})])]),
          ]),
          buildRow([
            buildCell([buildPara([buildSpan('셀 1', {})])]),
            buildCell([buildPara([buildSpan('셀 2', {})])]),
          ]),
        ],
        {
          align: 'left',
          defaultStroke: { kind: 'solid', pt: 1, color: '000000' },
          colWidths: [100, 100]
        }
      ),

      // 2. 표 정렬 테스트 - 중앙 정렬
      buildPara([buildSpan('2. 표 정렬: 중앙 정렬', { b: true })]),
      buildGrid(
        [
          buildRow([
            buildCell([buildPara([buildSpan('중앙 정렬 표', { b: true })])]),
            buildCell([buildPara([buildSpan('데이터', {})])]),
          ]),
          buildRow([
            buildCell([buildPara([buildSpan('셀 1', {})])]),
            buildCell([buildPara([buildSpan('셀 2', {})])]),
          ]),
        ],
        {
          align: 'center',
          defaultStroke: { kind: 'solid', pt: 1, color: '000000' },
          colWidths: [100, 100]
        }
      ),

      // 3. 표 정렬 테스트 - 오른쪽 정렬
      buildPara([buildSpan('3. 표 정렬: 오른쪽 정렬', { b: true })]),
      buildGrid(
        [
          buildRow([
            buildCell([buildPara([buildSpan('오른쪽 정렬 표', { b: true })])]),
            buildCell([buildPara([buildSpan('데이터', {})])]),
          ]),
          buildRow([
            buildCell([buildPara([buildSpan('셀 1', {})])]),
            buildCell([buildPara([buildSpan('셀 2', {})])]),
          ]),
        ],
        {
          align: 'right',
          defaultStroke: { kind: 'solid', pt: 1, color: '000000' },
          colWidths: [100, 100]
        }
      ),

      // 4. 표 선 스타일 테스트 - 각 변별 다른 스타일
      buildPara([buildSpan('4. 표 선 스타일: 각 변별 다른 스타일', { b: true })]),
      buildGrid(
        [
          buildRow([
            buildCell([buildPara([buildSpan('상단', {})])], {
              top: { kind: 'double', pt: 2, color: '0000FF' }
            }),
            buildCell([buildPara([buildSpan('하단', {})])], {
              bot: { kind: 'dash', pt: 1.5, color: 'FF0000' }
            }),
          ]),
          buildRow([
            buildCell([buildPara([buildSpan('좌측', {})])], {
              left: { kind: 'dot', pt: 1, color: '00FF00' }
            }),
            buildCell([buildPara([buildSpan('우측', {})])], {
              right: { kind: 'solid', pt: 3, color: 'FFFF00' }
            }),
          ]),
        ],
        { defaultStroke: { kind: 'solid', pt: 0.5, color: '808080' } }
      ),

      // 5. 텍스트 스타일 테스트 - 글꼴
      buildPara([buildSpan('5. 글꼴 테스트', { b: true })]),
      buildPara([
        buildSpan('배탕체', { font: 'Batang' }),
        buildSpan(' + ', {}),
        buildSpan('맑은고딕', { font: 'Malgun Gothic' }),
        buildSpan(' + ', {}),
        buildSpan('굴림체', { font: 'GulimChe' }),
      ]),

      // 6. 텍스트 스타일 테스트 - 글자 모양
      buildPara([buildSpan('6. 글자 모양 테스트', { b: true })]),
      buildPara([
        buildSpan('볼드', { b: true }),
        buildSpan(' + ', {}),
        buildSpan('이탤릭', { i: true }),
        buildSpan(' + ', {}),
        buildSpan('밑줄', { u: true }),
        buildSpan(' + ', {}),
        buildSpan('취소선', { s: true }),
      ]),

      // 7. 텍스트 스타일 테스트 - 글색
      buildPara([buildSpan('7. 글색 테스트', { b: true })]),
      buildPara([
        buildSpan('빨강', { color: 'FF0000' }),
        buildSpan(' + ', {}),
        buildSpan('파랑', { color: '0000FF' }),
        buildSpan(' + ', {}),
        buildSpan('초록', { color: '008000' }),
      ]),

      // 8. 텍스트 스타일 테스트 - 형광펜 (배경색)
      buildPara([buildSpan('8. 형광펜 테스트', { b: true })]),
      buildPara([
        buildSpan('노란 형광펜', { bg: 'FFFF00' }),
        buildSpan(' + ', {}),
        buildSpan('초록 형광펜', { bg: 'CCFFCC' }),
        buildSpan(' + ', {}),
        buildSpan('파랑 형광펜', { bg: 'CCCCFF' }),
      ]),

      // 9. 복합 스타일 테스트
      buildPara([buildSpan('9. 복합 스타일 테스트', { b: true })]),
      buildPara([
        buildSpan('볼드 + 빨강 + 노란 형광펜', {
          b: true,
          color: 'FF0000',
          bg: 'FFFF00'
        }),
      ]),
      buildPara([
        buildSpan('이탤릭 + 파랑 + 굴림체', {
          i: true,
          color: '0000FF',
          font: 'GulimChe'
        }),
      ]),
    ]),
  ]
);

async function runTest() {
  console.log('=== HWP 인코더 스타일 테스트 ===\n');

  // HWP 인코딩
  console.log('HWP 인코딩 중...');
  const hwpEncoder = new HwpEncoder();
  const hwpOutcome = await hwpEncoder.encode(doc);
  if (hwpOutcome.ok) {
    fs.writeFileSync('test-styling.hwp', hwpOutcome.data);
    console.log('✓ HWP 파일 생성 완료: test-styling.hwp');
    console.log(`  파일 크기: ${hwpOutcome.data.length} bytes\n`);
  } else {
    console.error('✗ HWP 인코딩 실패:', hwpOutcome.err);
    process.exit(1);
  }

  // HWPX 인코딩
  console.log('HWPX 인코딩 중...');
  const hwpxEncoder = new HwpxEncoder();
  const hwpxOutcome = await hwpxEncoder.encode(doc);
  if (hwpxOutcome.ok) {
    fs.writeFileSync('test-styling.hwpx', hwpxOutcome.data);
    console.log('✓ HWPX 파일 생성 완료: test-styling.hwpx');
    console.log(`  파일 크기: ${hwpxOutcome.data.length} bytes\n`);
  } else {
    console.error('✗ HWPX 인코딩 실패:', hwpxOutcome.err);
    process.exit(1);
  }

  console.log('=== 테스트 완료 ===');
  console.log('생성된 파일을 HWP 뷰어로 열어 다음을 확인하세요:');
  console.log('  1. 표 정렬: 왼쪽/중앙/오른쪽 정렬이 적용되는지');
  console.log('  2. 표 선 스타일: 각 변별 다른 선 종류/굵기/색상이 적용되는지');
  console.log('  3. 글꼴: 배탕체/맑은고딕/굴림체가 적용되는지');
  console.log('  4. 글자 모양: bold/italic/underline/strikethrough 가 적용되는지');
  console.log('  5. 글색: 빨강/파랑/초록이 적용되는지');
  console.log('  6. 형광펜: 노랑/초록/파랑 배경색이 적용되는지');
}

runTest().catch(console.error);
