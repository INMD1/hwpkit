# Javascript 변환기 사용 가이드

이 디렉토리에는 DOCX와 HWPX 간의 변환을 수행하는 Javascript 소스 코드가 포함되어 있습니다.

## 1. DOCX to HWPX (`docx-to-hwpx.js`)

이 스크립트는 **브라우저**와 **Node.js** 환경 모두에서 사용할 수 있습니다.

### 브라우저에서 사용하기 (HTML)

`docx-to-hwpx.js`는 `JSZip` 라이브러리에 의존합니다. HTML 파일에서 `JSZip`을 먼저 로드한 후 스크립트를 사용하세요.

#### 예제 코드 (`demo.html`)

```html
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>DOCX to HWPX Converter</title>
    <!-- 1. JSZip 라이브러리 로드 (필수) -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
    <!-- 2. 변환기 스크립트 로드 -->
    <script src="docx-to-hwpx.js"></script>
</head>
<body>
    <h2>DOCX to HWPX 변환기</h2>
    <input type="file" id="upload" accept=".docx" />
    <button onclick="convert()">변환하기</button>

    <script>
        async function convert() {
            const fileInput = document.getElementById('upload');
            const file = fileInput.files[0];
            if (!file) {
                alert('파일을 선택해주세요.');
                return;
            }

            try {
                // 3. 변환기 인스턴스 생성
                const converter = new DocxToHwpxConverter();

                // 4. 변환 실행 (Blob 반환)
                console.log('변환 시작...');
                const hwpxBlob = await converter.convert(file);
                
                // 5. 다운로드 처리
                const url = URL.createObjectURL(hwpxBlob);
                const a = document.createElement('a');
                a.href = url;
                a.download = file.name.replace('.docx', '.hwpx');
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
                
                console.log('변환 완료!');
            } catch (err) {
                console.error('변환 중 오류 발생:', err);
                alert('변환에 실패했습니다.');
            }
        }
    </script>
</body>
</html>
```

### Node.js에서 사용하기

Node.js 환경에서는 모듈로 불러와서 사용할 수 있습니다.

```javascript
const fs = require('fs');
const JSZip = global.JSZip = require('jszip'); // JSZip을 전역이나 스코프에 로드해야 할 수 있습니다.
const { DOMParser } = require('xmldom'); // XML 파서 필요
global.DOMParser = DOMParser; // 브라우저 환경 흉내

const DocxToHwpxConverter = require('./docx-to-hwpx.js');

async function main() {
    const inputPath = 'document.docx';
    const data = fs.readFileSync(inputPath);
    
    const converter = new DocxToHwpxConverter();
    const resultBlob = await converter.convert(data);
    
    // Blob을 Buffer로 변환하여 저장 (Node.js 버전에 따라 다를 수 있음)
    const buffer = Buffer.from(await resultBlob.arrayBuffer());
    fs.writeFileSync('output.hwpx', buffer);
}

main();
```

---

## 2. HWPX to DOCX (`hwpx-to-docx.js`)

이 스크립트는 현재 **Node.js 환경(CLI)**을 위해 설계되었습니다. 파일 시스템(`fs`) 접근이 필요하므로 브라우저에서 직접 실행할 수 없습니다.

### 사용법 (CLI)

```bash
node hwpx-to-docx.js <input.hwpx> [output.docx]
```

*   **input.hwpx**: 변환할 HWPX 파일 경로 또는 압축 해제된 폴더 경로
*   **output.docx**: (선택) 저장할 DOCX 파일 경로
