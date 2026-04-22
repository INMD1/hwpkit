import * as fs from 'fs';

// Create a simple 1x1 red PNG image
function createSimplePng(): Uint8Array {
  // Minimal valid PNG (1x1 red pixel)
  const png = new Uint8Array([
    // PNG signature
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
    // IHDR chunk
    0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
    0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, 0xDE,
    // IDAT chunk (compressed pixel data)
    0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54,
    0x08, 0xD7, 0x63, 0xF8, 0xFF, 0xFF, 0xFF, 0x00,
    0x05, 0xFE, 0x02, 0xFE, 0xDC, 0xCC, 0x59, 0xE7,
    // IEND chunk
    0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44,
    0xAE, 0x42, 0x60, 0x82
  ]);
  return png;
}

// Save the test image
const pngData = createSimplePng();
fs.writeFileSync('./test_image.png', pngData);
console.log('Created test_image.png');

// Now test encoding/decoding with different layouts
import { Pipeline } from './src/index';

async function testLayoutEncoding() {
  // Create HWPX with different image layouts by manipulating the XML directly
  const hwpxContent = `<?xml version="1.0" encoding="UTF-8"?>
<hp:HWPX>
  <hp:core>
    <hp:version major="1" minor="0"/>
  </hp:core>
  <hp:body>
    <hp:sec>
      <hp:para>
        <hp:run>
          <hp:t>테스트 텍스트</hp:t>
        </hp:run>
      </hp:para>
      <hp:para>
        <hp:run>
          <hp:pic textWrap="SQUARE" textFlow="BOTH_SIDES">
            <hp:sz width="1000" height="1000"/>
            <hp:pos flowWithText="1"/>
            <hp:img binaryItemIDRef="img1.png"/>
          </hp:pic>
        </hp:run>
      </hp:para>
      <hp:para>
        <hp:run>
          <hp:pic textWrap="TIGHT" textFlow="BOTH_SIDES">
            <hp:sz width="1000" height="1000"/>
            <hp:pos flowWithText="1"/>
            <hp:img binaryItemIDRef="img2.png"/>
          </hp:pic>
        </hp:run>
      </hp:para>
      <hp:para>
        <hp:run>
          <hp:pic textWrap="BEHIND" textFlow="NONE">
            <hp:sz width="1000" height="1000"/>
            <hp:pos flowWithText="0"/>
            <hp:img binaryItemIDRef="img3.png"/>
          </hp:pic>
        </hp:run>
      </hp:para>
    </hp:sec>
  </hp:body>
</hp:HWPX>`;

  console.log('HWPX content created (simplified for testing)');
  console.log('textWrap values in XML: SQUARE, TIGHT, BEHIND');
}

testLayoutEncoding();
