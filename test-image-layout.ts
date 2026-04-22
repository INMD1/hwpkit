import { Pipeline } from './src/index';
import * as fs from 'fs';

async function testImageLayout() {
  // First, let's use an existing HWPX file with images
  const inputPath = './data/samples/sample3_input.hwpx';

  if (!fs.existsSync(inputPath)) {
    console.log(`❌ Input file not found: ${inputPath}`);
    console.log('Looking for any sample files...');

    // Find any sample files
    const sampleDir = './data/samples/';
    if (fs.existsSync(sampleDir)) {
      const files = fs.readdirSync(sampleDir);
      console.log('Sample files:', files);
    }
    return;
  }

  console.log(`📄 Testing with: ${inputPath}`);
  const data = fs.readFileSync(inputPath);

  try {
    const pipeline = Pipeline.open(data, 'hwpx');

    // First, let's decode and see the document structure
    console.log('\n--- Decoding HWPX ---');
    const docResult = await pipeline.inspect();

    console.log('docResult.ok:', docResult.ok);
    console.log('docResult.data:', docResult.data);
    console.log('docResult.value:', docResult.value);

    if (docResult.ok && docResult.data) {
      console.log('✅ Decoded DOCX successfully');
      console.log('DocRoot keys:', Object.keys(docResult.data));
      console.log('DocRoot.tag:', docResult.data.tag);
      console.log('DocRoot.kids:', docResult.data.kids);
      console.log('DocRoot.sheets:', docResult.data.sheets);

      // Check for images
      const sheets = docResult.data.kids || docResult.data.sheets || [];
      sheets.forEach((sheet, sheetIdx) => {
        console.log(`\nSheet ${sheetIdx}:`);
        console.log('  sheet.tag:', sheet.tag);
        console.log('  sheet.kids type:', typeof sheet.kids, Array.isArray(sheet.kids) ? 'Array' : 'Object');

        const sheetKids = Array.isArray(sheet.kids) ? sheet.kids : Object.values(sheet.kids);
        sheetKids.forEach((kid: any, kidIdx: number) => {
          console.log(`  kid[${kidIdx}].tag:`, kid?.tag);

          if (kid?.tag === 'para') {
            kid.kids?.forEach((child: any, childIdx: number) => {
              console.log(`    child[${childIdx}].tag:`, child?.tag);
              if (child?.tag === 'img') {
                console.log(`      Image ${childIdx}: mime=${child.mime}, w=${child.w}, h=${child.h}, layout=`, child.layout);
              }
            });
          }
        });
        console.log(`\nSheet ${sheetIdx}:`);
        sheet.kids.forEach((kid, kidIdx) => {
          if (kid.tag === 'para') {
            kid.kids.forEach((child, childIdx) => {
              if (child.tag === 'img') {
                console.log(`  Image ${childIdx}: mime=${child.mime}, w=${child.wPt}, h=${child.hPt}, layout=`, child.layout);
              }
            });
          }
        });
      });

      // Convert to HWPX
      console.log('\n--- Converting to HWPX ---');
      const hwpxResult = await pipeline.to('hwpx');
      if (!hwpxResult.ok) {
        console.log(`❌ HWPX Encoding failed: ${hwpxResult.error}`);
        return;
      }

      console.log(`✅ HWPX created: ${hwpxResult.data.length} bytes`);
      fs.writeFileSync('./test_image.hwpx', hwpxResult.data);

      // Decode back
      console.log('\n--- Decoding HWPX back ---');
      const decodePipeline = Pipeline.open(hwpxResult.data, 'hwpx');
      const decodedDoc = await decodePipeline.inspect();

      if (decodedDoc.ok) {
        console.log('✅ Decoded HWPX successfully');

        // Check each sheet
        decodedDoc.value.kids.forEach((sheet, sheetIdx) => {
          console.log(`\nSheet ${sheetIdx}:`);

          // Check paragraphs
          sheet.kids.forEach((kid, kidIdx) => {
            if (kid.tag === 'para') {
              kid.kids.forEach((child, childIdx) => {
                if (child.tag === 'img') {
                  console.log(`  Image ${childIdx}: mime=${child.mime}, w=${child.wPt}, h=${child.hPt}, layout=`, child.layout);
                }
              });
            }
          });
        });
      } else {
        console.log(`❌ Decode failed: ${decodedDoc.error}`);
      }
    } else {
      console.log(`❌ DOCX Decode failed: ${docResult.error}`);
    }
  } catch (e: any) {
    console.error(`❌ EXCEPTION: ${e.message}`);
    console.error(e.stack);
  }
}

testImageLayout().catch(console.error);
