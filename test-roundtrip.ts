import { Pipeline } from './src/index';
import * as fs from 'fs';

async function testRoundtrip() {
  // First, check the original HWPX
  const inputPath = './data/samples/sample3_input.hwpx';

  if (!fs.existsSync(inputPath)) {
    console.log(`❌ Input file not found: ${inputPath}`);
    return;
  }

  console.log(`📄 Testing roundtrip with: ${inputPath}`);
  const hwpxData = fs.readFileSync(inputPath);

  // Check original HWPX structure
  console.log('\n--- Original HWPX ---');
  const origPipeline = Pipeline.open(hwpxData, 'hwpx');
  const origDoc = await origPipeline.inspect();
  if (origDoc.ok) {
    const origSheets = origDoc.data.kids || [];
    origSheets.forEach((sheet: any, sheetIdx: number) => {
      console.log(`Sheet ${sheetIdx}:`);
      const sheetKids = Array.isArray(sheet.kids) ? sheet.kids : Object.values(sheet.kids);
      sheetKids.forEach((kid: any, kidIdx: number) => {
        console.log(`  kid[${kidIdx}].tag:`, kid?.tag);
        if (kid?.tag === 'grid') {
          console.log(`    Grid found! rows=${kid.kids?.length}`);
          kid.kids?.forEach((row: any, rowIdx: number) => {
            row.kids?.forEach((cell: any, cellIdx: number) => {
              cell.kids?.forEach((para: any, paraIdx: number) => {
                para.kids?.forEach((child: any, childIdx: number) => {
                  console.log(`      Grid cell[${rowIdx}][${cellIdx}] para[${paraIdx}] child[${childIdx}].tag:`, child?.tag);
                  if (child?.tag === 'img') {
                    console.log(`        Image: w=${child.w}, h=${child.h}, layout=`, child.layout);
                  }
                });
              });
            });
          });
        } else if (kid?.tag === 'para') {
          kid.kids?.forEach((child: any, childIdx: number) => {
            console.log(`    child[${childIdx}].tag:`, child?.tag);
            if (child?.tag === 'img') {
              console.log(`      Image: w=${child.w}, h=${child.h}, layout=`, child.layout);
            }
          });
        }
      });
    });
  }

  try {
    // Step 1: HWPX → DOCX
    console.log('\n--- Step 1: HWPX → DOCX ---');
    const toDocxPipeline = Pipeline.open(hwpxData, 'hwpx');
    const docxResult = await toDocxPipeline.to('docx');

    if (!docxResult.ok) {
      console.log(`❌ HWPX → DOCX failed: ${docxResult.error}`);
      return;
    }

    console.log(`✅ HWPX → DOCX: ${docxResult.data.length} bytes`);

    // Check images in DOCX
    console.log('\n--- Check DOCX images ---');
    const docxPipeline = Pipeline.open(docxResult.data, 'docx');
    const docxDoc = await docxPipeline.inspect();
    if (docxDoc.ok) {
      const docxSheets = docxDoc.data.kids || [];
      docxSheets.forEach((sheet: any, sheetIdx: number) => {
        console.log(`Sheet ${sheetIdx}:`);
        console.log(`  sheet.kids type:`, typeof sheet.kids, Array.isArray(sheet.kids) ? 'Array' : 'Object');

        const sheetKids = Array.isArray(sheet.kids) ? sheet.kids : Object.values(sheet.kids);
        sheetKids.forEach((kid: any, kidIdx: number) => {
          console.log(`  kid[${kidIdx}].tag:`, kid?.tag);
          if (kid?.tag === 'grid') {
            console.log(`    Grid found! rows=${kid.kids?.length}`);
            kid.kids?.forEach((row: any, rowIdx: number) => {
              row.kids?.forEach((cell: any, cellIdx: number) => {
                cell.kids?.forEach((para: any, paraIdx: number) => {
                  para.kids?.forEach((child: any, childIdx: number) => {
                    console.log(`      Grid cell[${rowIdx}][${cellIdx}] para[${paraIdx}] child[${childIdx}].tag:`, child?.tag);
                    if (child?.tag === 'img') {
                      console.log(`        Image: w=${child.w}, h=${child.h}, layout=`, child.layout);
                    }
                  });
                });
              });
            });
          } else if (kid?.tag === 'para') {
            console.log(`    para.kids count:`, kid.kids?.length);
            kid.kids?.forEach((child: any, childIdx: number) => {
              console.log(`      child[${childIdx}].tag:`, child?.tag);
              if (child?.tag === 'img') {
                console.log(`        Image: w=${child.w}, h=${child.h}, layout=`, child.layout);
              }
            });
          }
        });
      });
    }

    // Step 2: DOCX → HWPX
    console.log('\n--- Step 2: DOCX → HWPX ---');
    const toHwpxPipeline = Pipeline.open(docxResult.data, 'docx');
    const hwpxResult = await toHwpxPipeline.to('hwpx');

    if (!hwpxResult.ok) {
      console.log(`❌ DOCX → HWPX failed: ${hwpxResult.warnings?.join(', ') || hwpxResult.error}`);
      return;
    }

    console.log(`✅ DOCX → HWPX: ${hwpxResult.data.length} bytes`);

    // Step 3: Decode the final HWPX to check image layout
    console.log('\n--- Step 3: Check final HWPX ---');
    const finalPipeline = Pipeline.open(hwpxResult.data, 'hwpx');
    const finalDoc = await finalPipeline.inspect();

    if (!finalDoc.ok) {
      console.log(`❌ Final decode failed: ${finalDoc.error}`);
      return;
    }

    // Check images in final document
    const sheets = finalDoc.data.kids || [];
    sheets.forEach((sheet: any, sheetIdx: number) => {
      console.log(`\nSheet ${sheetIdx}:`);
      const sheetKids = Array.isArray(sheet.kids) ? sheet.kids : Object.values(sheet.kids);
      sheetKids.forEach((kid: any) => {
        if (kid?.tag === 'para') {
          kid.kids?.forEach((child: any) => {
            if (child?.tag === 'img') {
              console.log(`  Image: w=${child.w}, h=${child.h}, layout=`, child.layout);
            }
          });
        }
      });
    });

  } catch (e: any) {
    console.error(`❌ EXCEPTION: ${e.message}`);
    console.error(e.stack);
  }
}

testRoundtrip().catch(console.error);
