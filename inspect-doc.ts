import { Pipeline, TreeWalker, countNodes } from './src/index';
import * as fs from 'fs';

async function inspectDoc(filePath: string, description: string) {
  console.log(`\n📄 ${description}`);
  console.log(`   File: ${filePath} (${fs.statSync(filePath).size} bytes)`);

  const data = fs.readFileSync(filePath);
  const pipeline = Pipeline.open(data);
  const result = await pipeline.inspect();

  if (!result.ok) {
    console.error(`   ❌ FAILED: ${result.error}`);
    return;
  }

  const doc = result.data;
  console.log(`   Meta: title="${doc.meta.title}", author="${doc.meta.author}"`);
  console.log(`   Sheets: ${doc.kids.length}`);

  const walker = new TreeWalker();
  let paraCount = 0, spanCount = 0, txtCount = 0, imgCount = 0, gridCount = 0;
  let totalChars = 0;

  walker.walk(doc, (node) => {
    switch (node.tag) {
      case 'para': paraCount++; break;
      case 'span': spanCount++; break;
      case 'txt': txtCount++; totalChars += node.content.length; break;
      case 'img': imgCount++; break;
      case 'grid': gridCount++; break;
    }
  });

  console.log(`   Paragraphs: ${paraCount}`);
  console.log(`   Spans: ${spanCount}`);
  console.log(`   Text nodes: ${txtCount}, Total chars: ${totalChars}`);
  console.log(`   Images: ${imgCount}`);
  console.log(`   Tables (grids): ${gridCount}`);

  if (result.warns.length > 0) {
    console.log(`   ⚠️  Warnings: ${result.warns.slice(0, 5).join(', ')}${result.warns.length > 5 ? '...' : ''}`);
  }
}

async function main() {
  // Input files
  await inspectDoc('./data/sample/sample1_input.hwp', 'Sample1 Input (HWP)');
  await inspectDoc('./data/sample/sample3_input.hwpx', 'Sample3 Input (HWPX)');
  await inspectDoc('./data/sample/sample4_input.docx', 'Sample4 Input (DOCX)');

  // Expected outputs
  await inspectDoc('./data/sample/sample1_output.docx', 'Sample1 Expected Output (DOCX)');
  await inspectDoc('./data/sample/sample1_output.hwpx', 'Sample1 Expected Output (HWPX)');
}

main().catch(console.error);
