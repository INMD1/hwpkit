import { Pipeline } from './src/index';
import * as fs from 'fs';

async function testDocxToHwp() {
  const inputPath = './data/samples/original.docx';
  console.log(`\n📄 Testing DOCX → HWP conversion from: ${inputPath}`);
  const data = fs.readFileSync(inputPath);

  try {
    const pipeline = Pipeline.open(data, 'docx');

    console.log('Attempting to convert to HWP...');
    const result = await pipeline.to('hwp');

    if (result.ok) {
      console.log(`✅ Success! HWP output: ${result.data.length} bytes`);
      fs.writeFileSync('./output_test.hwp', result.data);
      console.log('Saved to: ./output_test.hwp');
    } else {
      console.log(`❌ HWP Failed: ${result.error}`);
    }

    console.log('\nAttempting to convert to HWPX...');
    const resultHwpx = await pipeline.to('hwpx');
    if (resultHwpx.ok) {
      console.log(`✅ Success! HWPX output: ${resultHwpx.data.length} bytes`);
      fs.writeFileSync('./output_test.hwpx', resultHwpx.data);
      console.log('Saved to: ./output_test.hwpx');
    } else {
      console.log(`❌ HWPX Failed: ${resultHwpx.error}`);
    }

    return result.ok && resultHwpx.ok;
  } catch (e: any) {
    console.error(`❌ EXCEPTION: ${e.message}`);
    console.error(e.stack);
    return false;
  }
}

testDocxToHwp().catch(console.error);
