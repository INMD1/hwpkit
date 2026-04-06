import { Pipeline } from './src/index';
import * as fs from 'fs';

async function testDocxToHwp() {
  const inputPath = './data/sample/sample4_input.docx';
  console.log(`\n📄 Testing DOCX → HWP conversion from: ${inputPath}`);
  const data = fs.readFileSync(inputPath);

  try {
    const pipeline = Pipeline.open(data, 'docx');

    console.log('Attempting to convert to HWP...');
    const result = await pipeline.to('hwp');

    if (result.ok) {
      console.log(`✅ Success! HWP output: ${result.data.length} bytes`);

      // Save to file for verification
      fs.writeFileSync('./output_test.hwp', result.data);
      console.log('Saved to: ./output_test.hwp');

      // Verify by converting back to MD
      console.log('\n--- Verifying by converting back to MD ---');
      const verifyPipeline = Pipeline.open(result.data, 'hwp');
      const mdResult = await verifyPipeline.to('md');
      if (mdResult.ok) {
        const mdText = new TextDecoder().decode(mdResult.data);
        console.log(`Verification MD output: ${mdText.length} bytes`);
        console.log('First 500 chars:', mdText.substring(0, 500));
      } else {
        console.log(`Verification failed: ${mdResult.error}`);
      }
    } else {
      console.log(`❌ Failed: ${result.error}`);
    }

    return result.ok;
  } catch (e: any) {
    console.error(`❌ EXCEPTION: ${e.message}`);
    console.error(e.stack);
    return false;
  }
}

testDocxToHwp().catch(console.error);
