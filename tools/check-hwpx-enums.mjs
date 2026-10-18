// Usage: node tools/check-hwpx-enums.mjs FILE.hwpx [...]
// Checks enum attribute values in Contents/*.xml against the Hancom OWPML enum
// names in tests/fixtures/hancom-owpml-enums.json. Hancom matches names exactly.
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { SaxesParser } from 'saxes';

const files = process.argv.slice(2);
if (!files.length) throw new Error('Expected one or more .hwpx files');
const dictionary = JSON.parse(
  await readFile(new URL('../tests/fixtures/hancom-owpml-enums.json', import.meta.url), 'utf8'),
).attributes;
let checked = 0;
let violations = 0;
for (const path of files) {
  const zip = await JSZip.loadAsync(await readFile(path));
  for (const entry of Object.values(zip.files)) {
    if (!/^Contents\/.*\.xml$/.test(entry.name)) continue;
    const parser = new SaxesParser({ xmlns: true });
    parser.on('opentag', tag => {
      for (const attr of Object.values(tag.attributes)) {
        const allowed = dictionary[attr.local];
        if (!allowed || attr.local === 'value' || /^-?\d+$/.test(attr.value)) continue;
        checked++;
        if (allowed.includes(attr.value)) continue;
        violations++;
        console.log(`${path} ${entry.name}: <${tag.name} ${attr.local}="${attr.value}">`);
      }
    });
    parser.write(await entry.async('string')).close();
  }
}
console.log(`files ${files.length}, enum attributes ${checked}, violations ${violations}`);
process.exitCode = violations ? 1 : 0;
