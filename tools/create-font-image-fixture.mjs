import * as kit from '../dist/index.mjs';
import {mkdirSync,writeFileSync,readFileSync} from 'node:fs';
const dir=process.argv[2];
if (!dir) throw new Error('Usage: node tools/create-font-image-fixture.mjs OUTPUT_DIRECTORY');
mkdirSync(dir,{recursive:true});
const label='font-image';
const {buildRoot,buildSheet,buildPara,buildSpan,buildImg,buildGrid,buildRow,buildCell,registry,Pipeline}=kit;
const png='iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';
const doc=buildRoot({},[buildSheet([
 buildPara([buildSpan('FONT AND PICTURE ANCHOR CHECK',{font:'Arial',b:true,pt:16})]),
 buildPara([buildSpan('Regular 한글 ABC 123',{font:'HY견고딕',fontFaces:{hangul:'HY견고딕',latin:'Arial'},pt:12})]),
 buildPara([buildSpan('Bold 한글 ABC 123',{font:'HY견고딕',fontFaces:{hangul:'HY견고딕',latin:'Arial'},pt:12,b:true})]),
 buildPara([buildSpan('Italic 한글 ABC 123',{font:'바탕',fontFaces:{hangul:'바탕',latin:'Times New Roman'},pt:12,i:true})]),
 buildPara([buildSpan('BEFORE ',{font:'Arial',pt:12}),buildImg(png,'image/png',36,24),buildSpan(' AFTER',{font:'Arial',pt:12})],{align:'center'}),
 buildGrid([buildRow([buildCell([buildPara([buildSpan('CELL BEFORE ',{font:'Arial'}),buildImg(png,'image/png',24,16),buildSpan(' CELL AFTER',{font:'Arial',b:true})])])])]),
 buildPara([buildSpan('END OF FLOW',{font:'Arial',pt:12})]),
 buildPara([buildImg(png,'image/png',30,30,'paper-anchor',{wrap:'front',horzRelTo:'page',vertRelTo:'page',horzAlign:'left',vertAlign:'top',xPt:30,yPt:30})]),
])]);
for(const format of ['hwp','hwpx','docx']) {
 const out=await registry.getEncoder(format).encode(doc); if(!out.ok)throw new Error(out.error);
 writeFileSync(`${dir}/${label}-style.${format}`,out.data);
}
const demo=await Pipeline.open(new Uint8Array(readFileSync(new URL('../tests/fixtures/demo.docx', import.meta.url))),'docx').to('hwp');
if(!demo.ok)throw new Error(demo.error);writeFileSync(`${dir}/${label}-demo.hwp`,demo.data);
console.log(dir,label);
