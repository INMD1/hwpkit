import * as fs from 'fs';
import pako from 'pako';

const hwp = fs.readFileSync('/mnt/92b2cb7e-8f06-4e4f-bfde-a14c87f2f96c/Github/hwpkit/data/sample/sample1_input.hwp');
const SS = 512;
const buf = hwp.buffer.slice(hwp.byteOffset, hwp.byteOffset + hwp.length);

function u32LE(b: Buffer, off: number) {
  return b[off] | (b[off+1]<<8) | (b[off+2]<<16) | ((b[off+3]&0xFF)*0x1000000);
}
function i32LE(b: Buffer, off: number) {
  const v = u32LE(b,off);
  return v >= 0x80000000 ? v - 0x100000000 : v;
}
function u16LE(b: Buffer, off: number) { return b[off] | (b[off+1]<<8); }

const hdr = hwp.slice(0, 512);
const fatSec = u32LE(hdr as Buffer, 76);
const dirSec = u32LE(hdr as Buffer, 48);

function sectorSlice(sec: number) {
  const off = (sec+1)*SS;
  return hwp.slice(off, off+SS) as Buffer;
}

const fatSector = sectorSlice(fatSec);
function nextSec(sec: number) { return u32LE(fatSector as Buffer, (sec%128)*4); }

function readChain(startSec: number, maxB = 200000): Buffer {
  const chunks: Buffer[] = [];
  let sec = startSec, tot = 0;
  while (sec < 0xFFFFFFFE && tot < maxB) {
    chunks.push(sectorSlice(sec) as Buffer);
    tot += SS;
    sec = nextSec(sec);
  }
  return Buffer.concat(chunks);
}

const dirBuf = readChain(dirSec, 2048);
function readEntry(idx: number) {
  const base = idx*128;
  const nl = u16LE(dirBuf, base+64);
  let name = '';
  for (let i=0;i<(nl/2)-1&&i<32;i++) name += String.fromCharCode(u16LE(dirBuf, base+i*2));
  return {
    name, type: dirBuf[base+66],
    child: i32LE(dirBuf, base+76), right: i32LE(dirBuf, base+72),
    startSec: u32LE(dirBuf, base+116), size: u32LE(dirBuf, base+120)
  };
}

let docInfoE: any = null, sec0E: any = null;
function walk(idx: number, depth = 0) {
  if (idx < 0 || idx > 100) return;
  const e = readEntry(idx);
  if (depth < 5) console.log(' '.repeat(depth*2) + idx + ': "' + e.name + '" type=' + e.type + ' sz=' + e.size);
  if (e.name === 'DocInfo') docInfoE = e;
  if (e.name === 'Section0') sec0E = e;
  walk(e.child, depth+1);
  walk(e.right, depth);
}
walk(0);

if (!docInfoE || !sec0E) { console.log('NOT FOUND'); process.exit(1); }

const diC = readChain(docInfoE.startSec, docInfoE.size+512);
const s0C = readChain(sec0E.startSec, sec0E.size+512);
const diRaw = Buffer.from(pako.inflate(diC.slice(0,docInfoE.size)));
const s0Raw = Buffer.from(pako.inflate(s0C.slice(0,sec0E.size)));

function parseRecs(buf: Buffer) {
  const recs: any[] = []; let off = 0;
  while (off+4 <= buf.length) {
    const h = u32LE(buf,off), tag=h&0x3FF, lv=(h>>10)&0x3FF;
    let sz=(h>>>20)&0xFFF, hSz=4;
    if(sz===0xFFF){sz=u32LE(buf,off+4);hSz=8;}
    if(off+hSz+sz>buf.length) break;
    recs.push({tag,lv,sz,data:buf.slice(off+hSz,off+hSz+sz)});
    off+=hSz+sz;
  }
  return recs;
}

const s0 = parseRecs(s0Raw);
console.log('\nSection0 records:', s0.length);
for (const r of s0.slice(0, 50)) {
  const h = r.data.slice(0,Math.min(r.sz,50)).toString('hex');
  console.log(`tag=${r.tag} lv=${r.lv} sz=${r.sz} [${h}]`);
}
