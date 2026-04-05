#!/usr/bin/env node
/**
 * 간단한 PNG 아이콘 생성 스크립트 (Node.js 내장 모듈만 사용)
 * 색상: #0078d4 (Microsoft Blue)
 */
const zlib = require('zlib');
const fs = require('fs');
const path = require('path');

// CRC32 테이블
const crcTable = new Uint32Array(256);
for (let i = 0; i < 256; i++) {
  let c = i;
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  crcTable[i] = c;
}
function crc32(buf) {
  let crc = 0xffffffff;
  for (let i = 0; i < buf.length; i++) crc = crcTable[(crc ^ buf[i]) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}
function chunk(type, data) {
  const len = Buffer.alloc(4); len.writeUInt32BE(data.length, 0);
  const t = Buffer.from(type, 'ascii');
  const crcBuf = Buffer.alloc(4); crcBuf.writeUInt32BE(crc32(Buffer.concat([t, data])), 0);
  return Buffer.concat([len, t, data, crcBuf]);
}

/**
 * 단색 PNG 생성
 */
function createPNG(size, r, g, b) {
  const sig = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);

  const ihdrData = Buffer.alloc(13);
  ihdrData.writeUInt32BE(size, 0);
  ihdrData.writeUInt32BE(size, 4);
  ihdrData[8] = 8; ihdrData[9] = 2; // 8-bit RGB

  // 픽셀 데이터: filter(1) + RGB*size per row
  const rowLen = 1 + size * 3;
  const raw = Buffer.alloc(rowLen * size);
  for (let y = 0; y < size; y++) {
    raw[y * rowLen] = 0; // filter: None
    for (let x = 0; x < size; x++) {
      const o = y * rowLen + 1 + x * 3;
      raw[o] = r; raw[o + 1] = g; raw[o + 2] = b;
    }
  }
  const compressed = zlib.deflateSync(raw);

  return Buffer.concat([
    sig,
    chunk('IHDR', ihdrData),
    chunk('IDAT', compressed),
    chunk('IEND', Buffer.alloc(0)),
  ]);
}

const assetsDir = path.join(__dirname, '..', 'assets');
// Microsoft Blue: #0078d4 → r=0, g=120, b=212
const [r, g, b] = [0, 120, 212];

for (const size of [16, 32, 80]) {
  const outPath = path.join(assetsDir, `icon-${size}.png`);
  fs.writeFileSync(outPath, createPNG(size, r, g, b));
  console.log(`Created: ${outPath}`);
}
console.log('Done.');
