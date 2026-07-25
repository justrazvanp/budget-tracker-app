/**
 * Generates app icons (192x192, 512x512) as plain PNG files, using only Node's
 * built-in zlib for compression — no image libraries or external assets needed.
 * Run once with: node scripts/generate-icons.js
 */
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

const BG = [0x0f, 0x5c, 0x4a];       // dark green background
const COIN_FILL = [0xf5, 0xd4, 0x85]; // warm gold coin
const COIN_RIM = [0xc9, 0x9a, 0x2e];  // darker gold rim
const SUPERSAMPLE = 4;

function inCircle(x, y, cx, cy, r) {
  const dx = x - cx, dy = y - cy;
  return dx * dx + dy * dy <= r * r;
}

function renderHiRes(size) {
  const S = size * SUPERSAMPLE;
  const buf = new Uint8ClampedArray(S * S * 3);

  // background fill
  for (let i = 0; i < S * S; i++) {
    buf[i * 3] = BG[0];
    buf[i * 3 + 1] = BG[1];
    buf[i * 3 + 2] = BG[2];
  }

  // three overlapping coins suggesting a small stack of money
  const coins = [
    { cx: 0.34 * S, cy: 0.66 * S, r: 0.24 * S },
    { cx: 0.50 * S, cy: 0.50 * S, r: 0.24 * S },
    { cx: 0.66 * S, cy: 0.34 * S, r: 0.24 * S }
  ];

  for (const coin of coins) {
    for (let y = Math.floor(coin.cy - coin.r - 2); y <= coin.cy + coin.r + 2; y++) {
      if (y < 0 || y >= S) continue;
      for (let x = Math.floor(coin.cx - coin.r - 2); x <= coin.cx + coin.r + 2; x++) {
        if (x < 0 || x >= S) continue;
        const idx = (y * S + x) * 3;
        if (inCircle(x, y, coin.cx, coin.cy, coin.r)) {
          const isRim = !inCircle(x, y, coin.cx, coin.cy, coin.r * 0.82);
          const color = isRim ? COIN_RIM : COIN_FILL;
          buf[idx] = color[0];
          buf[idx + 1] = color[1];
          buf[idx + 2] = color[2];
        }
      }
    }
  }

  return buf;
}

function downsample(hiBuf, size) {
  const S = size * SUPERSAMPLE;
  const out = new Uint8ClampedArray(size * size * 4); // RGBA
  const factor = SUPERSAMPLE;
  for (let y = 0; y < size; y++) {
    for (let x = 0; x < size; x++) {
      let r = 0, g = 0, b = 0;
      for (let sy = 0; sy < factor; sy++) {
        for (let sx = 0; sx < factor; sx++) {
          const srcIdx = ((y * factor + sy) * S + (x * factor + sx)) * 3;
          r += hiBuf[srcIdx];
          g += hiBuf[srcIdx + 1];
          b += hiBuf[srcIdx + 2];
        }
      }
      const n = factor * factor;
      const dstIdx = (y * size + x) * 4;
      out[dstIdx] = Math.round(r / n);
      out[dstIdx + 1] = Math.round(g / n);
      out[dstIdx + 2] = Math.round(b / n);
      out[dstIdx + 3] = 255;
    }
  }
  return out;
}

function crc32(buf) {
  let c;
  const table = crc32.table || (crc32.table = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })());
  let crc = 0xffffffff;
  for (let i = 0; i < buf.length; i++) crc = table[(crc ^ buf[i]) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

function chunk(type, data) {
  const typeBuf = Buffer.from(type, 'ascii');
  const lenBuf = Buffer.alloc(4);
  lenBuf.writeUInt32BE(data.length, 0);
  const crcBuf = Buffer.alloc(4);
  crcBuf.writeUInt32BE(crc32(Buffer.concat([typeBuf, data])), 0);
  return Buffer.concat([lenBuf, typeBuf, data, crcBuf]);
}

function encodePNG(rgba, size) {
  const signature = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);

  const ihdrData = Buffer.alloc(13);
  ihdrData.writeUInt32BE(size, 0);
  ihdrData.writeUInt32BE(size, 4);
  ihdrData[8] = 8;  // bit depth
  ihdrData[9] = 6;  // color type RGBA
  ihdrData[10] = 0; // compression
  ihdrData[11] = 0; // filter
  ihdrData[12] = 0; // interlace
  const ihdr = chunk('IHDR', ihdrData);

  // raw scanlines, each prefixed with filter type 0 (none)
  const raw = Buffer.alloc(size * (size * 4 + 1));
  for (let y = 0; y < size; y++) {
    const rowStart = y * (size * 4 + 1);
    raw[rowStart] = 0;
    for (let x = 0; x < size * 4; x++) {
      raw[rowStart + 1 + x] = rgba[y * size * 4 + x];
    }
  }
  const compressed = zlib.deflateSync(raw, { level: 9 });
  const idat = chunk('IDAT', compressed);
  const iend = chunk('IEND', Buffer.alloc(0));

  return Buffer.concat([signature, ihdr, idat, iend]);
}

function generateIcon(size, outPath) {
  const hi = renderHiRes(size);
  const rgba = downsample(hi, size);
  const png = encodePNG(rgba, size);
  fs.writeFileSync(outPath, png);
  console.log('Wrote ' + outPath + ' (' + png.length + ' bytes)');
}

const outDir = path.join(__dirname, '..', 'icons');
fs.mkdirSync(outDir, { recursive: true });
generateIcon(192, path.join(outDir, 'icon-192.png'));
generateIcon(512, path.join(outDir, 'icon-512.png'));
