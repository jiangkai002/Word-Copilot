/**
 * 生成 manifest.xml 引用的占位图标（纯 Node，无依赖）：
 *   frontend/public/assets/icon-{16,32,80}.png
 * 蓝底白色圆点（Copilot 风格占位图），可随时替换为正式设计稿。
 *
 * 用法：node scripts/generate-icons.mjs
 */
import { writeFileSync, mkdirSync } from "node:fs";
import { deflateSync } from "node:zlib";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const SCRIPT_DIR = dirname(fileURLToPath(import.meta.url));
const OUT_DIR = resolve(SCRIPT_DIR, "../public/assets");

const BG = [15, 108, 189]; // #0F6CBD Office 主蓝
const FG = [255, 255, 255];

// ---------- PNG 编码（最小实现：8bit RGB、无压缩滤波） ----------

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n += 1) {
    let c = n;
    for (let k = 0; k < 8; k += 1) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i += 1) {
    c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  }
  return (c ^ 0xffffffff) >>> 0;
}

function chunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length);
  const typeBuf = Buffer.from(type, "ascii");
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(Buffer.concat([typeBuf, data])));
  return Buffer.concat([len, typeBuf, data, crc]);
}

function createPng(size, rgbAt) {
  const stride = size * 3;
  const raw = Buffer.alloc((stride + 1) * size);
  for (let y = 0; y < size; y += 1) {
    raw[y * (stride + 1)] = 0x00; // filter: None
    for (let x = 0; x < size; x += 1) {
      const [r, g, b] = rgbAt(x, y);
      const offset = y * (stride + 1) + 1 + x * 3;
      raw[offset] = r;
      raw[offset + 1] = g;
      raw[offset + 2] = b;
    }
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 2; // color type: truecolor RGB
  const signature = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  return Buffer.concat([
    signature,
    chunk("IHDR", ihdr),
    chunk("IDAT", deflateSync(raw)),
    chunk("IEND", Buffer.alloc(0)),
  ]);
}

// ---------- 绘制：蓝底 + 白色圆点 ----------

function iconPixels(size) {
  const center = (size - 1) / 2;
  const radius = size * 0.36;
  return (x, y) => {
    const dx = x - center;
    const dy = y - center;
    return dx * dx + dy * dy <= radius * radius ? FG : BG;
  };
}

mkdirSync(OUT_DIR, { recursive: true });
for (const size of [16, 32, 80]) {
  const png = createPng(size, iconPixels(size));
  const target = resolve(OUT_DIR, `icon-${size}.png`);
  writeFileSync(target, png);
  console.log(`已生成 ${target}（${png.length} 字节）`);
}
