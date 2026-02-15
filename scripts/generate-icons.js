/**
 * Generates simple PNG icon files for the Outlook Add-in manifest.
 * Creates solid-colored squares with a "C" letter at required sizes.
 * Uses only Node.js built-ins (zlib for PNG compression).
 */

const fs = require("fs");
const path = require("path");
const zlib = require("zlib");

const ICON_COLOR = { r: 99, g: 102, b: 241 }; // Indigo accent
const LETTER_COLOR = { r: 255, g: 255, b: 255 }; // White
const SIZES = [16, 32, 64, 80, 128];
const OUTPUT_DIR = path.resolve(__dirname, "..", "src", "assets");

/**
 * Compute CRC32 checksum for PNG chunk validation.
 * @param {Buffer} buf
 * @returns {number}
 */
function crc32(buf) {
  let crc = 0xffffffff;
  for (let i = 0; i < buf.length; i++) {
    crc ^= buf[i];
    for (let j = 0; j < 8; j++) {
      crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

/**
 * Create a single PNG chunk (length + type + data + crc).
 * @param {string} type
 * @param {Buffer} data
 * @returns {Buffer}
 */
function createChunk(type, data) {
  const length = Buffer.alloc(4);
  length.writeUInt32BE(data.length, 0);

  const typeBuffer = Buffer.from(type, "ascii");
  const crcInput = Buffer.concat([typeBuffer, data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(crcInput), 0);

  return Buffer.concat([length, typeBuffer, data, crc]);
}

/**
 * Simple bitmap font for the letter "C" (7x9 grid, 1=filled 0=empty).
 */
const LETTER_C = [
  [0, 0, 1, 1, 1, 0, 0],
  [0, 1, 1, 0, 1, 1, 0],
  [1, 1, 0, 0, 0, 1, 1],
  [1, 1, 0, 0, 0, 0, 0],
  [1, 1, 0, 0, 0, 0, 0],
  [1, 1, 0, 0, 0, 0, 0],
  [1, 1, 0, 0, 0, 1, 1],
  [0, 1, 1, 0, 1, 1, 0],
  [0, 0, 1, 1, 1, 0, 0],
];

/**
 * Generate a PNG file: solid color background with a "C" letter.
 * @param {number} size - Width and height in pixels
 * @returns {Buffer}
 */
function createIconPNG(size) {
  // PNG signature
  const signature = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);

  // IHDR chunk data
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8;  // bit depth
  ihdr[9] = 2;  // color type: RGB
  ihdr[10] = 0; // compression
  ihdr[11] = 0; // filter
  ihdr[12] = 0; // interlace

  // Pixel data: filter byte + RGB per pixel, per row
  const rowBytes = 1 + size * 3;
  const rawData = Buffer.alloc(size * rowBytes);

  // Letter placement: centered, scaled to ~50% of icon size
  const letterW = LETTER_C[0].length;
  const letterH = LETTER_C.length;
  const scale = Math.max(1, Math.floor(size * 0.5 / Math.max(letterW, letterH)));
  const scaledW = letterW * scale;
  const scaledH = letterH * scale;
  const offsetX = Math.floor((size - scaledW) / 2);
  const offsetY = Math.floor((size - scaledH) / 2);

  for (let y = 0; y < size; y++) {
    const rowOffset = y * rowBytes;
    rawData[rowOffset] = 0; // filter: none

    for (let x = 0; x < size; x++) {
      const pixOffset = rowOffset + 1 + x * 3;

      // Check if this pixel is part of the letter
      const lx = Math.floor((x - offsetX) / scale);
      const ly = Math.floor((y - offsetY) / scale);
      const isLetter =
        x >= offsetX && x < offsetX + scaledW &&
        y >= offsetY && y < offsetY + scaledH &&
        ly >= 0 && ly < letterH &&
        lx >= 0 && lx < letterW &&
        LETTER_C[ly][lx] === 1;

      const color = isLetter ? LETTER_COLOR : ICON_COLOR;
      rawData[pixOffset] = color.r;
      rawData[pixOffset + 1] = color.g;
      rawData[pixOffset + 2] = color.b;
    }
  }

  const compressed = zlib.deflateSync(rawData);

  return Buffer.concat([
    signature,
    createChunk("IHDR", ihdr),
    createChunk("IDAT", compressed),
    createChunk("IEND", Buffer.alloc(0)),
  ]);
}

// Generate icons
if (!fs.existsSync(OUTPUT_DIR)) {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

for (const size of SIZES) {
  const png = createIconPNG(size);
  const filePath = path.join(OUTPUT_DIR, `icon-${size}.png`);
  fs.writeFileSync(filePath, png);
  console.log(`Generated ${filePath} (${size}x${size}, ${png.length} bytes)`);
}

console.log("Icon generation complete.");
