/** Read image dimensions from raw binary data without browser APIs. */
export function getImageDimensions(buf: ArrayBuffer): { width: number; height: number } {
  const view = new DataView(buf);
  const u8 = new Uint8Array(buf);

  // PNG: bytes 16-23 contain width (4B BE) and height (4B BE)
  if (u8[0] === 0x89 && u8[1] === 0x50 && u8[2] === 0x4e && u8[3] === 0x47) {
    if (buf.byteLength >= 24) {
      return { width: view.getUint32(16), height: view.getUint32(20) };
    }
  }

  // JPEG: scan for SOF0/SOF2 marker (0xFF 0xC0 or 0xFF 0xC2)
  if (u8[0] === 0xff && u8[1] === 0xd8) {
    let offset = 2;
    while (offset + 9 < buf.byteLength) {
      if (u8[offset] !== 0xff) { offset++; continue; }
      const marker = u8[offset + 1];
      if (marker === 0xc0 || marker === 0xc2) {
        const height = view.getUint16(offset + 5);
        const width = view.getUint16(offset + 7);
        return { width, height };
      }
      // Skip this marker's segment
      const segLen = view.getUint16(offset + 2);
      offset += 2 + segLen;
    }
  }

  // GIF: bytes 6-9 contain width (2B LE) and height (2B LE)
  if (u8[0] === 0x47 && u8[1] === 0x49 && u8[2] === 0x46) {
    if (buf.byteLength >= 10) {
      return { width: view.getUint16(6, true), height: view.getUint16(8, true) };
    }
  }

  // BMP: bytes 18-25 contain width (4B LE) and height (4B LE)
  if (u8[0] === 0x42 && u8[1] === 0x4d) {
    if (buf.byteLength >= 26) {
      return { width: Math.abs(view.getInt32(18, true)), height: Math.abs(view.getInt32(22, true)) };
    }
  }

  // WebP (RIFF...WEBP)
  if (u8[0] === 0x52 && u8[1] === 0x49 && u8[2] === 0x46 && u8[3] === 0x46 &&
      u8[8] === 0x57 && u8[9] === 0x45 && u8[10] === 0x42 && u8[11] === 0x50) {
    // VP8 lossy
    if (u8[12] === 0x56 && u8[13] === 0x50 && u8[14] === 0x38 && u8[15] === 0x20) {
      if (buf.byteLength >= 30) {
        return { width: view.getUint16(26, true) & 0x3fff, height: view.getUint16(28, true) & 0x3fff };
      }
    }
    // VP8L lossless
    if (u8[12] === 0x56 && u8[13] === 0x50 && u8[14] === 0x38 && u8[15] === 0x4c) {
      if (buf.byteLength >= 25) {
        const bits = view.getUint32(21, true);
        return { width: (bits & 0x3fff) + 1, height: ((bits >> 14) & 0x3fff) + 1 };
      }
    }
  }

  // Fallback
  return { width: 400, height: 300 };
}

/** Map file extension to docx image type string. */
export function guessImageType(filename: string): string {
  const ext = (filename || "").split(".").pop()?.toLowerCase() || "";
  const map: Record<string, string> = {
    png: "png", jpg: "jpg", jpeg: "jpg", gif: "gif",
    bmp: "bmp", svg: "svg", webp: "png", tif: "png", tiff: "png",
  };
  return map[ext] || "png";
}
