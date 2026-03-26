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

  // JPEG: scan for SOF0/SOF2 marker
  if (u8[0] === 0xff && u8[1] === 0xd8) {
    let offset = 2;
    while (offset + 9 < buf.byteLength) {
      if (u8[offset] !== 0xff) { offset++; continue; }
      const marker = u8[offset + 1];
      if (marker === 0xc0 || marker === 0xc2) {
        return { width: view.getUint16(offset + 7), height: view.getUint16(offset + 5) };
      }
      const segLen = view.getUint16(offset + 2);
      offset += 2 + segLen;
    }
  }

  // GIF: bytes 6-9
  if (u8[0] === 0x47 && u8[1] === 0x49 && u8[2] === 0x46) {
    if (buf.byteLength >= 10) {
      return { width: view.getUint16(6, true), height: view.getUint16(8, true) };
    }
  }

  // BMP: bytes 18-25
  if (u8[0] === 0x42 && u8[1] === 0x4d) {
    if (buf.byteLength >= 26) {
      return { width: Math.abs(view.getInt32(18, true)), height: Math.abs(view.getInt32(22, true)) };
    }
  }

  // WebP
  if (u8[0] === 0x52 && u8[1] === 0x49 && u8[2] === 0x46 && u8[3] === 0x46 &&
      u8[8] === 0x57 && u8[9] === 0x45 && u8[10] === 0x42 && u8[11] === 0x50) {
    if (u8[12] === 0x56 && u8[13] === 0x50 && u8[14] === 0x38 && u8[15] === 0x20 && buf.byteLength >= 30) {
      return { width: view.getUint16(26, true) & 0x3fff, height: view.getUint16(28, true) & 0x3fff };
    }
    if (u8[12] === 0x56 && u8[13] === 0x50 && u8[14] === 0x38 && u8[15] === 0x4c && buf.byteLength >= 25) {
      const bits = view.getUint32(21, true);
      return { width: (bits & 0x3fff) + 1, height: ((bits >> 14) & 0x3fff) + 1 };
    }
  }

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

/** Guess image type from URL or Content-Type header. */
export function guessImageTypeFromUrl(url: string, contentType?: string): string {
  const ext = url.split("?")[0].split("#")[0].split(".").pop()?.toLowerCase() || "";
  const fromExt = guessImageType("img." + ext);
  if (fromExt !== "png" || ext === "png") return fromExt;
  if (contentType) {
    if (contentType.includes("jpeg") || contentType.includes("jpg")) return "jpg";
    if (contentType.includes("gif")) return "gif";
    if (contentType.includes("bmp")) return "bmp";
    if (contentType.includes("svg")) return "svg";
  }
  return "png";
}
