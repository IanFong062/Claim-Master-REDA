const TEXT_ENCODER = new TextEncoder();
const TEXT_DECODER = new TextDecoder("utf-8");

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n += 1) {
    let c = n;
    for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(bytes) {
  let c = 0xffffffff;
  for (const byte of bytes) c = CRC_TABLE[(c ^ byte) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function readUInt16(bytes, offset) {
  return bytes[offset] | (bytes[offset + 1] << 8);
}

function readUInt32(bytes, offset) {
  return (
    bytes[offset] |
    (bytes[offset + 1] << 8) |
    (bytes[offset + 2] << 16) |
    (bytes[offset + 3] << 24)
  ) >>> 0;
}

function writeUInt16(out, offset, value) {
  out[offset] = value & 0xff;
  out[offset + 1] = (value >>> 8) & 0xff;
}

function writeUInt32(out, offset, value) {
  out[offset] = value & 0xff;
  out[offset + 1] = (value >>> 8) & 0xff;
  out[offset + 2] = (value >>> 16) & 0xff;
  out[offset + 3] = (value >>> 24) & 0xff;
}

async function inflateRaw(data) {
  if (typeof DecompressionStream === "undefined") {
    throw new Error("This browser cannot decompress the template ZIP. Please use current Safari, Chrome, Edge, or Firefox.");
  }
  const stream = new Blob([data]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

async function parseZip(arrayBuffer) {
  const bytes = new Uint8Array(arrayBuffer);
  let eocd = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 66000); i -= 1) {
    if (readUInt32(bytes, i) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw new Error("Invalid XLSX ZIP: central directory was not found.");

  const entryCount = readUInt16(bytes, eocd + 10);
  let centralOffset = readUInt32(bytes, eocd + 16);
  const entries = [];

  for (let i = 0; i < entryCount; i += 1) {
    if (readUInt32(bytes, centralOffset) !== 0x02014b50) throw new Error("Invalid ZIP central directory.");
    const method = readUInt16(bytes, centralOffset + 10);
    const modTime = readUInt16(bytes, centralOffset + 12);
    const modDate = readUInt16(bytes, centralOffset + 14);
    const compressedSize = readUInt32(bytes, centralOffset + 20);
    const fileNameLength = readUInt16(bytes, centralOffset + 28);
    const extraLength = readUInt16(bytes, centralOffset + 30);
    const commentLength = readUInt16(bytes, centralOffset + 32);
    const localOffset = readUInt32(bytes, centralOffset + 42);
    const name = TEXT_DECODER.decode(bytes.slice(centralOffset + 46, centralOffset + 46 + fileNameLength));

    const localNameLength = readUInt16(bytes, localOffset + 26);
    const localExtraLength = readUInt16(bytes, localOffset + 28);
    const dataStart = localOffset + 30 + localNameLength + localExtraLength;
    const compressed = bytes.slice(dataStart, dataStart + compressedSize);
    const data = method === 0 ? compressed : await inflateRaw(compressed);
    entries.push({ name, data, modTime, modDate });
    centralOffset += 46 + fileNameLength + extraLength + commentLength;
  }
  return entries;
}

function createZip(entries) {
  const localParts = [];
  const centralParts = [];
  let offset = 0;

  for (const entry of entries) {
    const nameBytes = TEXT_ENCODER.encode(entry.name);
    const data = entry.data;
    const crc = crc32(data);
    const local = new Uint8Array(30 + nameBytes.length);
    writeUInt32(local, 0, 0x04034b50);
    writeUInt16(local, 4, 20);
    writeUInt16(local, 6, 0x0800);
    writeUInt16(local, 8, 0);
    writeUInt16(local, 10, entry.modTime || 0);
    writeUInt16(local, 12, entry.modDate || 0);
    writeUInt32(local, 14, crc);
    writeUInt32(local, 18, data.length);
    writeUInt32(local, 22, data.length);
    writeUInt16(local, 26, nameBytes.length);
    local.set(nameBytes, 30);
    localParts.push(local, data);

    const central = new Uint8Array(46 + nameBytes.length);
    writeUInt32(central, 0, 0x02014b50);
    writeUInt16(central, 4, 20);
    writeUInt16(central, 6, 20);
    writeUInt16(central, 8, 0x0800);
    writeUInt16(central, 10, 0);
    writeUInt16(central, 12, entry.modTime || 0);
    writeUInt16(central, 14, entry.modDate || 0);
    writeUInt32(central, 16, crc);
    writeUInt32(central, 20, data.length);
    writeUInt32(central, 24, data.length);
    writeUInt16(central, 28, nameBytes.length);
    writeUInt32(central, 42, offset);
    central.set(nameBytes, 46);
    centralParts.push(central);
    offset += local.length + data.length;
  }

  const centralOffset = offset;
  const centralSize = centralParts.reduce((sum, part) => sum + part.length, 0);
  const eocd = new Uint8Array(22);
  writeUInt32(eocd, 0, 0x06054b50);
  writeUInt16(eocd, 8, entries.length);
  writeUInt16(eocd, 10, entries.length);
  writeUInt32(eocd, 12, centralSize);
  writeUInt32(eocd, 16, centralOffset);

  const size = offset + centralSize + eocd.length;
  const output = new Uint8Array(size);
  let cursor = 0;
  for (const part of localParts) {
    output.set(part, cursor);
    cursor += part.length;
  }
  for (const part of centralParts) {
    output.set(part, cursor);
    cursor += part.length;
  }
  output.set(eocd, cursor);
  return output;
}

function excelSerialDate(yyyyMmDd) {
  if (!yyyyMmDd) return null;
  const [year, month, day] = yyyyMmDd.split("-").map(Number);
  const utc = Date.UTC(year, month - 1, day);
  return Math.round(utc / 86400000 + 25569);
}

function parseAddress(address) {
  const match = /^([A-Z]+)(\d+)$/.exec(address);
  if (!match) throw new Error(`Invalid cell address: ${address}`);
  let col = 0;
  for (const letter of match[1]) col = col * 26 + letter.charCodeAt(0) - 64;
  return { col, row: Number(match[2]) };
}

function findCell(row, address) {
  return Array.from(row.children).find((node) => node.localName === "c" && node.getAttribute("r") === address);
}

function ensureCell(doc, row, address) {
  let cell = findCell(row, address);
  if (cell) return cell;
  cell = doc.createElementNS(row.namespaceURI, "c");
  cell.setAttribute("r", address);
  const target = parseAddress(address).col;
  const before = Array.from(row.children).find((node) => {
    if (node.localName !== "c") return false;
    return parseAddress(node.getAttribute("r")).col > target;
  });
  row.insertBefore(cell, before || null);
  return cell;
}

function ensureRow(doc, sheetData, rowNumber) {
  let row = Array.from(sheetData.children).find((node) => node.localName === "row" && node.getAttribute("r") === String(rowNumber));
  if (row) return row;
  row = doc.createElementNS(sheetData.namespaceURI, "row");
  row.setAttribute("r", String(rowNumber));
  const before = Array.from(sheetData.children).find((node) => node.localName === "row" && Number(node.getAttribute("r")) > rowNumber);
  sheetData.insertBefore(row, before || null);
  return row;
}

function clearCell(cell) {
  Array.from(cell.children).forEach((child) => {
    if (["f", "v", "is"].includes(child.localName)) cell.removeChild(child);
  });
  cell.removeAttribute("t");
}

function setCellValue(doc, cell, value, styleId) {
  clearCell(cell);
  if (styleId !== undefined && styleId !== null) cell.setAttribute("s", String(styleId));
  if (value === null || value === undefined || value === "") return;
  if (typeof value === "number") {
    const v = doc.createElementNS(cell.namespaceURI, "v");
    v.textContent = String(value);
    cell.appendChild(v);
    return;
  }
  cell.setAttribute("t", "inlineStr");
  const is = doc.createElementNS(cell.namespaceURI, "is");
  const t = doc.createElementNS(cell.namespaceURI, "t");
  t.setAttribute("xml:space", "preserve");
  t.textContent = String(value);
  is.appendChild(t);
  cell.appendChild(is);
}

function appendStyle(stylesXml, baseStyleIds, numFmtId, formatCode) {
  const doc = new DOMParser().parseFromString(stylesXml, "application/xml");
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  let numFmts = Array.from(root.children).find((node) => node.localName === "numFmts");
  if (!numFmts) {
    numFmts = doc.createElementNS(ns, "numFmts");
    root.insertBefore(numFmts, root.firstChild);
  }
  const exists = Array.from(numFmts.children).some((node) => node.getAttribute("numFmtId") === String(numFmtId));
  if (!exists) {
    const numFmt = doc.createElementNS(ns, "numFmt");
    numFmt.setAttribute("numFmtId", String(numFmtId));
    numFmt.setAttribute("formatCode", formatCode);
    numFmts.appendChild(numFmt);
    numFmts.setAttribute("count", String(numFmts.children.length));
  }

  const cellXfs = Array.from(root.children).find((node) => node.localName === "cellXfs");
  const originalXfs = Array.from(cellXfs.children);
  const styleMap = {};
  for (const base of baseStyleIds) {
    const baseNode = originalXfs[Number(base)] || originalXfs[0];
    const clone = baseNode.cloneNode(true);
    clone.setAttribute("numFmtId", String(numFmtId));
    clone.setAttribute("applyNumberFormat", "1");
    cellXfs.appendChild(clone);
    styleMap[String(base)] = cellXfs.children.length - 1;
  }
  cellXfs.setAttribute("count", String(cellXfs.children.length));
  return { xml: new XMLSerializer().serializeToString(doc), styleMap };
}

function patchStyles(stylesXml) {
  const dateResult = appendStyle(stylesXml, ["8"], 165, "dd/mm/yyyy");
  const moneyResult = appendStyle(dateResult.xml, ["9", "10", "21", "40", "45", "48", "50"], 166, "0.00");
  return { xml: moneyResult.xml, dateStyles: dateResult.styleMap, moneyStyles: moneyResult.styleMap };
}

function assign(sheetDoc, address, value, styleMap) {
  const sheetData = Array.from(sheetDoc.documentElement.children).find((node) => node.localName === "sheetData");
  const parsed = parseAddress(address);
  const row = ensureRow(sheetDoc, sheetData, parsed.row);
  const cell = ensureCell(sheetDoc, row, address);
  const originalStyle = cell.getAttribute("s");
  const styleId = styleMap && originalStyle && styleMap[originalStyle] ? styleMap[originalStyle] : undefined;
  setCellValue(sheetDoc, cell, value, styleId);
}

function tripRows(mapping) {
  return mapping.tripSections.flatMap((section) =>
    Array.from({ length: section.endRow - section.startRow + 1 }, (_, i) => section.startRow + i)
  );
}

function money(value) {
  const n = Number(value || 0);
  return Number.isFinite(n) ? Math.round(n * 100) / 100 : 0;
}

function buildPatches(claim, mapping) {
  if (claim.trips.length > mapping.maxTrips) {
    throw new Error(`Template supports ${mapping.maxTrips} trips only. You entered ${claim.trips.length}.`);
  }
  const patches = [];
  const week = claim.weekEnding || "";
  patches.push(["B1", `Travelling Report from ${claim.weekStart || "_________"} to ${week || "__________"} (以一週計即星期一至星期日) `]);
  patches.push(["I1", `  僱員姓名 : ${claim.employeeName || ""}${claim.employeeNumber ? ` (${claim.employeeNumber})` : ""}`]);
  patches.push(["B2", `僱員住址: ${claim.address || ""}`]);
  patches.push(["B3", `Claim Date: ${claim.claimDate || ""}${claim.remarks ? ` | Remarks: ${claim.remarks}` : ""}`]);

  const rows = tripRows(mapping);
  claim.trips.forEach((trip, index) => {
    const row = rows[index];
    const transport = trip.transportType || "Other";
    const amount = money(trip.amount);
    const deduction = money(trip.deduction);
    const col = mapping.transportColumns[transport];
    const locSuffix = transport === "Other" ? ` [Other: ${amount.toFixed(2)}]` : "";
    patches.push([`B${row}`, { type: "date", value: excelSerialDate(trip.date) }]);
    patches.push([`C${row}`, `${trip.from || ""} → ${trip.to || ""}${locSuffix}`]);
    patches.push([`E${row}`, trip.timeFrom || ""]);
    patches.push([`F${row}`, trip.timeTo || ""]);
    patches.push([`G${row}`, trip.client || ""]);
    patches.push([`H${row}`, trip.projectCode || ""]);
    if (col) patches.push([`${col}${row}`, { type: "money", value: amount }]);
    patches.push([`M${row}`, { type: "money", value: amount }]);
    patches.push([`N${row}`, { type: "money", value: deduction }]);
  });

  const firstPageTrips = Math.min(claim.trips.length, mapping.tripSections[0].endRow - mapping.tripSections[0].startRow + 1);
  const total1 = claim.trips.slice(0, firstPageTrips).reduce((sum, trip) => sum + money(trip.amount) - money(trip.deduction), 0);
  const total2 = claim.trips.slice(firstPageTrips).reduce((sum, trip) => sum + money(trip.amount) - money(trip.deduction), 0);
  patches.push(["N29", { type: "money", value: money(total1) }]);
  patches.push(["N58", { type: "money", value: money(total2) }]);
  return patches;
}

export async function generateClaimWorkbook(claim, mapping) {
  const response = await fetch(mapping.templateFile);
  if (!response.ok) throw new Error("Template file cannot be loaded.");
  const entries = await parseZip(await response.arrayBuffer());
  const styleEntry = entries.find((entry) => entry.name === mapping.workbook.stylesXml);
  const sheetEntry = entries.find((entry) => entry.name === mapping.workbook.sheetXml.Sheet1);
  if (!styleEntry || !sheetEntry) throw new Error("Template workbook structure is not as expected.");

  const stylePatch = patchStyles(TEXT_DECODER.decode(styleEntry.data));
  styleEntry.data = TEXT_ENCODER.encode(stylePatch.xml);

  const sheetDoc = new DOMParser().parseFromString(TEXT_DECODER.decode(sheetEntry.data), "application/xml");
  for (const [address, payload] of buildPatches(claim, mapping)) {
    if (payload && typeof payload === "object" && payload.type === "date") {
      assign(sheetDoc, address, payload.value, stylePatch.dateStyles);
    } else if (payload && typeof payload === "object" && payload.type === "money") {
      assign(sheetDoc, address, payload.value, stylePatch.moneyStyles);
    } else {
      assign(sheetDoc, address, payload, null);
    }
  }
  sheetEntry.data = TEXT_ENCODER.encode(new XMLSerializer().serializeToString(sheetDoc));
  return new Blob([createZip(entries)], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
}

export function exportedTripRowCount(claim, mapping) {
  return Math.min(claim.trips.length, mapping.maxTrips);
}

export function calculateGrandTotal(claim) {
  return claim.trips.reduce((sum, trip) => sum + money(trip.amount) - money(trip.deduction), 0);
}

export function safeFileName(claim) {
  const date = claim.claimDate || new Date().toISOString().slice(0, 10);
  const employee = (claim.employeeName || "Employee").replace(/[\\/:*?"<>|]+/g, "").replace(/\s+/g, "_");
  return `Travelling_Claim_${date}_${employee}.xlsx`;
}
