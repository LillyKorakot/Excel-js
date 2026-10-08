import ExcelJS from "exceljs";

/**
 * excelBuilder.js
 * function กลางสำหรับสร้างไฟล์ Excel (.xlsx) ด้วย ExcelJS
 *  - หลาย sheet
 *  - dropdown (list จาก API / ค่าตายตัว) และ validation แบบอื่นๆ
 *  - สูตร (ใช้ placeholder อ้างคอลัมน์ด้วย key ได้)
 *  - อ่านไฟล์ที่ผู้ใช้กรอกกลับมาเป็น JSON
 */

const LIST_SHEET = "Lists";
const INVALID_SHEET_CHARS = /[\\/?*[\]:]/;

// ---------- helpers ----------

const colLetter = (n) => {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
};

// รายการ dropdown เป็น string หรือ object {label} / {name} ก็ได้
const labelOf = (item) =>
  item !== null && typeof item === "object"
    ? item.label ?? item.name ?? item.value
    : item;

// แปลงค่าจาก cell ให้เป็นค่าธรรมดา (จัดการสูตร, rich text, hyperlink)
const plain = (v) => {
  if (v === null || v === undefined) return null;
  if (v instanceof Date) return v;
  if (typeof v === "object") {
    if ("result" in v) return plain(v.result);
    if (Array.isArray(v.richText)) return v.richText.map((r) => r.text).join("");
    if ("text" in v) return v.text;
    if ("error" in v) return null;
  }
  return v;
};

// แทน placeholder ในสูตร:  {qty}*{price}  ->  C2*D2   และ {row} -> เลขแถว
const buildFormula = (tpl, row, colMap) => {
  const raw = typeof tpl === "function" ? tpl(row, colMap) : tpl;
  return String(raw)
    .replace(/^=/, "")
    .replace(/\{([^}]+)\}/g, (_, k) => {
      if (k === "row") return String(row);
      if (colMap[k]) return `${colMap[k]}${row}`;
      throw new Error(`สูตรอ้างถึงคอลัมน์ "${k}" ที่ไม่มีใน columns`);
    });
};

// ---------- main: สร้าง workbook ----------

/**
 * config = {
 *   lists: { depts: ["ขาย","บัญชี"], ... },      // ค่าสำหรับ dropdown (มาจาก API ได้)
 *   sheets: [{
 *     name: "ข้อมูล",
 *     tabColor: "FF1F4E78",           // optional
 *     hidden: false,                  // optional
 *     rowCount: 200,                  // จำนวนแถวที่ใส่ validation/สูตรไว้ล่วงหน้า
 *     freezeHeader: true,
 *     freezeColumns: 0,
 *     autoFilter: true,
 *     rows: [{ name: "สมชาย" }],      // ข้อมูลเริ่มต้น (optional)
 *     columns: [{
 *       header: "ชื่อ", key: "name", width: 25,
 *       required: true,               // ใส่ * ที่หัวคอลัมน์ และไม่ให้เว้นว่างใน validation
 *       numFmt: "#,##0.00",
 *       dropdown: { list: "depts" }           // อ้าง lists ด้านบน
 *              // หรือ { values: ["A","B"] }  // ค่าตายตัว
 *              // + { allowBlank, error, prompt }
 *       validation: { type: "whole", operator: "between", formulae: [1, 100] }, // validation อื่นๆ ของ ExcelJS
 *       formula: "{qty}*{price}",     // หรือ function (row, colMap) => string
 *     }],
 *     cells: [{ address: "F1", value: "รวม", formula: "SUM(C2:C500)", bold: true, numFmt: "#,##0" }],
 *   }]
 * }
 */
export function buildWorkbook({
  sheets,
  lists = {},
  listSheetState = "veryHidden", // "hidden" = ผู้ใช้ Unhide ดูได้
  creator = "Excel Template Builder",
  headerStyle = {},
} = {}) {
  if (!Array.isArray(sheets) || sheets.length === 0) {
    throw new Error("ต้องมี sheets อย่างน้อย 1 ชีท");
  }

  const wb = new ExcelJS.Workbook();
  wb.creator = creator;
  wb.created = new Date();

  // ตรวจชื่อชีท
  const seen = new Set();
  sheets.forEach((s) => {
    if (!s.name || s.name.length > 31 || INVALID_SHEET_CHARS.test(s.name)) {
      throw new Error(`ชื่อชีทไม่ถูกต้อง: "${s.name}" (ยาวไม่เกิน 31 ตัว และห้ามมี \\ / ? * [ ] :)`);
    }
    if (s.name === LIST_SHEET) throw new Error(`ชื่อชีท "${LIST_SHEET}" ถูกสงวนไว้`);
    if (seen.has(s.name)) throw new Error(`ชื่อชีทซ้ำ: "${s.name}"`);
    seen.add(s.name);
  });

  // 1) รวบรวมรายการ dropdown ทั้งหมด (รวมค่าตายตัว) ไว้ใน allLists
  const allLists = { ...lists };
  const listKeyOf = (si, col) => {
    const d = col.dropdown;
    if (!d) return null;
    if (d.list) {
      if (!allLists[d.list]) throw new Error(`ไม่พบ list "${d.list}" ใน config.lists`);
      return d.list;
    }
    if (Array.isArray(d.values)) {
      const k = `__s${si}_${col.key}`;
      allLists[k] = d.values;
      return k;
    }
    return null;
  };
  const colListKeys = sheets.map((s, si) => s.columns.map((c) => listKeyOf(si, c)));

  // 2) คำนวณช่วงเซลล์ของแต่ละ list ในชีท Lists (คอลัมน์ละ 1 list, เริ่มแถว 2)
  const listRefs = {};
  Object.keys(allLists).forEach((key, i) => {
    const L = colLetter(i + 1);
    const n = Math.max(allLists[key].length, 1);
    listRefs[key] = `${LIST_SHEET}!$${L}$2:$${L}$${n + 1}`;
  });

  // 3) สร้างแต่ละชีท
  sheets.forEach((cfg, si) => {
    const cols = cfg.columns ?? [];
    const colMap = {};
    cols.forEach((c, i) => (colMap[c.key] = colLetter(i + 1)));

    const ws = wb.addWorksheet(cfg.name, {
      state: cfg.hidden ? "hidden" : "visible",
      properties: cfg.tabColor ? { tabColor: { argb: cfg.tabColor } } : {},
      views:
        cfg.freezeHeader === false
          ? []
          : [{ state: "frozen", ySplit: 1, xSplit: cfg.freezeColumns ?? 0 }],
    });

    ws.columns = cols.map((c) => ({
      header: c.required ? `${c.header} *` : c.header,
      key: c.key,
      width: c.width ?? 18,
    }));

    // สไตล์หัวตาราง
    const hs = {
      font: { bold: true, color: { argb: "FFFFFFFF" } },
      fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F4E78" } },
      alignment: { vertical: "middle", horizontal: "center", wrapText: true },
      ...headerStyle,
    };
    const headerRow = ws.getRow(1);
    headerRow.height = 24;
    headerRow.eachCell((cell) => {
      cell.font = hs.font;
      cell.fill = hs.fill;
      cell.alignment = hs.alignment;
      cell.border = { bottom: { style: "thin" } };
    });

    if (cfg.autoFilter !== false && cols.length) {
      ws.autoFilter = { from: "A1", to: `${colLetter(cols.length)}1` };
    }

    // แถวข้อมูล + validation + สูตร
    const dataRows = cfg.rows ?? [];
    const totalRows = Math.max(cfg.rowCount ?? 200, dataRows.length);

    for (let r = 2; r <= totalRows + 1; r++) {
      cols.forEach((c, ci) => {
        const cell = ws.getCell(`${colMap[c.key]}${r}`);
        if (c.numFmt) cell.numFmt = c.numFmt;

        if (c.formula) {
          cell.value = { formula: buildFormula(c.formula, r, colMap) };
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF2F2F2" } };
          return; // คอลัมน์สูตรไม่ต้องมี validation
        }

        const v = dataRows[r - 2]?.[c.key];
        if (v !== undefined && v !== null) cell.value = v;

        const listKey = colListKeys[si][ci];
        if (listKey) {
          const d = c.dropdown;
          cell.dataValidation = {
            type: "list",
            allowBlank: d.allowBlank ?? !c.required,
            formulae: [listRefs[listKey]],
            showErrorMessage: true,
            errorStyle: "stop",
            errorTitle: "ค่าไม่ถูกต้อง",
            error: d.error ?? "กรุณาเลือกจากรายการที่กำหนด",
            ...(d.prompt && { showInputMessage: true, promptTitle: c.header, prompt: d.prompt }),
          };
        } else if (c.validation) {
          cell.dataValidation = {
            allowBlank: !c.required,
            showErrorMessage: true,
            errorTitle: "ค่าไม่ถูกต้อง",
            error: "ข้อมูลไม่ตรงตามเงื่อนไข",
            ...c.validation,
          };
        }
      });
    }

    // เซลล์พิเศษ เช่น ช่องสรุปผล / สูตรข้ามชีท
    (cfg.cells ?? []).forEach((x) => {
      const cell = ws.getCell(x.address);
      cell.value = x.formula
        ? { formula: String(x.formula).replace(/^=/, "") }
        : x.value;
      if (x.bold) cell.font = { bold: true };
      if (x.numFmt) cell.numFmt = x.numFmt;
    });
  });

  // 4) ชีทเก็บค่า dropdown (สร้างหลังสุดเพื่อให้อยู่ท้ายสุด)
  const listKeys = Object.keys(allLists);
  if (listKeys.length) {
    const ls = wb.addWorksheet(LIST_SHEET, { state: listSheetState });
    listKeys.forEach((key, i) => {
      ls.getCell(1, i + 1).value = key;
      allLists[key].forEach((item, j) => {
        ls.getCell(j + 2, i + 1).value = labelOf(item);
      });
    });
  }

  return wb;
}

// ---------- download (browser) ----------

export async function downloadWorkbook(config, fileName = "template.xlsx") {
  const wb = buildWorkbook(config);
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const { saveAs } = await import("file-saver");
  saveAs(blob, fileName);
}

// ---------- อ่านไฟล์ที่ผู้ใช้กรอกกลับมา ----------

/**
 * @param {ArrayBuffer} arrayBuffer  เช่นจาก await file.arrayBuffer()
 * @param {Array} sheets             config.sheets เดิม (ใช้ name + columns)
 * @returns {Object}  { [ชื่อชีท]: [ {key: value, ...}, ... ] }
 */
export async function readWorkbook(arrayBuffer, sheets) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(arrayBuffer);

  const out = {};
  for (const cfg of sheets) {
    const ws = wb.getWorksheet(cfg.name);
    if (!ws) {
      out[cfg.name] = null; // ผู้ใช้ลบ/เปลี่ยนชื่อชีท
      continue;
    }
    const rows = [];
    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber === 1) return;
      const obj = {};
      let hasData = false;
      cfg.columns.forEach((c, i) => {
        const v = plain(row.getCell(i + 1).value);
        obj[c.key] = v === "" ? null : v;
        // นับว่าแถวมีข้อมูลเฉพาะคอลัมน์ที่ไม่ใช่สูตร
        if (!c.formula && v !== null && v !== "") hasData = true;
      });
      if (hasData) rows.push(obj);
    });
    out[cfg.name] = rows;
  }
  return out;
}
