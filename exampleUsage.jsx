import { useState } from "react";
import { downloadWorkbook, readWorkbook } from "./excelBuilder";

// สร้าง config จากข้อมูลที่ได้จาก API
const makeConfig = (depts, statuses) => ({
  lists: {
    depts, // เช่น ["ขาย", "บัญชี", "IT"]  หรือ [{ label: "ขาย" }, ...]
    statuses,
  },
  sheets: [
    {
      name: "พนักงาน",
      tabColor: "FF1F4E78",
      rowCount: 500,
      columns: [
        { header: "ชื่อ", key: "name", width: 25, required: true },
        { header: "แผนก", key: "dept", width: 18, dropdown: { list: "depts", prompt: "เลือกแผนก" } },
        { header: "สถานะ", key: "status", width: 15, dropdown: { list: "statuses" } },
        // dropdown แบบค่าตายตัว ไม่ต้องประกาศใน lists
        { header: "ประเภทสัญญา", key: "contract", width: 16, dropdown: { values: ["ประจำ", "ชั่วคราว"] } },
        // validation ตัวเลข 0 - 100
        {
          header: "คะแนน",
          key: "score",
          width: 10,
          validation: { type: "decimal", operator: "between", formulae: [0, 100], error: "ใส่ 0 - 100" },
        },
      ],
    },
    {
      name: "ออเดอร์",
      rowCount: 500,
      columns: [
        { header: "สินค้า", key: "item", width: 25 },
        { header: "จำนวน", key: "qty", width: 10, validation: { type: "whole", operator: "greaterThan", formulae: [0] } },
        { header: "ราคา/หน่วย", key: "price", width: 14, numFmt: "#,##0.00" },
        // สูตร: อ้างคอลัมน์ด้วย {key} และ {row} = เลขแถวปัจจุบัน
        {
          header: "รวม",
          key: "total",
          width: 14,
          numFmt: "#,##0.00",
          formula: '=IF(OR({qty}="",{price}=""),"",{qty}*{price})',
        },
      ],
      // เซลล์สรุปผลด้านขวา (สูตรปกติ)
      cells: [
        { address: "G1", value: "ยอดรวมทั้งหมด", bold: true },
        { address: "H1", formula: "SUM(D2:D501)", bold: true, numFmt: "#,##0.00" },
      ],
    },
    {
      name: "สรุป",
      freezeHeader: false,
      columns: [
        { header: "หัวข้อ", key: "k", width: 25 },
        { header: "ค่า", key: "v", width: 18 },
      ],
      rowCount: 0,
      cells: [
        { address: "A2", value: "จำนวนพนักงาน" },
        // สูตรข้ามชีท
        { address: "B2", formula: "COUNTA('พนักงาน'!A2:A501)" },
        { address: "A3", value: "ยอดขายรวม" },
        { address: "B3", formula: "'ออเดอร์'!H1", numFmt: "#,##0.00" },
      ],
    },
  ],
});

export default function ExcelTemplate() {
  const [result, setResult] = useState(null);
  const [config, setConfig] = useState(null);

  const handleDownload = async () => {
    const [depts, statuses] = await Promise.all([
      fetch("/api/departments").then((r) => r.json()), // ["ขาย","บัญชี"]
      fetch("/api/statuses").then((r) => r.json()),
    ]);
    const cfg = makeConfig(depts, statuses);
    setConfig(cfg);
    await downloadWorkbook(cfg, "template.xlsx");
  };

  const handleUpload = async (e) => {
    const file = e.target.files[0];
    if (!file || !config) return;
    const data = await readWorkbook(await file.arrayBuffer(), config.sheets);
    setResult(data); // { "พนักงาน": [...], "ออเดอร์": [...], ... }
  };

  return (
    <div>
      <button onClick={handleDownload}>ดาวน์โหลด Template</button>
      <input type="file" accept=".xlsx" onChange={handleUpload} />
      <pre>{result && JSON.stringify(result, null, 2)}</pre>
    </div>
  );
}
