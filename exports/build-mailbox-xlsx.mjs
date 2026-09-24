import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const csvPath = process.argv[2];
const xlsxPath = process.argv[3];
const previewPath = process.argv[4];

if (!csvPath || !xlsxPath || !previewPath) {
  throw new Error("Usage: node build-mailbox-xlsx.mjs input.csv output.xlsx preview.png");
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let field = "";
  let quoted = false;

  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    const next = text[i + 1];

    if (quoted) {
      if (ch === '"' && next === '"') {
        field += '"';
        i++;
      } else if (ch === '"') {
        quoted = false;
      } else {
        field += ch;
      }
      continue;
    }

    if (ch === '"') {
      quoted = true;
    } else if (ch === ",") {
      row.push(field);
      field = "";
    } else if (ch === "\n") {
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else if (ch !== "\r") {
      field += ch;
    }
  }

  if (field.length || row.length) {
    row.push(field);
    rows.push(row);
  }

  return rows.filter((r) => r.some((c) => c !== ""));
}

function toNumber(value) {
  if (value === undefined || value === null || value === "") return null;
  const parsed = Number(String(value).replace(/,/g, ""));
  return Number.isFinite(parsed) ? parsed : null;
}

function toDate(value) {
  if (!value) return null;
  const normalized = String(value).replace(/\u202f/g, " ").replace(/\u00a0/g, " ");
  const parsed = new Date(normalized);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function colName(index) {
  let n = index + 1;
  let name = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

const csvText = await fs.readFile(csvPath, "utf8");
const parsed = parseCsv(csvText);
const headers = parsed[0];
const sourceRows = parsed.slice(1);
const numericHeaders = new Set(["ItemCount", "DeletedItemCount", "SizeBytes", "SizeGB"]);
const dateHeaders = new Set(["WhenMailboxCreated", "LastLogonTime"]);

const dataRows = sourceRows.map((row) =>
  headers.map((header, index) => {
    const value = row[index] ?? "";
    if (numericHeaders.has(header)) return toNumber(value);
    if (dateHeaders.has(header)) return toDate(value);
    return value === "" ? null : value;
  }),
);

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Mailbox Sizes");
sheet.showGridLines = false;

const allRows = [headers, ...dataRows];
sheet.getRange("A1").write(allRows);

const lastColumn = colName(headers.length - 1);
const lastRow = allRows.length;
const table = sheet.tables.add(`A1:${lastColumn}${lastRow}`, true, "MailboxSizes");
table.style = "TableStyleMedium2";
table.showFilterButton = true;

sheet.freezePanes.freezeRows(1);

sheet.getRange(`A1:${lastColumn}1`).format = {
  fill: "#1F4E78",
  font: { name: "Arial", bold: true, color: "#FFFFFF", size: 10 },
};
sheet.getRange(`A2:${lastColumn}${lastRow}`).format.font = { name: "Arial", size: 10 };
sheet.getRange(`A1:${lastColumn}${lastRow}`).format.verticalAlignment = "center";

const headerIndex = Object.fromEntries(headers.map((header, index) => [header, index]));
for (const header of ["ItemCount", "DeletedItemCount", "SizeBytes"]) {
  const c = colName(headerIndex[header]);
  sheet.getRange(`${c}2:${c}${lastRow}`).format.numberFormat = "#,##0";
}
sheet.getRange(`${colName(headerIndex.SizeGB)}2:${colName(headerIndex.SizeGB)}${lastRow}`).format.numberFormat = "#,##0.000";
for (const header of ["WhenMailboxCreated", "LastLogonTime"]) {
  const c = colName(headerIndex[header]);
  sheet.getRange(`${c}2:${c}${lastRow}`).format.numberFormat = "mm/dd/yy h:mm AM/PM";
}

sheet.getRange(`A1:${lastColumn}${lastRow}`).format.autofitColumns();
sheet.getRange("A:A").format.columnWidth = 28;
sheet.getRange("B:C").format.columnWidth = 34;
sheet.getRange("I:I").format.columnWidth = 30;
sheet.getRange("M:M").format.columnWidth = 24;

workbook.recalculate();

const inspect = await workbook.inspect({
  kind: "table",
  range: `A1:${lastColumn}6`,
  include: "values,formulas",
  tableMaxRows: 6,
  tableMaxCols: headers.length,
});
console.log(inspect.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
});
console.log(errors.ndjson);

const preview = await workbook.render({
  sheetName: "Mailbox Sizes",
  range: "A1:M25",
  scale: 1,
  format: "png",
});
await fs.mkdir(path.dirname(previewPath), { recursive: true });
await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(xlsxPath);

console.log(JSON.stringify({ xlsxPath, previewPath, rows: dataRows.length }));
