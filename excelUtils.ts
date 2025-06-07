import ExcelJS from "exceljs";
import fs from "fs";
import path from "path";

// ファイル名に使えない文字を置換（Windows 互換）
export const sanitize = (s: string) => s.replace(/[\/:*?"<>|]/g, "_");

export function toDisplayValue(
  cellValue: ExcelJS.CellValue
): string | number | boolean | null {
  if (cellValue === null || cellValue === undefined) return null;
  if (
    typeof cellValue === "string" ||
    typeof cellValue === "number" ||
    typeof cellValue === "boolean"
  )
    return cellValue;
  if (cellValue instanceof Date) return cellValue.toISOString();
  if (isFormula(cellValue)) {
    return toDisplayValue(cellValue.result);
  }
  if (isHyperlink(cellValue)) {
    return cellValue.text ?? cellValue.hyperlink;
  }
  if (isRichText(cellValue)) {
    return cellValue.richText.map((rt) => rt.text).join("");
  }
  if (isError(cellValue)) {
    return cellValue.error;
  }
  return null;
}

function isFormula(v: ExcelJS.CellValue): v is ExcelJS.CellFormulaValue {
  return typeof v === "object" && !!(v as ExcelJS.CellFormulaValue).formula;
}
function isHyperlink(v: ExcelJS.CellValue): v is ExcelJS.CellHyperlinkValue {
  return typeof v === "object" && !!(v as ExcelJS.CellHyperlinkValue).hyperlink;
}
function isRichText(v: ExcelJS.CellValue): v is ExcelJS.CellRichTextValue {
  return typeof v === "object" && !!(v as ExcelJS.CellRichTextValue).richText;
}
function isError(v: ExcelJS.CellValue): v is ExcelJS.CellErrorValue {
  return typeof v === "object" && !!(v as ExcelJS.CellErrorValue).error;
}

export async function convertBook(
  xlsxPath: string,
  dst: string
): Promise<void> {
  const bookName = sanitize(path.basename(xlsxPath, ".xlsx"));
  const reader = new ExcelJS.stream.xlsx.WorkbookReader(xlsxPath, {
    worksheets: "emit",
    entries: "emit",
    sharedStrings: "cache",
    hyperlinks: "ignore",
  });

  for await (const worksheet of reader) {
    const sheetName = sanitize((worksheet as any).name);
    const csvPath = path.join(dst, bookName, `${sheetName}.csv`);
    fs.mkdirSync(path.dirname(csvPath), { recursive: true });
    const out = fs.createWriteStream(csvPath, { encoding: "utf8" });

    for await (const row of worksheet) {
      const newRow = Array.isArray(row.values) ? row.values : [];
      const line = newRow
        .map((cell) => {
          const value = toDisplayValue(cell);
          if (value === null) return "";
          if (typeof value === "string") {
            return `"${value.replace(/"/g, '""')}"`;
          }
          return String(value);
        })
        .join(",");
      out.write(line + "\n");
    }

    out.end();
    console.log(`✓ ${xlsxPath} → ${csvPath}`);
  }
}
