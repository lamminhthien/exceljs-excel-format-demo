import { Worksheet, Row, Cell } from "exceljs";

export const centerAlignHeader = (worksheet: Worksheet): void => {
  const headerRow = worksheet.getRow(1);
  headerRow.eachCell((cell: Cell) => {
    cell.alignment = { vertical: "middle", horizontal: "center" };
  });
};

export const centerAlignRowVertically = (row: Row): void => {
  row.eachCell((cell: Cell) => {
    cell.alignment = {
      ...cell.alignment,
      vertical: "middle",
      indent: 1,
    };
  });
};

export const autoColumnWidth = (worksheet: Worksheet) => {
  worksheet.columns.forEach((column) => {
    let maxColumnLength = 0;
    if (column && typeof column.eachCell === "function") {
      column.eachCell({ includeEmpty: true }, (cell) => {
        maxColumnLength = Math.max(maxColumnLength, cell.value ? cell.value.toString().length + 5 : 0);
      });
    }
  });
  return worksheet; // for chaining.
};
