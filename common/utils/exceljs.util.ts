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
      // For column A: this is user_id column, we must set width different than other column
      column.letter === "A" ? (column.width = maxColumnLength + 2.1) : (column.width = (maxColumnLength + 2) * 1);
    }
  });
  return worksheet; // for chaining.
};
