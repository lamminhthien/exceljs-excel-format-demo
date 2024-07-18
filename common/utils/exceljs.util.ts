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
