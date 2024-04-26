import { Column, Workbook } from "exceljs";

import { faker, fakerVI } from "@faker-js/faker";
import saveAs from "file-saver";

export const exportBeautifulExcel = () => {
  // Initial new Excel Workbook
  const workbook = new Workbook();

  // Create Excel Sheet
  const wsStudentResult = workbook.addWorksheet("Student Result");

  // Initial Excel Column Headers. And can also apply style for each column.
  // Pay attention to key in your interface data
  const wsColumnHeaders: Partial<Column>[] = [
    {
      key: "name",
      header: "Name",
      width: 40,
      style: {
        fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FF0070C0" } },
        font: { bold: true, size: 14, color: { argb: "FFFFFFFF" } },
      },
    },
    {
      key: "class",
      header: "Class",
      style: {
        font: { bold: true, size: 14 },
      },
    },
    {
      key: "score",
      header: "Score",
    },
  ];

  // Add Column Header to Excel Sheet
  wsStudentResult.columns = wsColumnHeaders;

  // Initial Sample data
  const students = Array.from({ length: 100 }, generateStudentResult);

  // Add data to excel
  wsStudentResult.addRows(students);

  // Add conditional formatting data bar to column C (ref: C:C),
  // Please use patch-package and post-install for fix issue: The author of this module when they upgrade library, they forget to color attribute in Typescript🥲
  wsStudentResult.addConditionalFormatting({
    ref: `C:C`,
    rules: [
      {
        type: "dataBar",
        minLength: 0,
        maxLength: 100,
        gradient: true,
        border: false,
        priority: 1,
        color: { argb: "90EE90" },
        // cfvo mean the min value and max value for fill gradient color properly
        cfvo: [
          { type: "num", value: 0 },
          { type: "num", value: 100 },
        ],
      },
    ],
  });

  // Write data and make browser download file by file-saver
  workbook.xlsx.writeBuffer().then((buffer) => {
    saveAs(
      new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }),
      `student-result-exceljs.xlsx`
    );
  });
};

// Fake data by library fakerjs
// Pay attention to key in your interface data
export const generateStudentResult = () => {
  return {
    name: fakerVI.name.fullName(),
    class: faker.helpers.arrayElement(["Toán", "Sử", "Địa"]),
    score: faker.number.int({ min: 50, max: 100 }),
  };
};
