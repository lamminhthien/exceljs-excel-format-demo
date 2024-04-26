import { IStudent } from "@/interfaces/student";
import { Cell, Column, Workbook } from "exceljs";

import { faker, fakerVI } from "@faker-js/faker";
import saveAs from 'file-saver';

export const exportBeautifulExcel = () => {
  // Initial new Excel Workbook
  const workbook = new Workbook();

  // Create Excel Sheet
  const wsStudentResult = workbook.addWorksheet("Student Result");

  // Initial Excel Column Headers
  const wsColumnHeaders: Partial<Column>[] = [
    {
      key: "name",
      header: "Name",
    },
    {
      key: "class",
      header: "Class",
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

  // Writing data to file
  workbook.xlsx.writeBuffer().then(buffer => {
    saveAs(
      new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      }),
      `student-result-exceljs.xlsx`
    );
  });
};

export const generateStudentResult = () => {
  return {
    name: fakerVI.name,
    class: faker.helpers.arrayElement(["Toán", "Sử", "Địa"]),
    score: faker.number.int({ min: 50, max: 100 }),
  };
};
