import ExcelJS from "exceljs";
import saveAs from "file-saver";
import { autoColumnWidth, centerAlignHeader, centerAlignRowVertically } from "./exceljs.util";

interface Student {
  name: string;
  email: string;
  phone: string;
  course: string;
  startDate: string;
  classTime: string;
}

export const createStudentListExcel = (data: Student[]) => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Student List");

  // Define columns
  worksheet.columns = [
    { header: "Name", key: "name", width: 20 },
    { header: "Email", key: "email", width: 30 },
    { header: "Phone", key: "phone", width: 15 },
    { header: "Course", key: "course", width: 25 },
    { header: "Start Date", key: "startDate", width: 15 },
    { header: "Class Time", key: "classTime", width: 15 },
  ];

  // Add data to the worksheet
  worksheet.addRows(data);

  // Set custom row height and font sizes
  const ROW_HEIGHT = 30; // Increased from 25
  const HEADER_FONT_SIZE = 14;
  const DATA_FONT_SIZE = 12;

  worksheet.eachRow((row, rowNumber) => {
    row.height = ROW_HEIGHT;
    if (rowNumber === 1) {
      // Header row
      row.font = { size: HEADER_FONT_SIZE, bold: true, color: { argb: "FFFFFFFF" } };
    } else {
      // Data rows
      row.font = { size: DATA_FONT_SIZE, color: { argb: "FF000000" } };
    }
  });

  // Format header
  const headerRow = worksheet.getRow(1);
  headerRow.height = ROW_HEIGHT + 5; // Make header slightly taller
  headerRow.eachCell((cell: ExcelJS.Cell) => {
    if (cell.value) {
      cell.alignment = { vertical: "middle", horizontal: "center" };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF000080" }, // Dark blue
      };
    }
  });

  // Format data cells
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) {
      const nameCell = row.getCell("name");
      if (nameCell) {
        nameCell.font = { size: DATA_FONT_SIZE, color: { argb: "FF0000FF" } };
      }
      centerAlignRowVertically(row);
    }

    row.eachCell((cell) => {
      cell.border = {
        top: { style: "thin", color: { argb: "FF808080" } },
        left: { style: "thin", color: { argb: "FF808080" } },
        bottom: { style: "thin", color: { argb: "FF808080" } },
        right: { style: "thin", color: { argb: "FF808080" } },
      };
    });
  });

  // Add hyperlinks for email and phone
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) {
      const emailCell = row.getCell("email");
      const phoneCell = row.getCell("phone");

      if (emailCell && emailCell.value) {
        emailCell.value = {
          text: emailCell.value.toString(),
          hyperlink: `mailto:${emailCell.value}`,
        };
      }

      if (phoneCell && phoneCell.value) {
        phoneCell.value = {
          formula: `""&"${phoneCell.value}"`,
          hyperlink: `tel:${phoneCell.value}`,
        };
      }
    }
  });

  centerAlignHeader(worksheet);

  // Auto column width
  autoColumnWidth(worksheet);

  // Remove grid lines
  worksheet.properties.showGridLines = false;

  // Save the Excel file
  try {
    workbook.xlsx.writeBuffer().then((buffer) => {
      saveAs(
        new Blob([buffer], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        }),
        `student-result-exceljs.xlsx`
      );
    });
  } catch (error) {
    console.error("An error occurred:", error);
    throw error;
  }
};

export default createStudentListExcel;
