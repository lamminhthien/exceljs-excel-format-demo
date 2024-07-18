"use client";
import { exportBeautifulExcel } from "@/common/utils/export-beautiful-excel";
import createStudentListExcel from "@/common/utils/export-report-portal";

export default function Home() {
  // Sample data student
  const students = [
    {
      name: "John Doe",
      email: "johndoe@example.com",
      phone: "1234567890",
      course: "Basic English",
      startDate: "2024-08-01",
      classTime: "18:00-20:00",
    },
    {
      name: "Jane Smith",
      email: "janesmith@example.com",
      phone: "0987654321",
      course: "Python Programming",
      startDate: "2024-08-15",
      classTime: "19:30-21:30",
    },
    {
      name: "Bob Johnson",
      email: "bobjohnson@example.com",
      phone: "5555555555",
      course: "Online Marketing",
      startDate: "2024-09-01",
      classTime: "08:00-10:00",
    },
  ];

  return (
    <main className="flex min-h-screen flex-col items-center justify-between p-24">
      <button className="p-3 text-white bg-blue-700 rounded-lg" onClick={() => exportBeautifulExcel()}>
        Download Excel Formating Demo
      </button>
      <button className="p-3 text-white bg-blue-700 rounded-lg" onClick={() => createStudentListExcel(students)}>
        Download Excel Formating Report Demo
      </button>
    </main>
  );
}
