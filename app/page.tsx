"use client";
import { exportBeautifulExcel } from "@/common/utils/export-beautiful-excel";
import createStudentListExcel from "@/common/utils/export-report-portal";

export default function Home() {
  return (
    <main className="flex min-h-screen flex-col items-center justify-between p-24">
      <button className="p-3 text-white bg-blue-700 rounded-lg" onClick={() => exportBeautifulExcel()}>
        Download Excel Formating Demo
      </button>
      <button className="p-3 text-white bg-blue-700 rounded-lg" onClick={() => createStudentListExcel()}>
        Download Excel Formating Report Demo
      </button>
    </main>
  );
}
