const fs = require("fs");
const XLSX = require("xlsx");

// File input dan output
const inputFile = "data.xlsx";
const outputFile = "output.xlsx";

// Baca file Excel input
const workbook = XLSX.readFile(inputFile);
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// Cari kolom 'DATE'
const dateColumn = Object.keys(worksheet).filter((cellRef) => {
  const cell = worksheet[cellRef];
  return /^[A-Z]+[0-9]+$/.test(cellRef) && cell.v === "DATE";
})[0];

if (!dateColumn) {
  console.error('Kolom "DATE" tidak ditemukan!');
  process.exit(1);
}

// Dapatkan huruf kolom dari header 'DATE'
const dateColLetter = dateColumn.match(/[A-Z]+/)[0];

// Loop untuk konversi tanggal di seluruh kolom 'DATE'
Object.keys(worksheet).forEach((cellRef) => {
  if (cellRef.startsWith(dateColLetter) && /^[A-Z]+[0-9]+$/.test(cellRef)) {
    const cell = worksheet[cellRef];
    const rowNum = parseInt(cellRef.replace(/[^0-9]/g, ""));
    if (rowNum > 1 && cell.v) {
      const value = cell.v.toString();
      if (value.length === 6) {
        const year = value.slice(0, 4);
        const month = value.slice(4, 6);
        const formattedDate = `01/${month}/${year}`;
        worksheet[cellRef].v = formattedDate;
      }
    }
  }
});

// Simpan file hasil konversi
XLSX.writeFile(workbook, outputFile);
console.log(`Konversi selesai! File disimpan sebagai ${outputFile}`);
