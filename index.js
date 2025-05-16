const fs = require("fs");
const XLSX = require("xlsx");
const readlineSync = require("readline-sync");

// File input dan output
const inputFile = "input.xlsx";
const outputFile = "output.xlsx";

// Baca file Excel input
let workbook;
try {
  workbook = XLSX.readFile(inputFile);
} catch (error) {
  console.error(`Error membaca file input "${inputFile}": ${error.message}`);
  process.exit(1);
}

// Pilih sheet yang akan diproses
const sheetNames = workbook.SheetNames;
console.log("\nSheet yang tersedia:");
sheetNames.forEach((name, index) => {
  console.log(`${index + 1}. ${name}`);
});
const selectedSheetIndexes = readlineSync
  .question("\nMasukkan nomor sheet yang ingin diproses (pisahkan dengan koma jika lebih dari satu, misal: 1,3,5): ")
  .split(",")
  .map((idx) => parseInt(idx.trim()) - 1)
  .filter((idx) => idx >= 0 && idx < sheetNames.length);

if (selectedSheetIndexes.length === 0) {
  console.error("Tidak ada sheet yang dipilih. Program berhenti.");
  process.exit(1);
}

// Fungsi untuk mencari kolom berdasarkan nama header
function findColumnLetter(worksheet, headerName) {
  const headerCells = Object.keys(worksheet).filter((cellRef) => {
    const cell = worksheet[cellRef];
    return /^[A-Z]+1$/.test(cellRef) && cell.v && cell.v.toString().trim().toLowerCase() === headerName.toLowerCase();
  });
  if (headerCells.length > 0) {
    return headerCells[0].match(/[A-Z]+/)[0];
  }
  return null;
}

// Fungsi untuk format tanggal (YYYYMM -> 01/MM/YYYY)
function formatDateColumn(worksheet) {
  const dateColumnHeader = "months";
  const dateColLetter = findColumnLetter(worksheet, dateColumnHeader);

  if (!dateColLetter) {
    console.warn(`Kolom "${dateColumnHeader}" tidak ditemukan! Format tanggal dilewati.`);
    return;
  }

  console.log(`Kolom "${dateColumnHeader}" ditemukan di kolom ${dateColLetter}. Memproses tanggal...`);

  Object.keys(worksheet).forEach((cellRef) => {
    if (cellRef.startsWith(dateColLetter) && /^[A-Z]+[0-9]+$/.test(cellRef)) {
      const cell = worksheet[cellRef];
      const rowNum = parseInt(cellRef.replace(/[^0-9]/g, ""));

      if (rowNum > 1 && cell.v) {
        const value = cell.v.toString();
        if (value.length === 6 && /^\d+$/.test(value)) {
          const year = value.slice(0, 4);
          const month = value.slice(4, 6);
          const formattedDate = `01/${month}/${year}`;
          worksheet[cellRef].v = formattedDate;
          worksheet[cellRef].t = "s";
          console.log(`  Baris ${rowNum}: "${value}" -> "${formattedDate}"`);
        } else if (value.length !== 6 && /^\d+$/.test(value)) {
          console.warn(`  Baris ${rowNum}: Nilai "${value}" di kolom tanggal tidak memiliki format YYYYMM. Dilewati.`);
        }
      }
    }
  });
  console.log("Format tanggal selesai.");
}

// Fungsi untuk format angka
function formatNumberColumns(worksheet) {
  const targetHeaders = ["value usd", "CIF Total In USD", "qty", "Net KG Wt", "USD Qty Unit", "CIF KG Unit In USD"];
  const columnLettersToFormat = [];

  targetHeaders.forEach((header) => {
    const colLetter = findColumnLetter(worksheet, header);
    if (colLetter) {
      columnLettersToFormat.push(colLetter);
      console.log(`Kolom "${header}" ditemukan di kolom ${colLetter}. Akan diformat.`);
    } else {
      console.warn(`Kolom "${header}" tidak ditemukan! Tidak akan diformat.`);
    }
  });

  if (columnLettersToFormat.length === 0) {
    console.log("Tidak ada kolom target untuk format angka yang ditemukan.");
    return;
  }

  console.log("Memproses format angka...");
  Object.keys(worksheet).forEach((cellRef) => {
    const matchResult = cellRef.match(/[A-Z]+/);
    if (!matchResult) {
      console.warn(`  Peringatan: Sel "${cellRef}" tidak memiliki referensi kolom yang valid. Dilewati.`);
      return;
    }

    const currentCellColLetter = matchResult[0];
    const cell = worksheet[cellRef]; // Tambahkan ini

    // Cek apakah kolom saat ini adalah salah satu yang ingin diformat
    if (columnLettersToFormat.includes(currentCellColLetter) && /^[A-Z]+[0-9]+$/.test(cellRef)) {
      const rowNum = parseInt(cellRef.replace(/[^0-9]/g, ""));

      if (rowNum > 1 && cell.v !== undefined && cell.v !== null) {
        let valueStr = cell.v.toString().trim();

        if (valueStr.includes(",")) {
          let originalValue = valueStr;
          if (!valueStr.includes(".")) {
            valueStr = valueStr.replace(/,(?=[^,]*$)/, ".");
          }
          valueStr = valueStr.replace(/,/g, "");

          const numericValue = parseFloat(valueStr);
          if (!isNaN(numericValue)) {
            worksheet[cellRef].v = numericValue;
            worksheet[cellRef].t = "n";
            console.log(`  Baris ${rowNum}, Kolom ${currentCellColLetter}: "${originalValue}" -> ${numericValue}`);
          }
        } else if (typeof cell.v === "string" && !isNaN(parseFloat(cell.v)) && cell.t !== "n") {
          const numericValue = parseFloat(cell.v);
          worksheet[cellRef].v = numericValue;
          worksheet[cellRef].t = "n";
          console.log(`  Baris ${rowNum}, Kolom ${currentCellColLetter}: String "${cell.v}" dikonversi ke angka ${numericValue}`);
        }
      }
    }
  });
}

// Menu pilihan
console.log("\nPilih opsi:");
console.log("1. Format Tanggal (YYYYMM -> 01/MM/YYYY)");
console.log("2. Format Satuan Angka (misal: 22,345 -> 22.345)");

const choice = readlineSync.question("\nMasukkan pilihan (1 atau 2): ");

selectedSheetIndexes.forEach((index) => {
  const sheetName = sheetNames[index];
  const worksheet = workbook.Sheets[sheetName];

  console.log(`\nMemproses sheet "${sheetName}"...`);

  if (choice === "1") {
    formatDateColumn(worksheet);
  } else if (choice === "2") {
    formatNumberColumns(worksheet);
  } else {
    console.error("Pilihan tidak valid!");
    process.exit(1);
  }
});

// Simpan file hasil konversi
try {
  XLSX.writeFile(workbook, outputFile);
  console.log(`\nKonversi selesai! File disimpan sebagai ${outputFile}`);
} catch (error) {
  console.error(`Error menyimpan file output "${outputFile}": ${error.message}`);
  process.exit(1);
}
