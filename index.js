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
  console.error("Tidak ada sheet yang dipilih!");
  process.exit(1);
}

// Menu pilihan
console.log("\nPilih opsi:");
console.log("1. Format Tanggal (YYYYMM -> 01/MM/YYYY)");
console.log("2. Format Satuan Angka (misal: 22,345 -> 22.345)");
console.log("3. Konversi BUSINESS QUANTITY ke KG");

const choice = readlineSync.question("\nMasukkan pilihan (1, 2, atau 3): ");

selectedSheetIndexes.forEach((sheetIndex) => {
  const sheetName = sheetNames[sheetIndex];
  const worksheet = workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(worksheet, { defval: "-" });

  console.log(`\nMemproses sheet "${sheetName}"...`);

  if (choice === "1") {
    formatDateColumn(worksheet);
  } else if (choice === "2") {
    formatNumberColumns(worksheet);
  } else if (choice === "3") {
    data.forEach((row) => {
      const unitOfWeight = row["UNIT OF WEIGHT"] || "-";
      const businessQuantity = parseFloat(row["BUSINESS QUANTITY"]) || "-";
      const unitPrice = parseFloat(row["UNIT PRICE(USD)"]) || "-";
      const width = parseFloat(row["Width (cm)"]) || "-";
      const gsm = parseFloat(row["GSM"]) || "-";
      let result = "-";

      if (businessQuantity !== "-" && unitPrice !== "-" && width !== "-" && gsm !== "-") {
        switch (unitOfWeight.toUpperCase()) {
          case "MTR":
            result = (unitPrice * 1000) / (width * gsm);
            break;
          case "MTR2":
            result = (unitPrice * 1000) / gsm;
            break;
          case "YD":
            result = (unitPrice * 1000) / (width * gsm);
            break;
          case "GRM":
            result = businessQuantity / 1000;
            break;
          case "KG":
            result = businessQuantity;
            break;
          case "ROL":
          case "ROLL":
            result = businessQuantity / gsm;
            break;
          default:
            result = "-";
        }
      }
      row["BUSINESS QUANTITY (KG)"] = result;
    });

    // Update sheet
    const newWorksheet = XLSX.utils.json_to_sheet(data);
    workbook.Sheets[sheetName] = newWorksheet;
  } else {
    console.error("Pilihan tidak valid!");
    process.exit(1);
  }
});

// Simpan file hasil
XLSX.writeFile(workbook, outputFile);
console.log(`\nFile berhasil disimpan ke "${outputFile}"\n`);
