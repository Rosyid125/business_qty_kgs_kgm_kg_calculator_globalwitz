const fs = require("fs");
const XLSX = require("xlsx");
const readlineSync = require("readline-sync");
const path = require("path");

// Folder input dan output
const inputFolder = "./input";
const outputFolder = "./output";

// Fungsi untuk membaca file Excel dari folder input
function getInputFiles() {
  if (!fs.existsSync(inputFolder)) {
    console.error(`Folder input "${inputFolder}" tidak ditemukan!`);
    process.exit(1);
  }

  const files = fs.readdirSync(inputFolder).filter(file => 
    file.toLowerCase().endsWith('.xlsx') || file.toLowerCase().endsWith('.xls')
  );

  if (files.length === 0) {
    console.error(`Tidak ada file Excel (.xlsx/.xls) ditemukan di folder "${inputFolder}"`);
    process.exit(1);
  }

  return files;
}

// Fungsi untuk membaca semua kolom dari sheet dengan lebih efisien
function analyzeSheetColumns(worksheet) {
  console.log("Membaca kolom dari sheet...");
  
  // Ambil range sheet untuk mendapatkan header saja
  const range = XLSX.utils.decode_range(worksheet['!ref']);
  const headerRow = [];
  
  // Baca hanya baris pertama (header)
  for (let col = range.s.c; col <= range.e.c; col++) {
    const cellAddress = XLSX.utils.encode_cell({ r: range.s.r, c: col });
    const cell = worksheet[cellAddress];
    headerRow.push(cell ? (cell.v || `Column_${col + 1}`) : `Column_${col + 1}`);
  }

  if (headerRow.length === 0) {
    console.log("Sheet kosong atau tidak memiliki data.");
    return [];
  }

  console.log("\nKolom yang tersedia di sheet ini:");
  headerRow.forEach((col, index) => {
    console.log(`${index + 1}. ${col}`);
  });

  return headerRow;
}

// Fungsi untuk memilih kolom yang diperlukan
function selectRequiredColumns(availableColumns) {
  const requiredColumns = {
    unitOfWeight: null,
    businessQuantity: null,
    unitPrice: null,
    width: null,
    gsm: null
  };

  console.log("\n=== PEMILIHAN KOLOM UNTUK KONVERSI BUSINESS QUANTITY KE KG ===");
  console.log("Silakan pilih kolom yang sesuai untuk setiap parameter:");

  // Unit of Weight
  console.log("\n1. Kolom UNIT OF WEIGHT:");
  const unitWeightIndex = readlineSync.questionInt("Pilih nomor kolom untuk Unit of Weight: ") - 1;
  if (unitWeightIndex >= 0 && unitWeightIndex < availableColumns.length) {
    requiredColumns.unitOfWeight = availableColumns[unitWeightIndex];
  }

  // Business Quantity
  console.log("\n2. Kolom BUSINESS QUANTITY:");
  const businessQtyIndex = readlineSync.questionInt("Pilih nomor kolom untuk Business Quantity: ") - 1;
  if (businessQtyIndex >= 0 && businessQtyIndex < availableColumns.length) {
    requiredColumns.businessQuantity = availableColumns[businessQtyIndex];
  }

  // Unit Price
  console.log("\n3. Kolom UNIT PRICE (USD):");
  const unitPriceIndex = readlineSync.questionInt("Pilih nomor kolom untuk Unit Price (USD): ") - 1;
  if (unitPriceIndex >= 0 && unitPriceIndex < availableColumns.length) {
    requiredColumns.unitPrice = availableColumns[unitPriceIndex];
  }

  // Width
  console.log("\n4. Kolom WIDTH:");
  const widthIndex = readlineSync.questionInt("Pilih nomor kolom untuk Width: ") - 1;
  if (widthIndex >= 0 && widthIndex < availableColumns.length) {
    requiredColumns.width = availableColumns[widthIndex];
  }

  // GSM
  console.log("\n5. Kolom GSM:");
  const gsmIndex = readlineSync.questionInt("Pilih nomor kolom untuk GSM: ") - 1;
  if (gsmIndex >= 0 && gsmIndex < availableColumns.length) {
    requiredColumns.gsm = availableColumns[gsmIndex];
  }

  return requiredColumns;
}

// Fungsi untuk konversi Business Quantity ke KG
function convertBusinessQuantityToKG(data, columns) {
  console.log("\nMemulai konversi Business Quantity ke KG...");
  
  data.forEach((row, index) => {
    const unitOfWeight = columns.unitOfWeight ? (row[columns.unitOfWeight] || "-") : "-";
    const businessQuantity = columns.businessQuantity ? (parseFloat(row[columns.businessQuantity]) || 0) : 0;
    const unitPrice = columns.unitPrice ? (parseFloat(row[columns.unitPrice]) || 0) : 0;
    const width = columns.width ? (parseFloat(row[columns.width]) || 0) : 0;
    const gsm = columns.gsm ? (parseFloat(row[columns.gsm]) || 0) : 0;
    
    let result = "-";

    // Handle GRM/GR first - only needs businessQuantity (Fixed: should multiply by 1000, not divide)
    if ((unitOfWeight.toUpperCase() === "GRM" || unitOfWeight.toUpperCase() === "GR") && businessQuantity > 0) {
      result = businessQuantity * 1000; // GRM to KG = Business Quantity * 1000
    }
    // Handle KG/KGM/KGS - only needs businessQuantity  
    else if ((unitOfWeight.toUpperCase() === "KG" || unitOfWeight.toUpperCase() === "KGM" || unitOfWeight.toUpperCase() === "KGS") && businessQuantity > 0) {
      result = businessQuantity;
    }
    // Handle other units that need all parameters
    else if (businessQuantity > 0 && unitPrice > 0 && width > 0 && gsm > 0) {
      switch (unitOfWeight.toUpperCase()) {
        case "MTR":
          result = (unitPrice * 1000) / (width * gsm);
          break;
        case "MTK":
        case "MTR2":
          result = (unitPrice * 1000) / gsm;
          break;
        case "YD":
          result = ((unitPrice / 0.9144) * 1000) / (width * gsm);
          break;
        case "ROL":
        case "ROLL":
          result = businessQuantity / gsm;
          break;
      }
    }
    
    row["BUSINESS QUANTITY (KG)"] = result;
    
    // Show first 5 conversions as example
    if (index < 5) {
      console.log(`  Baris ${index + 2}: ${unitOfWeight} -> ${result} KG`);
    }
  });
  
  console.log(`Dst. untuk ${data.length} baris data.`);
  return data;
}

// Program utama
console.log("=== KONVERTER BUSINESS QUANTITY KE KG ===");

// 1. Baca file dari folder input
const inputFiles = getInputFiles();
console.log("\nFile Excel yang tersedia di folder input:");
inputFiles.forEach((file, index) => {
  console.log(`${index + 1}. ${file}`);
});

const selectedFileIndex = readlineSync.questionInt("\nPilih nomor file yang ingin diproses: ") - 1;
if (selectedFileIndex < 0 || selectedFileIndex >= inputFiles.length) {
  console.error("Pilihan file tidak valid!");
  process.exit(1);
}

const selectedFile = inputFiles[selectedFileIndex];
const inputFilePath = path.join(inputFolder, selectedFile);

// 2. Baca file Excel
let workbook;
try {
  console.log(`\nMembaca file: ${selectedFile}`);
  workbook = XLSX.readFile(inputFilePath);
} catch (error) {
  console.error(`Error membaca file "${selectedFile}": ${error.message}`);
  process.exit(1);
}

// 3. Pilih sheet yang akan diproses
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

// 4. Proses setiap sheet yang dipilih
selectedSheetIndexes.forEach((sheetIndex) => {
  const sheetName = sheetNames[sheetIndex];
  const worksheet = workbook.Sheets[sheetName];

  console.log(`\n=== MEMPROSES SHEET: "${sheetName}" ===`);
  
  // Analisis kolom yang tersedia
  const availableColumns = analyzeSheetColumns(worksheet);
  if (availableColumns.length === 0) {
    console.log(`Sheet "${sheetName}" kosong, dilewati.`);
    return;
  }

  // Pilih kolom yang diperlukan
  const selectedColumns = selectRequiredColumns(availableColumns);
  
  // Validasi minimal kolom yang diperlukan
  if (!selectedColumns.unitOfWeight || !selectedColumns.businessQuantity) {
    console.error("Kolom Unit of Weight dan Business Quantity wajib dipilih!");
    return;
  }

  // Baca data dan konversi
  const data = XLSX.utils.sheet_to_json(worksheet, { defval: "-" });
  const convertedData = convertBusinessQuantityToKG(data, selectedColumns);

  // Update sheet dengan data yang sudah dikonversi
  const newWorksheet = XLSX.utils.json_to_sheet(convertedData);
  workbook.Sheets[sheetName] = newWorksheet;
});

// 5. Simpan file hasil ke folder output
const outputFileName = `converted_${selectedFile}`;
const outputFilePath = path.join(outputFolder, outputFileName);

// Pastikan folder output ada
if (!fs.existsSync(outputFolder)) {
  fs.mkdirSync(outputFolder, { recursive: true });
}

XLSX.writeFile(workbook, outputFilePath);
console.log(`\n✅ File berhasil disimpan ke: "${outputFilePath}"`);
console.log("\n=== PROSES SELESAI ===");
