const ExcelJS = require("exceljs");
const QRCode = require("qrcode");
const fs = require("fs");

const options = {
    width: 300,
    height: 300,
    errorCorrectionLevel: 'H',
    type: 'png',
    quality: 1,
    margin: 1,
    color: {
        dark: '#000000',
        light: '#ffffff'
    }
};

async function generateQRFromExcel(filePath) {
  // Initialize Excel workbook and read data
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  // Select the first worksheet by its name
  const worksheet = workbook.getWorksheet("Sheet1"); // Change 'Sheet1' to your worksheet name if different

  if (!worksheet) {
    throw new Error("Worksheet not found");
  }

  // Extract key and value data from Excel
  const keys = worksheet.getRow(1).values.slice(1); // First row as keys
  const values = [];

  for (let i = 2; i <= worksheet.rowCount; i++) {
    const rowValues = worksheet.getRow(i).values.slice(1); // Get row values starting from second row
    const rowObject = {};

    keys.forEach((key, index) => {
      rowObject[key] = rowValues[index]; // Create key-value pair
    });

    values.push(rowObject);
  }

  // Generate QR codes for each key-value pair
  for (let i = 0; i < values.length; i++) {
    const qrData = JSON.stringify(values[i]); // Convert object to JSON string

    try {
      const qrCodePath = `qrcodes/qr_code_row_${i + 2}.png`; // Path to save QR code image (starting from second row)
      await QRCode.toFile(qrCodePath, qrData,options); // Generate QR code and save to file
      console.log(`QR code for Row ${i + 2} created at ${qrCodePath}`);
    } catch (error) {
      console.error(`Error generating QR code for Row ${i + 2}: ${error}`);
    }
  }
}

// Create directory to store QR codes
if (!fs.existsSync("qrcodes")) {
  fs.mkdirSync("qrcodes");
}

// Usage
const excelFilePath = "./excel1.xlsx";
generateQRFromExcel(excelFilePath).catch(console.error);