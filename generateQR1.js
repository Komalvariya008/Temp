const ExcelJS = require("exceljs");
const QRCode = require("qrcode");
const fs = require("fs");
const sharp = require("sharp");
const nodemailer = require("nodemailer");

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

  // Colors for foreground: Orange, White, Green
  const foregroundColors = ["#FF9933", "#FFFFFF", "#138808"];

  // Generate QR codes for each key-value pair
  for (let i = 0; i < values.length; i++) {
    const qrData = JSON.stringify(values[i]); // Convert object to JSON string

    try {
      const qrCodeSVG = await QRCode.toString(qrData, {
        type: "svg",
        color: {
          dark: foregroundColors[0], // Cycle through colors
          light: foregroundColors[2]
        },
      });

      const qrCodePath = `qrcodes/qr_code_row_${i + 2}.png`; // Path to save QR code image (starting from second row)

      // Convert SVG to PNG using sharp
      await sharp(Buffer.from(qrCodeSVG)).png().toFile(qrCodePath);

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
const excelFilePath = "Financial Sample.xlsx";
generateQRFromExcel(excelFilePath).catch(console.error);
