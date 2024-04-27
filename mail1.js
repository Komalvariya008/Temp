const ExcelJS = require("exceljs");
const { QRCodeCanvas } = require('@loskir/styled-qr-code-node'); // Import QRCodeCanvas
const fs = require("fs");
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

  // Generate QR codes for each key-value pair
  for (let i = 0; i < values.length; i++) {
    const qrData = JSON.stringify(values[i]); // Convert object to JSON string

    try {
      const qrCodePath = `qrcodes/qr_code_row_${i + 2}.png`; // Path to save QR code image (starting from second row)
      const qrCode = new QRCodeCanvas({
        data: qrData,
        color: { dark: "#000000", light: "#ffffff" }, // Set dark color to black and light color to white
        imageOptions: { "hideBackgroundDots": true, "imageSize": 0.4, "margin": 0 },
      });

      await qrCode.toFile(qrCodePath, 'png');
      console.log(`QR code for Row ${i + 2} created at ${qrCodePath}`);
      await sendEmail(values[i], qrCodePath); // Send email with QR code and rowData
    } catch (error) {
      console.error(`Error generating QR code for Row ${i + 2}: ${error}`);
    }
  }
}

// Create directory to store QR codes
if (!fs.existsSync("qrcodes")) {
  fs.mkdirSync("qrcodes");
}

// Nodemailer transporter configuration
const transporter = nodemailer.createTransport({
  service: "gmail", // E.g., "gmail"
  auth: {
    user: "komalvariya814@gmail.com",
    pass: "wnqc cjsj yvoo pxnc",
  },
});

async function sendEmail(rowData, qrCodePath) {
  // Email message
  const mailOptions = {
    from: "komalvariya814@gmail.com",
    to: rowData["EMAIL"],
    subject: "QR Code",
    text: "Please find your QR code attached.",
    attachments: [
      {
        filename: qrCodePath,
        path: qrCodePath,
      },
    ],
  };

  transporter.sendMail(mailOptions, (error, info) => {
    if (error) {
      console.log('Error occurred: ', error);
    } else {
      console.log('Email sent to ' + rowData["EMAIL"] + ': ' + info.response);
    }
  });
}

// Usage
const excelFilePath = "./excel2.xlsx";
generateQRFromExcel(excelFilePath).catch(console.error);
