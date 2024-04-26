// generateExcel.js
const Excel = require('excel4node');
const qrCode = require('qrcode');
const fs = require('fs');

// Function to generate QR code image
async function generateQRCode(data) {
    return new Promise((resolve, reject) => {
        qrCode.toDataURL(data, (err, url) => {
            if (err) {
                reject(err);
            } else {
                resolve(url);
            }
        });
    });
}

// Function to attach QR code image to Excel file
async function attachQRCodeToExcel(data) {
    try {
        // Create a new Excel workbook
        const wb = new Excel.Workbook();
        const ws = wb.addWorksheet('QR Code');

        // Generate QR code image
        const qrCodeDataURL = await generateQRCode(data);

        // Save QR code image to a file
        const qrCodeFileName = 'qrCode.png';
        fs.writeFileSync(qrCodeFileName, qrCodeDataURL.split(';base64,').pop(), { encoding: 'base64' });

        // Add image to worksheet
        ws.addImage({
            path: qrCodeFileName,
            type: 'picture',
            position: {
                type: 'twoCellAnchor',
                from: {
                    col: 2,
                    colOff: 0,
                    row: 2,
                    rowOff: 0,      
                },
                to: {
                    col: 5,
                    colOff: 0,
                    row: 10,
                    rowOff: 0,
                },
            },
        });

        // Save Excel workbook
        wb.write('qrCodeExcel.xlsx', (err, stats) => {
            if (err) {
                console.error('Error saving Excel file:', err);
            } else {
                console.log('Excel file saved successfully!');
            }
        });

        // Remove the generated QR code image file
        fs.unlinkSync(qrCodeFileName);
    } catch (error) {
        console.error('Error generating QR code:', error);
    }
}

// Generate Excel file with attached QR code
attachQRCodeToExcel('https://www.example.com');
