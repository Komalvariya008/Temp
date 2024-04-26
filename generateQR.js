// generateQR.js
const { createCanvas } = require('canvas');
const qrCode = require('qrcode');

const options = {
    width: 300,
    height: 300,
    errorCorrectionLevel: 'H',
    type: 'png',
    quality: 1,
    margin: 1,
    color: {
        dark: '#4267b2',
        light: '#ffffff'
    }
};

const data = "qwertyui";

qrCode.toDataURL(data, options, (err, url) => {
    if (err) throw err;
    console.log('Custom QR Code:', url);
});