const fs = require('fs');
const xlsx = require('xlsx');
const nodemailer = require('nodemailer');

// Read the Excel file
const workbook = xlsx.readFile('excel.xlsx');
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// Parse email addresses from the Excel file
const emails = xlsx.utils.sheet_to_json(worksheet, { header: 'A' });

// Nodemailer setup
const transporter = nodemailer.createTransport({
  service: 'Gmail',
  auth: {
    user: 'variyakomal008@gmail.com',
    pass: 'mpwi oebm ghje fzhf'
  }
});

// Email content
const mailOptions = {
  from: 'variyakomal008@gmail.com', // Sender address
  subject: 'Test Subject',
  text: 'Test Body'
};

// Function to send emails
function sendEmail(email) {
  mailOptions.to = email;
  transporter.sendMail(mailOptions, (error, info) => {
    if (error) {
      console.log('Error occurred: ', error);
    } else {
      console.log('Email sent to ' + email + ': ' + info.response);
    }
  });
}

// Send emails
emails.forEach(email => {
  sendEmail(email);
});
