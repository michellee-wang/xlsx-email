const xlsx = require('xlsx');
const nodemailer = require('nodemailer');

async function sendEmail(recipient, body) {
    const subject = 'Special Invitation: Dinner with Hack Club @ Coletta';

    let transporter = nodemailer.createTransport({
        service: 'gmail', // use your email service
        auth: {
            user: 'from@gmail.com', // your email
            pass: 'password' // your email password
        }
    });

    let mailOptions = {
        from: 'from@gmail.com',  // use your email service
        to: recipient, 
        subject: subject, 
        text: body, 
    };

    try {
        let info = await transporter.sendMail(mailOptions);
        console.log('Email sent: ' + info.response);
    } catch (error) {
        console.error('Error sending email: ' + error.message);
    }
}

function processSpreadsheet(filePath) {
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(sheet, {header:1}); 

    for (let i = 1; i < data.length; i++) {
        let row = data[i];
        sendEmail(row[1], row[2]);
    }
}

// Replace with the path to your spreadsheet
processSpreadsheet('DaysOfService.xlsx');
