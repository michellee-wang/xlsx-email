const xlsx = require('xlsx');
const nodemailer = require('nodemailer');

// Function to send email
async function sendEmail(recipient, subject, body) {
    let transporter = nodemailer.createTransport({
        service: 'gmail', // use your email service
        auth: {
            user: 'your-email@gmail.com', // your email
            pass: 'your-password' // your email password
        }
    });

    let mailOptions = {
        from: 'your-email@gmail.com', // sender address
        to: recipient, // list of receivers
        subject: subject, // Subject line
        text: body, // plain text body
    };

    try {
        let info = await transporter.sendMail(mailOptions);
        console.log('Email sent: ' + info.response);
    } catch (error) {
        console.error('Error sending email: ' + error.message);
    }
}

// Function to read spreadsheet and send emails
function processSpreadsheet(filePath) {
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(sheet);

    data.forEach(row => {
        // Assuming the spreadsheet has 'Email', 'Subject', and 'Body' columns
        sendEmail(row.Email, row.Subject, row.Body);
    });
}

// Replace with the path to your spreadsheet
processSpreadsheet('path/to/your/spreadsheet.xlsx');
