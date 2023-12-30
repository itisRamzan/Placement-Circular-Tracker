require('dotenv').config();
const cheerio = require('cheerio');
const fs = require('fs');
const XLSX = require('xlsx');
const cron = require('node-cron');
const nodemailer = require('nodemailer');
const websiteUrl = process.env.WEBSITE_URL;
const excelFilePath = process.env.EXCEL_FILE_PATH;
const emailUser = process.env.EMAIL_USER;
const emailPassword = process.env.EMAIL_PASSWORD;
const emailRecipient = process.env.EMAIL_RECIPIENT;

// Function to read existing data from the Excel file
const readExistingDataFromExcel = () => {
    try {
        const workbook = XLSX.readFile(excelFilePath);
        const worksheet = workbook.Sheets['Circulars'];
        return XLSX.utils.sheet_to_json(worksheet);
    } catch (error) {
        console.error('Error reading existing data:', error);
        return [];
    }
};

// Function to compare two arrays of objects based on a key
const arraysAreEqual = (arr1, arr2, key) => {
    const getKey = obj => obj[key];
    return arr1.map(getKey).join() === arr2.map(getKey).join();
};

// Function to fetch data and save it in the Excel file if it is updated
const fetchDataAndSaveToExcelIfNeeded = () => {
    const existingData = readExistingDataFromExcel();
    fetch(websiteUrl)
        .then(response => response.text())
        .then(html => {
            const $ = cheerio.load(html);
            const data = [];
            const dateRegex = /(\d{2})\.(\d{2})\.(\d{4})/;
            $('div[dir="ltr"] ul a').each((index, element) => {
                const circular = element.children[0]?.data;
                const link = element.attribs.href;
                const date = circular?.match(dateRegex)?.[0] || "No date found";
                data.push({
                    "Company Name": circular,
                    "Link": link || "No link found",
                    "Date": date
                });
            });
            if (!arraysAreEqual(existingData, data, 'circular')) {
                const workbook = XLSX.utils.book_new();
                const worksheet = XLSX.utils.json_to_sheet(data);
                XLSX.utils.book_append_sheet(workbook, worksheet, 'Circulars');
                XLSX.writeFile(workbook, excelFilePath);
                console.log('Excel file updated:', excelFilePath);
                sendEmailNotification();
            } else {
                console.log('No update found.');
            }
        })
        .catch(error => {
            console.error('Error:', error);
        });
};

// Function to send email notification
const sendEmailNotification = () => {
    const transporter = nodemailer.createTransport({
        service: 'gmail',
        auth: {
            user: emailUser,
            pass: emailPassword
        }
    });
    const mailOptions = {
        from: `"My Placement Circular Tracker" <${emailUser}>`,
        to: emailRecipient,
        subject: "New Company came for Placement",
        text: "New Company Circulars Added.",
        importance: 'high',  // Set email importance to high
        headers: {
            'X-Priority': '1',  // Set email priority to highest
            'X-MSMail-Priority': 'High',  // Set email priority for Outlook
        },
        attachments: [
            {
                filename: 'Placement Circulars.xlsx',
                path: excelFilePath
            }
        ]
    };

    transporter.sendMail(mailOptions, (error, info) => {
        if (error) {
            console.error('Error sending email:', error);
        } else {
            console.log('Email sent:', info.response);
        }
    });
};

// Function to create the Excel file if it doesn't exist
const createExcelFileIfNeeded = () => {
    if (!fs.existsSync(excelFilePath)) {
        const workbook = XLSX.utils.book_new();
        const worksheet = XLSX.utils.json_to_sheet([]);
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Circulars');
        XLSX.writeFile(workbook, excelFilePath);
        console.log('Excel file created:', excelFilePath);
    }
};

createExcelFileIfNeeded();
fetchDataAndSaveToExcelIfNeeded();
cron.schedule('0 */1 * * *', () => {
    fetchDataAndSaveToExcelIfNeeded();
});