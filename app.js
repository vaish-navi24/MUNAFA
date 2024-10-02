const express = require('express');
const app = express();
const path = require('path');
const ExcelJS = require('exceljs');
const fs = require('fs');

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Set up the view engine and static files
app.set('view engine', 'ejs');
app.use(express.static(path.join(__dirname, 'public')));

// Serve the main form page
app.get('/', (req, res) => {
    res.render('main');
});

// Handle form submission and append to Excel file
app.post('/post', async (req, res) => {
    const { name, age } = req.body;
    const filePath = path.join(__dirname, 'form-data.xlsx');

    // Create a new workbook or load
    // the existing one
    let workbook;
    if (fs.existsSync(filePath)) {
        workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);
        console.log("done");
    } else {
        // Create a new workbook if the file doesn't exist
        workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet('Form Data');
        sheet.columns = [
            { header: 'Name', key: 'name', width: 20 },
            { header: 'Age', key: 'age', width: 10 }
        ];
    }

    // Get the worksheet and add a new row
    const sheet = workbook.getWorksheet('Form Data');
    sheet.addRow({ name, age });

    // Save the Excel file with the new data
    await workbook.xlsx.writeFile(filePath);

    res.redirect('/');
});

// Start the server
app.listen(3000, () => {
    console.log('Server is running on http://localhost:3000');
});
