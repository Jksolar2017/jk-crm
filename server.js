// 🚀 Basic Setup
const express = require('express');
const multer = require('multer');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const ExcelJS = require('exceljs');
const bodyParser = require('body-parser');
const cors = require('cors');
const axios = require('axios');
const qs = require('qs');
require('dotenv').config();

const app = express();
const port = 3000;

app.use(cors());
app.use(express.static('public')); // your frontend files (html/css/js)
app.use(express.json());
app.use('/public', express.static(path.join(__dirname, 'public')));
app.use(express.urlencoded({ extended: true }));
app.use(express.static(__dirname));

// 📂 Paths Setup
const crmFolderPath = 'C:/Users/JK SOLAR/OneDrive/CRM_PWA';
const uploadsBasePath = path.join(crmFolderPath, 'uploads');
const excelFilePath = path.join(crmFolderPath, 'TempData.xlsx');
const stockFilePath = path.join(crmFolderPath, 'Stock Sheet.xlsx');
//const applicationTimelineFilePath = path.join(crmFolderPath, 'applicationTimeline.xlsx');

// 🛠 Middlewares
app.use(express.static('public'));
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

//excel path for timeline application
//const applicationExcelFilePath = 'C:/Users/JK SOLAR/OneDrive/CRM_PWA/applicationTimeline.xlsx';
const excelPath = 'C:/Users/JK SOLAR/OneDrive/CRM_PWA/Stock Sheet.xlsx';
const FILE_PATH = 'C:/Users/JK SOLAR/OneDrive/CRM_PWA/Stock Sheet.xlsx';

// Middleware to serve frontend
app.use(express.static('public'));
app.use(express.json());
app.use(express.urlencoded({ extended: true }));


// ✅ Multer instance for multiple file fields (aadhar, bill, etc.)
const multiUpload = multer({
  storage: multer.diskStorage({
    destination: function (req, file, cb) {
      const clientName = req.body.name.trim().toLowerCase().replace(/\s+/g, '_');
      const clientFolderPath = path.join(uploadsBasePath, clientName);

      // Ensure the client folder exists
      if (!fs.existsSync(clientFolderPath)) {
        fs.mkdirSync(clientFolderPath, { recursive: true });
        console.log(`✅ Folder created: ${clientFolderPath}`);
      } else {
        console.log(`📁 Folder already exists: ${clientFolderPath}`);
      }

      cb(null, clientFolderPath); // Use the correct folder path
    },
    filename: function (req, file, cb) {
      const ext = path.extname(file.originalname).toLowerCase();
      let newName = '';

      // Assign unique names based on field name
      switch (file.fieldname) {
        case 'aadharFront': newName = 'aadharfront' + ext; break;
        case 'aadharBack': newName = 'aadharback' + ext; break;
        case 'panCard': newName = 'pancard' + ext; break;
        case 'bill': newName = 'bill' + ext; break;
        case 'ownershipProof': newName = 'ownershipproof' + ext; break;
        case 'cancelCheque': newName = 'cancelcheque' + ext; break;
        case 'purchaseAgreement': newName = 'purchaseagreement' + ext; break;
        case 'netMeteringAgreement': newName = 'netmeteringagreement' + ext; break;
        default: newName = Date.now() + '-' + file.originalname; break;
      }

      cb(null, newName);
    }
  })
}).fields([
  { name: 'aadharFront', maxCount: 1 },
  { name: 'aadharBack', maxCount: 1 },
  { name: 'panCard', maxCount: 1 },
  { name: 'bill', maxCount: 1 },
  { name: 'ownershipProof', maxCount: 1 },
  { name: 'cancelCheque', maxCount: 1 },
  { name: 'purchaseAgreement', maxCount: 1 },
  { name: 'netMeteringAgreement', maxCount: 1 },
  { name: 'clientPhoto', maxCount: 1 }  // ✅ Add this line
]);


// ✅ Handle the form submission with multiple file uploads
app.post('/submit-client', multiUpload, async (req, res) => {

  const data = req.body;
  const files = req.files;
//moved
  if (req.body.photoData) {
    const base64Data = req.body.photoData.replace(/^data:image\/png;base64,/, '');
    const clientName = req.body.name.trim().toLowerCase().replace(/\s+/g, '_');
    const clientFolder = path.join(uploadsBasePath, clientName);
  
    if (!fs.existsSync(clientFolder)) {
      fs.mkdirSync(clientFolder, { recursive: true });
    }
  
    const imagePath = path.join(clientFolder, 'clientPhoto.png');
    fs.writeFileSync(imagePath, base64Data, 'base64');
  }
  const clientName = data.name.trim().toLowerCase().replace(/\s+/g, '_');
  const clientFolder = path.join(uploadsBasePath, clientName);

  // Create the client folder if it doesn't exist
  if (!fs.existsSync(clientFolder)) {
    fs.mkdirSync(clientFolder, { recursive: true });
    console.log(`✅ Folder created: ${clientFolder}`);
  } else {
    console.log(`📁 Folder already exists: ${clientFolder}`);
  }

  // Debugging the uploaded files object
  console.log('Form Data:', data);
  console.log('Uploaded Files:', files);

  // Prepare the data object to save
  const client = {
    date: data.date,
    name: data.name,
    address: data.address,
    mobile: data.mobile,
    email: data.email,
    kw: data.kw,
    advance: data.advance,
    totalCost: data.totalCost,
    aadharFront: files?.aadharFront?.[0]?.path || '',
    aadharBack: files?.aadharBack?.[0]?.path || '',
    panCard: files?.panCard?.[0]?.path || '',
    bill: files?.bill?.[0]?.path || '',
    ownershipProof: files?.ownershipProof?.[0]?.path || '',
    cancelCheque: files?.cancelCheque?.[0]?.path || '',
    purchaseAgreement: files?.purchaseAgreement?.[0]?.path || '',
    netMeteringAgreement: files?.netMeteringAgreement?.[0]?.path || ''
  };

  // Handle saving the data (for example, to an Excel file)
  let workbook, worksheet;
  if (fs.existsSync(excelFilePath)) {
    workbook = xlsx.readFile(excelFilePath);
    worksheet = workbook.Sheets[workbook.SheetNames[0]];
  } else {
    workbook = xlsx.utils.book_new();
    worksheet = xlsx.utils.aoa_to_sheet([[
      'Date', 'Name', 'Address', 'Mobile', 'Email', 'KW', 'Advance', 'Total Cost', 'Aadhar Front', 'Aadhar Back', 'Pan Card', 'Bill', 'Ownership Proof', 'Cancel Cheque', 'Purchase Agreement', 'Net Metering Agreement'
    ]]);
    xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
  }

  try {
    // Append new row to the sheet
    const sheetData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
    sheetData.push([
      client.date, client.name, client.address, client.mobile, client.email, client.kw,
      client.advance, client.totalCost, client.aadharFront, client.aadharBack, client.panCard, client.bill, client.ownershipProof, client.cancelCheque,
      client.purchaseAgreement, client.netMeteringAgreement
    ]);
  
    const updatedSheet = xlsx.utils.aoa_to_sheet(sheetData);
    workbook.Sheets[workbook.SheetNames[0]] = updatedSheet;
  
    // Write updated workbook to file
    xlsx.writeFile(workbook, excelFilePath);
    console.log('✅ Client data saved locally to Excel file.');
  
    res.send('✅ Client submitted successfully and files saved.');
    
  } catch (err) {
    console.error('❌ Error writing to Excel file:', err.message);
    res.status(500).send('⚠️ Could not save data. Make sure TempData.xlsx is closed and not open in Excel.');
  }
})

// ✅ Search route to find if the client folder exists
const uploadFolderPath = uploadsBasePath;

app.get('/search-client', (req, res) => {
  const name = req.query.name?.trim();

  if (!name) {
    return res.status(400).json({ error: 'No name provided' });
  }

  fs.readdir(uploadsBasePath, { withFileTypes: true }, (err, files) => {
    if (err) {
      console.error('Error reading uploads folder:', err);
      return res.status(500).json({ error: 'Internal Server Error' });
    }

    const folderNames = files
      .filter(dirent => dirent.isDirectory())
      .map(dirent => dirent.name.toLowerCase());

    const clientExists = folderNames.includes(name.toLowerCase());

    res.json({ found: clientExists });
  });
});

// Route to handle adding timeline event for a client
app.post('/add-timeline', express.json(), (req, res) => {
  const { clientName, event, eventDate, eventDescription, status } = req.body;

  const filePath = 'TempData.xlsx';
  const workbook = new ExcelJS.Workbook();

  // Check if TempData.xlsx already exists
  if (fs.existsSync(filePath)) {
    workbook.xlsx.readFile(filePath).then(() => {
      let worksheet = workbook.getWorksheet('Timeline'); // Look for the Timeline sheet

      // If the "Timeline" sheet doesn't exist, create it
      if (!worksheet) {
        worksheet = workbook.addWorksheet('Timeline');
        worksheet.addRow(['Client Name', 'Event', 'Event Date', 'Event Description', 'Status']);
      }

      // Add the new timeline event
      worksheet.addRow([clientName, event, eventDate, eventDescription, status]);

      // Save the Excel file with the new timeline event
      return workbook.xlsx.writeFile(filePath);
    }).then(() => {
      res.send('Timeline event added successfully!');
    }).catch((error) => {
      console.error('Error adding timeline event:', error);
      res.status(500).send('Error adding timeline event.');
    });
  } else {
    // If TempData.xlsx doesn't exist, create it and add Timeline sheet
    const worksheet = workbook.addWorksheet('Timeline');
    worksheet.addRow(['Client Name', 'Event', 'Event Date', 'Event Description', 'Status']);
    worksheet.addRow([clientName, event, eventDate, eventDescription, status]);
    
    workbook.xlsx.writeFile(filePath).then(() => {
      res.send('Timeline event added successfully!');
    }).catch((error) => {
      console.error('Error creating new Excel file:', error);
      res.status(500).send('Error creating new Excel file.');
    });
  }
});

// Multer setup for document uploads
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    // Ensure the 'uploads' folder exists
    const uploadDir = 'uploads';
    if (!fs.existsSync(uploadDir)) {
      fs.mkdirSync(uploadDir);
    }
    cb(null, uploadDir); // Set the folder where files will be uploaded
  },
  filename: (req, file, cb) => {
    cb(null, Date.now() + path.extname(file.originalname)); // Set unique filenames based on current timestamp
  }
});
const upload = multer({ storage: storage });

// New route for document upload (handles multiple files)
app.post('/submit-documents', upload.fields([
  { name: 'aadharFront', maxCount: 1 },
  { name: 'aadharBack', maxCount: 1 },
  { name: 'panCard', maxCount: 1 },
  { name: 'bill', maxCount: 1 },
  { name: 'ownershipProof', maxCount: 1 },
  { name: 'cancelCheque', maxCount: 1 },
  { name: 'purchaseAgreement', maxCount: 1 },
  { name: 'netMeteringAgreement', maxCount: 1 }
]), (req, res) => {
  if (!req.files) {
    return res.status(400).send('No files were uploaded.');
  }

  // You can access the uploaded files here
  console.log(req.files);  // Log all uploaded files (for debugging)

  // If needed, save the file information to a database or Excel file
  // Example: Save file paths or names in your database

  // Send success response
  res.send('Documents uploaded successfully!');
});


// Define the uploads folder path
const uploadFolder = path.join('C:', 'Users', 'JK SOLAR', 'OneDrive', 'CRM_PWA', 'uploads');

// List of required files
const requiredFiles = [
  'AadharFront.jpg',
  'AadharBack.jpg',
  'PanCard.jpg',
  'Bill.jpg',
  'OwnershipProof.jpg',
  'CancelCheque.jpg',
  'PurchaseAgreement.pdf',
  'NetMeteringAgreement.pdf'
];

// Serve static files from the uploads folder
app.use('/uploads', express.static(uploadFolder));

// Route to return file status for a specific client
// Define the Excel file path
// Route to return file status for a specific client
app.get('/file-status/:clientName', async (req, res) => {
  const clientName = req.params.clientName.trim().toLowerCase();
  const filePaths = [
    'AadharFront',
    'AadharBack',
    'PanCard',
    'Bill',
    'OwnershipProof',
    'CancelCheque',
    'PurchaseAgreement',
    'NetMeteringAgreement'
  ];

  try {
    console.log(`📂 Checking files for: ${clientName}`);

    let clientInfo = null;
    let fileStatus = [];

    if (fs.existsSync(excelFilePath)) {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(excelFilePath);
      const worksheet = workbook.getWorksheet('Client Data');

      worksheet.eachRow((row) => {
        const rowClientName = (row.getCell(2).value || '').toString().trim().toLowerCase(); // Ensure you’re trimming and handling cases
  console.log(`Row client name: '${rowClientName}' (Excel data) vs '${clientName}' (searched name)`); // Debugging log

        if (rowClientName === clientName) {
          clientInfo = {
            name: row.getCell(2).value || '',    // B
            address: row.getCell(3).value || '', // C
            mobile: row.getCell(4).value || '',  // D
            email: row.getCell(5).value || '',   // E
            kw: row.getCell(6).value || ''       // F
          };

          // Now, check file paths in columns I to P (Aadhar Front, Aadhar Back, etc.)
          filePaths.forEach((filePath, index) => {
            const filePathInExcel = row.getCell(9 + index).value?.toString().trim(); // Columns I to P
            const exists = fs.existsSync(filePathInExcel); // Check if file exists

            const label = filePath
              .replace(/([A-Z])/g, ' $1') // Add space before capital letters
              .replace(/^./, str => str.toUpperCase()) // Capitalize the first letter
              .trim();

            fileStatus.push({
              file: filePath,
              label: label,
              exists: exists
            });
          });
        }
      });
    }

    if (clientInfo) {
      return res.json({
        files: fileStatus,
        clientInfo: clientInfo || {}
      });
    } else {
      return res.status(404).json({ error: 'Client not found' });
    }

  } catch (err) {
    console.error('❌ Error in /file-status route:', err);
    return res.status(500).json({ error: 'Internal Server Error' });
  }
});


// 🧪 Just to view all status manually
app.get('/check-files', (req, res) => {
  res.send(`<pre>${JSON.stringify(topLevelFileStatus, null, 2)}</pre>`);
});

// 📅 View timeline of a client
app.get('/view-timeline/:clientName', async (req, res) => {
  const rawClientName = req.params.clientName || '';
  const clientName = rawClientName.trim().toLowerCase();
  const filePath = 'TempData.xlsx';

  console.log(`🔍 Looking for timeline for client: "${clientName}"`);

  if (!fs.existsSync(filePath)) {
    return res.status(404).json({ error: 'Excel file not found' });
  }

  const workbook = new ExcelJS.Workbook();

  try {
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.getWorksheet('Client Data');

    if (!worksheet) {
      return res.status(404).json({ error: 'Client Data sheet not found' });
    }

    const clientTimeline = [];

    worksheet.eachRow((row, rowNumber) => {
      const rowValues = row.values;
      const clientNameFromRow = rowValues[2]; // Column B = Name
      const event = rowValues[3]; // Column C = Event

      if (
        clientNameFromRow &&
        clientNameFromRow.toLowerCase().trim() === clientName &&
        event && event.trim() !== ''
      ) {
        clientTimeline.push({
          event: event,                     // C
          eventDate: rowValues[4],         // D
          eventDescription: rowValues[5],  // E
          status: rowValues[6]             // F
        });
      }
    });

    if (clientTimeline.length === 0) {
      return res.status(404).json({ error: 'No timeline events found for this client' });
    }

    res.json(clientTimeline);
  } catch (error) {
    console.error('❌ Error reading Excel file:', error);
    res.status(500).json({ error: 'Internal Server Error' });
  }
});

// 🧾 Get client info (basic details)
app.get('/client-info/:clientName', async (req, res) => {
  const clientName = req.params.clientName.trim().toLowerCase();
  const filePath = 'TempData.xlsx';

  if (!fs.existsSync(filePath)) {
    return res.status(404).json({ error: 'Excel file not found' });
  }

  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.getWorksheet('Client Data');

    if (!sheet) {
      return res.status(404).json({ error: 'Client Data sheet not found' });
    }

    let clientInfo = null;

    sheet.eachRow((row) => {
      const rowClientName = row.getCell(2).value?.toString().trim().toLowerCase(); // Column B

      if (rowClientName === clientName) {
        clientInfo = {
          name: row.getCell(2).value || '',     // B
          address: row.getCell(3).value || '',  // C
          mobile: row.getCell(4).value || '',   // D
          email: row.getCell(5).value || '',    // E
          kw: row.getCell(6).value || ''        // F
        };
      }
    });

    if (clientInfo) {
      res.json(clientInfo);
    } else {
      res.status(404).json({ error: 'Client not found' });
    }
  } catch (err) {
    console.error('❌ Error reading client info:', err);
    res.status(500).json({ error: 'Internal Server Error' });
  }
});

//aaplication timeline function
// Set up middleware for parsing form data
//aaplication timeline function
// Set up middleware for parsing form data
//aaplication timeline function
// Set up middleware for parsing form data
//aaplication timeline function
// Set up middleware for parsing form data
//aaplication timeline function
// Set up middleware for parsing form data

app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// Serve static files (like your HTML, CSS, etc.)
app.use(express.static('public'));

// Function to initialize Excel file if it doesn't exist
//function initializeApplicationExcelFile() {
  //if (!fs.existsSync(applicationExcelFilePath)) {
  //  const wb = xlsx.utils.book_new();
   // const ws = xlsx.utils.aoa_to_sheet([
   //   ['Applied KW', 'Applied on PM Surya', 'Application DHBVN', 'Load/Name Change/Reduction/New Connection']
   // ]);
   // xlsx.utils.book_append_sheet(wb, ws, 'Application Timeline');
   // xlsx.writeFile(wb, applicationExcelFilePath);
   // console.log('✅ Created new applicationTimeline.xlsx file');
//  } else {
 //   console.log('📂 applicationTimeline.xlsx already exists');
 // }


// Call the initializer when the server starts
//initializeApplicationExcelFile();

// 📝 Route to save timeline data (with upload.none() to parse FormData without files)
const Excel = require('exceljs');
const tempDataPath = 'C:/Users/JK SOLAR/OneDrive/CRM_PWA/TempData.xlsx';

app.post('/save-timeline', upload.none(), async (req, res) => {
  const { appliedKW, appliedOnPMSurya, applicationDHBVN, loadChangeReductionNewConnection, clientName } = req.body;

  if (!clientName) return res.status(400).json({ error: 'Client name missing' });

  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(tempDataPath);
    const sheet = workbook.getWorksheet('Client Data');

    let found = false;

    sheet.eachRow((row, rowNumber) => {
      const nameCell = row.getCell(2).value?.toString().toLowerCase().trim(); // Column B

      if (nameCell === clientName.toLowerCase().trim()) {
        row.getCell(17).value = appliedKW;                             // Column Q
        row.getCell(18).value = appliedOnPMSurya;                      // Column R
        row.getCell(19).value = applicationDHBVN;                      // Column S
        row.getCell(20).value = loadChangeReductionNewConnection;     // Column T
        found = true;
      }
    });

    if (!found) {
      return res.status(404).json({ error: 'Client not found in Excel' });
    }

    await workbook.xlsx.writeFile(tempDataPath);

    res.json({ message: 'Timeline data saved successfully!' });
  } catch (err) {
    console.error('Error saving timeline data:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});





// Serve the timeline data as JSON
// Route to fetch the application timeline for a specific client
app.get('/application-timeline/:clientName', async (req, res) => {
  const clientName = req.params.clientName.toLowerCase().trim();

  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(tempDataPath);
    const sheet = workbook.getWorksheet('Client Data');

    let result = null;

    sheet.eachRow((row) => {
      const nameCell = row.getCell(2).value?.toString().toLowerCase().trim(); // Column B

      if (nameCell === clientName) {
        result = {
          appliedKW: row.getCell(17)?.value || '',
          appliedOnPMSurya: row.getCell(18)?.value || '',
          applicationDHBVN: row.getCell(19)?.value || '',
          loadChangeReductionNewConnection: row.getCell(20)?.value || ''
        };
      }
    });

    if (result) {
      res.json(result);
    } else {
      res.status(404).json({ error: 'Client not found' });
    }
  } catch (err) {
    console.error('Error loading timeline data:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});





// Add this route in your Express server (server.js)

app.get('/file-status/:clientName', (req, res) => {
  const clientName = req.params.clientName.toLowerCase(); // Make sure client name is consistent
  const filePath = path.join(__dirname, 'uploads', clientName, 'TempData.xlsx'); // Update the path if necessary
  
  if (!fs.existsSync(filePath)) {
    return res.status(404).json({ message: 'Client not found' });
  }

  // Load the client data from the excel file (update if needed to match your logic)
  const workbook = XLSX.readFile(filePath);
  const worksheet = workbook.Sheets['Client Data'];
  const jsonData = XLSX.utils.sheet_to_json(worksheet);

  // Fetch client-specific data based on logic
  const clientData = jsonData.find(client => client.Name.toLowerCase() === clientName);
  
  if (!clientData) {
    return res.status(404).json({ message: 'No data found for the client' });
  }

  // Respond with client data
  res.json({
    clientInfo: {
      name: clientData.Name,
      address: clientData.Address,
      mobile: clientData.Mobile,
      email: clientData.Email,
      kw: clientData.KW
    },
    files: [
      { file: 'aadharfront.png', exists: fs.existsSync(path.join(filePath, 'aadharfront.png')) },
      { file: 'aadharback.png', exists: fs.existsSync(path.join(filePath, 'aadharback.png')) },
      { file: 'pancard.png', exists: fs.existsSync(path.join(filePath, 'pancard.png')) },
      { file: 'bill.png', exists: fs.existsSync(path.join(filePath, 'bill.png')) },
      { file: 'ownershipproof.png', exists: fs.existsSync(path.join(filePath, 'ownershipproof.png')) },
      { file: 'cancelcheque.png', exists: fs.existsSync(path.join(filePath, 'cancelcheque.png')) },
      { file: 'purchaseagreement.png', exists: fs.existsSync(path.join(filePath, 'purchaseagreement.png')) },
      { file: 'netmeteringagreement.png', exists: fs.existsSync(path.join(filePath, 'netmeteringagreement.png')) }
    ]
  });
});

// Sample route to get timeline data
// Correct route to get application timeline for a client
app.get('/application-timeline/:clientName', async (req, res) => {
  const rawClientName = req.params.clientName; // Use params for URL parameters
  const clientName = rawClientName ? rawClientName.trim().toLowerCase() : '';

  if (!clientName) {
    return res.status(400).json({ error: 'Client name is required' });
  }

  const filePath = 'C:/Users/JK SOLAR/OneDrive/CRM_PWA/TempData.xlsx'; // Path to the Excel file

  // Check if file exists
  if (!fs.existsSync(filePath)) {
    return res.status(404).json({ error: 'Excel file not found' });
  }

  const workbook = new ExcelJS.Workbook();
  
  try {
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.getWorksheet('Client Data');

    const clientTimeline = [];

    worksheet.eachRow((row) => {
      const rowValues = row.values;
      const clientNameFromRow = rowValues[2]; // Column B = Name
      const appliedKW = rowValues[17]; // Column Q = Applied KW
      const appliedOnPMSurya = rowValues[18]; // Column R = Applied on PM Surya
      const applicationDHBVN = rowValues[19]; // Column S = Application DHBVN
      const loadChangeReductionNewConnection = rowValues[20]; // Column T = Load/Name Change/Reduction/New Connection

      if (clientNameFromRow && clientNameFromRow.toLowerCase().trim() === clientName) {
        clientTimeline.push({
          appliedKW: appliedKW,
          appliedOnPMSurya: appliedOnPMSurya,
          applicationDHBVN: applicationDHBVN,
          loadChangeReductionNewConnection: loadChangeReductionNewConnection
        });
      }
    });

    if (clientTimeline.length === 0) {
      return res.status(404).json({ error: 'No application timeline found for this client' });
    }

    res.json(clientTimeline); // Send timeline data as response
  } catch (err) {
    console.error('❌ Error reading Excel file:', err);
    return res.status(500).json({ error: 'Internal Server Error' });
  }
});

//project status timeline excel
const projectFields = [
  "Civil", "Earthing", "EarthingCable", "Panel", "Inverter", "ACDB",
  "DCDB", "ACCable", "DCCable", "LA", "NetMetering"
];

app.post('/save-project-status', express.urlencoded({ extended: true }), async (req, res) => {
  const clientName = req.body.clientName?.toLowerCase().trim();
  if (!clientName) return res.status(400).json({ error: 'Client name missing' });

  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('C:/Users/JK SOLAR/OneDrive/CRM_PWA/TempData.xlsx');
    const sheet = workbook.getWorksheet('Client Data');

    let updated = false;

    sheet.eachRow((row) => {
      const nameCell = row.getCell(2).value?.toString().toLowerCase().trim(); // Column B
      if (nameCell === clientName) {
        // Columns U (21) to AE (31)
        projectFields.forEach((field, index) => {
          row.getCell(21 + index).value = req.body[field] || '';
        });
        updated = true;
      }
    });

    if (!updated) {
      return res.status(404).json({ error: 'Client not found in Excel' });
    }

    await workbook.xlsx.writeFile('C:/Users/JK SOLAR/OneDrive/CRM_PWA/TempData.xlsx');
    res.json({ message: 'Project status saved successfully' });
  } catch (err) {
    console.error('❌ Error saving project status:', err);
    res.status(500).json({ error: 'Internal Server Error' });
  }
});

app.get('/project-status/:clientName', async (req, res) => {
  const clientName = req.params.clientName.toLowerCase().trim();

  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('C:/Users/JK SOLAR/OneDrive/CRM_PWA/TempData.xlsx');
    const sheet = workbook.getWorksheet('Client Data');

    let result = null;

    sheet.eachRow(row => {
      const name = row.getCell(2).value?.toString().toLowerCase().trim();
      if (name === clientName) {
        result = {
          Civil: row.getCell(21)?.value || '',
          Earthing: row.getCell(22)?.value || '',
          EarthingCable: row.getCell(23)?.value || '',
          Panel: row.getCell(24)?.value || '',
          Inverter: row.getCell(25)?.value || '',
          ACDB: row.getCell(26)?.value || '',
          DCDB: row.getCell(27)?.value || '',
          ACCable: row.getCell(28)?.value || '',
          DCCable: row.getCell(29)?.value || '',
          LA: row.getCell(30)?.value || '',
          NetMetering: row.getCell(31)?.value || ''
        };
      }
    });

    if (!result) return res.status(404).json({ error: 'Client not found' });
    res.json({ status: result });
  } catch (err) {
    console.error('❌ Error loading project status:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

//Payemnt timeline Autofetch route
app.get('/payment-status/:clientName', async (req, res) => {
  const clientName = req.params.clientName.toLowerCase().trim();

  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('C:/Users/JK SOLAR/OneDrive/CRM_PWA/TempData.xlsx');
    const sheet = workbook.getWorksheet('Client Data');

    let result = null;

    sheet.eachRow((row) => {
      const name = row.getCell(2).value?.toString().toLowerCase().trim();
      if (name === clientName) {
        result = {
          totalCost: row.getCell(8)?.value || '',
          advance: row.getCell(7)?.value || '',
          projectStatus: {
            Civil: row.getCell(21)?.value || '',
            NetMetering: row.getCell(31)?.value || ''
          },
          saved: {
            installment2: row.getCell(32)?.value || '',
            installment3: row.getCell(33)?.value || '',
            finalPayment: row.getCell(34)?.value || '',
            balance: row.getCell(35)?.value || ''
          }
        };
      }
    });

    if (!result) return res.status(404).json({ error: 'Client not found' });
    res.json(result);

  } catch (err) {
    console.error('Error reading payment data:', err);
    res.status(500).json({ error: 'Internal Server Error' });
  }
});


//save payment data to excel
app.post('/save-payment-status', express.urlencoded({ extended: true }), async (req, res) => {
  const clientName = req.body.clientName?.toLowerCase().trim();
  if (!clientName) return res.status(400).json({ error: 'Client name missing' });

  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('C:/Users/JK SOLAR/OneDrive/CRM_PWA/TempData.xlsx');
    const sheet = workbook.getWorksheet('Client Data');

    let updated = false;

    sheet.eachRow(row => {
      const nameCell = row.getCell(2).value?.toString().toLowerCase().trim(); // Column B
      if (nameCell === clientName) {
        row.getCell(32).value = req.body.installment2 || '';
        row.getCell(33).value = req.body.installment3 || '';
        row.getCell(34).value = req.body.finalPayment || '';
        row.getCell(35).value = req.body.balance || '';
        updated = true;
      }
    });

    if (!updated) return res.status(404).json({ error: 'Client not found' });

    await workbook.xlsx.writeFile('C:/Users/JK SOLAR/OneDrive/CRM_PWA/TempData.xlsx');
    res.json({ message: 'Payment status saved successfully' });
  } catch (err) {
    console.error('Error saving payment status:', err);
    res.status(500).json({ error: 'Internal Server Error' });
  }
});


//Stocksheet
// Load workbook
async function loadWorkbook() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(FILE_PATH);
  return wb;
}



// Load workbook
// Load and Save Excel Helpers
async function loadWorkbook() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(FILE_PATH);
  return workbook;
}

async function saveWorkbook(workbook) {
  await workbook.xlsx.writeFile(FILE_PATH);
}

// ✅ Find or create row for material in "Stock April"
function findOrCreateMaterialRow(sheet, material) {
  for (let i = 3; i <= sheet.rowCount; i++) {
    const cell = sheet.getCell(`A${i}`);
    if (cell.value && cell.value.toString().toLowerCase() === material.toLowerCase()) {
      return sheet.getRow(i);
    }
  }
  const newRow = sheet.addRow([material]);
  return newRow;
}


// ✅ Stock In
app.post('/submit-stock-in', async (req, res) => {
  const { date, material, invoice, quantity } = req.body;
  if (!date || !material || !invoice || !quantity) {
    return res.status(400).send('Missing required fields');
  }

  try {
    const workbook = await loadWorkbook();

    const monthNumber = ('0' + (new Date(date).getMonth() + 1)).slice(-2);
    const stockSheetName = `Stock ${monthNumber}`;
    const stockInSheetName = `Stock In ${monthNumber}`;

    const sheetIn = getOrCreateSheet(workbook, stockInSheetName);
    const stockSheet = getOrCreateSheet(workbook, stockSheetName);

    // Save in stock in sheet
    sheetIn.addRow([date, material, invoice, quantity]);

    // Ensure date columns exist
    createMissingDates(stockSheet, date);

    // Retry finding date columns
    let dateCols = findDateColumns(stockSheet, date);

    // Retry fallback
    if (!dateCols) {
      await saveWorkbook(workbook);


      await workbook.xlsx.readFile(FILE_PATH);
      const retrySheet = workbook.getWorksheet(stockSheetName);
      createMissingDates(retrySheet, date); // try creating again just in case
      dateCols = findDateColumns(retrySheet, date);
    }

    if (!dateCols) {
      console.error('❌ Still cannot find columns after re-creating');
      return res.status(500).send('❌ Date columns missing.');
    }

    const materialRow = findOrCreateMaterialRow(stockSheet, material);
    const { inCol } = dateCols;

    const existingQty = parseFloat(materialRow.getCell(inCol).value) || 0;
    materialRow.getCell(inCol).value = existingQty + parseFloat(quantity);

    await updateCurrentStock(workbook, stockSheet, material);
    await saveWorkbook(workbook);


    res.send('✅ Stock In recorded & Stock Sheet updated');
  } catch (err) {
    console.error(err);
    res.status(500).send('❌ Error writing Excel');
  }
});



// ✅ Stock Out
app.post('/submit-stock-out', async (req, res) => {
  const { date, material, quantity, remarks } = req.body;
  if (!date || !material || !quantity) {
    return res.status(400).send('Missing required fields');
  }

  try {
    const workbook = await loadWorkbook();

    const monthNumber = ('0' + (new Date(date).getMonth() + 1)).slice(-2);
    const stockSheetName = `Stock ${monthNumber}`;
    const stockOutSheetName = `Stock Out ${monthNumber}`;

    const sheetOut = getOrCreateSheet(workbook, stockOutSheetName);
    const stockSheet = getOrCreateSheet(workbook, stockSheetName);

    // Save to stock out sheet
    sheetOut.addRow([date, material, quantity, remarks || '']);

    // Create missing date columns
    createMissingDates(stockSheet, date);

    // Retry finding date columns
    let dateCols = findDateColumns(stockSheet, date);

    if (!dateCols) {
      await saveWorkbook(workbook);


      await workbook.xlsx.readFile(FILE_PATH);
      const retrySheet = workbook.getWorksheet(stockSheetName);
      createMissingDates(retrySheet, date); // reapply
      dateCols = findDateColumns(retrySheet, date);
    }

    if (!dateCols) {
      console.error('❌ Still cannot find columns after re-creating');
      return res.status(500).send('❌ Date columns missing.');
    }

    const materialRow = findOrCreateMaterialRow(stockSheet, material);
    const { outCol } = dateCols;

    const existingOut = parseFloat(materialRow.getCell(outCol).value) || 0;
    materialRow.getCell(outCol).value = existingOut + parseFloat(quantity);

    await updateCurrentStock(workbook, stockSheet, material);
    await saveWorkbook(workbook);




    res.send('✅ Stock Out recorded & Stock Sheet updated');
  } catch (err) {
    console.error(err);
    res.status(500).send('❌ Error writing Excel');
  }
});



// ✅ Fix getOrCreateSheet
function getOrCreateSheet(workbook, sheetName) {
  let sheet = workbook.getWorksheet(sheetName);
  if (!sheet) {
    if (sheetName.includes('Stock In')) {
      sheet = workbook.addWorksheet(sheetName);
      sheet.addRow(['Date', 'Material', 'Invoice No.', 'Quantity']);
    } else if (sheetName.includes('Stock Out')) {
      sheet = workbook.addWorksheet(sheetName);
      sheet.addRow(['Date', 'Material', 'Quantity', 'Remarks']);
    } else if (sheetName.includes('Stock')) {
      sheet = workbook.addWorksheet(sheetName);
      setupStockSheet(sheet, new Date().getFullYear(), sheetName.split(' ')[1]); // ✨ correctly setup Stock sheet
    }
  }
  return sheet;
}





// ✅ Update Current Stock in Stock April (Column C)
// ✅ Update Current Stock correctly based on monthly sheets
async function updateCurrentStock(workbook, stockSheet, material) {
  const materialRow = findOrCreateMaterialRow(stockSheet, material);

  const openingStock = parseFloat(materialRow.getCell(2).value) || 0; // Column B

  let totalIn = 0;
  let totalOut = 0;

  for (let col = 5; col <= stockSheet.columnCount; col += 3) {
    const inVal = parseFloat(materialRow.getCell(col).value) || 0;
    const outVal = parseFloat(materialRow.getCell(col + 1).value) || 0;

    totalIn += inVal;
    totalOut += outVal;
  }

  const currentStock = openingStock + totalIn - totalOut;
  materialRow.getCell(3).value = currentStock; // Column C (Current Stock)

  updateMinStockAndHighlight(stockSheet, materialRow); 

  await saveWorkbook(workbook);



}

// 🧠 Helper to update Min Stock (D) and Red Color
function updateMinStockAndHighlight(sheet, row) {
  const openingStock = parseFloat(row.getCell(2).value) || 0;
  const closingStock = parseFloat(row.getCell(3).value) || 0;

  const minStock = +(openingStock * 0.10).toFixed(2);
  row.getCell(4).value = minStock; // Column D

  if (closingStock <= minStock) {
    const redFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF0000' } };
    row.getCell(1).fill = redFill;
    row.getCell(4).fill = redFill;
  } else {
    row.getCell(1).fill = {};
    row.getCell(4).fill = {};
  }
}


// ✅ Update Min Stock (10% of Opening) and highlight red if needed
function updateMinStockAndHighlight(sheet, row) {
  const openingStock = parseFloat(row.getCell(2).value) || 0;  // Column B
  const closingStock = parseFloat(row.getCell(3).value) || 0;  // Column C

  const minStock = +(openingStock * 0.10).toFixed(2);
  row.getCell(4).value = minStock; // Column D

  const redFill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFF0000' }
  };

  const clearFill = {
    type: 'pattern',
    pattern: 'none'
  };

  if (closingStock <= minStock) {
    row.getCell(1).fill = redFill;  // Column A (material)
    row.getCell(4).fill = redFill;  // Column D (Min stock)
  } else {
    row.getCell(1).fill = clearFill;
    row.getCell(4).fill = clearFill;
  }
}

// ✅ Setup Stock Sheet structure with proper A-D columns
function setupStockSheet(sheet, year, month) {
  // Row 1: Headings
  sheet.getCell('A1').value = 'Material';
  sheet.getCell('B1').value = 'Opening Stock';
  sheet.getCell('C1').value = 'Current Stock';
  sheet.getCell('D1').value = 'Min Stock';

  // Row 2: Leave A-D cells blank for now (only E onwards will have In/Out/Remarks)
  sheet.getCell('A2').value = '';
  sheet.getCell('B2').value = '';
  sheet.getCell('C2').value = '';
  sheet.getCell('D2').value = '';

  // Calculate how many days in the month
  let daysInMonth = monthDays[parseInt(month)];
  if (parseInt(month) === 2 && isLeapYear(year)) {
    daysInMonth = 29;
  }

  let startCol = 5; // Column E starts after A-D

  for (let day = 1; day <= daysInMonth; day++) {
    const dateStr = `${day.toString().padStart(2, '0')}-${month}-${year}`;

    // Merge 3 columns (In/Out/Remarks under date)
    sheet.mergeCells(1, startCol, 1, startCol + 2);
    sheet.getCell(1, startCol).value = dateStr;
    sheet.getCell(1, startCol).alignment = { vertical: 'middle', horizontal: 'center' };

    // Row 2 under each merged date columns
    sheet.getCell(2, startCol).value = 'In';
    sheet.getCell(2, startCol + 1).value = 'Out';
    sheet.getCell(2, startCol + 2).value = 'Remarks';

    startCol += 3;
  }
}


//find date column
function findDateColumns(sheet, targetDate) {
  const headerRow1 = sheet.getRow(1);
  const formattedTarget = `${String(new Date(targetDate).getDate()).padStart(2, '0')}-${String(new Date(targetDate).getMonth() + 1).padStart(2, '0')}-${new Date(targetDate).getFullYear()}`;

  for (let col = 5; col <= sheet.columnCount; col += 3) {
    const val = headerRow1.getCell(col).value;
    if (val && typeof val === 'string' && val.trim() === formattedTarget) {
      return {
        inCol: col,
        outCol: col + 1,
        remarksCol: col + 2
      };
    }
  }
  return null;
}



//find or create material row
function createMissingDates(sheet, targetDate) {
  const headerRow1 = sheet.getRow(1);
  const headerRow2 = sheet.getRow(2);

  const formatDate = (dateObj) =>
    `${String(dateObj.getDate()).padStart(2, '0')}-${String(dateObj.getMonth() + 1).padStart(2, '0')}-${dateObj.getFullYear()}`;

  const startCol = 5;
  const existingDates = new Set();

  for (let col = startCol; col <= sheet.columnCount; col += 3) {
    const val = headerRow1.getCell(col).value;
    if (val && typeof val === 'string' && !isNaN(new Date(val))) {
      existingDates.add(val);
    }
  }

  const target = new Date(targetDate);
  const targetMonth = target.getMonth();
  const targetYear = target.getFullYear();
  const daysInMonth = new Date(targetYear, targetMonth + 1, 0).getDate();

  let current = new Date(targetYear, targetMonth, 1);
  let insertCol = sheet.columnCount + 1;

  for (let day = 1; day <= daysInMonth; day++) {
    const dateStr = formatDate(current);
    if (!existingDates.has(dateStr)) {
      sheet.mergeCells(1, insertCol, 1, insertCol + 2);
      sheet.getCell(1, insertCol).value = dateStr;
      sheet.getCell(1, insertCol).alignment = { vertical: 'middle', horizontal: 'center' };

      headerRow2.getCell(insertCol).value = 'In';
      headerRow2.getCell(insertCol + 1).value = 'Out';
      headerRow2.getCell(insertCol + 2).value = 'Remarks';

      insertCol += 3;
    }
    current.setDate(current.getDate() + 1);
  }
}


//dates manually
const monthDays = {
  1: 31,
  2: 28, // We'll adjust for leap year separately
  3: 31,
  4: 30,
  5: 31,
  6: 30,
  7: 31,
  8: 31,
  9: 30,
  10: 31,
  11: 30,
  12: 31
};

// ✅ Helper to check leap year
function isLeapYear(year) {
  return (year % 4 === 0 && year % 100 !== 0) || (year % 400 === 0);
}

// ✅ Initialize monthly sheets if missing
async function initializeMonthlySheets(workbook, targetDate) {
  const month = ('0' + (targetDate.getMonth() + 1)).slice(-2); // "04"
  const year = targetDate.getFullYear();

  const stockSheetName = `Stock ${month}`;
  const stockInSheetName = `Stock In ${month}`;
  const stockOutSheetName = `Stock Out ${month}`;

  // Check if the sheets exist already
  let stockSheet = workbook.getWorksheet(stockSheetName);
  let stockInSheet = workbook.getWorksheet(stockInSheetName);
  let stockOutSheet = workbook.getWorksheet(stockOutSheetName);

  // If any missing, create
  if (!stockSheet) {
    stockSheet = workbook.addWorksheet(stockSheetName);
    setupStockSheet(stockSheet, year, month);
  }
  if (!stockInSheet) {
    stockInSheet = workbook.addWorksheet(stockInSheetName);
    stockInSheet.addRow(["Date", "Material", "Invoice No.", "Quantity"]);
  }
  if (!stockOutSheet) {
    stockOutSheet = workbook.addWorksheet(stockOutSheetName);
    stockOutSheet.addRow(["Date", "Material", "Quantity", "Remarks"]);
  }
}

// ✅ Setup Stock Sheet structure with proper A-D columns and merged headings
// ✅ Setup Stock Sheet structure with styles applied
function setupStockSheet(sheet, year, month) {
  // Row 1: A-D headers
  const headers = ['Material', 'Opening Stock', 'Current Stock', 'Min Stock'];
  const blueGreyFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D9D9D9' } }; // Blue-grey, Text2, lighter 80%
  for (let i = 0; i < headers.length; i++) {
    const cell = sheet.getCell(1, i + 1);
    cell.value = headers[i];
    cell.font = { bold: true };
    cell.fill = blueGreyFill;
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    sheet.mergeCells(1, i + 1, 2, i + 1); // Merge A1:A2, B1:B2, etc.
  }

  // Calculate days
  let daysInMonth = monthDays[parseInt(month)];
  if (parseInt(month) === 2 && isLeapYear(year)) daysInMonth = 29;

  let col = 5; // Starting at Column E

  for (let day = 1; day <= daysInMonth; day++) {
    const dateStr = `${String(day).padStart(2, '0')}-${month}-${year}`;

    // Merge Date headers
    sheet.mergeCells(1, col, 1, col + 2);
    const headerCell = sheet.getCell(1, col);
    headerCell.value = dateStr;
    headerCell.font = { bold: true };
    headerCell.alignment = { vertical: 'middle', horizontal: 'center' };

    // Row 2: In / Out / Remarks
    const greenFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E2F0D9' } }; // Green Accent 6 60% lighter
    const orangeFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F9CB9C' } }; // Orange Accent 2 60% lighter
    const blueGreyFillLight = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D9D9D9' } }; // Same as A-D headers

    const inCell = sheet.getCell(2, col);
    const outCell = sheet.getCell(2, col + 1);
    const remarksCell = sheet.getCell(2, col + 2);

    inCell.value = 'In';
    inCell.font = { bold: true };
    inCell.fill = greenFill;
    inCell.alignment = { vertical: 'middle', horizontal: 'center' };

    outCell.value = 'Out';
    outCell.font = { bold: true };
    outCell.fill = orangeFill;
    outCell.alignment = { vertical: 'middle', horizontal: 'center' };

    remarksCell.value = 'Remarks';
    remarksCell.font = { bold: true };
    remarksCell.fill = blueGreyFillLight;
    remarksCell.alignment = { vertical: 'middle', horizontal: 'center' };

    col += 3;
  }

  // ✅ Optional: Set Column Widths for neatness
  const widths = [20, 15, 15, 15];
  for (let i = 0; i < widths.length; i++) {
    sheet.getColumn(i + 1).width = widths[i];
  }
}


//search current stock
app.get('/search-current-stock', async (req, res) => {
  const { material, month } = req.query;
  if (!material || !month) return res.status(400).json({ error: 'Material and Month required' });

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(FILE_PATH);

  const sheet = wb.getWorksheet(`Stock ${month}`);
  if (!sheet) return res.status(404).json({ error: `Stock ${month} Sheet not found` });

  for (let i = 3; i <= sheet.rowCount; i++) {
    const mat = sheet.getCell(`A${i}`).value;
    if (mat && mat.toString().toLowerCase() === material.toLowerCase()) {
      const currentStock = sheet.getCell(`C${i}`).value || 0; // Column C = Current Stock
      return res.json({ material: mat, currentStock });
    }
  }

  return res.status(404).json({ error: 'Material not found in selected month.' });
});

//search min. stock
app.get('/search-min-stock', async (req, res) => {
  const { month } = req.query;
  if (!month) return res.status(400).json({ error: 'Month required' });

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(FILE_PATH);

  const sheet = wb.getWorksheet(`Stock ${month}`);
  if (!sheet) return res.status(404).json({ error: `Stock ${month} Sheet not found` });

  const result = [];

  for (let i = 3; i <= sheet.rowCount; i++) {
    const material = sheet.getCell(`A${i}`).value;
    const currentStock = parseFloat(sheet.getCell(`C${i}`).value) || 0;
    const minStock = parseFloat(sheet.getCell(`D${i}`).value) || 0;

    if (material && currentStock <= minStock) {
      result.push({ material, currentStock });
    }
  }

  res.json(result);
});

//search stock by date
// ✅ NEW: Search stock by date (ALL Materials)
app.get('/search-stock-by-date', async (req, res) => {
  const { date } = req.query;
  if (!date) return res.status(400).json({ error: 'Date is required' });

  const month = ('0' + (new Date(date).getMonth() + 1)).slice(-2); // Extract month
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(FILE_PATH);

  const sheet = wb.getWorksheet(`Stock ${month}`);
  if (!sheet) return res.status(404).json({ error: `Stock ${month} Sheet not found` });

  // Find the date columns
  const headerRow = sheet.getRow(1);
  const formattedDate = `${String(new Date(date).getDate()).padStart(2, '0')}-${month}-${new Date(date).getFullYear()}`;

  let foundCol = null;
  for (let col = 5; col <= sheet.columnCount; col += 3) {
    if (headerRow.getCell(col).value === formattedDate) {
      foundCol = col;
      break;
    }
  }

  if (!foundCol) return res.status(404).json({ error: 'Date not found in Sheet.' });

  // Now collect all materials and their in/out
  const result = [];

  for (let i = 3; i <= sheet.rowCount; i++) {
    const material = sheet.getCell(`A${i}`).value;
    if (material) {
      const inQty = sheet.getCell(i, foundCol).value || 0;
      const outQty = sheet.getCell(i, foundCol + 1).value || 0;

      if (inQty !== 0 || outQty !== 0) { // Only show if there is any entry
        result.push({ material, in: inQty, out: outQty });
      }
    }
  }

  res.json(result);
});
// Totals: sales revenue, payment received, plants installed
app.get('/api/getDashboardStats', async (req, res) => {
  try {
    res.set('Cache-Control', 'no-store'); // No cache for live updates

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('C:/Users/JK SOLAR/OneDrive/CRM_PWA/TempData.xlsx');
    const sheet = workbook.getWorksheet('Client Data');

    let totalSalesRevenue = 0;
    let totalBalance = 0;
    let plantsInstalled = 0;

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // Skip header row

      const totalCost = parseFloat(row.getCell(8).value) || 0;    // Total Cost (H)
      const balanceRaw = row.getCell(35).value;                   // Balance (AI)

      // Always add total cost
      totalSalesRevenue += totalCost;

      let balance = null;
      if (balanceRaw !== null && balanceRaw !== '' && balanceRaw !== '-' && balanceRaw !== '.') {
        if (typeof balanceRaw === 'string') {
          balance = parseFloat(balanceRaw.trim()) || 0;
        } else if (typeof balanceRaw === 'number') {
          balance = balanceRaw;
        }

        // Add balance to total balance sum
        totalBalance += balance;

        // 🛡️ Plants Installed: Only if balance === 0 exactly
        if (balance === 0) {
          plantsInstalled += 1;
        }
      }
    });

    // ✨ Final Payment calculation
    const totalPaymentReceived = totalSalesRevenue - totalBalance;

    res.json({ totalSalesRevenue, totalPaymentReceived, plantsInstalled });
  } catch (error) {
    console.error('❌ Error in getDashboardStats:', error);
    res.status(500).json({ error: 'Internal Server Error' });
  }
});




// 2️⃣ Pie-chart data
app.get('/api/getPieData', async (req, res) => {
  try {
    res.set('Cache-Control', 'no-store'); // No cache for live updates

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('C:/Users/JK SOLAR/OneDrive/CRM_PWA/TempData.xlsx');
    const sheet = workbook.getWorksheet('Client Data');

    let totalAmount = 0; // Sum of Total Cost
    let totalBalance = 0; // Sum of Balance

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // Skip header row

      const totalCost = parseFloat(row.getCell(8).value) || 0;    // Total Cost (H)
      const balanceRaw = row.getCell(35).value;                   // Balance (AI)

      totalAmount += totalCost; // Sum all Total Cost

      let balance = 0;
      if (balanceRaw !== null && balanceRaw !== '' && balanceRaw !== '-' && balanceRaw !== '.') {
        if (typeof balanceRaw === 'string') {
          balance = parseFloat(balanceRaw.trim()) || 0;
        } else if (typeof balanceRaw === 'number') {
          balance = balanceRaw;
        }
      }

      totalBalance += balance; // Sum all valid balances
    });

    res.json({ totalAmount, totalBalance });
  } catch (error) {
    console.error('❌ Error in getPieData:', error);
    res.status(500).json({ error: 'Internal Server Error' });
  }
});

app.get('/api/getApplicationStatusData', async (req, res) => {
  try {
    res.set('Cache-Control', 'no-store'); // Disable cache

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('C:/Users/JK SOLAR/OneDrive/CRM_PWA/TempData.xlsx');
    const sheet = workbook.getWorksheet('Client Data');

    let applicationApplied = 0;
    let applicationPending = 0;

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // Skip header

      const applicationCell = row.getCell(19).value; // Column S is 19th column

      if (applicationCell !== null && applicationCell !== '' && applicationCell.toString().toLowerCase() !== 'no') {
        applicationApplied += 1;
      } else {
        applicationPending += 1;
      }
    });

    res.json({ applied: applicationApplied, pending: applicationPending });
  } catch (error) {
    console.error('❌ Error in getApplicationStatusData:', error);
    res.status(500).json({ error: 'Internal Server Error' });
  }
});

app.get('/api/getBarGraphData', async (req, res) => {
  try {
    res.set('Cache-Control', 'no-store');

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('C:/Users/JK SOLAR/OneDrive/CRM_PWA/TempData.xlsx');
    const sheet = workbook.getWorksheet('Client Data');

    let totalCostSum = 0;
    let advanceSum = 0;
    let secondInstallmentReceivedSum = 0;
    let finalInstallmentReceivedSum = 0;

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // skip header

      const totalCost = parseFloat(row.getCell(8).value) || 0;    // H
      const advance = parseFloat(row.getCell(7).value) || 0;      // G
      const secondInstallment = parseFloat(row.getCell(32).value) || 0; // AF
      const finalInstallment = parseFloat(row.getCell(34).value) || 0; // AH

      totalCostSum += totalCost;
      advanceSum += advance;
      secondInstallmentReceivedSum += secondInstallment;
      finalInstallmentReceivedSum += finalInstallment;
    });

    // Now calculate according to your final logic
    const sixtyPercentOfTotalCost = 0.6 * totalCostSum;

    const secondInstallmentDue = totalCostSum - (advanceSum + sixtyPercentOfTotalCost);
    const finalInstallmentDue = totalCostSum - (advanceSum + secondInstallmentReceivedSum);

    res.json({
      totalCostSum,
      advanceSum,
      secondInstallmentReceivedSum,
      secondInstallmentDue,
      finalInstallmentReceivedSum,
      finalInstallmentDue
    });
  } catch (error) {
    console.error('❌ Error in getBarGraphData:', error);
    res.status(500).json({ error: 'Internal Server Error' });
  }
});

app.get('/api/getPaymentsData', async (req, res) => {
  try {
    res.set('Cache-Control', 'no-store');

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('C:/Users/JK SOLAR/OneDrive/CRM_PWA/TempData.xlsx');
    const sheet = workbook.getWorksheet('Client Data');

    const payments = [];

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // Skip header

      const customerName = row.getCell(2).value || '';
      const totalCost = parseFloat(row.getCell(8).value) || 0;
      const advance = parseFloat(row.getCell(7).value) || 0;
      const secondInstallment = row.getCell(32).value;
      const finalInstallment = row.getCell(34).value;
      const balance = parseFloat(row.getCell(35).value) || 0;

      // Only include rows where balance is not zero
      if (balance !== 0) {
        payments.push({
          customerName,
          totalCost,
          advance,
          secondInstallment: (secondInstallment === null || secondInstallment === '' || secondInstallment === undefined) ? 'Due' : secondInstallment,
          finalInstallment: (finalInstallment === null || finalInstallment === '' || finalInstallment === undefined) ? 'Due' : finalInstallment,
          balance
        });
      }
    });
    
    res.json({ payments });
  } catch (error) {
    console.error('❌ Error fetching Payments Data:', error);
    res.status(500).json({ error: 'Internal Server Error' });
  }
});


//dateYes bro
app.post('/api/addTask', async (req, res) => {
  try {
      const { date, time, description } = req.body;
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile('TempData.xlsx');
      const sheet = workbook.getWorksheet('Client Data');

      // Find first empty row in AK (column 37)
      let rowToUse;
      sheet.eachRow((row, rowNumber) => {
          if (!row.getCell(37).value && !rowToUse) {
              rowToUse = rowNumber;
          }
      });

      if (!rowToUse) {
          rowToUse = sheet.lastRow.number + 1;
      }

      sheet.getRow(rowToUse).getCell(37).value = date;         // AK
      sheet.getRow(rowToUse).getCell(38).value = time;         // AL
      sheet.getRow(rowToUse).getCell(39).value = description;  // AM

      await workbook.xlsx.writeFile('TempData.xlsx');
      res.json({ success: true });
  } catch (error) {
      console.error('❌ Error adding task:', error);
      res.status(500).json({ error: 'Failed to add task' });
  }
});

app.post('/api/addTask', async (req, res) => {
  const { date, time, description } = req.body;
  const workbook = new Excel.Workbook();
  await workbook.xlsx.readFile('TempData.xlsx');
  const worksheet = workbook.getWorksheet('Client Data');

  const nextRow = worksheet.lastRow.number + 1;

  worksheet.getCell(`AK${nextRow}`).value = date;
  worksheet.getCell(`AL${nextRow}`).value = time;
  worksheet.getCell(`AM${nextRow}`).value = description;

  await workbook.xlsx.writeFile('TempData.xlsx');
  res.json({ success: true });
});

app.get('/api/getTasks', async (req, res) => {
  const workbook = new Excel.Workbook();
  await workbook.xlsx.readFile('TempData.xlsx');
  const worksheet = workbook.getWorksheet('Client Data');

  const tasks = [];
  worksheet.eachRow((row, rowNumber) => {
    const date = row.getCell(37).value;
    const time = row.getCell(38).value;
    const description = row.getCell(39).value;

    if (date && time && description) {
      tasks.push({ date, time, description }); // ✅ use 'description' key to match frontend expectation
    }
  });

  res.json({ tasks });
});


app.post('/api/addReminder', async (req, res) => {
  try {
      const { date, text } = req.body;
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile('TempData.xlsx');
      const sheet = workbook.getWorksheet('Client Data');

      // Find first empty row in AN (column 40)
      let rowToUse;
      sheet.eachRow((row, rowNumber) => {
          if (!row.getCell(40).value && !rowToUse) {
              rowToUse = rowNumber;
          }
      });

      if (!rowToUse) {
          rowToUse = sheet.lastRow.number + 1;
      }

      sheet.getRow(rowToUse).getCell(40).value = date;     // AN
      sheet.getRow(rowToUse).getCell(41).value = text;     // AO

      await workbook.xlsx.writeFile('TempData.xlsx');
      res.json({ success: true });
  } catch (error) {
      console.error('❌ Error adding reminder:', error);
      res.status(500).json({ error: 'Failed to add reminder' });
  }
});

app.get('/api/getReminders', async (req, res) => {
  try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile('TempData.xlsx');
      const sheet = workbook.getWorksheet('Client Data');
      const reminders = [];

      sheet.eachRow((row, rowNumber) => {
          const date = row.getCell(40).value;
          const text = row.getCell(41).value;

          if (date && text) {
              reminders.push({ date, text });
          }
      });

      res.json({ reminders });
  } catch (error) {
      console.error('❌ Error fetching reminders:', error);
      res.status(500).json({ error: 'Failed to fetch reminders' });
  }
});

// Delete Task
app.post('/api/deleteTask', async (req, res) => {
  const { date, time, description } = req.body;

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('tempdata.xlsx');

  const sheet = workbook.getWorksheet('Client Data'); // ✅ Use correct sheet name
  if (!sheet) {
    return res.status(404).json({ success: false, message: 'Sheet not found' });
  }

  let rowToDelete = null;

  for (let i = 2; i <= sheet.rowCount; i++) {
    const row = sheet.getRow(i);
    const d = row.getCell(37).value?.toString().trim(); // AK
    const t = row.getCell(38).value?.toString().trim(); // AL
    const desc = row.getCell(39).value?.toString().trim(); // AM

    if (d === date && t === time && desc === description) {
      rowToDelete = i;
      break;
    }
  }

  if (rowToDelete) {
    sheet.spliceRows(rowToDelete, 1);
    await workbook.xlsx.writeFile('tempdata.xlsx');
    res.json({ success: true });
  } else {
    res.json({ success: false, message: 'Task not found' });
  }
});




// Delete Reminder
app.post('/api/deleteReminder', async (req, res) => {
  try {
    const { date, text } = req.body;
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('TempData.xlsx');
    const sheet = workbook.getWorksheet('Client Data');

    sheet.eachRow((row) => {
      const rowDate = row.getCell(40).value;
      const rowText = row.getCell(41).value;

      if (rowDate === date && rowText === text) {
        row.getCell(40).value = null;
        row.getCell(41).value = null;
      }
    });

    await workbook.xlsx.writeFile('TempData.xlsx');
    res.json({ success: true });
  } catch (error) {
    console.error('❌ Error deleting reminder:', error);
    res.status(500).json({ error: 'Failed to delete reminder' });
  }
});

//JK SOLAR LEADSssss

app.get('/get-next-refno', async (req, res) => {
  const filePath = path.join(__dirname, 'leads.xlsx');
  const workbook = new ExcelJS.Workbook();

  try {
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.getWorksheet(1);

    let lastRef = null;

    // Loop through all rows to find the last reference number
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // skip header
      const refCell = row.getCell(5).value;
      if (refCell && typeof refCell === 'string' && /^[A-Z]\d{4}$/.test(refCell)) {
        lastRef = refCell;
      }
    });

    let nextRef = 'A0001';

    if (lastRef) {
      const letter = lastRef.charAt(0);
      const number = parseInt(lastRef.slice(1));
      if (number < 9999) {
        nextRef = letter + (number + 1).toString().padStart(4, '0');
      } else {
        const nextChar = String.fromCharCode(letter.charCodeAt(0) + 1);
        nextRef = nextChar + '0001';
      }
    }

    res.send(nextRef);
  } catch (err) {
    console.error("Error generating next ref no:", err);
    res.status(500).send('Error');
  }
});


app.post('/save-lead', async (req, res) => {
  const filePath = path.join(__dirname, 'leads.xlsx');
  const { date, name, address, mobile, refno, kw, reference } = req.body;
  const workbook = new ExcelJS.Workbook();

  try {
    // Create file if not exists
    if (!fs.existsSync(filePath)) {
      const newSheet = workbook.addWorksheet('Leads');
      newSheet.addRow(['Date', 'Consumer Name', 'Address', 'Mobile No.', 'Ref No.', 'KW', 'Reference']);
      await workbook.xlsx.writeFile(filePath);
    }

    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.getWorksheet(1); // 'Leads'

    sheet.addRow([date, name, address, mobile, refno, kw, reference]);

    await workbook.xlsx.writeFile(filePath);
    res.sendStatus(200);
  } catch (err) {
    console.error('Error saving lead:', err);
    res.status(500).send('Failed to save');
  }
});

app.get('/get-leads', async (req, res) => {
  const filePath = path.join(__dirname, 'leads.xlsx');
  const workbook = new ExcelJS.Workbook();

  try {
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.getWorksheet(1);

    const data = [];
    sheet.eachRow((row, rowNumber) => {
      const rowData = row.values.slice(1); // drop first blank
      if (rowNumber === 1 || rowData[7]?.toString().toLowerCase() !== 'no') {
        data.push(rowData);
      }
    });

    res.json(data);
  } catch (err) {
    console.error('Error reading leads:', err);
    res.status(500).send('Failed to read');
  }
});


app.post('/update-lead', async (req, res) => {
  const filePath = path.join(__dirname, 'leads.xlsx');
  const { field, rowIndex, value } = req.body;

  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.getWorksheet(1);

    const fieldMap = {
      call: 8,
      proposal: 9,
      meeting: 10,
      reminder: 11,
      status: 12,
      final: 13
    };

    const row = sheet.getRow(rowIndex + 2); // +2 for 0-index + header row
    const colIndex = fieldMap[field];

    if (row && colIndex) {
      row.getCell(colIndex).value = value;
      row.commit();
      await workbook.xlsx.writeFile(filePath);
      res.sendStatus(200);
    } else {
      res.status(400).send("Invalid row/column");
    }
  } catch (err) {
    console.error("Error updating lead:", err);
    res.status(500).send("Update failed");
  }
});

app.post('/delete-lead', async (req, res) => {
  const filePath = path.join(__dirname, 'leads.xlsx');
  const { rowIndex } = req.body;

  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.getWorksheet(1);

    const actualRow = rowIndex + 2; // Because headers + 0-indexed
    
    await workbook.xlsx.writeFile(filePath);
    res.sendStatus(200);
  } catch (err) {
    console.error("Error deleting lead:", err);
    res.status(500).send("Delete failed");
  }
});

//neww proposall
const proposalExcelPath = path.join(__dirname, 'proposal.xlsx');

app.post('/save-proposal', async (req, res) => {
  const data = req.body;
  const filePath = proposalExcelPath;

  try {
    const workbook = new ExcelJS.Workbook();

    if (fs.existsSync(filePath)) {
      await workbook.xlsx.readFile(filePath);
    } else {
      const ws = workbook.addWorksheet('Proposals');
      ws.addRow(['Ref', 'Date', 'Subsidy', 'KW', 'Address', 'State', 'City', 'To Whom', 'Mobile', 'Price', 'Panel Brand', 'Panel Wp', 'Inverter Brand']);
    }

    const ws = workbook.getWorksheet('Proposals');

    // Auto-generate reference number like 01P, 02P
    const lastRow = ws.lastRow;
    let newNumber = 1;
    if (lastRow && lastRow.getCell(1).value && lastRow.getCell(1).value.toString().endsWith('P')) {
      const lastRef = lastRow.getCell(1).value.toString().replace('P', '');
      newNumber = parseInt(lastRef) + 1;
    }
    const newRef = `${String(newNumber).padStart(2, '0')}P`;

    ws.addRow([
      newRef,
      data.date,
      data.subsidy,
      data.kw,
      data.address,
      data.state,
      data.city,
      data.toWhom,
      data.mobile,
      data.price,
      data.panelBrand,
      data.panelWp,
      data.inverterBrand
    ]);

    await workbook.xlsx.writeFile(filePath);
    res.json({ success: true, ref: newRef });

  } catch (err) {
    console.error('❌ Proposal save error:', err);
    res.status(500).json({ success: false, error: 'Proposal saving failed' });
  }
});

//Proposal
app.post('/submit-proposal', upload.none(), (req, res) => {
  const filePath = 'C:/Users/JK SOLAR/OneDrive/CRM_PWA/proposal.xlsx';

  let workbook, worksheet, data;
  if (fs.existsSync(filePath)) {
    workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];

    if (!sheetName) {
      worksheet = xlsx.utils.json_to_sheet([]);
      xlsx.utils.book_append_sheet(workbook, worksheet, 'Proposals');
      data = [];
    } else {
      worksheet = workbook.Sheets[sheetName];
      data = xlsx.utils.sheet_to_json(worksheet);
    }
  } else {
    workbook = xlsx.utils.book_new();
    worksheet = xlsx.utils.json_to_sheet([]);
    xlsx.utils.book_append_sheet(workbook, worksheet, 'Proposals');
    data = [];
  }

  const maxRef = data.reduce((max, row) => {
    const ref = row.Ref;
    if (typeof ref === 'string' && ref.endsWith('P')) {
      const num = parseInt(ref.replace('P', ''));
      return isNaN(num) ? max : Math.max(max, num);
    }
    return max;
  }, 0);

  const newRef = (maxRef + 1).toString().padStart(2, '0') + 'P';

  const newRow = {
    Ref: newRef,
    Date: req.body.date,
    Subsidy: req.body.subsidy,
    KW: req.body.kw,
    Address: req.body.address,
    State: req.body.state,
    City: req.body.city,
    "To Whom": req.body.toWhom,
    "Mobile No": req.body.mobile,
    Price: req.body.price,
    "Panel Brand": req.body.panelBrand,
    "Panel Wp": req.body.panelWp,
    "Inverter Brand": req.body.inverterBrand
  };

  data.push(newRow);
  const newSheet = xlsx.utils.json_to_sheet(data, {
    header: [
      'Ref', 'Date', 'Subsidy', 'KW', 'Address', 'State', 'City', 'To Whom',
      'Mobile No', 'Price', 'Panel Brand', 'Panel Wp', 'Inverter Brand'
    ]
  });

  const newWb = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(newWb, newSheet, 'Proposals');
  xlsx.writeFile(newWb, filePath);
  console.log("➡️ Proposal submission triggered with data:", req.body);
  res.json({ success: true, ref: newRef });
});


//Proposal fetching

app.post('/submit-proposal', upload.none(), (req, res) => {
  try {
    const filePath = path.join(__dirname, 'TempData.xlsx');

    let workbook, sheet, data;
    if (fs.existsSync(filePath)) {
      workbook = xlsx.readFile(filePath);
      sheet = workbook.Sheets[workbook.SheetNames[0]];
      data = xlsx.utils.sheet_to_json(sheet);
    } else {
      workbook = xlsx.utils.book_new();
      data = [];
    }

    // Generate next ref like 01P, 02P
    const maxNum = data.reduce((max, row) => {
      const match = typeof row.Ref === 'string' && row.Ref.match(/^(\d+)P$/);
      if (match) {
        const num = parseInt(match[1]);
        return Math.max(max, num);
      }
      return max;
    }, 0);
    const nextRef = `${(maxNum + 1).toString().padStart(2, '0')}P`;

    // Create new row
    const newRow = {
      Ref: nextRef,
      Date: new Date().toLocaleDateString('en-IN'),
      To: req.body.to,
      Mobile: req.body.mobile,
      KW: req.body.kw,
      Price: req.body.price,
      WP: req.body.wp,
      PanelBrand: req.body.panelBrand,
      InverterBrand: req.body.inverterBrand
    };

    // Add to data and save
    data.push(newRow);
    const newSheet = xlsx.utils.json_to_sheet(data, {
      header: ['Ref', 'Date', 'To', 'Mobile', 'KW', 'Price', 'WP', 'PanelBrand', 'InverterBrand']
    });
    xlsx.utils.book_append_sheet(workbook, newSheet, 'Proposals');
    xlsx.writeFile(workbook, filePath);

    res.json({ success: true, ref: nextRef });
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, error: err.message });
  }
});



// Route: Get Proposal by Ref

app.get('/get-proposal', async (req, res) => {
  const ref = req.query.ref;
  const filePath = path.join(__dirname, 'proposal.xlsx');

  if (!ref) return res.status(400).json({ success: false, error: 'Missing reference number' });

  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.getWorksheet('Proposals');

    let matchedRow;
    sheet.eachRow((row, rowNumber) => {
      if (row.getCell(1).value === ref) {
        matchedRow = row;
      }
    });

    if (!matchedRow) return res.status(404).json({ success: false, error: 'Proposal not found' });

    const headers = sheet.getRow(1).values.slice(1); // Remove empty index 0
    const values = matchedRow.values.slice(1);
    const proposalData = {};

    headers.forEach((header, i) => {
      proposalData[header.trim().toLowerCase().replace(/\s+/g, '')] = values[i];
    });

    res.json({ success: true, data: proposalData });

  } catch (err) {
    console.error('❌ Error reading proposal.xlsx:', err);
    res.status(500).json({ success: false, error: 'Error reading Excel' });
  }
});

const puppeteer = require('puppeteer');

app.get('/generate-pdf', async (req, res) => {
  const ref = req.query.ref;
  const previewURL = `http://localhost:3000/proposal-preview.html?ref=${ref}`;

  try {
    const browser = await puppeteer.launch({ headless: 'new' });
    const page = await browser.newPage();
    await page.goto(previewURL, { waitUntil: 'networkidle0' });

    const pdfBuffer = await page.pdf({
      format: 'A4',
      printBackground: true,
      margin: { top: '10mm', bottom: '10mm', left: '10mm', right: '10mm' }
    });

    await browser.close();

    res.set({
      'Content-Type': 'application/pdf',
      'Content-Disposition': `attachment; filename="JK_Solar_Proposal_${ref}.pdf"`
    });

    res.send(pdfBuffer);

  } catch (err) {
    console.error('PDF Generation Error:', err);
    res.status(500).send('Failed to generate PDF');
  }
});

app.get('/get-proposals', (req, res) => {
  const filePath = 'C:/Users/JK SOLAR/OneDrive/CRM_PWA/proposal.xlsx';

  try {
    if (!fs.existsSync(filePath)) {
      return res.json([]);
    }

    const workbook = xlsx.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = xlsx.utils.sheet_to_json(sheet);
    res.json(data);
  } catch (err) {
    console.error('Error reading proposal.xlsx:', err);
    res.status(500).json({ error: 'Failed to read proposal data.' });
  }
});



const proposalPath = path.join(__dirname, 'proposal.xlsx');
const leadsPath = path.join(__dirname, 'leads.xlsx');

async function transferProposalToLeads() {
  const proposalWorkbook = new ExcelJS.Workbook();
  await proposalWorkbook.xlsx.readFile(proposalPath);
  const proposalSheet = proposalWorkbook.worksheets[0];

  const leadsWorkbook = new ExcelJS.Workbook();
  await leadsWorkbook.xlsx.readFile(leadsPath);
  const leadsSheet = leadsWorkbook.worksheets[0];

  for (let i = 2; i <= proposalSheet.rowCount; i++) {
    const rowProposal = proposalSheet.getRow(i);
    const proposalKW = String(rowProposal.getCell('D').value).trim();    // KW
    const proposalName = String(rowProposal.getCell('H').value).trim();  // Name
    const proposalMobile = String(rowProposal.getCell('I').value).trim(); // Mobile

    let replaced = false;

    for (let j = 2; j <= leadsSheet.rowCount; j++) {
      const rowLead = leadsSheet.getRow(j);
      const leadKW = String(rowLead.getCell(6).value).trim();     // F
      const leadName = String(rowLead.getCell(2).value).trim();   // B
      const leadMobile = String(rowLead.getCell(4).value).trim(); // D

      if (proposalKW === leadKW && proposalName === leadName && proposalMobile === leadMobile) {
        leadsSheet.spliceRows(j, 1); // delete old row
        leadsSheet.insertRow(j, [
          rowProposal.getCell('B').value, // Consumer Name → A
          rowProposal.getCell('H').value, // Address       → B
          rowProposal.getCell('E').value, // State         → C
          rowProposal.getCell('I').value, // Mobile No     → D
          rowProposal.getCell('A').value, // Ref No        → E
          rowProposal.getCell('D').value, // Date          → F
          '',                              // Reference     → G
          '',                              // H
  'Sent'                           // I ← NEW!
        ]);
        console.log(`🔁 Replaced row ${j} for ${proposalName}, ${proposalMobile}, ${proposalKW}`);
        replaced = true;
        break;
      }
    }

    if (!replaced) {
      leadsSheet.addRow([
        rowProposal.getCell('B').value, // Consumer Name → A
        rowProposal.getCell('H').value, // Address       → B
        rowProposal.getCell('E').value, // State         → C
        rowProposal.getCell('I').value, // Mobile No     → D
        rowProposal.getCell('A').value, // Ref No        → E
        rowProposal.getCell('D').value, // Date          → F
        '',                              // Reference     → G
       
      ]);
      console.log(`🆕 Added new row for ${proposalName}, ${proposalMobile}, ${proposalKW}`);
    }
  }

  await leadsWorkbook.xlsx.writeFile(leadsPath);
  console.log("✅ Leads file updated successfully.");
}

// Initial run
transferProposalToLeads().catch(console.error);

// Auto-watch for changes
fs.watchFile(proposalPath, { interval: 2000 }, (curr, prev) => {
  if (curr.mtime !== prev.mtime) {
    console.log("📄 Proposal file changed. Updating leads...");
    transferProposalToLeads().catch(console.error);
  }
});


const notesFilePath = "C:\\Users\\JK SOLAR\\OneDrive\\CRM_PWA\\TempData.xlsx";

// 🔹 GET notes
app.get('/api/getNotes', async (req, res) => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile("C:\\Users\\JK SOLAR\\OneDrive\\CRM_PWA\\TempData.xlsx");
  const sheet = workbook.getWorksheet('Client Data');
  const notes = [];

  // START from row 2 to skip header
  for (let i = 2; i <= sheet.rowCount; i++) {
    const note = sheet.getRow(i).getCell(38).value; // Column AL
    if (note) {
      notes.push(note.toString());
    }
  }

  res.json({ notes });
});


// 🔹 ADD note
app.post('/api/addNote', async (req, res) => {
  const { note } = req.body;
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(notesFilePath);
  const sheet = workbook.getWorksheet('Client Data');

  let rowToWrite = sheet.actualRowCount + 1;

  // Find the next empty row in column AL
  for (let i = 2; i <= sheet.rowCount + 100; i++) {
    const cell = sheet.getRow(i).getCell(38).value;
    if (!cell || cell === '') {
      rowToWrite = i;
      break;
    }
  }

  sheet.getRow(rowToWrite).getCell(38).value = note;
  await workbook.xlsx.writeFile(notesFilePath);

  res.sendStatus(200);
});



// 🔹 DELETE note by index
app.post('/api/deleteNote', async (req, res) => {
  const { index } = req.body;
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(notesFilePath);
  const sheet = workbook.getWorksheet('Client Data');

  let rowIndex = 0, count = 0;

  sheet.eachRow((row, i) => {
    if (row.getCell(38).value) {
      if (count === index) rowIndex = i;
      count++;
    }
  });

  if (rowIndex > 0) {
    sheet.getRow(rowIndex).getCell(38).value = null;
    await workbook.xlsx.writeFile(notesFilePath);
  }

  res.sendStatus(200);
});


// STEP 1: Redirect to Microsoft Login
app.get('/login', (req, res) => {
  const url = `https://login.microsoftonline.com/785fd7e9-594d-4549-91b9-9372f7295962/oauth2/v2.0/authorize?client_id=89a49313-0f16-44c3-9f71-cf96eab166ad&response_type=code&redirect_uri=https://www.jksolarpower.com/auth/callback&response_mode=query&scope=offline_access%20Files.ReadWrite%20User.Read`;
  
  console.log("🔗 Redirecting to Microsoft with URL:", url); // Debug print
  res.redirect(url);

});

// STEP 2: Microsoft redirects back here after login
app.get('/auth/callback', async (req, res) => {
  const code = req.query.code;

  try {
    const response = await axios.post(
      `https://login.microsoftonline.com/785fd7e9-594d-4549-91b9-9372f7295962/oauth2/v2.0/token`,
      qs.stringify({
        client_id: '89a49313-0f16-44c3-9f71-cf96eab166ad',
        scope: 'offline_access Files.ReadWrite User.Read',
        code,
        redirect_uri: 'https://www.jksolarpower.com/auth/callback',
        grant_type: 'authorization_code',
        client_secret: 'a0958e75-6be9-45cd-a1e9-e0a436769602',
      }),
      {
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
        },
      }
    );

    const { access_token, refresh_token } = response.data;
    console.log("✅ Access Token:", access_token);
    console.log("♻️ Refresh Token:", refresh_token);

    // For now, just show a success page
    res.send('Login successful! Token received. Check console.');
  } catch (err) {
    console.error('❌ Token error:', err.response.data);
    res.status(500).send('Failed to get access token');
  }
});



// 🚀 Start server
app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);

});
