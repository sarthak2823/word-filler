// Import required libraries
const express = require('express');
const multer = require('multer'); // for handling file uploads
const XLSX = require('xlsx'); // for working with Excel files
const Docxtemplater = require('docxtemplater'); // for filling in Word document templates
const fs = require('fs'); // for working with files

// Create a new Express application
const app = express();

// Create a new Multer instance for handling file uploads
const upload = multer({
  storage: multer.diskStorage({
    destination: (req, file, cb) => {
      cb(null, "uploads");
    },
    filename: (req, file, cb) => {
      cb(null, file.originalname);
    },
  }),
  fileFilter: (req, file, cb) => {
    if (
      file.mimetype !== "application/vnd.openxmlformats-officedocument.wordprocessingml.document" &&
      file.mimetype !== "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ) {
      cb(new Error("Invalid file type."));
    } else {
      cb(null, true);
    }
  },
});


// Set up a route to handle the form submission
app.post('/fill', upload.fields([{ name: 'wordDoc', maxCount: 1 }, { name: 'excelFile', maxCount: 1 }]), (req, res) => {
  // Get the Word document and Excel file streams from the request
  const wordDoc = req.files.wordDoc[0];
  const excelFile = req.files.excelFile[0];
  console.log(wordDoc,excelFile);
  let fileReader = new FileReader();
  let excelData;
        fileReader.readAsBinaryString(excelFile);
        fileReader.onload = (event)=>{
         let data = event.target.result;
         let workbook = XLSX.read(data,{type:"binary"});
         console.log(workbook);
         workbook.SheetNames.forEach(sheet => {
              let rowObject = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheet]);
              console.log(rowObject);
              excelData= JSON.stringify(rowObject,undefined,4);
              console.log(excelData);
         });
        }
  // Create a new Docxtemplater instance with the Word document stream
  // const template = new Docxtemplater();
  // template.loadZip(wordDoc);

  // // Read the Excel file into a workbook object
  // const workbook = XLSX.read(excelFile);

  // // Get the worksheet object for the first sheet
  // const worksheet = workbook.Sheets[workbook.SheetNames[0]];

  // // Convert the worksheet data to a 2D array of cells
  // const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

  // // Get the tags (placeholders) in the Word document
  // const tags = template.getTags();

  // // Loop through each tag in the Word document
  // for (let i = 0; i < tags.length; i++) {
  //   const tag = tags[i];
  //   const tagStart = tag.position[0];
  //   const tagEnd = tag.position[1];
  //   const tagValue = tag.rawValue;

  //   // If the tag value is '____', it's a placeholder to be filled in
  //   if (tagValue === '____') {
  //     // Get the row of data that corresponds to this placeholder
  //     const row = data[tagStart];

  //     // Get the value from the corresponding cell in the Excel file
  //     const value = row[tagEnd];

  //     // Set the value of the placeholder to the value from the Excel file
  //     template.setDataField(tag, value);
  //   }
  // }

  // // Generate the output document as a Node.js buffer
  // const output = template.getZip().generate({
  //   type: 'nodebuffer',
  //   mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  // });

  // // Write the output buffer to a file
  // fs.writeFileSync('output.docx', output);

  // // Download the output file as a response to the user's request
  // res.download('output.docx');
});

// Start the server on port 3000
app.listen(3000, () => {
  console.log('Server started on port 3000');
});


