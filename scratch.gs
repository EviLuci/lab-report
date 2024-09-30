// function onOpen() {
//   SpreadsheetApp.getUi()
//       .createMenu('Custom Menu')
//       .addItem('Generate Dynamic Report', 'createDynamicLabReport')
//       .addToUi();
// }

// var sheet = SpreadsheetApp.getActiveSpreadsheet();
// var formDataSheet = sheet.getSheetByName("Data");
// var testTemplate = sheet.getSheetByName("Templates");
// var header = sheet.getSheetByName("Header");
// var footer = sheet.getSheetByName("Footer");
// var lastRow = formDataSheet.getLastRow();
// var data = formDataSheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];

// var referenceRanges = {
//   'TLC': { referenceLow: 4000, referenceHigh: 11000 },
//   'Neutrophil': { referenceLow: 45, referenceHigh: 75 },
//   'Lymphocyte': { referenceLow: 20, referenceHigh: 45 },
//   'Monocyte': { referenceLow: 2, referenceHigh: 8 },
//   'Eosinophil': { referenceLow: 1, referenceHigh: 6 },
//   'Basophil': { referenceLow: 0, referenceHigh: 2 },
//   'Haemoglobin': { referenceLow: 13, referenceHigh: 18 },
//   'Platelets': { referenceLow: 150000, referenceHigh: 400000 },
//   'RBC': { referenceLow: 3.8, referenceHigh: 5.9 },
//   'PCV': { referenceLow: 36, referenceHigh: 54 },
//   'MCV': { referenceLow: 76, referenceHigh: 96 },
//   'MCH': { referenceLow: 26, referenceHigh: 34 },
//   'MCHC': { referenceLow: 31, referenceHigh: 36 },
//   'UREA': { referenceLow: 10.0, referenceHigh: 45.0 },
//   'SGPT': { referenceHigh: 45.0 }, // Considered <45
//   'SGOT': { referenceHigh: 35 }, // Considered <35
//   'Creatinine': { referenceLow: 0.4, referenceHigh: 1.4 },
//   'TotalCholesterol': { referenceHigh: 200 }, // Considered <200
//   'Triglyceride': { referenceHigh: 150.0 }, // Considered <150
//   'BloodSugar': { referenceLow: 70, referenceHigh: 110 }
// };

// var testColumns = {
//   'TLC': data[1],
//   'Neutrophil': data[2],
//   'Lymphocyte': data[3],
//   'Monocyte': data[4],
//   'Eosinophil': data[5],
//   'Basophil': data[6],
//   'Haemoglobin': data[7],
//   'Platelets': data[8],
//   'RBC': data[9],
//   'PCV': data[10],
//   'MCV': data[11],
//   'MCH': data[12],
//   'MCHC': data[13],
//   'UREA': data[14],
//   'SGPT': data[15],
//   'SGOT': data[16],
//   'Creatinine': data[17],
//   'TotalCholesterol': data[18],
//   'Triglyceride': data[19],
//   'BloodSugar': data[20]
// };

// function insertHaematologyTable(doc) {

//   var body = doc.getBody();

//   // Insert Test Tables
//   var tableData = testTemplate.getRange("A3:E19").getValues();
//     // Insert the table into the Google Docs template
//   var table = body.appendTable(tableData);
//   for (var i = 0; i < table.getNumRows(); i++) {
//     var row = table.getRow(i);

//     for (var j = 0; j < row.getNumCells(); j++) {
//       var cell = row.getCell(j);
//       var text = cell.editAsText();

//       // Customize cell styles
//       text.setFontSize(i === 0 ? 10 : 9);  // Set font size
//       text.setFontFamily('Arial');  // Set font family
//       text.setBold(i === 0);  // Bold header row
//       text.setForegroundColor(i === 0 ? '#FFFFFF' : '#000000');

//       cell.setBackgroundColor(i === 0 ? '#01417B' : '#FFFFFF');
//       cell.setPaddingTop(0);
//       cell.setPaddingBottom(0);
//       cell.setPaddingLeft(0);
//       cell.setPaddingRight(0);

//       // Center-align the text in the cell
//       if (j > 0) {
//         cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
//       }
//       if (j ===0 ) {
//         cell.setPaddingLeft(5);
//       }
//     }
//   }
//   var numRows = table.getNumRows();
//   var numCols = table.getRow(0).getNumCells();

//   table.setBorderWidth(0);

//   for (var test in testColumns) {
//     if (testColumns.hasOwnProperty(test)) {
//       var result = testColumns[test];
//       var reference = referenceRanges[test];

//       // Check if the placeholder exists before attempting to replace or format
//       var placeholderText = '<<' + test + '>>';
//       var placeholderFound = body.findText(placeholderText);

//       if (placeholderFound) {
//         // Replace placeholder with the actual result
//         body.replaceText(placeholderText, result);

//         // Apply conditional formatting (bold) if the result is outside the reference range
//         var resultTextElement = body.findText(result.toString());
//         if (reference.referenceLow !== undefined && reference.referenceHigh !== undefined) {
//           if (result < reference.referenceLow || result > reference.referenceHigh) {
//             if (resultTextElement) {
//               var text = resultTextElement.getElement().asText();
//               text.setBold(resultTextElement.getStartOffset(), resultTextElement.getEndOffsetInclusive(), true);
//             }
//           }
//         } else if (reference.referenceHigh !== undefined && result > reference.referenceHigh) {
//           if (resultTextElement) {
//             var text = resultTextElement.getElement().asText();
//             text.setBold(resultTextElement.getStartOffset(), resultTextElement.getEndOffsetInclusive(), true);
//           }
//         }
//       } else {
//         Logger.log('Placeholder not found for: ' + test);
//       }
//     }
//   }

// }

// function createDynamicLabReport() {
//   var templateId = '1N_kiEdMubYED3JhIHgx3iheHrOsW1-WRzeQirS03S0k';
//   var docCopy = DriveApp.getFileById(templateId).makeCopy().getId();
//   var doc = DocumentApp.openById(docCopy);
//   var body = doc.getBody();

//   // Get the last row of form data
//   var lastRow = formDataSheet.getLastRow();
//   var formData = formDataSheet.getRange(lastRow, 1, 1, formDataSheet.getLastColumn()).getValues()[0];

//   // Access the document's header
//   var header = doc.getHeader();

//   if (header) {
//     header.replaceText('{{Patient_Name}}', 'Sujan Koju');  // Example placeholder replacement
//     header.replaceText('{{Age}}', '23');
//     header.replaceText('{{Gender}}', 'M');
//     header.replaceText('{{Address}}', 'Byasi');
//     header.replaceText('{{Referral}}', 'Hero');
//     header.replaceText('{{Collection_Date}}', '2020/12/25');
//     header.replaceText('{{Dispatch_Date}}', '2020/12/25');
//     header.replaceText('{{Sample_No}}', '1111111');
//     header.replaceText('{{Source}}', 'Gods knows how');
//   } else {
//     Logger.log("No header found in the document.");
//   }

//   insertHaematologyTable(doc);

//   // Save and generate PDF
//   doc.saveAndClose();
//   Utilities.sleep(1000);
//   var pdfFile = DriveApp.createFile(doc.getAs('application/pdf'));
//   pdfFile.setName('Lab_Report_' + data[0] + '.pdf');
//   Logger.log('PDF created: ' + pdfFile.getUrl());
//   // Clean up: delete the temporary Google Docs file if not needed
//   DriveApp.getFileById(docCopy).setTrashed(true);
// }

// function generateLabReportPDF() {

//   var templateId = '1WJ325efU5U-YyhgPceY4QpXdwu1HqgvqQ-25kifmjiY';
//   var docCopy = DriveApp.getFileById(templateId).makeCopy().getId();
//   var pdfDoc = DocumentApp.openById(docCopy);
//   var body = pdfDoc.getBody();

//   // Replace placeholders with actual data
//   body.replaceText('{{PatientName}}', data[0]);  // Example placeholder replacement
//   // body.replaceText('{{TestName}}', row[1]);
//   // body.replaceText('{{TestResult}}', row[2]);

//   for (var test in testColumns) {
//     if (testColumns.hasOwnProperty(test)) {
//       var result = testColumns[test];
//       var reference = referenceRanges[test];

//       // Check if the placeholder exists before attempting to replace or format
//       var placeholderText = '<<' + test + '>>';
//       var placeholderFound = body.findText(placeholderText);

//       if (placeholderFound) {
//         // Replace placeholder with the actual result
//         body.replaceText(placeholderText, result);

//         // Apply conditional formatting (bold) if the result is outside the reference range
//         var resultTextElement = body.findText(result.toString());
//         if (reference.referenceLow !== undefined && reference.referenceHigh !== undefined) {
//           if (result < reference.referenceLow || result > reference.referenceHigh) {
//             if (resultTextElement) {
//               var text = resultTextElement.getElement().asText();
//               text.setBold(resultTextElement.getStartOffset(), resultTextElement.getEndOffsetInclusive(), true);
//             }
//           }
//         } else if (reference.referenceHigh !== undefined && result > reference.referenceHigh) {
//           if (resultTextElement) {
//             var text = resultTextElement.getElement().asText();
//             text.setBold(resultTextElement.getStartOffset(), resultTextElement.getEndOffsetInclusive(), true);
//           }
//         }
//       } else {
//         Logger.log('Placeholder not found for: ' + test);
//       }
//     }
//   }

//   // Save and generate PDF
//   pdfDoc.saveAndClose();
//   Utilities.sleep(1000);
//   var pdfFile = DriveApp.createFile(pdfDoc.getAs('application/pdf'));
//   pdfFile.setName('Lab_Report_' + data[0] + '.pdf');

//   // Optionally, email the PDF to the patient or doctor
//   // MailApp.sendEmail(row[5], 'Your Lab Test Report', 'Please find your lab test report attached.', {
//   //     attachments: [pdfFile]
//   // });

//   // Clean up: delete the temporary Google Docs file if not needed
//   DriveApp.getFileById(docCopy).setTrashed(true);
// }

// function importTableFromSheets() {
//   // IDs for the Google Sheets and Docs
//   var docId = '1CTexMynaaNx2r_yaCNJ0Scra8dzG7z0XWba2BtVLAHk';

//   // Copy the range values (you can also use `getValues()` if you want to manipulate the data)
//   var values = testTemplate.getRange("A3:E19").getValues();

//   // Open the Google Docs
//   var doc = DocumentApp.openById(docId);
//   var body = doc.getBody();

//   // Create a table in Google Docs
//   var table = body.appendTable(values);

//   // Optional: Set the table's alignment, styles, etc.
//   table.setBorderWidth(1);  // Set border width
//   table.setBorderColor('#000000');  // Set border color
//   table.setTableAlignment(DocumentApp.HorizontalAlignment.CENTER);  // Center the table

//   doc.saveAndClose();
// }

// function insertTestTables(doc, testType, formData, testTemplate) {
//   var body = doc.getBody();

//   // Insert the title of the test
//   body.appendParagraph(testType + ' - GENERAL TEST REPORT').setAlignment(DocumentApp.HorizontalAlignment.CENTER).setBold(true);

//   // Retrieve the table definition from the sheet
//   var testRange = testTemplate.getRange('A3:E19'); // Adjust range for test table template
//   var testTableValues = testRange.getValues();

//   // Insert the table into the document
//   var table = body.appendTable();
//   for (var i = 0; i < testTableValues.length; i++) {
//     var row = table.appendTableRow();
//     for (var j = 0; j < testTableValues[i].length; j++) {
//       var cellText = testTableValues[i][j];
//       // if (cellText.startsWith('<<') && cellText.endsWith('>>')) {
//       //   // Replace placeholder with actual form data
//       //   var placeholderIndex = parseInt(cellText.replace('<<', '').replace('>>', '')) - 1;
//       //   cellText = formData[placeholderIndex];
//       // }
//       row.appendTableCell(cellText);
//     }
//   }

//   // Conditional formatting based on reference range
//   applyConditionalFormatting(table, formData, testTemplate);
// }

// function applyConditionalFormatting(table, formData, testTemplate) {
//   // Iterate through table cells and apply bold formatting based on reference range
//   var rows = table.getNumRows();
//   for (var i = 1; i < rows; i++) { // Start from 1 to skip header row
//     var cellValue = parseFloat(table.getRow(i).getCell(1).getText()); // Adjust for result cell
//     var referenceRange = testTemplate.getRange(i, 4).getValue(); // Adjust range for reference range
//     if (cellValue > referenceRange) {
//       table.getRow(i).getCell(1).setBold(true); // Apply bold formatting
//     }
//   }
// }

// function insertSummary(body, staticContentSheet) {
//   var summaryText = staticContentSheet.getRange('B1').getValue(); // Adjust range
//   body.appendParagraph('\nNOTE:').setBold(true);
//   body.appendParagraph(summaryText);
// }

// // Too much pain in the ass to implement image insertion (Invalid Image Data!!!!!!!!!!!)
// function insertSignatureAndName(doc, imageFileId, signatureName) {
//   var body = doc.getBody();

//   try {
//     // Fetch the image file from Drive
//     var imageFile = DriveApp.getFileById(imageFileId);
//     Logger.log('Image file retrieved: ' + imageFile.getName());

//     // Get the image as a blob
//     var imageBlob = imageFile.getBlob();
//     Logger.log('Image blob created.');

//     // Convert the image blob to a proper format
//     var imageData = imageBlob.getDataAsString();
//     var imageBlobFormatted = Utilities.newBlob(imageData, imageBlob.getContentType(), imageFile.getName());

//     // Insert the image into the document
//     body.appendImage(imageBlobFormatted);
//     body.insertInlineImage(imageBlob);
//     Logger.log('Image inserted into document.');

//     // Add a new line after the image
//     body.appendParagraph('\n');

//     // Insert the name of the individual
//     body.appendParagraph('Signature: ' + signatureName)
//         .setBold(true)
//         .setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
//     Logger.log('Signature name added to document.');

//   } catch (e) {
//     Logger.log('Error inserting signature and name: ' + e.message);
//   }
// }

// var labTests = {
//   Hematology: {
//     templateTableIndex: 0,
//     tests: [
//       { placeholder: '<<TLC>>', result: formData[1]},
//       { placeholder: '<<Neutrophil>>', result: formData[2]},
//       { placeholder: '<<Lymphocyte>>', result: formData[3]},
//       { placeholder: '<<Monocyte>>', result: formData[4]},
//       { placeholder: '<<Eosinophil>>', result: formData[5]},
//       { placeholder: '<<Basophil>>', result: formData[6]},
//       { placeholder: '<<Hemoglobin>>', result: formData[7]},
//       { placeholder: '<<Platelets>>', result: formData[8]},
//       { placeholder: '<<RBC>>', result: formData[9]},
//       { placeholder: '<<PCV>>', result: formData[10]},
//       { placeholder: '<<MCV>>', result: formData[11]},
//       { placeholder: '<<MCH>>', result: formData[12]},
//       { placeholder: '<<MCHC>>', result: formData[13]}
//     ]
//   },
//   Lipid_Profile: {
//     templateTableIndex: 1,
//     tests: [
//       { placeholder: '<<Triglyceride>>', result: formData[1]},
//       { placeholder: '<<Cholesterol>>', result: formData[2]},
//       { placeholder: '<<HDL Cholesterol>>', result: formData[3]},
//       { placeholder: '<<LDL Cholesterol>>', result: formData[4]},
//       { placeholder: '<<VLDL>>', result: formData[5]},
//     ]
//   }
// };

// // Data structure to hold test info
// var hematologyTests = [
//   { placeholder: '<<TLC>>', result: formData[9]},
//   { placeholder: '<<Neutrophil>>', result: formData[2]},
//   { placeholder: '<<Lymphocyte>>', result: formData[3]},
//   { placeholder: '<<Monocyte>>', result: formData[4]},
//   { placeholder: '<<Eosinophil>>', result: formData[5]},
//   { placeholder: '<<Basophil>>', result: formData[6]},
//   { placeholder: '<<Hemoglobin>>', result: formData[7]},
//   { placeholder: '<<Platelets>>', result: formData[8]},
//   { placeholder: '<<RBC>>', result: formData[9]},
//   { placeholder: '<<PCV>>', result: formData[10]},
//   { placeholder: '<<MCV>>', result: formData[11]},
//   { placeholder: '<<MCH>>', result: formData[12]},
//   { placeholder: '<<MCHC>>', result: formData[13]}
// ];

// return Table with signature
// function getSignatureFromTemplate(tableIndex) {
//   const tableTemplateDoc;

//   try {
//     // Open the template document using the global variable
//     tableTemplateDoc = DocumentApp.openById(SIGNATURE_TEMPLATE_DOC_ID);
//   } catch (e) {
//     Logger.log('Error opening table template document: ' + e.message);
//     return null;
//   }

//   // Get all tables from the document
//   const tables = tableTemplateDoc.getBody().getTables();

//   // Ensure that the specified table index exists
//   if (tableIndex < 0 || tableIndex >= tables.length) {
//     Logger.log('Invalid table index: ' + tableIndex);
//     return null;
//   }

//   // Return the table at the specified index
//   return tables[tableIndex].copy();
// }
