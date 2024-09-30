// Constants
const TEMPLATE_DOC_ID = "1ZM4UQq8ayT1Zyxq3_Ns8mw89P-oVGuo7XZiX3ZbGeJM";
const SIGNATURE_TEMPLATE_DOC_ID =
  "1ry6GTSd8qNDMolxCyn_f9DniXnhDN3gWAYR0nwFl7xQ";
const REPORT_FOLDER_ID = "19r23CO8xVQ50cl7KxT0cnir9ebedilKT";
const REPORT_TEMPLATE_ID = "1N_kiEdMubYED3JhIHgx3iheHrOsW1-WRzeQirS03S0k";

// Custom menu creation
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Custom Menu")
    .addItem("Generate Dynamic Report", "createDynamicLabReport")
    .addToUi();
}

// Main function to create dynamic lab report
function createDynamicLabReport() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const formDataSheet = sheet.getSheetByName("Data");
  const lastRow = formDataSheet.getLastRow();
  const formData = formDataSheet
    .getRange(lastRow, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  try {
    const patientFolder = createPatientFolder(formData[0]);
    const docCopyId = createReportCopy(patientFolder, formData[0]);
    const doc = DocumentApp.openById(docCopyId);

    replacePlaceholdersInHeader(doc, formData);
    processTableInsertion(doc, formData);
    appendSignatureTable(doc);

    doc.saveAndClose();

    const pdfFile = saveAsPdf(docCopyId, formData[0]);
    movePdfToFolder(pdfFile, patientFolder);

    Logger.log("PDF created: " + pdfFile.getUrl());
  } catch (error) {
    Logger.log("Error in createDynamicLabReport: " + error.toString());
  }
}

// Helper functions
function createPatientFolder(patientName) {
  const reportFolder = DriveApp.getFolderById(REPORT_FOLDER_ID);
  return reportFolder.createFolder(patientName + "_Report");
}

function createReportCopy(folder, patientName) {
  const docCopy = DriveApp.getFileById(REPORT_TEMPLATE_ID).makeCopy();
  docCopy.setName("Lab_Report_" + patientName).moveTo(folder);
  return docCopy.getId();
}

function saveAsPdf(docId, fileName) {
  const doc = DocumentApp.openById(docId);
  const pdfFile = DriveApp.createFile(doc.getAs("application/pdf"));
  pdfFile.setName("Lab_Report_" + fileName);
  return pdfFile;
}

function movePdfToFolder(pdfFile, folder) {
  DriveApp.getFileById(pdfFile.getId()).moveTo(folder);
}

function replacePlaceholdersInHeader(doc, formData) {
  const header = doc.getHeader();
  if (!header) {
    Logger.log("No header found in the document.");
    return;
  }

  const placeholders = {
    "{{Patient_Name}}": formData[0], // Assuming patient name is the first item
    "{{Age}}": formData[1],
    "{{Gender}}": formData[2],
    "{{Address}}": formData[3],
    "{{Referral}}": formData[4],
    "{{Collection_Date}}": formData[5],
    "{{Dispatch_Date}}": formData[6],
    "{{Sample_No}}": formData[7],
    "{{Source}}": formData[8],
  };

  Object.entries(placeholders).forEach(([placeholder, value]) => {
    header.replaceText(placeholder, value || "");
  });
}

function getTableFromTemplate(tableIndex, templateDocId) {
  try {
    const tableTemplateDoc = DocumentApp.openById(templateDocId);
    const tables = tableTemplateDoc.getBody().getTables();

    if (tableIndex < 0 || tableIndex >= tables.length) {
      throw new Error("Invalid table index: " + tableIndex);
    }

    return tables[tableIndex].copy();
  } catch (e) {
    Logger.log("Error in getTableFromTemplate: " + e.message);
    return null;
  }
}

function checkIfEmpty(value) {
  return (
    value == null ||
    value === "" ||
    (typeof value === "string" && value.trim() === "") ||
    (Array.isArray(value) && value.length === 0) ||
    (typeof value === "object" && Object.keys(value).length === 0)
  );
}

function insertTable(doc, testType, labTest) {
  const body = doc.getBody();

  const heading = body.insertParagraph(
    body.getNumChildren(),
    testType.toUpperCase()
  );
  heading
    .setHeading(DocumentApp.ParagraphHeading.HEADING3)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .setBold(true);

  const testTable = getTableFromTemplate(
    labTest.templateTableIndex,
    TEMPLATE_DOC_ID
  );

  if (!testTable) {
    Logger.log(`Failed to retrieve the ${testType} table.`);
    return;
  }

  for (let row = 1; row < testTable.getNumRows(); row++) {
    const cell = testTable.getCell(row, 1);
    const cellText = cell.getText();

    const test = labTest.tests.find((t) => t.placeholder === cellText);
    if (test) {
      if (checkIfEmpty(test.result)) {
        for (let col = 0; col < testTable.getNumColumns(); col++) {
          testTable.getCell(row, col).clear();
        }
        testTable.hideRow(row);
      } else {
        cell.setText(test.result);
        applyFormatting(testTable, row, test.result);
      }
    }
  }

  body.appendTable(testTable);

  if (labTest.isLastTest !== true) {
    body.appendPageBreak();
  }
}

function applyFormatting(testTable, row, result) {
  const rangeCol = testTable.getCell(row, 3).getText();
  if (!checkIfEmpty(rangeCol)) {
    const [refLow, refHigh] = rangeCol
      .split("-")
      .map((v) => parseFloat(v.trim()));
    if (result < refLow || result > refHigh) {
      testTable.getCell(row, 1).setBold(true);
    }
  }
}

function processTableInsertion(doc, formData) {
  const selectedTests = JSON.parse(formData[0]);
  const labTests = getLabTests(formData);

  selectedTests.forEach((selectedTest, index) => {
    if (labTests[selectedTest]) {
      const labTest = labTests[selectedTest];
      labTest.isLastTest = index === selectedTests.length - 1;
      insertTable(doc, selectedTest, labTest);
    }
  });
}

function getLabTests(formData) {
  // You may want to adjust the indices based on your actual form data structure
  return {
    Hematology: {
      templateTableIndex: 0,
      tests: [
        { placeholder: "<<TLC>>", result: formData[9] },
        { placeholder: "<<Neutrophil>>", result: formData[10] },
        // ... other hematology tests ...
      ],
    },
    Lipid_Profile: {
      templateTableIndex: 1,
      tests: [
        { placeholder: "<<Triglyceride>>", result: formData[11] },
        { placeholder: "<<Cholesterol>>", result: formData[12] },
        // ... other lipid profile tests ...
      ],
    },
    // ... other test types ...
  };
}

function appendSignatureTable(doc) {
  const body = doc.getBody();
  const signatureTable = getTableFromTemplate(0, SIGNATURE_TEMPLATE_DOC_ID);
  body.appendTable(signatureTable);
}

function onSheetChange(e) {
  if (isScriptRunning()) {
    Logger.log("Script is already running. Exiting.");
    return;
  }

  setScriptRunning(true);

  try {
    const changeType = e.changeType;

    if (changeType === "EDIT") {
      Logger.log("Manual Edit Detected: Aborting Script");
    } else if (changeType === "INSERT_ROW") {
      createDynamicLabReport();
    }
  } catch (error) {
    Logger.log("Error in onSheetChange: " + error.toString());
  } finally {
    setScriptRunning(false);
  }
}

function isScriptRunning() {
  return (
    PropertiesService.getScriptProperties().getProperty("isRunning") === "true"
  );
}

function setScriptRunning(isRunning) {
  PropertiesService.getScriptProperties().setProperty(
    "isRunning",
    isRunning.toString()
  );
}

function setupTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((trigger) => ScriptApp.deleteTrigger(trigger));

  ScriptApp.newTrigger("onSheetChange")
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onChange()
    .create();

  Logger.log("Trigger set up successfully");
}
