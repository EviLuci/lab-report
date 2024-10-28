/** This file is for testig new logic and features only */

function test() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const testSheet = sheet.getSheetByName("Data");
  const lastRow = testSheet.getLastRow();
  const formData = testSheet
    .getRange(lastRow, 1, 1, testSheet.getLastColumn())
    .getValues()[0];

  const doc = DocumentApp.openById(
    "1qnOPvNES1c4QViP9Ec62VJIvy5jByieRMNezoqW1ysY"
  );
  const body = doc.getBody();

  Logger.log(formData[10]);
  const labTests = JSON.parse(formData[10]);
  Logger.log(labTests);
}

function testingTableHeight() {
  const tableTemplateDoc = DocumentApp.openById(TEMPLATE_DOC_ID);
  const tables = tableTemplateDoc.getBody().getTables();
  test_table = tables[2];
  const rowheight_ofTHeTable = test_table.getRow(0).getMinimumHeight();

  const doc = DocumentApp.openById(
    "1qnOPvNES1c4QViP9Ec62VJIvy5jByieRMNezoqW1ysY"
  );
  const body = doc.getBody();
  const height_ofTHeDoc = body.getPageHeight();
  Logger.log(rowheight_ofTHeTable);
}

// process table insertion according to the test selected
function newprocessTableInsertion() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const testSheet = sheet.getSheetByName("Data");
  const lastRow = testSheet.getLastRow();
  const formData = testSheet
    .getRange(lastRow, 1, 1, testSheet.getLastColumn())
    .getValues()[0];

  const selectedTests = JSON.parse(formData[10]);
  const labTests = getLabTests(formData);
  const departmentTests = {};

  selectedTests.forEach((selectedTest) => {
    const labTest = labTests[selectedTest];
    if (labTest) {
      const department = labTest.department;
      // Logger.log(department);
      if (!departmentTests[department]) {
        departmentTests[department] = [];
      }
      departmentTests[department].push(labTest);
    }
  });

  // Insert tables by department
  const departments = Object.keys(departmentTests);
  const totalTests = selectedTests.length;
  let currentTestCount = 0;

  const page_height = body.getPageHeight();
  const contentHeight = body.getText().length * 0.7; // Estimate current content height
  const remainingHeight = page_height - contentHeight; // Approximation
  Logger.log(contentHeight + "&" + remainingHeight + "&" + estimatedHeight);

  departments.forEach((department, index) => {
    const labTestsInDepartment = departmentTests[department];

    // Insert all tests for this department
    labTestsInDepartment.forEach((labTestInDepartment) => {
      const rowCount = labTestInDepartment.tests.length + 1; // Include header row
      const tableHeight = calculateTableHeight(rowCount);
      const estimatedSpaceNeeded = tableHeight + 120;

      Logger.log(rowCount + "&" + tableHeight + "&" + estimatedSpaceNeeded);

      // // Check if there’s enough space; insert page break if not
      // if (!hasEnoughSpace(body, estimatedSpaceNeeded)) {
      //   body.appendPageBreak();
      // }
      // // Logger.log("table inserted")
      // insertTable(body, labTestInDepartment, formData);
      // currentTestCount++; // Increment after each table is inserted
    });

    // Check if this is the last test
    // const isLastTest = currentTestCount === totalTests;

    // // Append signature to the table
    // appendSignatureTable(body, 0, formData[9], isLastTest);

    // // Only insert a page break if this is not the last department
    // if (index < departments.length - 1) {
    //   body.appendPageBreak();
    // }
  });

  // selectedTests.forEach((selectedTest, index) => {
  //   if (labTests[selectedTest]) {
  //     const labTest = labTests[selectedTest];
  //     labTest.isLastTest = index === selectedTests.length - 1;
  //     insertTable(doc, labTest, formData);
  //   }
  // });
  // Logger.log(departmentTests);
}
