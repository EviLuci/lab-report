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

// process table insertion according to the test selected
function newprocessTableInsertion(doc, formData) {
  const selectedTests = JSON.parse(formData[10]);
  const labTests = getLabTests(formData);
  const departmentTests = {};

  selectedTests.forEach((selectedTest) => {
    const labTest = labTests[selectedTest];
    if (labTest) {
      const department = labTest.department;
      if (!departmentTests[department]) {
        departmentTests[department] = [];
      }
      departmentTests[department].push(labTest);
    }
  });

  selectedTests.forEach((selectedTest, index) => {
    if (labTests[selectedTest]) {
      const labTest = labTests[selectedTest];
      labTest.isLastTest = index === selectedTests.length - 1;
      insertTable(doc, labTest, formData);
    }
  });
}

// Function to check if there’s enough space for a table on the current page
function hasEnoughSpace(body, estimatedHeight) {
  const remainingHeight =
    DocumentApp.PageSize.A4.getHeight() -
    body.getCursor().getSurroundingTextOffset(); // Approximation
  return remainingHeight >= estimatedHeight;
}
