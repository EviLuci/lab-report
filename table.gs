// Return indexed table from table template doc
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

// process table insertion according to the test selected
function processTableInsertion(body, formData) {
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

  departments.forEach((department, index) => {
    const labTestsInDepartment = departmentTests[department];

    // Insert all tests for this department
    labTestsInDepartment.forEach((labTestInDepartment) => {
      const rowCount = labTestInDepartment.tests.length + 1; // Include header row
      const tableHeight = calculateTableHeight(rowCount);
      const estimatedSpaceNeeded = tableHeight + 120;

      const contentHeight = body.getText().length * 0.7; // Estimate current content height
      const remainingHeight = PAGE_HEIGHT - (contentHeight % PAGE_HEIGHT); // Approximation
      Logger.log(
        contentHeight + "&" + remainingHeight + "&" + estimatedSpaceNeeded
      );
      // Check if there’s enough space; insert page break if not
      if (!hasEnoughSpace(body, estimatedSpaceNeeded)) {
        Logger.log("space not enough");
        body.appendPageBreak();
      }
      // Logger.log("table inserted")
      insertTable(body, labTestInDepartment, formData);
      currentTestCount++; // Increment after each table is inserted
    });

    // Check if this is the last test
    const isLastTest = currentTestCount === totalTests;

    // Append signature to the table
    appendSignatureTable(body, 0, formData[9], isLastTest);

    // Only insert a page break if this is not the last department
    if (index < departments.length - 1) {
      body.appendPageBreak();
    }
  });

  // const selectedTests = JSON.parse(formData[10]);
  // const labTests = getLabTests(formData);

  // selectedTests.forEach((selectedTest, index) => {
  //   if (labTests[selectedTest]) {
  //     const labTest = labTests[selectedTest];
  //     labTest.isLastTest = index === selectedTests.length - 1;
  //     insertTable(doc, labTest, formData);
  //   }
  // });
}

// For inserting test table
function insertTable(body, labTest, formData) {
  // const body = doc.getBody();

  // Insert heading before the table
  const heading = body.insertParagraph(
    body.getNumChildren(),
    labTest.tableHeading
  );
  heading
    .setHeading(DocumentApp.ParagraphHeading.HEADING2)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .setBold(true)
    .setFontSize(12);

  const testTable = getTableFromTemplate(
    labTest.templateTableIndex,
    TEMPLATE_DOC_ID
  );

  // Check if the table was successfully retrieved
  if (!testTable) {
    Logger.log("Failed to retrieve the ${testType} table.");
    return;
  }

  // Iterate over table cells and replace placeholders with formData
  for (let row = 1; row < testTable.getNumRows(); row++) {
    const cell = testTable.getCell(row, 1);
    const cellText = cell.getText();

    // Find the corresponding test from the test objects
    const test = labTest.tests.find((t) => t.placeholder === cellText);
    if (test) {
      if (checkIfEmpty(test.result)) {
        if (row === testTable.getNumRows() - 1) {
          // If it's the last row, don't delete it, just clear its contents
          var numCells = testTable.getRow(row).getNumCells();
          for (var col = 0; col < numCells; col++) {
            testTable.getRow(row).getCell(col).clear();
          }
        } else {
          testTable.removeRow(row);
          row--; // Adjust the row index after deletion
        }
      } else {
        let result = test.result;
        // Format single-digit numbers as two digits
        if (!isNaN(result) && Number.isInteger(result) && result < 10) {
          result = "0" + result;
        }
        // Replace placeholder with result
        cell.setText(result);

        // const rangeCol = testTable.getCell(row, 3).getText();
        // const rangeColSplit = rangeCol.split('-');
        // const refLow = parseFloat(rangeColSplit[0].trim());
        // const refHigh = parseFloat(rangeColSplit[1].trim());

        // // Apply bold if the result is out of the reference range
        // if (test.result < refLow || test.result > refHigh) {
        //   cell.setBold(true);
        // }
        applyFormatting(testTable, row, result);
      }
    }
  }

  // Append the updated table to the document
  body.appendTable(testTable);

  // Append comment
  // if (formData[76] === 'Yes') {
  //     insertComment(doc, formData);
  // }

  // Append comment
  if (formData[76] === "Custom") {
    insertComment(body, formData);
  } else if (formData[76] === "Yes") {
    if (labTest.commentTableIndex) {
      appendCommentTable(body, labTest.commentTableIndex);
    } else {
      Logger.log("No comment index");
    }
  } else {
    Logger.log("No comment");
  }

  // If this isn't the last test type, add a page break
  // if (labTest.isLastTest !== true) {
  //   body.appendPageBreak();
  // }
}

// apply formatting based on reference range
function applyFormatting(testTable, row, result) {
  const headerRow = testTable.getRow(0);
  let referenceColIndex = null;

  // Search for the "Reference Range" column in the header row
  for (let col = 0; col < headerRow.getNumCells(); col++) {
    const headerText = headerRow.getCell(col).getText().trim();
    if (headerText === "Reference Range") {
      referenceColIndex = col;
      break;
    }
  }

  // Proceed with formatting only if the "Reference Range" column is found
  if (referenceColIndex !== null) {
    const rangeCol = testTable.getCell(row, referenceColIndex).getText();
    if (!checkIfEmpty(rangeCol)) {
      const [refLow, refHigh] = rangeCol
        .split("-")
        .map((v) => parseFloat(v.trim()));
      if (result < refLow || result > refHigh) {
        testTable.getCell(row, 1).setBold(true);
      }
    }
  }
}

// append signature table from signature template
function appendSignatureTable(
  body,
  signatureTableIndex,
  isSpecialData,
  isLastTest
) {
  // const body = doc.getBody();

  // If this is the last test type, add END OF REPORT
  if (isLastTest) {
    body
      .appendParagraph("** END OF REPORT **")
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
      .setBold(true)
      .setFontSize(10);
  }

  const signatureTable = getTableFromTemplate(
    signatureTableIndex,
    SIGNATURE_TEMPLATE_DOC_ID
  );
  body.appendTable(signatureTable);
  if (isSpecialData === "Yes") {
    body
      .appendParagraph(
        "*This sample was processed at Medi Quest Laboratory Clinic."
      )
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
      .setBold(false)
      .setFontSize(10);
  }
}

// append comment table from comment template
function appendCommentTable(body, commentTableIndex) {
  // const body = doc.getBody();

  const commentTable = getTableFromTemplate(
    commentTableIndex,
    COMMENT_TEMPLATE_ID
  );
  body.appendTable(commentTable);
}

// for comment
function insertComment(body, formData) {
  // const body = doc.getBody();
  const commentTable = getTableFromTemplate(0, COMMENT_TEMPLATE_ID);
  body.appendTable(commentTable);
  body.replaceText("{{Comment}}", formData[77]);
}
