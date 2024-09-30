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
function processTableInsertion(doc, formData) {
  const selectedTests = JSON.parse(formData[10]);
  const labTests = getLabTests(formData);

  selectedTests.forEach((selectedTest, index) => {
    if (labTests[selectedTest]) {
      const labTest = labTests[selectedTest];
      labTest.isLastTest = index === selectedTests.length - 1;
      insertTable(doc, labTest, formData);
    }
  });
}

// For inserting test table
function insertTable(doc, labTest, formData) {
  const body = doc.getBody();

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
        // Replace placeholder with result
        cell.setText(test.result);

        // const rangeCol = testTable.getCell(row, 3).getText();
        // const rangeColSplit = rangeCol.split('-');
        // const refLow = parseFloat(rangeColSplit[0].trim());
        // const refHigh = parseFloat(rangeColSplit[1].trim());

        // // Apply bold if the result is out of the reference range
        // if (test.result < refLow || test.result > refHigh) {
        //   cell.setBold(true);
        // }
        applyFormatting(testTable, row, test.result);
      }
    }
  }

  // Append the updated table to the document
  body.appendTable(testTable);

  // Append comment
  if (formData[76] === "Yes") {
    insertComment(doc, formData);
  }

  // Append signature to the table
  appendSignatureTable(doc, 0, formData[9]);

  // If this isn't the last test type, add a page break
  if (labTest.isLastTest !== true) {
    body.appendPageBreak();
  }
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
function appendSignatureTable(doc, signatureTableIndex, isSpecialData) {
  const body = doc.getBody();

  body
    .appendParagraph("** END OF REPORT **")
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .setBold(true)
    .setFontSize(10);

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

// for comment
function insertComment(doc, formData) {
  const body = doc.getBody();
  const commentTable = getTableFromTemplate(0, COMMENT_TEMPLATE_ID);
  body.appendTable(commentTable);
  body.replaceText("{{Comment}}", formData[77]);
}
