
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
    .addItem('Fetch and Insert Data', 'fetchAndInsertData')
    .addItem('Remove Empty Rows (1st Sheet)', 'removeEmptyRows')
    .addItem('Process Verses', 'processVerses')
    .addItem('Mark Duplicates', 'markDuplicates')    
    .addToUi();
}

function fetchAndInsertData() {
  // Very Useful - Don't modify
  // Fetches all verse details from Onlinepj.in website and
  // Puts in Verses Sheet add a empty row between chapters

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = spreadsheet.getSheetByName('Chapters');
  const targetSheet = spreadsheet.getSheetByName('Verses');
  
  const richTextValues = sourceSheet.getRange('B1:B' + sourceSheet.getLastRow()).getRichTextValues();
  const statusValues = sourceSheet.getRange('C1:C' + sourceSheet.getLastRow()).getValues();

  let targetRow = targetSheet.getLastRow() + 1;

  richTextValues.forEach((richTextValue, index) => {
    const url = richTextValue[0].getLinkUrl(); // Get the hyperlink URL
    if (url && !statusValues[index][0]) { // Check if URL is not empty and corresponding C cell is empty
      try {
        const response = UrlFetchApp.fetch(url);
        const html = response.getContentText();
        const content = parseHTML(html);
        
        const chapterNumber = index + 1; // Incremented number (chapter number)

        // Collect data to insert
        let dataToInsert = [];

        content.forEach((item) => {
          const lines = item.split('\n');
          lines.forEach((line) => {
            if (line.trim() !== '') {
              let cleanLine = line.replace(/&nbsp;/g, ' ').trim(); // Replace &nbsp; with space and trim
              dataToInsert.push([chapterNumber, '', cleanLine]);
            }
          });
        });

        if (dataToInsert.length > 0) {
          targetSheet.insertRows(targetRow, dataToInsert.length); // Insert required rows
          targetSheet.getRange(targetRow, 2, dataToInsert.length, 3).setValues(dataToInsert); // Set values in columns B, C, and D
          targetRow += dataToInsert.length;
        }

        // Mark the corresponding C cell in Chapters sheet as 'Done'
        sourceSheet.getRange(index + 1, 3).setValue('Done');
      } catch (error) {
        Logger.log('Error fetching or parsing URL: ' + url);
      }
    }
  });
}

function parseHTML(html) {
  // Use a regular expression to extract content within elements with class 'article-intro' and itemprop 'description'
  const regex = /<blockquote class="article-intro" itemprop="description">([\s\S]*?)<\/blockquote>/g;
  const matches = [];
  let match;
  
  while ((match = regex.exec(html)) !== null) {
    const content = match[1].replace(/<[^>]*>/g, ''); // Remove HTML tags
    matches.push(content.trim());
  }
  
  return matches;
}

function removeEmptyRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getDataRange();
  const values = range.getValues();
  
  let rowsToDelete = [];
  
  // Identify empty rows
  values.forEach((row, index) => {
    if (row.join('').trim() === '') {
      rowsToDelete.push(index + 1); // +1 because sheet rows are 1-indexed
    }
  });
  
  // Delete rows in reverse order to avoid indexing issues
  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    sheet.deleteRow(rowsToDelete[i]);
  }
}


function processVerses() {
  // Useful function
  // This function is used to put verse number in left cell of a verse like
  // The left adjacent cell would contain chpater number
  // The right adjacent cell would contain the verse

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = spreadsheet.getSheetByName('Verses');
  
  const lastRow = targetSheet.getLastRow();
  const verses = targetSheet.getRange('D1:D' + lastRow).getValues();

  verses.forEach((verse, index) => {
    if (typeof verse[0] === 'string' && verse[0].trim() !== '') {
      let verseText = verse[0].replace(/&nbsp;/g, ' ').trim(); // Replace &nbsp; with space and trim
      let verseNumber = extractVerseNumber(verseText);

      if (!verseNumber && verseText !== '') {
        verseNumber = '0'; // Add 0 if verse number does not start with a number and the verse text is not empty
      }

      if (verseNumber) {
        targetSheet.getRange(index + 1, 3).setValue(verseNumber); // Set verse number in column C
      }
      targetSheet.getRange(index + 1, 4).setValue(verseText); // Update the verse text after cleaning
    } else {
      Logger.log('Non-string value found at row ' + (index + 1));
    }
  });
}

function extractVerseNumber(verse) {
  const match = String(verse).match(/^(\d+)[.,\s]/); // Ensure verse is a string
  return match ? match[1] : null;
}


function markDuplicates() {

  // This function was used to mark dupicate rows in Test Sheet only  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Check if the active sheet is 'Test'
  if (sheet.getName() !== 'Test') {
    // return; // Exit the function if it's not the 'Test' sheet
  }
  
  var data = sheet.getRange("A:A").getValues(); // Get all values in column A
  var output = [];
  var seen = {};
  
  for (var i = 0; i < data.length; i++) {
    var value = data[i][0];
    if (value !== "") {
      if (seen[value]) {
        output.push(["duplicate"]);
      } else {
        seen[value] = true;
        output.push([""]);
      }
    } else {
      output.push([""]);
    }
  }
  
  sheet.getRange(1, 2, output.length, 1).setValues(output); // Put the results in column B
  
  var range = sheet.getRange("B:B");
  var conditionalFormatRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('duplicate')
    .setBackground("#FF0000")
    .setRanges([range])
    .build();
  var rules = sheet.getConditionalFormatRules();
  rules.push(conditionalFormatRule);
  sheet.setConditionalFormatRules(rules);
}



function combineNoVerseDataIntoOneRow() {

  // This function is used to combine the data before a sura's first verse start
  // Then put all the combined data into the first row
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var rowCount = 0;

  for (var i = 0; i < data.length; i++) {
    if (rowCount > 2) {
      break; // Safety break if more than 3 combinations are performed
    }

    var bValue = data[i][1];
    var cValue = data[i][2];

    if (typeof bValue === 'number' && bValue > 0 && cValue === 0) {
      var startRow = i;
      var endRow = i;
      var combinedData = '';

      // Combine data for successive rows with the same B, C values
      while (endRow < data.length && data[endRow][1] === bValue && data[endRow][2] === 0) {
        combinedData += data[endRow][3] + "\n\n";
        endRow++;
        
        // Break if some other data is found in other columns
        if (data[endRow] && (data[endRow][1] !== bValue || data[endRow][2] !== 0)) {
          break;
        }
      }

      if (endRow > startRow + 1) {
        // Insert a new row after the last of these rows
        sheet.insertRowAfter(endRow);

        // Put the combined data in the new row's C cell
        sheet.getRange(endRow + 1, 4).setValue(combinedData);

        // Increase the rowCount for safety check
        rowCount++;

        // Adjust the loop counter to continue after the newly inserted row
        i = endRow;
      }
    }
  }
}

