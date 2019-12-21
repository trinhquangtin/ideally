function prettyReport() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = spreadsheet.getSheetByName('Your Jira Issues');
  var targetSheet = spreadsheet.getSheetByName('report');
  var data = sourceSheet.getDataRange().getValues();
  
  targetSheet.clear();
  targetSheet.setFrozenRows(1);
  targetSheet.setFrozenColumns(4);
  var range = targetSheet.getRange(1, 1, 100, 100);
  range.setBorder(true, true, true, true, true, true, 'white', SpreadsheetApp.BorderStyle.SOLID);
//  SpreadsheetApp.setSpreadsheetTheme(SpreadsheetApp.AutoFillSeries);
  
  var countMerge = 2;
  var qc1 = 3, qc2 = 4, devstart = column('K'), devend = column('L'), devprogress = column('M'),
      qc1start = column('O'), qc1end = column('P'), qc1progress = column('Q'), 
      qc2start = column('O'), qc2end = column('O'), qc2progress = column('O'); 
  
  // add header
  headerContents = [' ', 'jira ID', 'Title', 'Owner', 'Start date plan', 'End date plan', 'Progress %', ' Start date actual', 'End date actual'];
  targetSheet.appendRow(headerContents);
  
  // pretty header
  var headerCells = targetSheet.getRange("B1:I1");
  
  // Sets borders on the top and bottom, but leaves the left and right unchanged
  // Also sets the color to "red", and the border to "DASHED".
  //  headerCells.setBorder(true, null, true, null, false, false, "red", SpreadsheetApp.BorderStyle.DASHED);
  headerCells.setBackground("#DDD");
  var bold = SpreadsheetApp.newTextStyle()
    .setBold(true)
    .build();
  headerCells.setTextStyle(bold);
  
  for (var i = 1; i < data.length; i++) {
    targetSheet.appendRow([' ', data[i][0], data[i][1], data[i][2], data[i][devstart], data[i][devend], data[i][devprogress]]);
    
    targetSheet.appendRow([' ', ' ', ' ', data[i][qc1]]);
    targetSheet.appendRow([' ', ' ', ' ', data[i][qc2]]);
    
    var iMerge = countMerge + 2;
    var mergeA = 'B' + countMerge + ':B' + iMerge; // A2:A4, A5:A7
    var mergeB = 'C' + countMerge + ':C' + iMerge;
    targetSheet.getRange(mergeA).merge();
    targetSheet.getRange(mergeB).merge();
    var borderCount = iMerge - 1;
    
    // pretty row
    var rowCells = targetSheet.getRange("B" + iMerge + ":I" + iMerge);
    rowCells.setVerticalAlignment('middle');
    rowCells.setWrap(true);
    rowCells.setBorder(true, null, true, null, null, null, "#ddd", SpreadsheetApp.BorderStyle.SOLID);
    var rowCells = targetSheet.getRange("B" + countMerge + ":I" + countMerge);
    rowCells.setVerticalAlignment('middle');
    rowCells.setWrap(true);
    rowCells.setBorder(true, null, true, null, null, null, "#ddd", SpreadsheetApp.BorderStyle.SOLID);
    var rowCells = targetSheet.getRange("B" + borderCount + ":I" + borderCount);
    rowCells.setVerticalAlignment('middle');
    rowCells.setWrap(true);
    rowCells.setBorder(true, null, true, null, null, null, "#ddd", SpreadsheetApp.BorderStyle.SOLID);
    
    countMerge += 3;
  }
}

function column(char) {
 return char.charCodeAt(0) - 65; 
}
