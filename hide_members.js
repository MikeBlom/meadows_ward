function hideRow(sName, minAge, maxAge, columnNum) {
  nightlyRowHide("Members", 11, 100, 5);
}

nightlyRowHide = function(sName, minAge, maxAge, columnNum) {
  var ss = SpreadsheetApp.openById('1K_hK9Q8AOSfHxZn5Uu-rKDRwBR3QboKRT1mTUtp2o_I');
  var s = ss.getSheetByName(sName);
  var range = s.getDataRange();
  var numRows = range.getNumRows();
  s.showRows(1, numRows);
  var hideRow = [];
  var lastHideRow = 0;
  for (var i = 1; i <= numRows; i++) {
    var cell = s.getRange(i, columnNum);
    var currentValue = cell.getValues();
    if( currentValue <= minAge || currentValue >= maxAge) {
      var lastInArray = hideRow[hideRow.length - 1];
      if(lastInArray != undefined && i - lastInArray[lastHideRow] == lastHideRow ){
        lastInArray[lastHideRow] += 1;
      } else {
        lastHideRow = i;
        var newHash = {};
        newHash[lastHideRow] = 1;
        hideRow.push(newHash);
      }
    }
  }

  hideRow.forEach(function(row, i){
    var rowToHide = parseInt(Object.keys(row)[0]);
    var numOfRows = row[rowToHide];
    s.hideRows(rowToHide, numOfRows);
  })
}
