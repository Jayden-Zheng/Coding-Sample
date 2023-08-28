/*
1. The function findAndReplaceByCountry() performs a "find and replace" for all the organization's address fetched from online sources. and if it contains any country name on the country list (another tab), they would be replaced by only the country name.
2. Similarly, the function findAndReplaceByCountry() works on the states within USA instead.
3. There are hardcoded parts in this code, including the Sheets' names, column indexes, and regex expressions.
4. To use only one function, please comment out the other in the defining and calling parts.
*/

function findAndReplaceByCountry() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName("TAB CONTAINING EACH INDIVIDUAL SPONSOR COMPANY");   // Hardcoded tab name
    var countryListSheet = spreadsheet.getSheetByName("TAB WITH COUNTRY NAMES");   // Hardcoded tab name
  
    var lastRowCountryList = countryListSheet.getLastRow();
    var columnIndex = 6;                                                  // Hardcoded column index
    // Extracting the part on the workbook that contains all the country names
    var countryList = countryListSheet.getRange(2, columnIndex, lastRowCountryList - 1, 1).getValues();
  
    countryList = countryList.filter(function (country) {
      return country[0] !== "";
    });
    // Convert the 2D list into a flat 1D array
    var countryNameArray = countryList.map(function (row) {
      return row[0];
    });
  
    // Checking: Verify the array is correctly set up: Output countryNameArray to cell J14
    countryListSheet.getRange("J14").setValue(countryNameArray.join(","));
  
  // Use an array to store all location info
    var columnC = sheet.getRange("C1:C" + sheet.getLastRow());           // Hardcoded output column
    var values = columnC.getValues();
  
  // for all cells selected, replace with the country name on countryNameArray
    for (var i = 0; i < values.length; i++) {
      var cellValue = values[i][0];
  
      for (var j = 0; j < countryList.length; j++) {
        var countryName = countryList[j][0];
        var regexPattern = new RegExp(".*" + countryName.trim() + ".*", "i");     // Hardcoded regex
  
        if (regexPattern.test(cellValue)) {
          var newValue = countryName;                                    // Hardcoded output
          columnC.getCell(i + 1, 1).setValue(newValue);
          break;
        }
      }
    }
  }
  
  
  function findAndReplaceByState() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName("Sheet name 1");   // Hardcoded tab name
    var countryListSheet = spreadsheet.getSheetByName("Sheet name 2");   // Hardcoded tab name
  
    var lastRowCountryList = countryListSheet.getLastRow();
    var stateNameList = countryListSheet.getRange(2, 7, lastRowCountryList, 1).getValues();
    var stateAbbreviationList = countryListSheet.getRange(2, 8, lastRowCountryList, 1).getValues();
  
    stateNameList = stateNameList.filter(function (state) {
      return state[0] !== "";
    });
  
    stateAbbreviationList = stateAbbreviationList.filter(function (abbr) {
      return abbr[0] !== "";
    });  
  
   
     var stateNameArray = stateNameList.map(function (row) {
      return row[0];
    });
  
     var stateAbbreviationArray = stateAbbreviationList.map(function (row) {
      return row[0];
    });
  
    // Checking: Verify the array is correctly set up: Output countryNameArray to some random non-impact cells
    countryListSheet.getRange("J15").setValue(stateNameArray.join(","));
    countryListSheet.getRange("J16").setValue(stateAbbreviationArray.join(","));
    countryListSheet.getRange("J17").setValue(lastRowCountryList);
  
    var columnC = sheet.getRange("C1:C" + sheet.getLastRow());
    var values = columnC.getValues();
  
  
  // State Name Matching
    for (var i = 0; i < values.length; i++) {
      var cellValue = values[i][0];
  
      for (var j = 0; j < stateNameList.length; j++) {
        var stateName = stateNameList[j][0];
  
        var regexPattern = new RegExp(".*" + stateName + ".*", "i");
  
        if (regexPattern.test(cellValue)) {
          var newValue = stateName + ", USA";
          columnC.getCell(i + 1, 1).setValue(newValue);
          break;
        }
      }
    }
  
    // State Abbreviation Matching
    for (var i = 0; i < values.length; i++) {
      var cellValue = values[i][0];
  
      for (var j = 0; j < stateAbbreviationList.length; j++) {
        var stateAbbreviation = stateAbbreviationList[j][0].toUpperCase();
  
        var regexPattern = new RegExp(".*" + stateAbbreviation + ".*");
  
        if (regexPattern.test(cellValue)) {
          var stateIndex = j;
          var stateName = stateNameList[stateIndex][0];
          var newValue = stateName + ", USA";
          columnC.getCell(i + 1, 1).setValue(newValue);
          break;
        }
      }
    }
  }
  
  // These functions must be runned on the custom menu on the workbook window since they directly change data on the workbook.
  function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Custom Menu')
      .addItem('findAndReplaceCountryAndState', 'runFunction')
      .addToUi();
  }
  
  function runFunction() {
    findAndReplaceByCountry();
    findAndReplaceByState();
  }
  
  