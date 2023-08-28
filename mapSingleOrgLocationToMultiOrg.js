/* This function is related to 2 sheets. One with all the rows, where each column A cell contains a list of sponsor organizations separated by "|"; the other with a location to each individual sponsor organization on each row.
The function does the work of filling in the locations for each clinical trials based on the information on the sheet with one location for each organization. It should be applied in the output cell.
*/

function getCountryList(companyNames, locationSheetName) {
    if (!companyNames || !locationSheetName) {
      return "";
    }
   
    var orgLocationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(locationSheetName);
    var orgLocationData = orgLocationSheet.getDataRange().getValues();
   
    // Store individual locations matched into a set for uniqueness
    var countrySet = new Set();
   
    companyNames.split("|").forEach(function(companyName) {
      orgLocationData.forEach(function(row) {
        if (row[0] === companyName) {
          countrySet.add(row[2].trim());
        }
      });
    });
   
    var countries = Array.from(countrySet);
   
    return countries.join("|");
  }
  