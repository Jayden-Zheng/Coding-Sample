
/* PLEASE READ:
Hard-coded elements: as shown in the code.
This function performs the following operations:
For each clinical trial (row) in column A in tab 1, extract each individual sponsor organization name, and output them uniquely each in one row in the column A of tab 2.
Note: Please check the hard-coded part before running function! It can overwrite existing sheet.
*/


function extractSponsors() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sponsorsSheet = ss.getSheetByName('TAB 1');                               // hard-coded
    var mappingSheet = ss.getSheetByName('TAB 2');                                // hard-coded
  
    var sponsorsData = sponsorsSheet.getRange('A2:A').getValues();                // hard-coded
  
    var sponsorsSet = new Set();
  
    for (var i = 0; i < sponsorsData.length; i++) {
      var sponsors = sponsorsData[i][0];
      if (sponsors !== "") {
        var sponsorsArr = sponsors.split('|');
        for (var j = 0; j < sponsorsArr.length; j++) {
          var sponsor = sponsorsArr[j].trim();
          if (!sponsorsSet.has(sponsor)) {
            sponsorsSet.add(sponsor);
          }
        }
      }
    }
  
    var sponsorsList = Array.from(sponsorsSet);
  
    if (sponsorsList.length > 0) {
      mappingSheet.getRange(1, 1, sponsorsList.length, 1).setValues(sponsorsList.map(function(sponsor) {
        return [sponsor];
      }));
      SpreadsheetApp.flush();
      Logger.log(sponsorsList.length + ' sponsors extracted successfully.');
    } else {
      Logger.log('No sponsors found.');
    }
  }
  