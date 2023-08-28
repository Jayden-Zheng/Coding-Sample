/* The function has the following arguments:
range: the range in the tab containing the countries table, where a certain portion of each column contains all the country names for the specific region;
columnIndex: the column corresponding to the desired areas and phase;
trialType: the trial type desired.
displayType: the format in which data is displayed in the output table. Options include "Count", "Location Set", "Location & Count", and "Sponsor List". The data displayed is intuitive.
The function performs the calculation based on the conditions and requirements provided as arguments.
*/

function sumMatchingValues(range, columnIndex, trialType, displayType) {
    if (!range || columnIndex === null || !trialType || !displayType) {
      return 0; // Return 0 if range is null or empty
    }
  
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName('COUNTRIES');      // hardcoded tab name
    var regionToMatch = sheet.getRange(range).getValues().flat().filter(String);
  
    var data = spreadsheet.getSheetByName('SPONSORS').getDataRange().getValues();
    var filteredData = data.filter(function(row) {
      return trialType.includes(row[2]) && row[columnIndex];          // hardcoded the column index for trialType
    });
   
    if (displayType === "Count") {
      var sum = 0;
  // for each row
      filteredData.forEach(function(row) {
  // find matches, add to the count, and avoid accidental mistaken matching for special cases
        if (regionToMatch.some(function(text) {return row[1].includes(text) && !(text==="India" && row[1] === "Indiana, USA") && !(text=== "Mexico" && row[1] ==="New Mexico, USA"); })) {
          sum += row[columnIndex];
        }
      });
      return sum;
  
  
    } else if (displayType === "Location Set") {
      var locationSet = new Set();
      filteredData.forEach(function(row) {
  // find the matching rows and get the location set
        if (regionToMatch.some(function(text) {return row[1].includes(text) && !(text==="India" && row[1] === "Indiana, USA") && !(text=== "Mexico" && row[1] ==="New Mexico, USA"); })) {
          var locations = row[1].split("|");
  // for each location in the location list for the clinical trial
          locations.forEach(function(location) {
  // if the location INCLUDES any country name from the region country list, add it to the set for output
            if (regionToMatch.some(function(text) {return location.includes(text) && !(text==="India" && location === "Indiana, USA") && !(text=== "Mexico" && location === "New Mexico, USA"); })) {
              locationSet.add(location.trim());
            }
          });
        }
      });
      let locationArray = [...locationSet];
      locationArray = locationArray.map(location => {
  // For USA states, remove the abbreviations since it is unlikely to cause confusions for business analysis purposes.
        if (location.includes(", USA")) {
          return location.replace(", USA", "");
        }
        return location;
      });
      return locationArray.join(" | ");
  
  
    } else if (displayType === "Location & Count") {
  // Perform similar operations, only that the output tracks the time that each country appears
      var locationCount = {};
  
      filteredData.forEach(function(row) {
        if (regionToMatch.some(function(text) { return row[1].includes(text) && !(text === "India" && row[1] === "Indiana, USA") && !(text=== "Mexico" && row[1] ==="New Mexico, USA"); })) {
          var locations = row[1].split("|");
          locations.forEach(function(location) {
            if (regionToMatch.some(function(text) { return location.includes(text) && !(text === "India" && location === "Indiana, USA") && !(text === "Mexico" && location === "New Mexico, USA"); })) {
              var trimmedLocation = location.trim();
              if (locationCount[trimmedLocation]) {
                locationCount[trimmedLocation]++;
              } else {
                locationCount[trimmedLocation] = 1;
              }
            }
          });
        }
      });
      var result = [];
      for (var key in locationCount) {
        if (locationCount.hasOwnProperty(key)) {
          result.push(key + " (" + locationCount[key] + ")");
        }
      }
  
      // Remove ", USA" from keys
      result = result.map(function(item) {
        if (item.includes(", USA")) {
          return item.replace(", USA", "");
        }
        return item;
      });
  
      return result.join(" | ");
  
  
  
    } else if (displayType === "Sponsor List") {
      var sponsorSet = new Set();
      filteredData.forEach(function(row) {
        if (regionToMatch.some(function(text) {return row[1].includes(text) && !(text==="India" && row[1] === "Indiana, USA") && !(text=== "Mexico" && row[1] ==="New Mexico, USA"); })) {
          sponsorSet.add(row[0]);
        }
      });
      return Array.from(sponsorSet).join(", ");
  } else {
      return 0;
    }
  }
  