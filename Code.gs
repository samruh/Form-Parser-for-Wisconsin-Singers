function onOpen() {
  var UI = SpreadsheetApp.getUi();
  UI.createMenu("Singer's Tools")
  .addItem("Sort Responses", "parseAll")
  .addSeparator()
  .addItem("Clear CM/PM Tabs", "cleanSheet")
  .addToUi();
}

/* IDEAS:
- Add section for checking if email has been used more than once and change to update mode and not change mode
    -> Might actually not matter -> might want most updated response
- Maybe try adding a coulmn thats a link to a different spreadsheet so the formatted sheets can immediately be coppied over
- How to store version info of multiple submits?
- Make a new section for each new choir director submission -> duplicate certain section of template




*/

// Main function to parse all the information
function parseAll(){
  
  // Make a bunch of variables
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // The entire spreadsheet
  var templatePM = spreadsheet.getSheetByName("Template - PM"); // the PM template sheet
  var templateCM = spreadsheet.getSheetByName("Template - CM"); // the CM template sheet
  var activeSheet = spreadsheet.getActiveSheet(); // The active sheet (should always be the form responses)
  var answerRange  = activeSheet.getDataRange(); // Get the entire range the has info from the active sheet
  var answerVals = answerRange.getValues(); // Get the raw values in a 2D Array
  
  // main loop to run through each row of data
  // Start at 3 because of the questions row and the tag row and row/column counts start at 1
  for (var i = 3; i < answerRange.getHeight() + 1; i++){

    // Start by checking if the sheet needs to be made new or just updated
    // i-1 is to accomodate the fact that ranges start at 1 and arrays start at 0
    // the toLowerCase is simply to make sure weird capitalization in the school/location cell won't break the code
    var newPM = spreadsheet.getSheetByName("PM " + answerVals[i-1][1].toLowerCase());
    var newCM = spreadsheet.getSheetByName("CM " + answerVals[i-1][1].toLowerCase());
    
    // get the checkbox Range (really just one cell, but it is a range object)
    var checkbox = activeSheet.getRange(i, 1);
    var toParse = !(checkbox.isChecked()) && (answerVals[i-1][1] !== "")
  
    // Logic to check is the person who answered is also a choir director
    var isChoirDir = true;
    
    // If the sheets already exist, then update them instead of making new ones
    // else make sure the "done box" is not checked and the school/location cell is filled in
    // if both are satisified then copy the template and change it's name
    if (newPM != null && toParse){
      if (isChoirDir) {
        var choirDirRange = spreadsheet.getRangeByName("ChoirDir");
        var choirFormMap = getChoirInfo(answerVals[1], answerVals[i-1], i);
        updateInfoWithNewChoir_(choirDirRange, newPM, i, choirFormMap);
      } else {
        var formMap = getInfo_(answerVals[1], answerVals[i-1]);
        sortInfo_(newPM, formMap);
      }
    } else if (toParse){
      templatePM.activate();
      newPM = spreadsheet.duplicateActiveSheet();
      newPM.setName("PM " + answerVals[i-1][1].toLowerCase());
      var formMap = getInfo_(answerVals[1], answerVals[i-1]);
      sortInfo_(newPM, formMap);
    }
    
    
    
    // Same logic as above just for the CM page
    if (newCM != null && toParse){
      //updateCM_(newCM);
    } else if (toParse){
      templateCM.activate();
      newCM = spreadsheet.duplicateActiveSheet();
      newCM.setName("CM " + answerVals[i-1][1].toLowerCase());
      var formMap = getInfo_(answerVals[1], answerVals[i-1]);
      sortInfo_(newPM, formMap);
    }
  }
  
  // Move the two templates and the form responses to the front of the sheet list
  activeSheet.activate();
  spreadsheet.moveActiveSheet(1);
  templateCM.activate();
  spreadsheet.moveActiveSheet(1);
  templatePM.activate();
  spreadsheet.moveActiveSheet(1);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName("Form Responses 1")); // Just to make sure the active sheet is always the Form Response page when the script is finished
}




function getInfo_(tagRow, infoRow){
  var formMap = new Map();
  for (var i=3; i < infoRow.length; i++){
    if (tagRow[i] !== "n/a"){
      formMap.set(tagRow[i].toLowerCase(), infoRow[i]);
    }
  }
  return formMap;
}

function getChoirInfo_(tagRow, infoRow, nameNum){
  var formMap = new Map();
  for (var i=3; i < infoRow.length; i++){
    if (tagRow[i] !== "n/a"){
      if (tagRow[i].endsWith("~")) {
        formMap.set(tagRow[i].toLowerCase() + nameNum, infoRow[i]);
      } else {
        formMap.set(tagRow[i].toLowerCase(), infoRow[i]);
      }
    }
  }
  return formMap;
}



function sortInfo_(newSheet, formDataMap){
  var testSheet = SpreadsheetApp.getActiveSheet();
  var testRange = testSheet.getDataRange();
  var testValues = testRange.getValues();
  var testNamedRanges = testSheet.getNamedRanges()
  
  // Begin by getting an array of the named ranges from the template and make a new name filter
  // ** NamedRanges are different than ranges!
  var namedRanges = newSheet.getNamedRanges();
  var rangeNameFilter = "'" + newSheet.getName() + "'!";
  
  namedRanges.forEach(function(specNamedRange) {
    var specRange = specNamedRange.getRange();
    var mapRangeName = specNamedRange.getName().replace(rangeNameFilter, "").toLowerCase();
    if (formDataMap.has(mapRangeName)) {
      var rangeVals = specRange.getValues();
      rangeVals[1][0] = formDataMap.get(mapRangeName);
      specRange.setValues(rangeVals);
    }
  });
  
  // might not want this, but we will see
  newSheet.autoResizeRows(1, newSheet.getMaxRows());
}


// TODO: Figure out a decent way to update the pages
// This will generally function the same as create page, except it will check if it is choirDir info and will update accordingly
function updateInfo_(currSheet){
  // Temporary //
  // Will eventaully handle updating and already existing sheet
  //SpreadsheetApp.getActiveSpreadsheet().deleteSheet(currSheet);
}


// TODO: Update this method to make all the new namedRanges in the proper place with "nameNum" attached to the end of it
function updateInfoWithNewChoir_(choirDirRange, newSheet, nameNum, choirFormMap){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  //setNamedRange(name, range)
  var newSheetRange = newSheet.getDataRange();
  choirDirRange.copyFormatToRange(newSheet, 1, choirDirRange.getWidth(), newSheetRange.getHeight() + 1, newSheetRange.getHeight() + 1 + choirDirRange.getHeight());
  choirDirRange.copyValuesToRange(newSheet, 1, choirDirRange.getWidth(), newSheetRange.getHeight() + 1, newSheetRange.getHeight() + 1 + choirDirRange.getHeight());
  sortInfo_(newSheet, choirFormMap);
}



// This function will delete all the CM/PM tabs for easier clean up after they have been coppied away
function cleanSheet(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var tabs = spreadsheet.getSheets();
  for (var i = 0; i < tabs.length; i++){
    if (tabs[i].getSheetName().startsWith("PM") || tabs[i].getSheetName().startsWith("CM")){
      spreadsheet.deleteSheet(tabs[i]);
    }
  }
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName("Form Responses 1"));
}