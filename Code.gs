function onOpen() {
  var UI = SpreadsheetApp.getUi();
  UI.createMenu("Singer's Tool")
  .addItem("Sort Responses", "parseAll")
  .addSeparator()
  .addItem("Clear CM/PM Tabs", "cleanSheet")
  .addToUi();
}

/* IDEAS:
- Add section for checking if email has been used more than once and change to update mode and not change mode
    -> Might actually not matter -> might want most updated response
- Maybe try adding a coulmn thats a link to a different spreadsheet so the formatted sheets can immediately be coppied over





*/

// Main function to parse all the information
function parseAll(){
  // Get the current spreadsheet and range
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var templatePM = spreadsheet.getSheetByName("Template - PM");
  var templateCM = spreadsheet.getSheetByName("Template - CM");
  var activeSheet = spreadsheet.getActiveSheet();
  var answerRange  = activeSheet.getDataRange();
  var answerVals = answerRange.getValues();
  for (var i = 1; i < answerRange.getHeight(); i++){
    // Start by checking if the sheet needs to be made new or just updated
    var newPM = spreadsheet.getSheetByName("PM " + answerVals[i][0]);
    var newCM = spreadsheet.getSheetByName("CM " + answerVals[i][0]);
  
    // If the sheets already exist, then update them instead of making new ones
    if (newPM != null){
      updatePM_(newPM);
    } else if (!(answerVals[i][0] === "")){
      templatePM.activate();
      newPM = spreadsheet.duplicateActiveSheet();
      newPM.setName("PM " + answerVals[i][0]);
      createPM_(newPM, answerVals[i]);
    }
    if (newCM != null){
      updateCM_(newCM);
    } else if (!(answerVals[i][0] === "")){
      templateCM.activate();
      newCM = spreadsheet.duplicateActiveSheet();
      newCM.setName("CM " + answerVals[i][0]);
      createCM_(newPM);
    }
  }
  
  // Move the two templates and the form responses to the front of the sheet list
  activeSheet.activate();
  spreadsheet.moveActiveSheet(1);
  templateCM.activate();
  spreadsheet.moveActiveSheet(1);
  templatePM.activate();
  spreadsheet.moveActiveSheet(1);
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
}

function updatePM_(currSheet){
  // Temporary //
  // Will eventaully handle updating and already existing sheet
  SpreadsheetApp.getActiveSpreadsheet().deleteSheet(currSheet);
}

function createPM_(newSheet, formVals){
  var testSheet = SpreadsheetApp.getActiveSheet();
  var testRange = testSheet.getDataRange();
  var testValues = testRange.getValues();
 
  // Begin by getting the range and values for the new spreadsheet
  var range = newSheet.getDataRange();
  var values = range.getValues();
  var skipTime = 0;
  for(var i = 0; i < formVals.length - 1; i++){
    if (i == 1){
      skipTime++;
    }
    values[1][i]  = formVals[skipTime];
    skipTime++;
  }
  range.setValues(values);
  newSheet.autoResizeColumn(5);
}

function updateCM_(currSheet){
  // Temporary //
  // Will eventaully handle updating and already existing sheet
  SpreadsheetApp.getActiveSpreadsheet().deleteSheet(currSheet);
}

function createCM_(newSheet){
  
}








// MIGHT NOT USE THESE!! STILL NOT SURE!
// This will return a map of key: value pairs specific to sponsor info
function getSponsorInfo_(){
  
}


// This will return a map of key: value pairs specific to tech info
function getTechInfo_(){
  
}


// This will return a map of key: value pairs specific to director info
function getDirectorInfo_(){
  
}