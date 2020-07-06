function onOpen() {
  var UI = SpreadsheetApp.getUi();
  UI.createMenu("Singer's Tools")
  .addItem("Sort Responses", "parseAll")
  .addSeparator()
  .addItem("Clear CM/PM Tabs", "cleanSheet")
  .addSeparator()
  .addItem("Archive Done Tabs", "archiveFinishedRows")
  .addToUi();
}

/* IDEAS:
- How to store version info of multiple submits?
- Make a new section for each new choir director submission -> duplicate certain section of template
     -> Possibly compare A1 notations to figure out if a certain namedRange is within the ChoirDir range
- add more special formatted cells in sort info -> possibly make new method that gets called containg siwtch statement??
- put archived data into some sort of order -> maybe alphabetical or in order of time or something else entirely
- add in error checking to make sure the user is on the right sheet when starting to parse info and when starting to archive data
- do test run with actual form responses
- make actual form template finally -> ask Josh about his template too
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
    var toParse = !(checkbox.isChecked()) && (answerVals[i-1][1] !== "");
    
  
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
    
    // update the checkbox to true
    //checkbox.setValue(true);
  }
  
  // Move the two templates and the form responses to the front of the sheet list
  activeSheet.activate();
  spreadsheet.moveActiveSheet(1);
  templateCM.activate();
  spreadsheet.moveActiveSheet(1);
  templatePM.activate();
  spreadsheet.moveActiveSheet(1);
  spreadsheet.getSheetByName("Archive").activate();
  spreadsheet.moveActiveSheet(1);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName("Form Responses 1")); // Just to make sure the active sheet is always the Form Response page when the script is finished
}




function getInfo_(tagRow, infoRow){
  var formMap = new Map();
  for (var i=2; i < infoRow.length; i++){
    if (tagRow[i] !== "n/a"){
      formMap.set(tagRow[i].toLowerCase(), infoRow[i]);
    }
  }
  return formMap;
}

function getChoirInfo_(tagRow, infoRow, nameNum){
  var formMap = new Map();
  for (var i=2; i < infoRow.length; i++){
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
    Logger.log(mapRangeName);
    if (formDataMap.has(mapRangeName)) {
      var rangeVals = specRange.getValues();
      Logger.log("Data for name above was found in map");
      rangeVals[1][0] = formDataMap.get(mapRangeName);
      specRange.setValues(rangeVals);
      if (mapRangeName.includes("_")){
        Logger.log("Made it inside of here!");
        switch (mapRangeName.split("_")[1]) {
          case "date": 
            specRange.setNumberFormat("dddd, mmmm d yyy at h:mm am/pm");
            break;
          default:
            throw new Error('Something went horribly wrong!!!');
        }
      }
    }
  });
  
  // might not want this, but we will see
  newSheet.autoResizeColumns(1, newSheet.getMaxColumns());
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


// This function will move all the rows marked done to an archive sheet
function archiveFinishedRows(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var formResponse = spreadsheet.getActiveSheet();
  var formResponseRange = formResponse.getDataRange();
  var formResponseValues = formResponseRange.getValues();
  var archive = spreadsheet.getSheetByName("Archive");
  var archiveNextRange = archive.getRange(archive.getDataRange().getHeight() + 1, 1, archive.getMaxRows(), formResponseRange.getWidth());
  var archiveNextValues = archiveNextRange.getValues();
  
  
  var archivePos = 0;
  for (var i = 3; i < formResponseRange.getHeight() + 1; i++) {
    var checkbox = formResponse.getRange(i - archivePos, 1);
    if (checkbox.isChecked()) {
      archiveNextValues[archivePos] = formResponseValues[i-1];
      formResponse.deleteRow(i - archivePos);
      archivePos ++;  
    }
  }
  archiveNextRange.setValues(archiveNextValues);
  
  
  var finishedArchiveRange = archive.getRange(3, 1, archive.getDataRange().getHeight() - 2, 1);
  for (var i = 1; i < finishedArchiveRange.getHeight() + 1; i++){
    var currCell = finishedArchiveRange.getCell(i,1);
    var criteria = SpreadsheetApp.DataValidationCriteria.CHECKBOX;
    var rule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .build();
    currCell.setDataValidation(rule);
  }
  
}