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
- ask Josh about his template
*/

// Main function to parse all the information
function parseAll(){
  
  // Make a bunch of variables
  var UI = SpreadsheetApp.getUi(); // The UI for this Document
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // The entire spreadsheet
  var templatePM = spreadsheet.getSheetByName("Template - PM"); // the PM template sheet
  var templateCM = spreadsheet.getSheetByName("Template - CM"); // the CM template sheet
  var activeSheet = spreadsheet.getActiveSheet(); // The active sheet (should always be the form responses)
  var answerRange = activeSheet.getDataRange(); // Get the entire range the has info from the active sheet
  var answerVals = answerRange.getValues(); // Get the raw values in a 2D Array
  var didParse = false; // This will be set to true if any information has been parsed, used to speed up code when no parse happens
  
  // This is purely to ensure that the user can't break the program by sorting from a non-"form response" sheet
  if (activeSheet.getName() !== "Form Responses 1"){
    UI.alert("ERROR:\nPlease make sure \"Form Responses 1\" is the active sheet.");
    return;
  }
 
  // main loop to run through each row of data
  // Start at 3 because of the questions row and the tag row and row/column counts start at 1
  for (var i = 3; i < answerRange.getHeight() + 1; i++){
    
    // get the checkbox Range (really just one cell, but it is a range object)
    // Now check to make to if the current row needs to be processed or not
    var checkbox = activeSheet.getRange(i, 1);
    var toParse = !(checkbox.isChecked()) && (answerVals[i-1][1] !== "");
    
    if (toParse){
      didParse = true; // This will make sure to rearrange the pages after finishing parsing
      
      // Start by checking if the sheet needs to be made new or just updated
      // i-1 is to accomodate the fact that ranges start at 1 and arrays start at 0
      // the toLowerCase is simply to make sure weird capitalization in the school/location cell won't break the code
      var newPM = spreadsheet.getSheetByName("PM " + answerVals[i-1][1].toLowerCase());
      var newCM = spreadsheet.getSheetByName("CM " + answerVals[i-1][1].toLowerCase());
      
      // Call function to check for sponsor, tech, and choir director
      var answerLength = answerRange.getWidth();
      var answerArr = getPersonType_(activeSheet.getRange(2, 1, 1, answerLength), activeSheet.getRange(i, 1, 1, answerLength));
      var isSponsor = answerArr[0];
      var isTech = answerArr[1];
      var isChoirDir = answerArr[2];
      
      // If the sheets already exist, then update them instead of making new ones
      // if neither exist then copy the template and change it's name
      if (newPM != null && newCM != null){
        var formMap;
        if (isChoirDir) {
          var choirDirRange = spreadsheet.getRangeByName("allchoirranges");
          formMap = getInfo_(answerVals[1], answerVals[i-1], i);
          updateInfoWithNewChoir_(choirDirRange, newPM, i, formMap, answerArr);
          
          // sortInfo for the CM page since the CM doesn't need choir director info
          sortInfo_(newCM, formMap, answerArr, i, false);
        } else {
          // get the info into a map
          formMap = getInfo_(answerVals[1], answerVals[i-1], -1);
          
          // call sortInfo for the pm page and the cm page
          sortInfo_(newPM, formMap, answerArr, -1, false);
          sortInfo_(newCM, formMap, answerArr, -1, false);
        }
      } else if ((newPM !=  null) ? !(newCM != null) : (newCM != null)) {
        UI.alert("ERROR:\nThere is one of the two tabs for the " + answerVals[i-1][1].toLowerCase() + " show missing and one not.\n" +
        "Try deleting the exisiting tab for the  " + answerVals[i-1][1].toLowerCase() + " show and unchecking all rows for that show.");
        return;
      } else {
        // Get the info
        formMap = getInfo_(answerVals[1], answerVals[i-1]);
        
        // Make the new PM sheet
        templatePM.activate();
        newPM = spreadsheet.duplicateActiveSheet();
        newPM.setName("PM " + answerVals[i-1][1].toLowerCase());
        sortInfo_(newPM, formMap, answerArr, -1, false);
        
        // Make the new CM sheet
        templateCM.activate();
        newCM = spreadsheet.duplicateActiveSheet();
        newCM.setName("CM " + answerVals[i-1][1].toLowerCase());
        sortInfo_(newCM, formMap, answerArr, -1, false);
      }

      // update the checkbox to true
      checkbox.setValue(true);
    }
  }
  
  if (didParse) {
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
}

// This is a helper function designed to check what type of person is filling out the form
function getPersonType_(tagRange, dataRange){
  // Array to tell who is filling out the form -> sponsor, tech, director
  var personTypeArr = [false, false, false];
  var dataValues = dataRange.getValues();
  var tagValues = tagRange.getValues();
  for (var i = 0; i < tagValues[0].length; i++){
    if (tagValues[0][i].includes("sponsor-") || tagValues[0][i].includes("tech-") || tagValues[0][i].includes("director-")){
      if (dataValues[0][i].toLowerCase().includes("sponsor")){
        personTypeArr[0] = true;
      }
      if (dataValues[0][i].toLowerCase().includes("tech")){
        personTypeArr[1] = true;
      }
      if (dataValues[0][i].toLowerCase().includes("director")){
        personTypeArr[2] = true;
      }
    }
  }
  return personTypeArr;
}


function getInfo_(tagRow, infoRow, nameNum){
  var formMap = new Map();
  for (var i=2; i < infoRow.length; i++){
    if (tagRow[i] !== "n/a"){
      if (nameNum > 1) formMap.set(tagRow[i].toLowerCase() + nameNum, infoRow[i]);
      else formMap.set(tagRow[i].toLowerCase(), infoRow[i]);
    }
  }
  return formMap;
}


function sortInfo_(newSheet, formDataMap, personType, nameNum, isChoirSort){
  var testSheet = SpreadsheetApp.getActiveSheet();
  var testRange = testSheet.getDataRange();
  var testValues = testRange.getValues();
  var testNamedRanges = testSheet.getNamedRanges();
  
  // Begin by getting an array of the named ranges from the template and make a new name filter
  // ** NamedRanges are different than ranges!
  var namedRanges = newSheet.getNamedRanges();
  var rangeNameFilter = "'" + newSheet.getName() + "'!";
  
  namedRanges.forEach(function(specNamedRange) {
    var specRange = specNamedRange.getRange();
    var mapRangeName = specNamedRange.getName().replace(rangeNameFilter, "").toLowerCase();
    var suffix = "";
    var hasUnder = false;
    if (mapRangeName.includes("_")){
      suffix = mapRangeName.split("_")[1];
      mapRangeName = mapRangeName.split("_")[0];
      hasUnder = true;
    }
    
    if (nameNum > 0){
      if (isChoirSort) {
        //var specRangeA1Num = specRange.getA1Notation().split(":")[0].match(/\d+/g)[0];
        if (hasUnder && ( suffix.match(/\d+/g)[0] !== "")){
          mapRangeName = mapRangeName + String.toString(suffix.match(/\d+/g)[0]);
          suffix = String.toString(suffix.match(/[a-zA-Z]+/g)[0]);
          
        } else if (hasUnder){
          mapRangeName = mapRangNum + String.toString(nameNum);
        }
      } else {
        mapRangeName = mapRangeNum + String.toString(nameNum);
      }
    }
    
    if (formDataMap.has(mapRangeName)) {
      var formData = formDataMap.get(mapRangeName);
      if (hasUnder){
        switch (suffix) {
          case "sponsor":
            if (personType[0]) setValues_(specRange,formData);
            break;
          case "tech":
            if (personType[1]) setValues_(specRange,formData);
            break;
          case "choir":
            if (personType[2]) setValues_(specRange,formData);
            break;
          case "date": 
            setValues_(specRange,formData);
            specRange.setNumberFormat("dddd, mmmm d yyy at h:mm am/pm");
            break;
          default:
            break;
        }
      } else {
        setValues_(specRange, formData);
      }
    }
  });
}

function setValues_(specRange, formData){
   var rangeVals = specRange.getValues();
   rangeVals[1][0] = formData;
   specRange.setValues(rangeVals);
}


// TODO: Update this method to make all the new namedRanges in the proper place with "nameNum" attached to the end of it
function updateInfoWithNewChoir_(choirDirRange, newSheet, nameNum, choirFormMap, personType){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var newSheetRange = newSheet.getDataRange();
  choirDirRange.copyFormatToRange(newSheet, 1, choirDirRange.getWidth(), newSheetRange.getHeight() + 2, newSheetRange.getHeight() + 2 + choirDirRange.getHeight());
  choirDirRange.copyValuesToRange(newSheet, 1, choirDirRange.getWidth(), newSheetRange.getHeight() + 2, newSheetRange.getHeight() + 2 + choirDirRange.getHeight());
  var wholeRange = newSheet.getDataRange();
  
  var newRange = newSheet.getRange(wholeRange.getHeight() - choirDirRange.getHeight() + 2, 1, choirDirRange.getHeight(), choirDirRange.getWidth());
  spreadsheet.setNamedRange("allchoirranges" + nameNum, newRange);
  var newRangeA1 = newRange.getA1Notation();
  var oldRangeA1 = choirDirRange.getA1Notation();
  var templatePM = spreadsheet.getSheetByName("Template - PM");
  var templateNames = templatePM.getNamedRanges();
  templateNames.forEach(function(specNamedRange) {
    var specRange = specNamedRange.getRange();
    var specRangeA1 = specRange.getA1Notation();
    var specRangeA1TopLeft = specRangeA1.split(":")[0];
    var oldRangeA1TopLeft = oldRangeA1.split(":")[0];
    if (parseInt(specRangeA1TopLeft.match(/\d+/g)[0]) > parseInt(oldRangeA1TopLeft.match(/\d+/g)[0])) {
      Logger.log(specNamedRange.getName() + "\n" + specRangeA1TopLeft.match(/\d+/g)[0] + " > " + oldRangeA1TopLeft.match(/\d+/g)[0]);
      var newRangeA1TopLeft = newRangeA1.split(":")[0];
      var A1TopNumDif = parseInt(specRangeA1TopLeft.match(/\d+/g)[0]) - parseInt(oldRangeA1TopLeft.match(/\d+/g)[0]);
      var A1TopLetDif = specRangeA1TopLeft.match(/[a-zA-Z]+/g)[0].charCodeAt(0) - oldRangeA1TopLeft.match(/[a-zA-Z]+/g)[0].charCodeAt(0);
      var A1BottomNumDif = specRange.getHeight() - 1;
      var A1BottomLetDif = specRange.getWidth() - 1;
      var newA1TopLeftNum = parseInt(newRangeA1TopLeft.match(/\d+/g)[0] , 10) + A1TopNumDif;
      var newA1TopLeftLet = newRangeA1TopLeft.match(/[a-zA-Z]+/g)[0].charCodeAt(0) + A1TopLetDif;
      var newA1BottomRightNum = newA1TopLeftNum + A1BottomNumDif;
      var newA1BottomRightLet = newA1TopLeftLet + A1BottomLetDif;
      var newA1TopLeft = String.fromCharCode(newA1TopLeftLet) + newA1TopLeftNum.toString();
      var newA1BottomRight = String.fromCharCode(newA1BottomRightLet) + newA1BottomRightNum.toString();
      var newA1 = newA1TopLeft + ":" + newA1BottomRight;
      Logger.log(newA1 + "\n");
      spreadsheet.setNamedRange(specNamedRange.getName() + nameNum, newSheet.getRange(newA1));
      // Just need to get range with newA1 and give it the name from specRange with nameNum on the end
    }
  });
  sortInfo_(newSheet, choirFormMap, personType, nameNum, true);
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
  
  if (formResponse.getName() !== "Form Responses 1"){
    throw new Error("Please make sure \"Form Responses 1\" is the active sheet.");
  }
  
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