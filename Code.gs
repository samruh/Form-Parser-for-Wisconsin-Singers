// This automatically gets called on opening the spreadsheet
function onOpen() {
  var UI = SpreadsheetApp.getUi();
  
  // create the custom menu to call the functions in this program
  UI.createMenu("Singer's Tools")
  .addItem("Sort Responses", "parseAll")
  .addSeparator()
  .addItem("Clear CM/PM Tabs", "cleanSheet")
  .addSeparator()
  .addItem("Archive Done Tabs", "archiveFinishedRows")
  .addToUi();
}


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
        
        // This handles adding a new choir director section to hold multiple instances of the choir director data
        if (isChoirDir) {
          var choirDirRange = spreadsheet.getRangeByName("allchoirranges");
          formMap = getInfo_(answerVals[1], answerVals[i-1], i);
          updateInfoWithNewChoir_(choirDirRange, newPM, i, formMap, answerArr); // This function calls sort for the PM tabs after adding the new choir section
          
          // sortInfo for the CM page since the CM doesn't need choir director info
          sortInfo_(newCM, formMap, answerArr, i, -1);
        } else {
          // get the info into a map
          formMap = getInfo_(answerVals[1], answerVals[i-1], -1);
          
          // call sortInfo for the pm page and the cm page
          sortInfo_(newPM, formMap, answerArr, -1, -1);
          sortInfo_(newCM, formMap, answerArr, -1, -1);
        }
      } else if ((newPM !=  null) ? !(newCM != null) : (newCM != null)) {
        // There should never be a time when one of the two pages for a given show exist and one does not
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
        sortInfo_(newPM, formMap, answerArr, -1, -1);
        
        // Make the new CM sheet
        templateCM.activate();
        newCM = spreadsheet.duplicateActiveSheet();
        newCM.setName("CM " + answerVals[i-1][1].toLowerCase());
        sortInfo_(newCM, formMap, answerArr, -1, -1);
      }

      // update the checkbox to true
      checkbox.setValue(true);
    }
  }
  
  if (didParse) {
    // Move the two templates,the archive, and the form responses to the front of the sheet list
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
  // Array to tell who is filling out the form -> [0] = sponsor, [1] = tech, [2] = director
  var personTypeArr = [false, false, false];
  var dataValues = dataRange.getValues();
  var tagValues = tagRange.getValues();
  
  // Loop through and check for every sponsor, tech, and director and change the array values accordingly
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

// This is a helper function that sorts all the row's data into a map with that tag row
function getInfo_(tagRow, infoRow, nameNum){
  var formMap = new Map();
  for (var i=2; i < infoRow.length; i++){
    if (tagRow[i] !== "n/a" && infoRow[i] !== ""){
      
      // This will handle setting up the map for general cases and for the choir director case
      if (nameNum > 1) formMap.set(tagRow[i].toLowerCase() + nameNum, infoRow[i]);
      else formMap.set(tagRow[i].toLowerCase(), infoRow[i]);
    }
  }
  return formMap;
}

// This is the main helper function that sorts all the information from the form responses to the new sheet
function sortInfo_(newSheet, formDataMap, personType, nameNum, choirA1){  
  // Begin by getting an array of the named ranges from the template and make a new name filter
  // ** NamedRanges are different than ranges!
  var namedRanges = newSheet.getNamedRanges();
  var rangeNameFilter = "'" + newSheet.getName() + "'!";
  
  // Loop through all the namedRanges and put data into the correct named ranges
  namedRanges.forEach(function(specNamedRange) {
    var specRange = specNamedRange.getRange();
    var mapRangeName = specNamedRange.getName().replace(rangeNameFilter, "").toLowerCase();
    var suffix = "";
    var hasUnder = false;
    
    // Check if the namedRange has a special tag and if it does then seperate it
    if (mapRangeName.includes("_")){
      suffix = mapRangeName.split("_")[1];
      mapRangeName = mapRangeName.split("_")[0];
      hasUnder = true;
    }
    
    // Check if this is a choir sort or not
    if (nameNum > 0){
      // Check if this is a cm or pm update (only the pm update will need to worry about the choir section)
      if (choirA1 > 0) {
        var specA1 = specRange.getA1Notation().split(":")[0].match(/\d+/g)[0];
        // Check if the current named range is within the choir range or not
        if (parseInt(specA1) > parseInt(choirA1)) {
          var suffixNum = suffix.match(/\d+/g);
          // See if the named range has a suffix and if the suffix has a number or not, if it has both then move the number to the main word and remove it from the suffix
          if (hasUnder && (suffixNum != null)){
            mapRangeName = mapRangeName + suffixNum[0].toString();
            suffix = (suffix.match(/[a-zA-Z]+/g)[0]).toString();
          } 
        } else {
          mapRangeName = mapRangeName + nameNum.toString();
        }
      } else {
        mapRangeName = mapRangeName + nameNum.toString();
      }
    }
    
    // Now check if the range name actuall has data in the map we created earlier
    if (formDataMap.has(mapRangeName)) {
      var formData = formDataMap.get(mapRangeName);
      
      // Check if a special tag was found and handle it correctly
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

// simple little helper function to write the info to the new sheet
function setValues_(specRange, formData){
   var rangeVals = specRange.getValues();
   rangeVals[1][0] = formData;
   specRange.setValues(rangeVals);
}


// This helper function will add a new copy of the choir range and set each of the named ranges within it correctly based on A1 notation calculations 
function updateInfoWithNewChoir_(choirDirRange, newSheet, nameNum, choirFormMap, personType){
  // Start by copying all the format and info from the template to the new sheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var newSheetRange = newSheet.getDataRange();
  choirDirRange.copyFormatToRange(newSheet, 1, choirDirRange.getWidth(), newSheetRange.getHeight() + 2, newSheetRange.getHeight() + 2 + choirDirRange.getHeight());
  choirDirRange.copyValuesToRange(newSheet, 1, choirDirRange.getWidth(), newSheetRange.getHeight() + 2, newSheetRange.getHeight() + 2 + choirDirRange.getHeight());
  var wholeRange = newSheet.getDataRange();
  
  // Get the range on the newly coppied choir range
  var newRange = newSheet.getRange(wholeRange.getHeight() - choirDirRange.getHeight() + 2, 1, choirDirRange.getHeight(), choirDirRange.getWidth());
  spreadsheet.setNamedRange("allchoirranges" + nameNum, newRange);
  var newRangeA1 = newRange.getA1Notation();
  var oldRangeA1 = choirDirRange.getA1Notation();
  var templatePM = spreadsheet.getSheetByName("Template - PM");
  var templateNames = templatePM.getNamedRanges();
  
  // look through each namedRange in the template and for everyone that is within the choir range add a new named range with nameNum on the end
  templateNames.forEach(function(specNamedRange) {
    var specRange = specNamedRange.getRange();
    var specRangeA1 = specRange.getA1Notation();
    var specRangeA1TopLeft = specRangeA1.split(":")[0];
    var oldRangeA1TopLeft = oldRangeA1.split(":")[0];
    if (parseInt(specRangeA1TopLeft.match(/\d+/g)[0]) > parseInt(oldRangeA1TopLeft.match(/\d+/g)[0])) {
      
      // This is a lot of variables that basically find the correct, relative place of the namedRanges from the template in the new choir range copy
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
      
      // Make the new namedRange, finally
      spreadsheet.setNamedRange(specNamedRange.getName() + nameNum, newSheet.getRange(newA1));
    }
  });
  
  // Call sort info to process the PM sheet
  sortInfo_(newSheet, choirFormMap, personType, nameNum, oldRangeA1.split(":")[0].match(/\d+/g)[0]);
}


// This function will delete all the CM/PM tabs for easier clean up after they have been coppied away
function cleanSheet(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var tabs = spreadsheet.getSheets();
  
  // loop through all the tabs and delete any of them that start with PM or CM
  for (var i = 0; i < tabs.length; i++){
    if (tabs[i].getSheetName().startsWith("PM") || tabs[i].getSheetName().startsWith("CM")){
      spreadsheet.deleteSheet(tabs[i]);
    }
  }
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName("Form Responses 1")); // Just to encrouage keeping Form Responses 1 as the active sheet
}


// This function will move all the rows marked done to an archive sheet
function archiveFinishedRows(){
  
  // Make a bunch of variables for use in the upcoming loops
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var formResponse = spreadsheet.getActiveSheet();
  var formResponseRange = formResponse.getDataRange();
  var formResponseValues = formResponseRange.getValues();
  var archive = spreadsheet.getSheetByName("Archive");
  var archiveNextRange = archive.getRange(archive.getDataRange().getHeight() + 1, 1, archive.getMaxRows(), formResponseRange.getWidth());
  var archiveNextValues = archiveNextRange.getValues();
  
  // Error checking to make sure the user is on the correct page when trying to archive information
  if (formResponse.getName() !== "Form Responses 1"){
    var UI = SpreadsheetApp.getUi();
    UI.alert("ERROR:\nPlease make sure \"Form Responses 1\" is the active sheet.");
    return;
  }
  
  // Loop through the form respones and archive every row that is marked as done
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
  
  // Make sure that all the checkboxes correctly copy over to the archive
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