function putRegistrations() {
  //Get the spreadsheets and data from the spreadsheets
  var sa = SpreadsheetApp.getActiveSpreadsheet();
  var slateSS = sa.getSheetByName("Form Data");
  var putSS = sa.getSheetByName("Registration List");
  var slateDT = getDataValues(slateSS.getName());
  var putDT = getDataValues(putSS.getName());
  var additions = [];
  
  //Loop through all of the data coming from Slate
  for(var i = 1; i < slateDT.length; i ++) {
    var bannerid = slateDT[i][0];
    var match = "No Match";
    var j = 1;
    
    //Loop through all of the grading data unless there is a match
    while(j < putDT.length && match == "No Match") {
      var matchID = putDT[j][0].trim();
      
      var helping = 1;

      if(bannerid.trim() != matchID){ //If the BannerIDs don't match, move on
        j++;
      } else {
        match = "Match"; //If the BannerIDs do match, stop the loop
      }
      var helper = 1;
    }
    
    var help = 1;
    
    if(match == "No Match"){ //If there was no match, add it to the additions array without the last three columns
      additions.push(moveColumns(slateDT[i]));
    }
  }
  
  help = 2;
  
  if(additions.length > 0) {
    //Paste these into the bottom of the first page starting in column 2
    putSS.getRange(putDT.length + 1, 1, additions.length, additions[0].length).setValues(additions);
  
    //Sort the entire sheet by name including the first row.
    putDT = getDataValues(putSS.getName());
    var sortRng = putSS.getRange(2, 1, putDT.length, putDT[0].length);
    sortRng.sort(2);
  }
}

function moveColumns(studentRow) {
  var helper = 1;
  
  //Determine payment status and put it in the correct field
  if(studentRow[2] != ""){
    studentRow[2] = "FA Request";
  } else {
    studentRow[2] = studentRow[3];
  }
  studentRow.splice(3, 1);
  helper = 1;
  
  //Create the hometown column and put it in the correct field
  studentRow[11] = studentRow[11] + ", " + studentRow[12];
  studentRow.splice(12, 1);
  helper = 1;
  
  //Move the unable activities in front of the able activities
  var v = checkDefined(studentRow.splice(35, 1));
  studentRow.splice(13, 0, v[0]);
  
  //Remove the CEEB Code
  studentRow.splice(12, 1);
  studentRow.splice(14, 17);
  studentRow.splice(15, 10);
  helper = 1;
  
  return studentRow;
}