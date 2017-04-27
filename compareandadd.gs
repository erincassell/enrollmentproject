function compareandadd() {
  //Get the spreadsheets and data from the spreadsheets
  var sa = SpreadsheetApp.getActiveSpreadsheet();
  var slateSS = sa.getSheetByName("From Slate");
  var gradingSS = sa.getSheets()[0];
  var slateDT = getDataValues(slateSS.getName());
  var gradingDT = getDataValues(gradingSS.getName());
  var additions = [];
  
  //Loop through all of the data coming from Slate
  for(var i = 20; i < slateDT.length; i ++) {
    var bannerid = slateDT[i][0];
    var match = "No Match";
    var j = 1;
    
    //Loop through all of the grading data unless there is a match
    while(j < gradingDT.length && match == "No Match") {
      var matchID = gradingDT[j][1].trim();

      if(bannerid.trim() != gradingDT[j][1].trim()){ //If the BannerIDs don't match, move on
        j++;
      } else {
        match = "Match"; //If the BannerIDs do match, stop the loop
      }
    }
    
    if(match == "No Match"){ //If there was no match, add it to the additions array without the last three columns
      additions.push(slateDT[i].splice(0, 7));
    }
  }
  
  if(additions.length > 0) {
    //Paste these into the bottom of the first page starting in column 2
    gradingSS.getRange(gradingDT.length + 1, 2, additions.length, additions[0].length).setValues(additions);
  
    //Sort the entire sheet by name including the first row.
    gradingDT = getDataValues(gradingSS.getName());
    var sortRng = gradingSS.getRange(2, 1, gradingDT.length, gradingDT[0].length);
    sortRng.sort(3);
  }
}
