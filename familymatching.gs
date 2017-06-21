function completeMatches()
{
  var activityCol; //location of the activity column
  activityCol = prepSheet(); //preps the spreadsheet for processing and returns the activity column
  mapStudents(activityCol); //passes in the activity column to complete the student mapping
}

function prepSheet() {
  
  //Set the variables
  var sa, shtFirst, shtSecond, shtThird, shtFour;
  var rngData, famData, headerData, studData;
  var lastCol, lastRow, activities, cntStud, gender, unable, ceeb;
  var afterActivities, combined;
  var arrayHeader, arrayColumn;
  
  //Get spreadsheet information
  sa = SpreadsheetApp.getActiveSpreadsheet();
  lastCol = SpreadsheetApp.setActiveSheet(sa.getSheets()[0]).getLastColumn();
  lastRow = SpreadsheetApp.setActiveSheet(sa.getSheets()[0]).getLastRow();
  
  //Set sheet information for the shets will be using
  shtFirst = sa.setActiveSheet(sa.getSheetByName("PRE Master List"));
  shtSecond = sa.setActiveSheet(sa.getSheetByName("Family Groups"));
  
  //Delete the student processing and family processing sheets if they exsist
  if(sa.getSheetByName("Student Processing"))
  {
    sa.deleteSheet(sa.getSheetByName("Student Processing"));
  }
  
  if(sa.getSheetByName("Family Processing"))
  {
    sa.deleteSheet(sa.getSheetByName("Family Processing"));
  }
  
  //Insert the new sheets
  sa.insertSheet(2).setName("Student Processing"); //Puts the Student Processing in the third slot
  sa.insertSheet(3).setName("Family Processing"); //Puts the Family Processing in the fourth slot

  //Clear the sheets in case this is a re-run
  shtThird = sa.setActiveSheet(sa.getSheetByName("Student Processing"));
  shtThird.clear();
  shtFour = sa.setActiveSheet(sa.getSheetByName("Family Processing"));
  shtFour.clear();

  //Copy the student data from the first sheet to the third sheet
  rngData = shtFirst.getRange(1, 1, shtFirst.getLastRow(), shtFirst.getLastColumn());
  rngData.copyTo(shtThird.getRange(1,1), {contentsOnly:true});

  //Remove the students that are withdrawn by looking at the last column
  objLastCol = shtThird.getRange(2, shtThird.getLastColumn(), shtThird.getLastRow()-1, 1).getValues();
  for(var j = 0; j < objLastCol.length; j++)
  {
    if(objLastCol[j][0] != "") //looks at last column
    {
      shtThird.deleteRow(j+2);
    }
  }
  
  //Insert two columns to line up correctly
  shtThird.insertColumnAfter(6); //Puts a blank column at G (eventually becomes H)
  shtThird.insertColumnAfter(6); //Puts a blank column at H (eventually becomes I)
  shtThird.insertColumnAfter(14); //Puts a blank column at O (eventually becomes P)

  //Delete the first column
  shtThird.deleteColumn(1);
  
  //Find the column location for different headers
  arrayHeader = shtThird.getRange(1, 1, 1, shtThird.getLastColumn()).getValues();
  var i = 0;
  while(i <= shtThird.getLastColumn()) {
    switch(arrayHeader[0][i])
    {
      case "Able_Activities":
        activities = i + 1;
        break;
      case "Gender":
        gender = i + 1;
        break;
      case "Unable_Activities":
        unable = i + 1;
        break;
      case "CEEB Code":
        ceeb = i + 1;
        break;
    }
    i = i + 1;
  }
  
  //Sort the third sheet before getting rid of columns
  var rngData = shtThird.getRange(2, 1, shtThird.getLastRow()-1, shtThird.getLastColumn()); //everything starting at the second row.
  //rngData.sort(gender); //Sort the range by gender
  
  //Remove everything after the activities
  rngData = shtThird.getRange(1, activities+8, shtThird.getLastRow(), shtThird.getLastColumn()-activities); //add 8 because we want the seven columns to the right of the activities
  rngData.clearContent();
  
  i = activities+8;
  while(i <= shtThird.getLastColumn())
  {
    shtThird.deleteColumn(i);
    i = i+1;
  }
  
  //Add a new column header at the end for the Family if it does not exist
  shtThird.getRange(1, activities+8).setValue("Family");
  
  //Copy the family data over to sheet #4
  shtThird.getRange(1, activities+1, 1, 6).copyTo(shtFour.getRange("B1"), {contentsOnly:true}); //Get the six families sheet three and copy
  shtSecond.getRange(2, 1, shtSecond.getLastRow(), 1).copyTo(shtFour.getRange("A2"), {contentsOnly:true}); //Get the families and copy
  famData = shtSecond.getRange(2, 2, shtSecond.getLastRow()-1, shtSecond.getLastColumn()-2).getValues(); //Load the family values into the famData array
  headerData = shtFour.getRange(1, 2, 1, shtFour.getLastColumn()-1).getValues(); //Load the header information into the array
  var values = [["Male", "Female", "Total", "Combined"]]; //New values that need to be put on the sheet
  
  //Add in the Male/Female and Total Column
  shtFour.getRange(1, shtFour.getLastColumn()+1, 1, 4).setValues(values); //Put the column headers in
  shtFour.getRange(2, shtFour.getLastColumn()-1).setFormulaR1C1("=SUM(R[0]C[-2]:R[0]C[-1])"); //Put in the function in the second row, last column
  shtFour.getRange(2, shtFour.getLastColumn()-1).copyTo(shtFour.getRange(3, shtFour.getLastColumn()-1, shtFour.getLastRow()-2, 1)); //Fill the function down
  
  //Get the combined number
  combined = shtFour.getLastColumn();
  
  //Copy the activities from sht2 to the combined column of sht4
  shtFour.getRange(2, combined, shtFour.getLastRow()-1, 1).setValues(shtSecond.getRange(2, shtSecond.getLastColumn(), shtSecond.getLastRow()-1, 1).getValues());
  
    //Loop through all of the family data array
  for(var k=0; k < famData.length; k++) //for the height of the array
  {
    //Loop through each activity in the family
    for(var m = 1; m < 4; m++)
    {
      //Loop through each header item to mark those.
      for(var n = 0; n < 6; n++)
      {
        //Compare with each header item and put the name if it matches
        if(famData[k][m].trim() == headerData[0][n].trim())
        {
          shtFour.getRange(k+2, n+2).setValue(headerData[0][n]);
        }
       }
    }
  }
  
  return activities;
}

function mapStudents(activityCol) {
  
  //Define and set variables
  var sa, shtThree, shtFour;
  var activity, activityCol;
  var i, f, offset;
  var arrayActivityTot;
  
  //Get spreadsheet information
  sa = SpreadsheetApp.getActiveSpreadsheet();
  shtThree = sa.getSheetByName("Student Processing");
  shtFour = sa.getSheetByName("Family Processing");
  
  //activityCol = 22; //only for debugging purposes

  f = 0;
  //Loop through for the number of activities that need to be processed
  arrayActivityTot = shtThree.getRange(shtThree.getLastRow(), activityCol+1, 1, 6).getValues();
  while(f <= 6) 
  {
    //Set the default values
    activity = arrayActivityTot[0][0]; //first activity total from the array
    offset = 1;

    //Go through the activities and process through them finding the next smallest to process    
    if(f == 1) //Inserting a higher count item in the second position
    {
      offset = 4;
    }
    else
    {
      for(var i = 1; i < 6; i++)
      {
        if(arrayActivityTot[0][i] < activity) //Compares the next activity with the current activity number
        //if(shtThree.getRange(shtThree.getLastRow(), activityCol+i).getValue() < activity)
        {
          activity = arrayActivityTot[0][i];
          offset = i+1;
        }
      }
    }

    familyMatch(offset, activityCol);
    arrayActivityTot[0][offset-1] = 1000; //Set the array value to 1000, so it is not processed again.
    f++;
  }
}

function familyMatch(offset, activityCol)
{
  //Define and set variables
  var sa, shtThree, shtFour, shtTwo;
  var lastCol, lastRow;
  var activityCol, strUnable, unable;
  var rngStudents, rngCnts, srtCol, rngFamilies;
  var stuData, famData, thresholdData;
  var i, j, k, offset, inc, chkUnable;
  var male, female, activities;
  
  //Get spreadsheet information
  sa = SpreadsheetApp.getActiveSpreadsheet();
  shtThree = sa.getSheetByName("Student Processing");
  shtFour = sa.getSheetByName("Family Processing");
  shtTwo = sa.getSheetByName("Thresholds");
  
  //******DEBUGGING ONLY
  //activityCol = 22;
  //offset = 3
  
  //Get the data for the thresholds
  thresholdData = shtTwo.getRange(2, 1, 1, 3).getValues();
  
  //Set the sort column based on the activity column and the offset passed-in
  srtCol = activityCol+offset; 
  
  //Load the sheet (no first row and no last row) into a range and sort
  rngStudents = shtThree.getRange(2, 1, shtThree.getLastRow()-2, shtThree.getLastColumn());
  rngStudents.sort([{column: srtCol, ascending: false}, {column: 16, ascending: true}, {column: shtThree.getLastColumn()-2, ascending: true}]); //Sort based on the activity and the student's total activity count
  
  //Sort the families
  rngFamilies = shtFour.getRange(2, 1, shtFour.getLastRow()-1, shtFour.getLastColumn());
  rngFamilies.sort([{column: offset+1, ascending: false}, {column: shtFour.getLastColumn()-1, ascending: true}]); //by activity and then total students in the family
  
  //Load the data arrays
  stuData = rngStudents.getValues();
  famData = rngFamilies.getValues();
  
  j = 0;
  lastCol = shtThree.getLastColumn(); //25
  unable = activityCol-2;

  //While we are still looking at the the students with that activity
  //Working with an array, so have to remember to work beginning with 0 rather than 1
  while(stuData[j][activityCol+offset-1] != "")
  {
    //Check to see if the student is in a family. 
    if(stuData[j][lastCol-1] == "") //If the family field is blank, then continue
    {
      k = 0;
      //Loop through the family data while there are still families for the item or until the student has been put in a family or until the family is full
      while(famData[k][offset] != "" && stuData[j][lastCol-1] == "" && famData[k][9] < thresholdData[0][2])
      {
        //Check the male/female
        male = famData[k][7];
        female = famData[k][8];

        switch(stuData[j][9]) //student sex column
        {
          case "M":
              male++;
              inc = "M";
              break;
          case "F":
              female++;
              inc = "F";
              break;
        }
        if(male > thresholdData[0][0] || female > thresholdData[0][1])
        {
        }
        else
        {
          //Proceed with further checks
          activities = famData[k][10]; //Gets the family activities
          
          //Check to see if the students unable activity is in the family list
          if(stuData[j][unable] != "")
          {
            chkUnable = checkUnable(stuData[j][unable], activities);
          }
          else
          {
            chkUnable = "OK";
          }
            
          if(chkUnable == "OK")
          {
            //Put the family in the student's family column
            stuData[j][lastCol-1] = famData[k][0];
            rngStudents.getCell(j+1, lastCol).setValue(stuData[j][lastCol-1]);
            
            //Increment the male or female
            if(inc == "M")
            {
              famData[k][7]++;
              rngFamilies.getCell(k+1, 8).setValue(famData[k][7]);
            }
            else
            {
              famData[k][8]++;
              rngFamilies.getCell(k+1, 9).setValue(famData[k][8]);
            }
          }
        }
        k++;
        
        //Sort the families to bring the next available to the top
        rngFamilies.sort([{column: offset+1, ascending: false}, {column: shtFour.getLastColumn()-1, ascending: true}]);
        
        //reload the family data
        famData = rngFamilies.getValues();
      }
    }
    j++;
  }
}

function checkUnable(unableActivities, familyActivities)
{
  var unableActivities, familyActivities;
  var comma, passes, i, j, k, OK;
  var unableOne, unableTwo;
  
  //See if there is more than one activity
  comma = unableActivities.indexOf(",");
  if(comma > 0) //If there is more than one unable activity
  {
    unableOne = unableActivities.substr(0, comma).trim(); //Split the activities
    unableTwo = unableActivities.substr(comma+1, unableActivities.length).trim();
    if(familyActivities.indexOf(unableOne) < 0 && familyActivities.indexOf(unableTwo) < 0)
    {
      OK = "OK";
    }
    else
    {
      OK = "No";
    }
  }
  else
  {
    unableOne = unableActivities.trim(); //Set the first to unableActivities
    if(familyActivities.indexOf(unableOne) < 0)
    {
      OK = "OK";
    }
    else
    {
      OK = "No";
    }
  }
  return OK;
}