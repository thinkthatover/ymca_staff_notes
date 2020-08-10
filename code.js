
function onOpen(e){
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var ssArray = spreadsheet.getSheets();
    spreadsheet.setActiveSheet(ssArray[0]);
  
  //  var sheet = spreadsheet.getActiveSheet()  
  //  var ui = SpreadsheetApp.getUi();
  //  ui.createMenu('Staff Notes Management')
  //      .addItem('Clean Spreadsheets', 'cleanPrompt')
  //      .addSeparator()
  //      .addItem('Update Front Page Notes', 'menuItem2')
  //      .addToUi();
  }
  
  function menuItem2(){updateRecents()}
  
  
  function cleanPrompt () {
    //menu item controlling the cleanup procedure
    var ui = SpreadsheetApp.getUi();
    
    var result = ui.prompt(
      'Enter the number of days back you\'d like to start deleting entries from:',
      ui.ButtonSet.OK_CANCEL)
    
    var button = result.getSelectedButton();
    var text = result.getResponseText();
    if (button == ui.Button.OK){
      if (!(+text >= 0) || (+text < 7)){
        ui.alert('Please enter a number greater than 6')
      }
      else{
        cleanSheets(+text)
        ui.alert('Clean as a whistle!')
        
      }
      
    }
  }
  
    
  function updateRecents(){
  //populates instruction page w/ most recent messages from management
    
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getActiveSheet();
    var ssArray = spreadsheet.getSheets();
    var sourceSheet = "Facility/General"
    var targetSheet = "Instructions"
    var managerList = ["Manager1",
                       "Manager2"]
    var recentsList = [] //[message, name, date]
    var currentList = []
    var numberOfPosts = 6
    
    //populate recentsList
    spreadsheet.setActiveSheet(ssArray[findSheet(targetSheet)]);
    sheet = spreadsheet.getActiveSheet();
    var noteTable= sheet.getRange(9,1,numberOfPosts,3).getValues()
    for (row = 0; row < numberOfPosts; row++){
      currentList.push([noteTable[row][0], noteTable[row][1], noteTable[row][2]])
    }
    
    
    //get column info
    spreadsheet.setActiveSheet(ssArray[findSheet(sourceSheet)]);
    sheet = spreadsheet.getActiveSheet();
    var headers = sheet.getRange(1,1,1, sheet.getLastColumn()).getValues();
    var nameCol = headers[0].indexOf("From") + 1;
    var messageCol = headers[0].indexOf("Description") + 1;
    var dateCol = headers[0].indexOf("Date") + 1;
    
    
    //iterate through messages,capture important ones
    for (j = sheet.getLastRow(); j > 1; j--){
      var name = sheet.getRange(j, nameCol,1,1).getValues()[0][0]
      var date = sheet.getRange(j, dateCol, 1,1).getValues()[0][0]
      
      if (include(managerList, name) && recentsList.length < numberOfPosts){
        var aList = [sheet.getRange(j, messageCol).getValues()[0][0], name, date]
        recentsList.push(aList)
      }
    }
    
    //compare entries between lists
    var newList = currentList.slice(0);
    recentsList.forEach(function (item, index){
      cbool = true;
      currentList.forEach(function (item2, index2){
        if (item[0] === item2[0]){
          cbool = false;
        }
      })
      if (cbool) {
        newList.pop()
        newList.unshift(item)
      }
    })
    
    //post messages
    spreadsheet.setActiveSheet(ssArray[findSheet(targetSheet)])
    sheet = spreadsheet.getActiveSheet();
    
    for (m = 0; m < newList.length; m++){
      sheet.getRange((2 + m),1).setValue(newList[m][0])
      sheet.getRange((2 + m),2).setValue(newList[m][1])
      sheet.getRange((2 + m),3).setValue(newList[m][2])
    }
    
  }

  
  function cleanSheets(days){
  //moves sheets older than *days* to completed entries sheet 
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var ssArray = spreadsheet.getSheets();
    var ssToClean = ["Facility/General",
                     "Daxko/Billing",
                     "Missed Hours",
                    "Trainer Notes"]
    if (days === undefined){
      days = 14
    }
      
    //get array w/ sheets we want
    var cleanArray = ssArray.filter(function(sheet){
      if (include(ssToClean, sheet.getName())) {
        return true
      } else{
        return false
      }})
    
    
    //loop over sheet array and clean each one
    for (i = 0; i < cleanArray.length; i++){
      spreadsheet.setActiveSheet(cleanArray[i]);
      var sheet = spreadsheet.getActiveSheet();
      var headers = sheet.getRange(1,1,1, sheet.getLastColumn()).getValues()
      var compCol = headers[0].indexOf("Completed") + 1;
      var dateCol = headers[0].indexOf("Date") + 1;
      var today = new Date()
      var compWatchArr = ["No","Urgent"]
      var checkDate = today.subtractDays(days)  //change # of days back to delete here
      
      //check each column for completed entries greater than two weeks old
      for (j = sheet.getLastRow(); j > 1; j--) {
        var compVal = sheet.getRange(j, compCol).getValue()
        var dateVal = sheet.getRange(j, dateCol).getValue()
        var stampDate = new Date(dateVal)
        
        if (!(include(compWatchArr, compVal)) && (stampDate < checkDate)){
          Logger.log(sheet.getName() + " " + j + ", " + stampDate)
          
          var sheetNameToMoveTheRowTo = "Completed Entries";
          var targetSheet = spreadsheet.getSheetByName(sheetNameToMoveTheRowTo);
          var targetRange = targetSheet.getRange(targetSheet.getLastRow() + 1, 3);
          var dateVal = sheet.getRange(j, (headers[0].indexOf("Date") + 1)).getValue()
          var targetRow = targetSheet.getLastRow() + 1;
          var range = sheet.getRange(j,1);
          
          targetSheet.getRange(targetRow, 1).setValue(sheet.getName())
          targetSheet.getRange(targetRow, 2).setValue(dateVal)
          
          sheet.getRange(range.getRow(), 1, 1, sheet.getLastColumn()).moveTo(targetRange);
          sheet.deleteRow(j); //delete empty row
        }//if
      }//rowloop
  //    Utilities.sleep(500)     
  //    sheet.setFrozenRows(1);
    }//sheet loop
    spreadsheet.setActiveSheet(ssArray[0])
  }
  
  function formsubmit(){
  //  activated when personal training form is submitted, adds entry to trainer notes, emails trainers, etc.
    
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    formSheet = spreadsheet.getSheetByName('PT Responses')
    trainerSheet = spreadsheet.getSheetByName('Trainer Notes')
    spreadsheet.setActiveSheet(formSheet)
    var sheet = spreadsheet.getActiveSheet();
    var headers = sheet.getRange(1,1,1, sheet.getLastColumn()).getValues()
    var noteArray = ['PT Form']
    
    var row = sheet.getLastRow() 
    var Name = sheet.getRange(row, headers[0].indexOf("Name") + 1).getValue()
    var trainer = sheet.getRange(row, headers[0].indexOf("Desired Trainer") + 1).getValue()
    if (trainer !== ""){
      var Tag = trainer
    }
    else{
      var Tag = "New Client"
    }
    
    var injury = sheet.getRange(row, headers[0].indexOf("Injury") + 1).getValue()
    if (injury !== ""){
      var injury = ", Injury: " + injury  
    }
    
    var Contact = sheet.getRange(row, headers[0].indexOf("Contact Info") + 1).getValue()
    var goal = sheet.getRange(row, headers[0].indexOf("Training Goals") + 1).getValue()
    var weekday = sheet.getRange(row, headers[0].indexOf("Availability [Weekday]") + 1).getValue()
    var weekend = sheet.getRange(row, headers[0].indexOf("Availability [Weekend]") + 1).getValue()
    
    var available = ", Available: "
    
    if (weekday !== ""){
      available = available + "weekday " + weekday.toLowerCase()
    }
    else if (weekend !== ""){
      available = available + "weekend " + weekend.toLowerCase()
    }
    else{
      available = ", follow up for availability"
    }
    
    var startDate = sheet.getRange(row, headers[0].indexOf("Ideal Start Date") + 1).getValue()
    if (startDate !== ""){
      startDate = new Date(startDate)
      startDate = ",  Start date: " + startDate.getMonth() + "/" + startDate.getDate()
    }
    var otherNotes = sheet.getRange(row, headers[0].indexOf("Other Details (specific day/availability, etc.)") + 1).getValue()
    
    if (otherNotes!== ""){
      otherNotes = ", Other Notes: " + otherNotes
    }
    var Description = "Goal: " + goal + injury + available + startDate + otherNotes 
    
    noteArray.push(Tag, Name, Contact, Description)
    spreadsheet.setActiveSheet(trainerSheet)
    var sheet = spreadsheet.getActiveSheet();
    
    var tarRow = sheet.getLastRow() + 1
    
    for (i=noteArray.length;i>0;i--){
      sheet.getRange(tarRow, i).setValue(noteArray[i-1])
    }
    
    var activestr = "A" + tarRow.toString()
    sheet.setActiveSelection(activestr)
    checkEdit()
    
  }
  
  
  
  function checkEdit() {
  // Controls features that occur when cells are edited, 
  
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getActiveSheet();
    var shName = sheet.getName()
    
    //get edit location
    var actRng = sheet.getActiveRange();
    var editCol = actRng.getColumn();
    var editRow = actRng.getRowIndex(); 
    
   //Get Column indexes
    var headers = sheet.getRange(1,1,1, sheet.getLastColumn()).getValues(); //getRange(row, column, numRows, numColumns)
    var respCol = headers[0].indexOf("Response") + 1;
    var compCol = headers[0].indexOf("Completed") + 1;
    var dateCol = headers[0].indexOf("Date") + 1;
    var tagCol = headers[0].indexOf("Tag") + 1;
    var nameCol = headers[0].indexOf("From") + 1;
    var messageCol = headers[0].indexOf("Description") + 1;
      //check sheet has a completed section
    Logger.log(compCol)
    Logger.log(nameCol)
  
    
    if (compCol != 0 || tagCol != 0){                  //Only edit sheets w/ Completed/tag columns
      var compString = sheet.getRange(editRow, compCol).getValues()[0][0];
    
    
    //mark rows as addressed if response added
    if (editCol == respCol && compString !== "Yes"){
        sheet.getRange(editRow, compCol).setValue("Addressed");
    }
    
    Logger.log('datecol isnt 0 = ' + dateCol)
    Logger.log('editrow isnt 1 = ' + editRow)
    Logger.log("Edit Col is tagcol or namecol = " + editCol)
    Logger.log('editCol = ' + editCol)
    Logger.log('nameCol = ' + nameCol)
    
    //timestamps and marks each new row entry when a tag is added
    if (dateCol > 0 && (editCol == tagCol || editCol == nameCol) && editRow !== 1){   //if date header exists, edited row not in header) 
      sheet.getRange(editRow, dateCol).setValue(Utilities.formatDate(new Date(), "GMT-5", "MM/dd/yyyy"));
      sheet.getRange(editRow, compCol).setValue("No");
    }
    
    urgentNote()
    }
  }
  
  
  function urgentNote() {
    //checks if a note is marked "urgent" and sends contetns to managers or trainers
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getActiveSheet()
    var shName = sheet.getName()
    
  
    //place of edit info
    var actRng = sheet.getActiveRange();
    var editCol = actRng.getColumn();
    var editRow = actRng.getRowIndex();
    
    //Get Column indexes for editRow
    var headers = sheet.getRange(1,1,1, sheet.getLastColumn()).getValues(); //getRange(row, column, numRows, numColumns)
    var respCol = headers[0].indexOf("Response") + 1;
    var compCol = headers[0].indexOf("Completed") + 1;
    var dateCol = headers[0].indexOf("Date") + 1;
    var tagCol = headers[0].indexOf("Tag") + 1;
    var nameCol = headers[0].indexOf("From") + 1;
    var messageCol = headers[0].indexOf("Description") + 1;
    
    //Get Column Values
    var compString = sheet.getRange(editRow, compCol).getValues()[0][0];
    var tagString = sheet.getRange(editRow, tagCol).getValues()[0][0]
    var nameString = sheet.getRange(editRow, nameCol).getValues()[0][0]
    var messageString = sheet.getRange(editRow, messageCol).getValues()[0][0]
  
    
  //logic for sending a message to trainers
    if (shName == 'Facility/General' || shName == 'Trainer Notes'){
    var checkcol = sheet.getLastColumn() + 1
    var checkval = sheet.getRange(editRow,checkcol).getValues()[0][0]
    Logger.log(checkval)
    if ((tagString == "Trainer Name") && (messageString !== "") && (editCol == messageCol || editCol == tagCol || editCol == nameCol)) {
      if (checkval != 1){
        var message = messageString + ' - ' + nameString
        MailApp.sendEmail('traineremail@gmail.com', 'Update from Leominster YMCA Staff Notes', messageString)
        sheet.getRange(editRow,checkcol).setValue('1')
      }
    }
    }
    
    //logic for sending a message to managers
    if ((editCol ==  compCol || editCol == tagCol) && compString === "Urgent"){
      var numberOfManagers = 2  //change this w/ number of managers
      var managerList = spreadsheet.getSheetByName('Data Sheet').getRange(5, 7, numberOfManagers, 1).getValues()
      var managerEmailList = spreadsheet.getSheetByName('Data Sheet').getRange(5, 8, numberOfManagers, 1).getValues()
      
      managerList = cleanList(managerList)
      managerEmailList = cleanList(managerEmailList)
      
      var managerIndex = managerList.indexOf(tagString)
      
      //checks if manager
      if (managerIndex > -1){
        var manager = managerList[managerIndex]
        var email = managerEmailList[managerIndex]
        var message = messageString + ' - ' + nameString
        sendTheEmail(email, message)
      } 
    }
  }
  
  
  
  function sendTheEmail(email, message){
    //sends email w/ message to recipeint
    var ui = SpreadsheetApp.getUi();
    var result = ui.alert(
      'You have created an Urgent message for a Manager.',
      'Would you like to send an email to notify them?',
      ui.ButtonSet.YES_NO
    )
    if (result == ui.Button.YES) {
      
      if (email === 'testemail@gmail.com'){
        ui.alert(
          'The manager you have selected is not currently subscribed to email messages'
          )
      } else{
      var subjectline = 'Update from Staff Notes'
      Logger.log(message)
      MailApp.sendEmail(email, subjectline, message)
      }
    }
  }
  
  
  
  //***********************
  //HELPER FUNCTIONS
  //***********************
    
  //range.getValues() makes every item a list in the returned list [[Person1],[Person2]]
  //this returns the list w/o nested list items so array.indexOf('keyword') will work
  
  
  function cleanList(dList){
    //multiple cells return nested arrays, this function returns an array for 1D cell selections (ie: 3x1),  

    var cList = []
    for (i = 0; i < dList.length; i++){
      cList.push(dList[i][0])
    }
    return cList
  }
  
  
  //replaces Array.includes(obj) cuz Google scripts runs on JS 1.6
  function include(arr,obj) {
    return (arr.indexOf(obj) != -1);
  }
  
  //returns index of sheet in given spreadsheet (heads up there's already a google sheets method for this)
  function findSheet(sheetname){
    var ssArray = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    for (i = 0; i < ssArray.length; i++){
      if (ssArray[i].getSheetName() === sheetname){
        return i
      }
    }
    Logger.log("No sheet " + sheetname + "found")
  }
  
  //checks if str2Ch begins with chStr
  function includesString(chStr, str2Ch){
    return (chStr == str2Ch.slice(0,chStr.length))
   }
  

  function testFuncs(){
    Logger.log(findSheet("Instructions"))
  }

//test function
function dumyfun(){
    cleanSheets(9)
}
  
  Date.prototype.subtractDays = function(d) {  
                  this.setTime(this.getTime() - (d*24*60*60*1000));  
                  return this;  
              } 