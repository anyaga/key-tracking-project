/*
Goals:
-Alumni/Peoples who have reached expiration date
-Renew key?
-Edit currentFormToClass to deal with wrong dates
-Edit currentFormToClass to deal with no advisor
-1st Year PhD  advisor is Heather
*/

class keyRecord {
  constructor(first,last,andrewID,advisor,dept,keys,room,givenDate,expDate)  {
    this.firstName = first;
    this.lastName = last;
    this.andrewID = andrewID;
    this.advisor = advisor;
    this.dept = dept;
    this.keys = keys;
    this.room = room;
    this.givenDate = givenDate;
    this.expDate = expDate;
  }
  //Basic constructor functions
  getFirstName(){
    return this.firstName
  }
  getLastName(){
    return this.lastName
  }
  getAndrewID(){
    if(this.andrewID == ""){
      return this.firstName +"_"+ this.lastName+ "_no_andrew_id"
    }
    return this.andrewID
  }
  getAdvisor(){
    return this.advisor
  }
  getDepartment(){
    return this.dept
  }
  getKeys(){
    return this.keys
  }
  getRoom(){
    return this.room
  }
  getGivenDate(){
    return this.givenDate
  }
  getExpirationDate(){
    return this.expDate
  }
  //--------------------------------------------------------
  getName() {
    return this.firstName + this.lastName
  }
  addKey(key,room) {
    this.keys.push(key)
    this.room.push(room)
  }
}

//Parsing the old code sheets
function parseOldKeySheet(){
  var allEntries = new Map();//keys = andrewID, entries = keyRecord
  const oldSS = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1Ox2aaN6ZUc3Hzm6mwCiQtm4A1K5Ey-vXYGCHOvOG-cQ/edit')
  sheets = oldSS.getSheets()

  //All the data
  for(var i = 0; i < sheets.length; i++){
    //Figure out the end range <--or how to avoid going over empty setion
    var sheet = sheets[i];
    var firstNameArr = sheet.getRange('D2:D113').getValues(); ////////Standardize!!!!
    var lastNameArr  = sheet.getRange('C2:C113').getValues();
    var andrewIDArr  = sheet.getRange('F2:F113').getValues();
    var advisorArr   = sheet.getRange('G2:G113').getValues();
    var deptArr      = sheet.getRange('I2:I113').getValues();
    var keysArr      = sheet.getRange('J2:J113').getValues();
    var roomArr      = sheet.getRange('K2:K113').getValues();

    for(var i = 0; i < 110; i++){ //Replace with length of the rows.
      var firstName = firstNameArr[i][0];
      var lastName = lastNameArr[i][0];
      var andrewID = andrewIDArr[i][0];
      var advisor = advisorArr[i][0];
      var dept = deptArr[i][0];
      var keys = keysArr[i][0];
      var room = roomArr[i][0];

      //No AndrewIDs
      if(andrewID == ""){
        var bEntry = new keyRecord(firstName,lastName,andrewID,
                      [advisor], dept,[keys],[room]);
        allEntries.set(firstName.concat(lastName,"-no-andrewID"),bEntry);
      }
      //No Entries or New AndrewID
      else if((allEntries.size== 0) || (!allEntries.has(andrewID))) {
        var newEntry = new keyRecord(firstName,lastName,andrewID,
                      [advisor], dept,[keys],[room],null,null); //Given and exp date not given
        allEntries.set(andrewID,newEntry);
      } 
      //Adding a key to existing record
      else {//<---Overwriting old values
        var entry = allEntries.get(andrewID); 
        allEntries.delete(andrewID); 
        entry.addKey(keys,room,advisor);
        allEntries.set(andrewID,entry); 
      }
    }
  }
  //Eventually, will add the new form once it starts to populate
  return allEntries;
}

function validKey(key) {
  //Some error with Key formating
  if(!key.includes("4501-")){
    //Key doesnt have dash, need to correct
    if(key.includes("4501")){
      end_i = key.length - 1;
      i = end_i;
      key_copy = key;
      base = "";
      add = "";
      while((i != -1) || (key_copy.slice(0,i) != "4501")){//(key[i] != "-")){
        base = key[i].concat(base);
        i--;
     }
    } 
    else return "invalid key";
  } return key;
}

function validRoomNum(num){
  const floorOpt = ["C","B","A","1","2","3","4","a","b","c"]; //find better way to deal with no Cap
  floor = num[0];
  digits = num.slice(1,num.length);
  validFloor = false;
  for(opt in floorOpt){if(opt == floor) validFloor = true;}
  //Is the first digit the floor number
  if(!validFloor) return false;
  //valid lenght of digits (3 digits to be room num in Doherty)
  else if(digits.length != 3) return false;
  //Are the digits(remaining room num) a valid number
  else if(parseInt(digits) == NaN) return false;
  else return true;
}

//make sure the second half is actually a number!!!!
function validRoom(room){
  roomNum = 0
  if(room.includes("DH ")){
    return  validRoomNum(room.slice(3,room.length)) ? room : "invalid room";
  } 
  else if(room.includes("Doherty")){
    roomNum = room.slice(7,room.length);
    while (roomNum[0] == " "){
      roomNum = roomNum.slice(1,roomNum.length);
    }
    return validRoomNum(roomNum) ? "DH ".concat(roomNum): "invalid room";
  } 
  else if(room.includes("DH")) {
    roomNum = room.slice(2,room.length);
    return validRoomNum(roomNum)? "DH ".concat(roomNum): "invalid room";
  } 
  else return "invalid room"
}

//Valid date!!!!!
function validDate(){
  return
}
/*
Translate form response to spreadsheet format (with keyReponse class)
*/
function currentFormToClass() { 
  var allEntries = parseOldKeySheet();
  //Array of form responses
  var firstName,lastName, advisor,andrewID, key,room,givenDate,expDate,ques,answ,dept;
  const form1 = FormApp.openByUrl(
    'https://docs.google.com/forms/d/1fPmkuLoWQXsgwz1ruQw3rkGO93eN1PrUEUINBaV4MBc/edit');
  var allResp = form1.getResponses();

  //Individual responses
  for(const resp of allResp) {
    //All the questions and response stores in an item
    for(item of resp.getItemResponses()){
      ques = item.getItem().getTitle();
      answ = item.getResponse();
      if(ques == "First Name:"){
        firstName = answ;
      } else if(ques == "Last Name:") {
        lastName = answ;
      } else if(ques == "Advisor:") {
        advisor = answ;
      } else if(ques == "andrewID:") {
        andrewID = answ;
      } else if(ques == "Key Number:") {
        key = validKey(answ);
      } else if(ques == "Room (Include Building and Room Number) Ex: DH 3213A") {
        room = validRoom(answ);
      } else if(ques == "What date were you given the key/key access?") {
        givenDate = answ;
      } else if(ques == "What date will you lose acess? (Typically expected graduation date)") {
        expDate = answ;
      } else if(ques == "Are you a part of the Chemical Engineering Department?") {
        if(answ == "Yes"){
          dept = "Chemical Engineering";
        } else if("No"){
          dept = "Other Department";
        } else{
          dept = answ;
        }
      }
    }
    if(!allEntries.has(andrewID)){
      var newEntry = new keyRecord(firstName,lastName,andrewID,[advisor],
                                                    dept,[key],[room],givenDate,expDate);
      allEntries.set(andrewID,newEntry);
    } else {
        var Entry = allEntries.get(andrewID);
        allEntries.delete(andrewID);
        Entry.addKey(key,room,advisor);
        allEntries.set(andrewID,Entry);
    }
  }
  return allEntries
}

function fillSheets(){
  allEntries = currentFormToClass()
  // const dataSS = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1O5FjZKW40jHhIe6TZKNXx0WLLyvf75qEnEa95kIAkiA/edit?usp=sharing')
  const dataSS = SpreadsheetApp.getActiveSpreadsheet()

  //Recalculate when ever there is a change (change in what?????)
  // const interval = dataSS.setRecalculationInterval(
  //   SpreadsheetApp.RecalculationInterval.ON_CHANGE,
  // )
  
  const allSheets = dataSS.getSheets()
  const template_sheet = allSheets[allSheets.length - 1]

  //Delete all the previous year sheets
  allSheets.forEach((sheet) => {
    if((sheet.getSheetName() != "Main") && (sheet.getSheetName() != "Template")){
      dataSS.deleteSheet(sheet)
    }
  })

  //Get the years from all the entries (map) in an array
  const years = []
  allEntries.forEach((entryRecord) =>{
    exp = entryRecord.getExpirationDate()
   
    if(exp != null){
      var date = new Date(entryRecord.getExpirationDate())
      var entry_yr = date.getFullYear()
      if(!years.includes(entry_yr)){
        years.push(entry_yr)
      }
    } 
  })
  years.push(NaN)
  years.sort().reverse()//sort years array in descending order

  //Create sheets with the given years
  for(i = 0; i < years.length; i++ ){
    //Create new sheet
    var new_sheet = dataSS.insertSheet((years[i]).toFixed(0), i+1, {template: template_sheet})
    //Name the new sheet
    if(i == 0) {
      new_sheet.getRange("A1").setValue('Unknown Expiration')
    } else {
      new_sheet.getRange("A1").setValue((`Expiration: ${years[i]} `))
    }
  }

  //Add entry to the different sheets
  allEntries.forEach((entryRecord) => {
    var d = entryRecord.getExpirationDate()
    var yr1 = NaN 
    if(d != null){
      var date_1 = new Date(d)
      yr1 = date_1.getFullYear()
    } 

    var new_sheet = dataSS.getSheetByName(yr1)
    var keys     = entryRecord.getKeys()
    var rooms    = entryRecord.getRoom()
    var advisors = entryRecord.getAdvisor()
    var a = advisors
    var adv = a[0]
    for(i = 1; i < a.length; i++){
      adv = adv + ", " + a[i]
    }

    for(i = 0; i < keys.length; i++){
      new_sheet.appendRow([ 
        entryRecord.getExpirationDate(), entryRecord.getAndrewID(),
        entryRecord.getLastName(), entryRecord.getFirstName(),
        adv,  entryRecord.getDepartment(),
        keys[i], rooms[i], entryRecord.getGivenDate()
      ])
    }
  })
}

function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .createMenu('Custom Menu')  
      .addItem('Show sidebar', 'sidebarHome')
      .addItem('Show Form', 'sidebarAdd')
      .addToUi();
}

function sidebarHome() {
  var html = HtmlService.createHtmlOutputFromFile('home_sidebar')
      .setTitle('Keys Project Home');
  SpreadsheetApp.getUi().showSidebar(html);
}


//https://developers.google.com/apps-script/guides/html#code.gs_1
function sidebarAdd() {
  var html = HtmlService.createHtmlOutputFromFile('add_sidebar')
    .setWidth(400)
    .setHeight(650)
  SpreadsheetApp.getUi()
    .showModalDialog(html, 'Add New Entry') 

}

function sidebarEdit() {
  var html = HtmlService.createHtmlOutputFromFile('edit_sidebar')
    .setWidth(400)
    .setHeight(650)
  SpreadsheetApp.getUi()
    .showModalDialog(html, 'Edit Entry') 

}

function getFirstName() {
  return "Bubbly"
}

function processInputs(fname, lname, advisor, andrewID, 
                      keyNum, roomNum, givenDate, loseDate) {
  // Process the inputs here
  Logger.log('Input 1: ' + fname);
  Logger.log('Input 2: ' + lname);
  Logger.log('Input 3: ' + advisor);
  Logger.log('Input 4: ' + andrewID);
  Logger.log('Input 5: ' + keyNum);
  Logger.log('Input 6: ' + roomNum);
  Logger.log('Input 7: ' + givenDate);
  Logger.log('Input 8: ' + loseDate);
}





function isDateInFrame(start, end,date){
  if(date == null || date == undefined) return false
  return start.getTime() <= date.getTime() 
      && date.getTime()  <= end.getTime()
}

function isExpired(curr,date){
  if(date == null || date == undefined) return false
  return curr.getTime() > date.getTime()
}









function analysis(){
  // const dataSS = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1O5FjZKW40jHhIe6TZKNXx0WLLyvf75qEnEa95kIAkiA/edit?usp=sharing')
  const dataSS = SpreadsheetApp.getActiveSpreadsheet()
  var allEntries = currentFormToClass()
  
  //ADD CONDIITON FOR ALL ENTRIES
  var currDate   = new Date()
  var sixDate = new Date(currDate)
  sixDate.setMonth(sixDate.getMonth() + 6)
  var threeDate = new Date(currDate)
  threeDate.setMonth(threeDate.getMonth() + 3)
  var oneDate = new Date(currDate)
  oneDate.setMonth(oneDate.getMonth() + 1) 

  var andrew_one   = []
  var andrew_three = []
  var andrew_six   = []
  var expired_list = []
  var unknown_list = []
  allEntries.forEach((entryRecord) => {
    var expiration = new Date(entryRecord.getExpirationDate())
    if(isDateInFrame(currDate,oneDate,expiration)){
      andrew_one.push(entryRecord.getAndrewID())
    } else if (isDateInFrame(currDate,threeDate,expiration)){
      andrew_three.push(entryRecord.getAndrewID())
    } else if(isDateInFrame(currDate,sixDate,expiration)){
      andrew_six.push(entryRecord.getAndrewID())
    } else if(isExpired(currDate,expiration)){
      expired_list.push(entryRecord.getAndrewID())
    } else {
      unknown_list.push(entryRecord.getAndrewID()) 
    }
  })
  
  const sheets = dataSS.getSheets()
  const mainSheet = sheets[0]

  var six     = mainSheet.getRange("B8:B")
  var six_values = six.getValues()
  for(var i = 0; i < six_values.length; i++){
    if(i < andrew_six.length){
      mainSheet.getRange(8+i,2).setValue(andrew_six[i])
    }
  }

  var three   = mainSheet.getRange("D8:D")
  var three_values = three.getValues()
  for(var i = 0; i < three_values.length; i++){
    if(i < andrew_three.length){
      mainSheet.getRange(8+i,4).setValue(andrew_three[i])
    }
  }

  var one     = mainSheet.getRange("F8:F")
  var one_values = one.getValues()
  for(var i = 0; i < one_values.length; i++){
    if(i < andrew_one.length){
      mainSheet.getRange(8+i,6).setValue(andrew_one[i])
    }
  }


  var expired = mainSheet.getRange("H8:H")
  var expired_values = expired.getValues()
  for(var i = 0; i < expired_values.length; i++){
    if(i < expired_list.length){
      mainSheet.getRange(8+i,8).setValue(expired_list[i])
    }
  }

  var unk     = mainSheet.getRange("J8:J")
  var unk_values = unk.getValues()
  for(var i = 0; i < unk_values.length; i++){
    if(i < unknown_list.length){
      mainSheet.getRange(8+i,10).setValue(unknown_list[i])
    }
  }





  //If it is within 6 months of expiration






  //Sort through all the sections
  //Later->Add new Sheet at the end of each year
  //1.Add all the new form responses
  //2.Add form checked out to each student
  //3.Add form to show outstanding keys
  //4.Add form to show current keys
  /**Toast message should be sent if deadline 
   * is being approachd **/

}



//allEntries = parseKeySheet()