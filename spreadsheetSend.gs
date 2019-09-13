function onOpen() {

  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Custom Menu Item")
      .addItem("Custom Option...","checkForSend")
      .addToUi();
};

function send(startRow,endRow){
  //var startingRow = 4 // headers above
  var submittedCol = 10
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('<Name of worksheet>');

  var name = ""
  var email = ""
  var class = ""
  var account = ""
  var link = ""
  var desc = ""
  var quant = ""
  var price = ""
  var date_needed = ""
  var mesg = ""

  mesg = mesg + '<html><img src="<validpath_image>" heigth="100" width="125">'
  mesg = mesg + '<h1>Summary</h1>'
  mesg + mesg + '<h2>The following will be ordered:</h2>'
  mesg = mesg + '<head><style> table, th, td {border: 1px solid black;border-collapse: collapse;}</style></head>'
  mesg = mesg + "<table><tr>\n"
  mesg = mesg + '<th valign="top" align="left">Name</th><th valign="top" align="left">Email</th><th valign="top" align="left">Class</th><th valign="top" align="left">Account</th><th valign="top" align="left">Link</th><th valign="top" align="left">Description</th><th valign="top" align="left">Quantity</th><th valign="top" align="left">Price</th><th valign="top" align="left">Line Total</th><th valign="top" align="left">Needed By</th></tr>\n'


  mesg = mesg + '<tr>\n'
  var cc_emails = ''
  var total = 0
  var bSend = 'False'
  while (startRow < endRow){
      bSend = 'True'
      mesg = mesg + '<tr>\n'
      //Browser.msgBox(startRow +" - "+ endRow)
      mesg = mesg + '<td valign="top">' + sheet.getRange(startRow,1).getValue() + '</td>'
      var cell = sheet.getRange(startRow,1)
      cell.setBackground('#E9E4E6')

      mesg = mesg + '<td valign="top">' + sheet.getRange(startRow,2).getValue() + '</td>'
      cc_emails = cc_emails + sheet.getRange(startRow,2).getValue() + ','
      var cell = sheet.getRange(startRow,2)
      cell.setBackground('#E9E4E6')

      mesg = mesg + '<td valign="top">' + sheet.getRange(startRow,3).getValue() + '</td>'
      var cell = sheet.getRange(startRow,3)
      cell.setBackground('#E9E4E6')

      mesg = mesg + '<td valign="top">' + sheet.getRange(startRow,4).getValue() + '</td>'
      var cell = sheet.getRange(startRow,4)
      cell.setBackground('#E9E4E6')

      mesg = mesg + '<td valign="top">' + sheet.getRange(startRow,5).getValue() + '</td>'
      var cell = sheet.getRange(startRow,5)
      cell.setBackground('#E9E4E6')

      mesg = mesg + '<td valign="top">' + sheet.getRange(startRow,6).getValue() + '</td>'
      var cell = sheet.getRange(startRow,6)
      cell.setBackground('#E9E4E6')

      mesg = mesg + '<td valign="top">' + sheet.getRange(startRow,7).getValue() + '</td>'
      var cell = sheet.getRange(startRow,7)
      cell.setBackground('#E9E4E6')

      mesg = mesg + '<td valign="top">' + '$' + sheet.getRange(startRow,8).getValue() + '</td>'
      var cell = sheet.getRange(startRow,8)
      cell.setBackground('#E9E4E6')

      //multiply quantity and price
      total = total + (sheet.getRange(startRow,7).getValue() * sheet.getRange(startRow,8).getValue())
      mesg = mesg + '<td valign="top">' + '$' + (sheet.getRange(startRow,7).getValue() * sheet.getRange(startRow,8).getValue()) + '</td>'
      var cell = sheet.getRange(startRow,9)
      cell.setBackground('#E9E4E6')

     // var date = sheet.getRange(startRow,9).getValue()
      mesg = mesg + '<td valign="top">' + String(sheet.getRange(startRow,9).getValue()).slice(0,-23) + '</td></tr>'
      var cell = sheet.getRange(startRow,10)

      cell.setBackground('#E9E4E6')

      var cell = sheet.getRange(startRow,submittedCol)
      cell.setValue('yes')
      startRow += 1
    }
  mesg = mesg + '<th valign="top" align="left"> </th><th valign="top" align="left"> </th><th valign="top" align="left"> </th><th valign="top" align="left"> </th><th valign="top" align="left"> </th><th valign="top" align="left"> </th><th valign="top" align="left"> </th><th valign="top" align="left"> </th><th valign="top" align="left"><font color="red">$'+total+'</th><th valign="top" align="left"> </th></tr>\n'

  mesg = mesg + '</table></html>'
  cc_emails = cc_emails.slice(0,-1)



  var to_emails = '<valid_email>,<valid_email>'
  //var to_emails = ''
  if (bSend == 'True'){
    MailApp.sendEmail({
        to: to_emails,
        cc: cc_emails,
        subject: 'Email Subject Goes Here ' + Utilities.formatDate(new Date(), "GMT-4", "MMM dd yyyy' 'HH:mm:ss' '"),
        htmlBody: mesg,
        name: '<valid_email>'
      });
  }
  else{
    Browser.msgBox('Nothing new to submit')
 }



};
function onEdit(e){
  var activeRange = e.source.getActiveRange();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('<Name of worksheet>');
  //var cell = sheet.getRange(activeRange)
  if(e.value != "<Missing Info>") { // add to cell
    var cell = sheet.getActiveCell();
    cell.setBackground("white");
  }
};
function checkForSend(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('<Name of worksheet>');
  var startRow = getStartingRow()
  var endRow = getLastRow()

  var submittedCol = 10
  var canSubmit = "True"
  var start = startRow
  //var end = endRow
  while (start < endRow){
    for (var i = 1;i<submittedCol;i++){
      cell = sheet.getRange(start,i)
      if (String(cell.getValue()).trim() == "" || String(cell.getValue()).trim() == "<Missing Info>"){
        cell.setValue("<Missing Info>")
        cell.setBackground("#FFC300")
        canSubmit = "False"
      }
    }
    start += 1
  }
  if (canSubmit == "False"){
    Browser.msgBox("Please provide missing info before submitting")
  }
  else{
    send(startRow,endRow)
  }
};

function getLastRow(){
  var nonHeaderRow = 4 // column headers are above this 4 index of 4
  var submittedCol = 10
  var row = getStartingRow()
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('<Name of worksheet>');

  //start with 1st cell in starting row
  var allCellsInRow = "x"
  var cell = ""
  while (allCellsInRow != ""){
    allCellsInRow = ""
    for (var i = 1;i<submittedCol;i++){
      cell = sheet.getRange(row,i)
      if (String(cell.getValue()).trim() == ""){

      }
      allCellsInRow = allCellsInRow + String(cell.getValue()).trim()
      //cell.setBackground("orange")
    }
    row += 1
  }
  return row - 1

};
function getStartingRow(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('<Name of worksheet>');
  var nonHeaderRow = 4 // column headers are above this 4 index of 4
  var submittedCol = 10

  // check first valid row and the submitted col to see if cell == 'yes'
  var cell = sheet.getRange(nonHeaderRow,submittedCol)
  while (cell.getValue() == 'yes'){
    cell = sheet.getRange(nonHeaderRow+=1,submittedCol)

  }

  return cell.getRow()

};
function findEmptyCellsInRow(cell){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('<Name of worksheet>');
  var startRow = cell.getRow() //
  var submittedCol = 10
  var startCol = 1
  // check first valid row and the submitted col to see if cell == 'yes'
  // go through each cell and check for end of table and whether or not cells are empty

  var cell = sheet.getRange(startRow,1)

  while (startCol <= 10){
    if (cell.isBlank() || (String(entry).trim() == "" )){
      cell.setValue('Fill in before submitting')
      cell.setBackground("#EC512A")
    }

  }


};

function reset(){
  var startingRow = 4
  var submittedCol = 10
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('<Name of worksheet>');
  // set submitted cell to yes
  var cell = sheet.getRange(startingRow,submittedCol)
  cell.setValue('')
  // now set whole row to background grey.
  for (var i = 1; i <= submittedCol; i++) {

    var cell = sheet.getRange(startingRow,i)
    cell.setBackground("white")
  }
};
