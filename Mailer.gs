
//Automatic email execution
function autoMail(){

  var reportSheet = SpreadsheetApp.getActiveSpreadsheet();
  var mailService = GmailApp;

  //Grab sheet properties
  var configSheet = reportSheet.getSheetByName("Report Configuration");
  var configEmailCellVal = configSheet.getRange("B21");
  var configReportCellVal = configSheet.getRange("B17");
  var configLinkCellVal = configSheet.getRange("B22");

  //Grab dashboard
  var dashboard = reportSheet.getSheetByName("Dashboard");

  //Grab cell value (i.e. email)
  var email = configEmailCellVal.getDisplayValue();

  //Sheet Name
  var reportName = configReportCellVal.getDisplayValue();

  var body = "The "+reportName+" Report is ready at: \n"+ configLinkCellVal.getDisplayValue();
  mailService.sendEmail(email,reportName + " Monthy Report", body);

}

//Add email menu item
function sendEmail(e){
  SpreadsheetApp.getUi()
  .createMenu("Email Spreadsheet")
  .addItem("Send Email",'exeMail').addToUi();

}

//Manual email execution
function exeMail(){

  var reportSheet = SpreadsheetApp.getActiveSpreadsheet();
  var mailService = GmailApp;

  //Grab sheet properties
  var configSheet = reportSheet.getSheetByName("Report Configuration");
  var configEmailCellVal = configSheet.getRange("B21");
  var configReportCellVal = configSheet.getRange("B17");
  var configLinkCellVal = configSheet.getRange("B22");

  //Grab dashboard
  var dashboard = reportSheet.getSheetByName("Dashboard");

  //Grab cell value (i.e. email)
  var email = configEmailCellVal.getDisplayValue();

  //Sheet Name
  var reportName = configReportCellVal.getDisplayValue();

  var body = "The "+reportName+" Report is ready at: \n"+ configLinkCellVal.getDisplayValue();

  //Prompted on send to show recipient email
  var promptWindow = SpreadsheetApp.getUi();
  if(email == ''){
    promptWindow.alert("An email needs to be entered in the configuration sethings");
  }else{
    promptWindow.alert("Email has been sent to "+ email);
    mailService.sendEmail(email,reportName + " Monthy Report", body);
  }


}

