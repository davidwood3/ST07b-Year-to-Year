var ssURL="https://docs.google.com/spreadsheets/d/1CdlJt6MOx0OZv47D4uhs59Pvcudbf-Y0_abi6Ss1EuU/edit#gid=0";

function onOpen() {
  // get active spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // create menu
  var menu = [{name: "Totals By Source", functionName: "buildTotalsBySource"},
              {name: "Subtoals By Month", functionName: "buildsubTotals"},
              {name: "Format Data Entry Sheet", functionName: "formatDataEntered"},
              {name: "Backup Spreadsheet on Demand", functionName: "makeCopy"}
             ];

  // add to menu
  ss.addMenu("Smith Team", menu);  
}

function formatDataEntered() {
  var ss = SpreadsheetApp.openByUrl(ssURL);
  var dataEnteredSheet = ss.getSheetByName("DataEntered");
  
  var lr = dataEnteredSheet.getLastRow();
  
  // if month number is in column no need to run this script
  if (dataEnteredSheet.getRange(lr,14).getValue() != "") {
    Logger.log("Skipped");
    return;
  }
  
  dataEnteredSheet.getRange("N2").setFormula("=Month(G2)");
  dataEnteredSheet.getRange("O2").setFormula('=ArrayFormula(text(date(2019,N2,1),"mmmm"))');
  

  var fillDownRange1 = dataEnteredSheet.getRange(2,14 ,lr-1);
  var fillDownRange2 = dataEnteredSheet.getRange(2,15 ,lr-1);
  dataEnteredSheet.getRange("N2").copyTo(fillDownRange1);
  dataEnteredSheet.getRange("O2").copyTo(fillDownRange2);
  

  // turn background color to light green if a subtotal.
  dataEnteredSheet.getRange(1,1,1,dataEnteredSheet.getLastColumn()).setBackgroundRGB(0, 255, 255);
  dataEnteredSheet.getRange(1,1,1,dataEnteredSheet.getLastColumn()).setFontWeight("bold");
  
  dataEnteredSheet.setColumnWidth(1,190);    // Property
  dataEnteredSheet.setColumnWidth(2,70);     // Street     
  dataEnteredSheet.setColumnWidth(3,200);    // Source
  dataEnteredSheet.setColumnWidth(4,150);    // Listing or Buyer
  dataEnteredSheet.setColumnWidth(5,130);    // Agent
  dataEnteredSheet.setColumnWidth(6,130);    // Contract Date
  dataEnteredSheet.setColumnWidth(7,130);    // COE
  dataEnteredSheet.setColumnWidth(8,140);    // Price
  dataEnteredSheet.setColumnWidth(9,140);    // Commission %
  dataEnteredSheet.setColumnWidth(10,110);   // Robins %
  dataEnteredSheet.setColumnWidth(11,120);   // GCI
  dataEnteredSheet.setColumnWidth(12,120);   // Robins Net CI
  dataEnteredSheet.setColumnWidth(13,90);    // TC Fee
  dataEnteredSheet.setColumnWidth(14,100);   // Month Name
  
  dataEnteredSheet.getRange('H:H').setNumberFormat("$#,##0.00;$(#,##0.00)");   // Price
  dataEnteredSheet.getRange('I:I').setNumberFormat('0.00');                    // Commission %
  dataEnteredSheet.getRange('J:J').setNumberFormat('00');                      // Robins %
  dataEnteredSheet.getRange('K:K').setNumberFormat("$#,##0.00;$(#,##0.00)");   // GCI
  dataEnteredSheet.getRange('L:L').setNumberFormat("$#,##0.00;$(#,##0.00)");   // Robins Net CI
  dataEnteredSheet.getRange('M:M').setNumberFormat("$#,##0.00;$(#,##0.00)");   // TC Fee
  
  //ScriptApp.newTrigger('onChange')
  //    .forSpreadsheet(ss)
  //    .onChange()
  //    .create();
}

