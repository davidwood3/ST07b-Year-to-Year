var ssURL="https://docs.google.com/spreadsheets/d/1CdlJt6MOx0OZv47D4uhs59Pvcudbf-Y0_abi6Ss1EuU/edit#gid=0";

function onOpen() {
  // get active spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // create menu
  var menu = [{name: "Build Year to Year Totals", functionName: "buildTotals"},
              {name: "Format Monthly Transaction Summary", functionName: "formatMonthlyTransactionSummary"},
              {name: "Backup Spreadsheet on Demand", functionName: "makeBackup"}
             ];

  // add to menu
  ss.addMenu("Smith Team", menu);  
}

function makeBackup() {

  var timeZone = Session.getScriptTimeZone();

  // generates the timestamp and stores in variable formattedDate as year-month-date hour-minute-second
  var formattedDate = Utilities.formatDate(new Date(), timeZone , "yyyy-MM-dd' 'HH:mm:ss");
    
  // getting name of the original file and appending the word "copy" followed by the timestamp stored in formattedDate
  var saveAs = SpreadsheetApp.getActiveSpreadsheet().getName() + " Copy " + formattedDate;
  
  // getting the destination folder by their ID
  var destinationFolder = DriveApp.getFolderById("1RilTIR2PsIuMtH3lgvyh7itk9X_OGckl");
  
  DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).makeCopy(saveAs, destinationFolder);
}

function formatMonthlyTransactionSummary() {
  var ss = SpreadsheetApp.openByUrl(ssURL);
  var monthSummary = ss.getSheetByName("2020 Monthly Summary");
  
  var lr = monthSummary.getLastRow();
  
  monthSummary.setColumnWidth(1,190);    // Month
  monthSummary.setColumnWidth(2,90);     // Appointments Made     
  monthSummary.setColumnWidth(3,90);     // Contracts
  monthSummary.setColumnWidth(4,150);    // Conversion Rate
  monthSummary.setColumnWidth(5,110);    // Listings Closed
  monthSummary.setColumnWidth(6,110);    // Buyers Closed
  monthSummary.setColumnWidth(7,150);    // Volume
  monthSummary.setColumnWidth(8,150);    // GCI
  monthSummary.setColumnWidth(9,150);    // Robins Net CI
  monthSummary.setColumnWidth(10,150);   // Transaction Fee
  
  monthSummary.getRange('G:G').setNumberFormat("$#,##0.00;$(#,##0.00)");   // Volume
  monthSummary.getRange('D:D').setNumberFormat('0.00');                    // Commission %
  monthSummary.getRange('H:H').setNumberFormat("$#,##0.00;$(#,##0.00)");   // GCI
  monthSummary.getRange('I:I').setNumberFormat("$#,##0.00;$(#,##0.00)");   // Robins Net CI
  monthSummary.getRange('J:J').setNumberFormat("$#,##0.00;$(#,##0.00)");   // TC Fee
  
  monthSummary.setFrozenRows(1);
  monthSummary.setFrozenRows(2);
  monthSummary.setFrozenRows(3);
  
  
  
}

