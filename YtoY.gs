var ssURL="https://docs.google.com/spreadsheets/d/1CdlJt6MOx0OZv47D4uhs59Pvcudbf-Y0_abi6Ss1EuU/edit#gid=0";

Array.prototype.insert = function ( index, item ) {
   this.splice( index, 0, item );
};

var ss = SpreadsheetApp.openByUrl(ssURL);
  var prevYearApptWS = ss.getSheetByName("2019 Appointments");
  var currYearApptWS = ss.getSheetByName("2020 Appointments");
  var prevYearTransWS = ss.getSheetByName("2019 Transactions");
  var currYearTransWS = ss.getSheetByName("2020 Transactions");
  var monthlySummaryWS = ss.getSheetByName("2020 Monthly Summary");

  var prevYear = [];
  var currYear = [];
  var monArr = [];
  var monListingClosed = 0;
  var monBuyersClosed = 0;

function buildTotals2() {
  
                         
  var arr = [];
  var monthsArr = [];
  
  var gtArr = [];
  var gtCount = 0;
  var gtPrice = 0;
  var gtGCI = 0;
  var gtRobinsNet = 0;
  var gtTCFee = 0;
}

 function BuildYearArray2()  {
    var monthNumber = 0   // January
    var yearCOE = '2019'  // these two values will be function arguments
    
    // get appointments and contracts from '2019 Appointments'
    var active_range = prevYearApptWS.getRange(6,1,8,12);
    arr = active_range.getValues();  
       
   for (var x=1; x<=1; x++)  {
     prevYear.push([0,0]);   // array 2 dimensions
   }
    
      prevYear[0][0] = "January " + yearCOE;
      prevYear[0][1] = arr[0][9];      // Get Monthly Appointments for Month and put in array
      prevYear[0][2] = arr[0][10];     // Get Monthly Contracts for Month and put in array
      prevYear[0][3] = parseFloat((arr[0][10] / arr[0][9]) * 100).toFixed(2);
      
    
    
    // Now get 2019 Transaction values from '2019 Transaction' sheet.
    active_range = prevYearTransWS.getRange(3, 1, prevYearTransWS.getLastRow(), prevYearTransWS.getLastColumn());
    arr = active_range.getValues();
   
    // Determine if COE date matches for month and year to scan for
    for (var i = 0; i < arr.length; i++) {
      var COEDate = new Date(arr[i][6]);      // [6] is COE Date
      if ((COEDate.getMonth() == monthNumber) && (COEDate.getYear() == yearCOE)) { 
         monArr.push(arr[i]);
         if (arr[i][3] == "LISTING") { monListingClosed += 1 }
         if (arr[i][3] == "BUYER") {monBuyersClosed += 1 }
       
       }  // if
    } // for
   
   // Now loop through monArr to get totals
   var monVolume = 0;
   var monGCI = 0;
   var monRobinsNet = 0;
   var monTransFee = 0;
   
   for (var i = 0; i < monArr.length; i++) {
     monVolume += monArr[i][7];         // price
     monGCI += monArr[i][10];           // GCI
     monRobinsNet += monArr[i][11];     // Robin's Net
     monTransFee += monArr[i][12];      // Tranaction Fee
   }
   
   //Logger.log(monTransFee);
   
   //for (var i = 0; i < monArr.length; i++) {
   //  Logger.log(monArr[i][0]);
   //}
  
    prevYear[0][4] = monListingClosed;
    prevYear[0][5] = monBuyersClosed;   
    prevYear[0][6] = parseFloat(monVolume).toFixed(2);
    prevYear[0][7] = parseFloat(monGCI).toFixed(2);
    prevYear[0][8] = parseFloat(monRobinsNet).toFixed(2);
    prevYear[0][9] = parseFloat(monTransFee).toFixed(2);
    
    Logger.log(prevYear[0]); 
    var _rows = prevYear.length;
    var _cols = prevYear[0].length;
    Logger.log(_rows);
    Logger.log(_cols);
    monthlySummaryWS.getRange(4, 1, 1, _cols).setValues(prevYear);
    Logger.log("Wrote to SS");
  }
