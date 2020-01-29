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

  var wholeArray = [];
  var _Year = [];
  var monArr = [];
  var monListingClosed = 0;
  var monBuyersClosed = 0;

function buildTotals() {
  
                         
  var arr = [];
  var monthsArr = [];
  
  var currOrPrev = '';
  var currYear = '2020';
  var prevYear = '2019';
  
  
  wholeArray = wholeArray.concat(BuildYearArray('January', 0, currYear, "curr"));
  wholeArray = wholeArray.concat(BuildYearArray('January', 0, prevYear, "prev"));
  //Logger.log(wholeArray);
  //BuildDelta();
  var _Rows = wholeArray.length;
  var _Cols = wholeArray[0].length;
  var _Range = monthlySummaryWS.getRange(4,1,_Rows,_Cols)
  _Range.setValues(wholeArray);
}

 function BuildYearArray(monthName, monthNumber, yearCOE, prevCurr)  {
    _Year = [];
    monArr = [];
   
    // get appointments and contracts from Appointments'
    var active_range = currYearApptWS.getRange(6,1,8,12);
    if  (prevCurr === "prev") {
      active_range = prevYearApptWS.getRange(6,1,8,12);
    }  
    arr = active_range.getValues();  
     //Logger.log(arr);
    //for (var x=1; x<=1; x++)  {
     _Year.push([0,0]);   // array 2 dimensions
    //}
   
    _Year[0][0] = monthName + ' ' + yearCOE;
    _Year[0][1] = arr[0][9];      // Get Monthly Appointments for Month and put in array
    _Year[0][2] = arr[0][10];     // Get Monthly Contracts for Month and put in array
    _Year[0][3] = parseFloat((arr[0][10] / arr[0][9]) * 100).toFixed(2);
      
    
    
    // Now get 2019 Transaction values from '2019 Transaction' sheet.
     var active_range = currYearTransWS.getRange(3,1,currYearTransWS.getLastRow(), currYearTransWS.getLastColumn());
     if  (prevCurr === "prev") {
        active_range = prevYearTransWS.getRange(3,1,prevYearTransWS.getLastRow(), prevYearTransWS.getLastColumn());
      }
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
   
    _Year[0][4] = monListingClosed;
    _Year[0][5] = monBuyersClosed;   
    _Year[0][6] = parseFloat(monVolume).toFixed(2);
    _Year[0][7] = parseFloat(monGCI).toFixed(2);
    _Year[0][8] = parseFloat(monRobinsNet).toFixed(2);
    _Year[0][9] = parseFloat(monTransFee).toFixed(2);
    
     return _Year;
  }
               

               
