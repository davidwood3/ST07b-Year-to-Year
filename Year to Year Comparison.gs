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
  var deltaArr = [];
  var apptArr = [];

function buildTotals() {
  var arr = [];
  var monthsArr = [];

  var currOrPrev = '';
  var currYear = '2020';
  var prevYear = '2019';
  
  
  wholeArray = wholeArray.concat(BuildYearArray('January', 0, currYear, "curr"));
  wholeArray = wholeArray.concat(BuildYearArray('January', 0, prevYear, "prev"));
  wholeArray = wholeArray.concat(BuildDelta('January',0, currYear, prevYear));
  wholeArray = wholeArray.concat(BuildYearArray('February', 1, currYear, "curr"));
  wholeArray = wholeArray.concat(BuildYearArray('February', 1, prevYear, "prev"));
  wholeArray = wholeArray.concat(BuildDelta('February',1, currYear, prevYear));
  
  var _Rows = wholeArray.length;
  var _Cols = wholeArray[0].length;
  var _Range = monthlySummaryWS.getRange(4,1,_Rows,_Cols)
  _Range.setValues(wholeArray);
  
  formatSheet();
}

 function BuildYearArray(monthName, monthNumber, yearCOE, prevCurr)  {
    _Year = [];
    monArr = [];
    apptArr = [];
    
    // get appointments and contracts from Appointments'
    var active_range = currYearApptWS.getRange(6,1,12,currYearApptWS.getLastColumn());
    if  (prevCurr === "prev") {
      active_range = prevYearApptWS.getRange(6,1,12,prevYearApptWS.getLastColumn());
    }  
    arr = active_range.getValues();  
       
    arr.push([0][0]);
    _Year.push([0,0]);   // array 2 dimensions
      
    for (var x = 0; x < arr.length-1; x++) {
       var str = arr[x][0].toString();
       var apptDate = new Date(arr[x][0]);      // [0] is COE Date 
      
       if ((apptDate.getMonth() == monthNumber) && (apptDate.getYear() == yearCOE) && (prevCurr == "prev")) {
         apptArr.push(arr[x]);
       }
       if ((apptDate.getMonth() == monthNumber) && (apptDate.getYear() == yearCOE) && (prevCurr == "curr")) {
         apptArr.push(arr[x]);
       }
    }
   
   // if either prevArr or currArray has a value store values in _Year array
   if (apptArr != null) {
      _Year[0][0] = monthName + ' ' + yearCOE;
      _Year[0][1] = apptArr[0][9];      // Get Monthly Appointments for Month and put in array
      _Year[0][2] = apptArr[0][10];     // Get Monthly Contracts for Month and put in array
      _Year[0][3] = parseFloat((apptArr[0][10] / apptArr[0][9]) * 100).toFixed(2);
   }  
   
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
               
function BuildDelta(monthName, monthNumber, currYear, prevYear)  {
  deltaArr = [];
  
  var summaryValues = monthlySummaryWS.getRange(4,1,monthlySummaryWS.getLastRow(),monthlySummaryWS.getLastColumn()).getValues();
    
  for (var x = 0; x < summaryValues.length; x++) {
    var deltaDate = new Date(summaryValues[x][0]);
    if ((deltaDate.getYear() == prevYear) && (deltaDate.getMonth()== monthNumber)) {
      var prevArr = summaryValues[x];
    }
    if ((deltaDate.getYear() == currYear) && (deltaDate.getMonth() == monthNumber)) {
      var currArr = summaryValues[x];
    }
  }
  
  deltaArr.push([0,0]);   // make array 2 dimensions
  
  deltaArr[0][0] = "MOM Delta";
  deltaArr[0][1] = currArr[1] - prevArr[1];                  // Appointments Made
  deltaArr[0][2] = currArr[2] - prevArr[2];                  // Contracts
  deltaArr[0][3] = "";                                       // Conversion Rate
  deltaArr[0][4] = currArr[4] - prevArr[4];                  // Listings Closed
  deltaArr[0][5] = currArr[5] - prevArr[5];                  // Buyers Closed
  deltaArr[0][6] = (currArr[6] - prevArr[6]).toFixed(2);     // Volume
  deltaArr[0][7] = (currArr[7] - prevArr[7]).toFixed(2);     // GCI
  deltaArr[0][8] = (currArr[8] - prevArr[8]).toFixed(2);     // Robins Net
  deltaArr[0][9] = (currArr[9] - prevArr[9]).toFixed(2);     // TC Fee
  
  return deltaArr;
}
               
function formatSheet()  {
  var data = monthlySummaryWS.getDataRange();
 
  dataValues = data.getValues();
  
  for (i=1; i < data.getLastRow(); i++) {
    str = dataValues[i][0].toString()
    if (str.indexOf("MOM Delta")>-1) {
      monthlySummaryWS.getRange(i+1,1,1,10).setBackgroundRGB(152, 252, 152);
    }  else {
      monthlySummaryWS.getRange(i+1,1,1,10).setBackgroundRGB(255, 255, 255);
    } // if
  } // for
}