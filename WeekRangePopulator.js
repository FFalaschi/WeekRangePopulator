function addDateRanges() {
  //Get The Spreadsheet I'm In
  var sheet = SpreadsheetApp.getActiveSheet();
  //get the cell I'm in
  var row = sheet.getActiveCell().getRow();
  var column = sheet.getActiveCell().getColumn();
  //Get Year
  var year = getYear();
  //GetStartingWeek
  var week = getStartingWeek();
  for( var i = 0; i <= (weeksInYear(year) - week); i++) {
    activeCell = sheet.getRange(row, column + i);
    activeCell.setValue(getDateRangeOfWeek(Number(week)+i));
  }
  Logger.log("Completed");
}

//Prompt that asks the user to enter the year
function getYear() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Enter the Year:');
  
  if (response.getSelectedButton() == ui.Button.OK) {
    return response.getResponseText(); 
  } else {
    var d = new Date();
    var n = d.getFullYear();
    return n;
  }
}

//Prompt that asks the user to enter the starting week in #
function getStartingWeek() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Enter the Starting Week Number:');
  
  if (response.getSelectedButton() == ui.Button.OK) {
    return response.getResponseText();
  }
}

function getDateRangeOfWeek(weekNo){
		var d1 = new Date();
		var numOfdaysPastSinceLastMonday = eval(7 - eval( d1.getDay() + 1 ));
		d1.setDate(d1.getDate() - numOfdaysPastSinceLastMonday);
		var weekNoToday = getWeek();
		var weeksInTheFuture = eval( weekNo - weekNoToday);
		d1.setDate(d1.getDate() + eval( 7 * weeksInTheFuture ));
		var rangeIsFrom =  eval(d1.getMonth()+1)   +"/"  +  d1.getDate()  + "/"  +   d1.getFullYear();
		d1.setDate(d1.getDate() + 6);		
		var rangeIsTo = eval(d1.getMonth()+1)   +"/"  +  d1.getDate()  + "/"  +   d1.getFullYear() ;
		return rangeIsFrom + " to "+rangeIsTo;
};

function getWeek() {
  // Create a copy of this date object  
  var target  = new Date();
  var dateInMs = target.getTime();
  
  // ISO week date weeks start on monday, so correct the day number  
  var dayNr   = (target.getDay() + 6) % 7;  
  
  // Set the target to the thursday of this week so the  
  // target date is in the right year  
  target.setDate(target.getDate() - dayNr + 3);  
  
  // ISO 8601 states that week 1 is the week with january 4th in it  
  var jan4    = new Date(target.getFullYear(), 0, 4);  
  
  // Number of days between target date and january 4th  
  var dayDiff = (target - jan4) / 86400000;    
  
  if(new Date(target.getFullYear(), 0, 1).getDay() < 5) {
    // Calculate week number: Week 1 (january 4th) plus the    
    // number of weeks between target date and january 4th
    return 1 + Math.ceil(dayDiff / 7);    
  }
  else {  // jan 4th is on the next week (so next week is week 1)
    return Math.ceil(dayDiff / 7);
  }
}

function getWeekNumber(d) {
  // Copy date so don't modify original
  d = new Date(+d);
  d.setHours(0,0,0);
  // Set to nearest Thursday: current date + 4 - current day number
  // Make Sunday's day number 7
  d.setDate(d.getDate() + 4 - (d.getDay()||7));
  // Get first day of year
  var yearStart = new Date(d.getFullYear(),0,1);
  // Calculate full weeks to nearest Thursday
  var weekNo = Math.ceil(( ( (d - yearStart) / 86400000) + 1)/7)
  // Return array of year and week number
  return [d.getFullYear(), weekNo];
}

//Returns the amount of weeks in a given year. I.e: 2015 = 53, 2016 = 52
function weeksInYear(year) {
  var month = 11, day = 31, week;
  
  // Find week that 31 Dec is in. If is first week, reduce date until
  // get previous week.
  do {
    d = new Date(year, month, day--);
    week = getWeekNumber(d)[1];
  } while (week == 1);
  
  return week;
}