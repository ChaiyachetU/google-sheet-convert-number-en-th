///////////////////////////////////////////////////////////////////////////////////////////
//Selected cell and change to Thai Number or Arabic Number///////////////////// 
function onOpen() {

  SpreadsheetApp.getUi().createMenu("🔢Thai/Arabic Number")
                         .addItem("Change to Thai Number", "changeToThaiNumber")
                         .addItem("Change to Arabic Number", "changeToArabicNumber")
                         .addToUi();

}

// Change selection to thai number
function changeToThaiNumber() {
  
  var ui = SpreadsheetApp.getUi();
  
  var result = ui.alert(
               'Please confirm',
               'Are you sure you want to change to Thai Number?',
               ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    
    // User clicked "Yes".
    // Get the active spreadsheet
    var sheet = SpreadsheetApp.getActiveSheet();
    
    var selectionValues = sheet.getSelection().getActiveRange().getDisplayValues();
    
    var thaiNumberValues = selectionValues.map(function(numbers) {
    
      return numbers.map(function(number) { return THAINUMBER(number) })
      
    });
    
    // Set values to selection
    sheet.getSelection().getActiveRange().setValues(thaiNumberValues);
    
  } else {
  
    // User clicked "No" or X in the title bar.
    return;
  
  }
  
}

// Change selection to arabic number
function changeToArabicNumber() {

  var ui = SpreadsheetApp.getUi();
  
  var result = ui.alert(
               'Please confirm',
               'Are you sure you want to change to Arabic Number?',
               ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {

    // Get the active spreadsheet
    var sheet = SpreadsheetApp.getActiveSheet();
    
    var selectionValues = sheet.getSelection().getActiveRange().getDisplayValues();
    
    var arabicNumberValues = selectionValues.map(function(numbers) {
    
      return numbers.map(function(number) { return ARABICNUMBER(number) })
      
    });
    
    // Set values to selection
    sheet.getSelection().getActiveRange().setValues(arabicNumberValues);
    
  } else {
  
    // User clicked "No" or X in the title bar.
    return;
  
  }
}
///////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////
/**
 * แสดงผลจากตัวเลขอารบิกเป็นตัวเลขไทย
 *
 * @param {number} arabicNumber The arabic number to convert
 * @param {number} decimal The decimal of number to convert
 * @return {string} thaiNumber ตัวเลขไทยจากเลขอารบิก
 * @customfunction
 */
function THAINUMBER(arabicNumber, decimal) {
  
  if (typeof(arabicNumber) === 'string') {
    
    var thaiNumber = arabicNumberToThaiNumber(arabicNumber);
    
    return thaiNumber;
    
  } else {
  
    if (typeof(decimal) === 'undefined') decimal = 0;
  
    var arabicNumberWithDeciaml = arabicNumber.toFixed(decimal); //toFixed() method return string type.
    
    var arabicNumberWithComma = numberWithCommas(arabicNumberWithDeciaml);
    
    var thaiNumber = arabicNumberToThaiNumber(arabicNumberWithComma);
    
    return thaiNumber;
  
  }
  
}
///////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////
/**
 * แสดงผลจากตัวเลขไทยเป็นตัวเลขอารบิก.
 *
 * @param {string} thaiNumber The thai number to convert
 * @return {number} arabicNumber ตัวเลขอารบิกจากเลขไทย
 * @customfunction
 */
function ARABICNUMBER(thaiNumber) {
  var arabicNumber = thaiNumberToArabicNumber(thaiNumber);
  return arabicNumber;
}
///////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////
//set comma to number and return to string
function numberWithCommas(number) {

    var parts = number.split(".");
    
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
    
    return parts.join(".");
    
}

//set arabic number to thai number format
function arabicNumberToThaiNumber(number) {

	number = number.replace(/0/gi,'๐');
	number = number.replace(/1/gi,'๑');
	number = number.replace(/2/gi,'๒');
	number = number.replace(/3/gi,'๓');
	number = number.replace(/4/gi,'๔');
	number = number.replace(/5/gi,'๕');
	number = number.replace(/6/gi,'๖');
	number = number.replace(/7/gi,'๗');
	number = number.replace(/8/gi,'๘');
	number = number.replace(/9/gi,'๙');
	return number;
    
}

//set thai number to arabic number format
function thaiNumberToArabicNumber(number) {
	number = number.replace(/๐/gi,'0');
	number = number.replace(/๑/gi,'1');
	number = number.replace(/๒/gi,'2');
	number = number.replace(/๓/gi,'3');
	number = number.replace(/๔/gi,'4');
	number = number.replace(/๕/gi,'5');
	number = number.replace(/๖/gi,'6');
	number = number.replace(/๗/gi,'7');
	number = number.replace(/๘/gi,'8');
	number = number.replace(/๙/gi,'9');
    number = number.replace(/,/gi,'');
	return number;
}
///////////////////////////////////////////////////////////////////////////////////////////
