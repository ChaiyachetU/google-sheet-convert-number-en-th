///////////////////////////////////////////////////////////////////////////////////////////
/**
 * แสดงผลจากตัวเลขอารบิกเป็นตัวเลขไทย
 *
 * @param {number} input the number to convert
 * @param {number} decimal the decimal of number to convert
 * @return ตัวเลขไทยจากเลขอารบิก
 * @customfunction
 */
function THAINUMBER(input, decimal) {
  
  if (typeof(input) === 'string') {
    
    var output = arabicNumberToThaiNumber(input);
    
    return output;
    
  } else {
  
    if (typeof(decimal) === 'undefined') decimal = 0;
  
    var inputWithDeciaml = input.toFixed(decimal); //toFixed() method return string type.
    
    var inputWithComma = numberWithCommas(inputWithDeciaml);
    
    var output = arabicNumberToThaiNumber(inputWithComma);
    
    return output;
  
  }
  
}

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
///////////////////////////////////////////////////////////////////////////////////////////


///////////////////////////////////////////////////////////////////////////////////////////
/**
 * แสดงผลจากตัวเลขไทยเป็นตัวเลขอารบิก.
 *
 * @param {number} input the number to convert.
 * @return The arabic number format.
 * @customfunction
 */
function ARABICNUMBER(input) {
  var output = thaiNumberToArabicNumber(input);
  return output;
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

  // Get the active spreadsheet
  var sheet = SpreadsheetApp.getActiveSheet();
  
  var selectionValues = sheet.getSelection().getActiveRange().getDisplayValues();
  
  var thaiNumberValues = selectionValues.map(function(numbers) {
  
    return numbers.map(function(number) { return THAINUMBER(number) })
    
  });
  
  // Set values to selection
  sheet.getSelection().getActiveRange().setValues(thaiNumberValues);
}

// Change selection to arabic number
function changeToArabicNumber() {

  // Get the active spreadsheet
  var sheet = SpreadsheetApp.getActiveSheet();
  
  var selectionValues = sheet.getSelection().getActiveRange().getDisplayValues();
  
  var arabicNumberValues = selectionValues.map(function(numbers) {
  
    return numbers.map(function(number) { return ARABICNUMBER(number) })
    
  });
  
  // Set values to selection
  sheet.getSelection().getActiveRange().setValues(arabicNumberValues);

}
///////////////////////////////////////////////////////////////////////////////////////////
