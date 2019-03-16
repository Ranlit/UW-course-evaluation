

/**
 * Add a custom menu to the spreadsheet.
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Prepare sheet...', functionName: 'prepareSheet_'},
    {name: 'Generate profile...', functionName: 'generateProfile_'}
  ];
  spreadsheet.addMenu('CourseRatings', menuItems);
}

// Some "global constants" (not really constants but wtv)
// Here is where all data about majors and plans should be (should find a better way later)
  var majors = [['CS','ACTSC','engineering omegalul']];
  var majorlen = majors[0].length;
  var csCourses = [['M135'], ['M136'], ['M137'],['M138'],['M235'],['M237'],['M239'],['S230'],['S231'],
                   ['CS135'],['CS136'],['CS240'],['CS241'],['CS245'],['CS246'],['CS251'],['CS341'],['CS350']];
  var actscCourses = [['MTHELbs']];
  var engCourses = [['retardedshit101'],['retardedshit102']];
  var courses = [csCourses, actscCourses, engCourses];


/**
 * A function that adds majors and some initial data to the spreadsheet.
 */

/** The length of the outer array in the 2D array is always the number of rows. 
 * The inner array is always the number of columns. 
 * But, you can have a 2D array, where the inner arrays are of a different length. Then, you'd get an error.
 */
function prepareSheet_() {
  var sheet = SpreadsheetApp.getActiveSheet().setName('Settings');
  sheet.getRange(1, 1, 1, majorlen).setValues(majors).setFontWeight('bold');
                 
  for (i = 0; i < courses.length; i++) {
                 sheet.getRange(2, i+1, courses[i].length, 1).setValues(courses[i]);
  }
  
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, majorlen);
}


/**
 * Creates a new sheet containing all must-take courses according to the major
 * on the "Settings" sheet that the user selected. Generate all cell inputs 
 * with data validation and conditional formating as well.
 */
function generateProfile_() {
  var spreadsheet = SpreadsheetApp.getActive();
  var settingsSheet = spreadsheet.getSheetByName('Settings');
  settingsSheet.activate();

  // Prompt the user for his name
  var name = Browser.inputBox('Generate profile',
      'Please enter Kimi no Na wa',
      Browser.Buttons.OK_CANCEL);
  
  if (selectedCol == 'cancel') {
    return;
  }
  
  // Prompt the user for a col number.
  var selectedCol = Browser.inputBox('Generate profile',
      'Please enter the column number representing your major field of study',
      Browser.Buttons.OK_CANCEL);
  
  if (selectedCol == 'cancel') {
    return;
  }
  
  var colNumber = Number(selectedCol);
  if (isNaN(colNumber) ||
      colNumber > settingsSheet.getLastColumn()) {
    Browser.msgBox('Error',
        Utilities.formatString('Col "%s" is not valid.', selectedCol),
        Browser.Buttons.OK);
    return;
  }

  // amount represents the amount of courses in the field of study
  var amount = courses[colNumber-1].length


  // Create a new sheet
  var sheetName = name;
  var evalSheet = spreadsheet.getSheetByName(sheetName);
  if (evalSheet) {
    evalSheet.clear();
    evalSheet.activate();
  } else {
    evalSheet =
        spreadsheet.insertSheet(sheetName, spreadsheet.getNumSheets());
  }
  
  // Retrieve the courses
  var row = settingsSheet.getRange(2, colNumber, courses[colNumber-1].length, 1);
  // Copy to the first column of new sheet
  var lstcourses = evalSheet.getRange(2, 1, amount, 1);
  row.copyTo(lstcourses);

  // Format the new sheet
  evalSheet.getRange('A1:1').setFontWeight('bold');

  evalSheet.setRowHeights(2, amount, 30);
  evalSheet.setColumnWidth(1, 100);
  evalSheet.setColumnWidth(7, 500);
  
  evalSheet.getRange('B1').setValue('Usefulness');
  evalSheet.getRange('C1').setValue('Easiness');
  evalSheet.getRange('D1').setValue('Interestingness');
  evalSheet.getRange('E1').setValue('Skipable');
  evalSheet.getRange('F1').setValue('100-able?');
  evalSheet.getRange('G1').setValue('Further comments');
  
  evalSheet.getRange('A1:A').setVerticalAlignment('center');
  evalSheet.getRange('B1:1').setVerticalAlignment('center');
  evalSheet.getRange('B1:1').setHorizontalAlignment('center');

  evalSheet.setFrozenColumns(1);
  evalSheet.setFrozenRows(1);
  
  // Data validation stuff
  var nbrange = SpreadsheetApp.newDataValidation().requireValueInList([1,2,3,4,5]).setAllowInvalid(false).build();
  evalSheet.getRange(2, 3, amount, 3).setDataValidation(nbrange);
  
  var nbrange2 = SpreadsheetApp.newDataValidation().requireNumberBetween(0, 1).setAllowInvalid(false).build();
  evalSheet.getRange(2, 2, amount, 1).setNumberFormat('0.00').setDataValidation(nbrange2);
  
  var enforceCheckbox = SpreadsheetApp.newDataValidation();
  enforceCheckbox.requireCheckbox();
  enforceCheckbox.setAllowInvalid(false);
  enforceCheckbox.build();
  evalSheet.getRange(2, 6, amount, 1).setDataValidation(enforceCheckbox);
  
  // Conditional formatting stuff
  var color1 = "#C8E0F2";
  var color2 = "#A1D1F4";
  var color3 = "#84C4F3";
  var color4 = "#60B5F3";
  var color5 = "#39A5F4";
  
  var range = evalSheet.getRange(2, 3, amount, 3);
  var rules = evalSheet.getConditionalFormatRules();
  
  var rule1 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(1)
    .setBackground(color1)
    .setRanges([range])
    .build();
  rules.push(rule1);
  
  var rule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(2)
    .setBackground(color2)
    .setRanges([range])
    .build();
  rules.push(rule2);
  
  var rule3 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(3)
    .setBackground(color3)
    .setRanges([range])
    .build();
  rules.push(rule3);
  
  var rule4 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(4)
    .setBackground(color4)
    .setRanges([range])
    .build();
  rules.push(rule4);
  
  var rule5 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(5)
    .setBackground(color5)
    .setRanges([range])
    .build();
  rules.push(rule5);
  
  evalSheet.setConditionalFormatRules(rules);
  
  var range2 = evalSheet.getRange(2, 2, amount, 1);
  var rules2 = evalSheet.getConditionalFormatRules();
  
  var rule6 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(0.00,0.20)
    .setBackground(color1)
    .setRanges([range2])
    .build();
  rules2.push(rule6);
  
  var rule7 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(0.20,0.40)
    .setBackground(color2)
    .setRanges([range2])
    .build();
  rules2.push(rule7);
  
  var rule8 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(0.40,0.60)
    .setBackground(color3)
    .setRanges([range2])
    .build();
  rules2.push(rule8);
  
  var rule9 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(0.60,0.80)
    .setBackground(color4)
    .setRanges([range2])
    .build();
  rules2.push(rule9);
  
  var rule10 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(0.80,1.00)
    .setBackground(color5)
    .setRanges([range2])
    .build();
  rules2.push(rule10);
  
  evalSheet.setConditionalFormatRules(rules2);


  SpreadsheetApp.flush();
}
