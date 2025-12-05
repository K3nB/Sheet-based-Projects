/**
 * @OnlyCurrentDoc
 */
/**
 * The onOpen function runs automatically every time the spreadsheet is opened.
 * It's used here to create a custom menu bar item.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Custom Tools') // The name of your custom menu
      .addItem('Open Web App', 'openUrl') // The menu item name and the function to run
      .addToUi();
}

/**
 * This function opens the specified URL in a new browser tab.
 */
function openUrl() {
  const url = 'https://script.google.com/macros/s/AKfycbw5GK78tpVSaG7ezQgf1LZ2cJtmWIP2nnq61QF6rbFYQDxwctiTruZPSy4P3GoJrSgopQ/exec';
  const html = `<script>window.open('${url}', '_blank'); google.script.host.close();</script>`;
  
  // Display a small dialog with the JavaScript to execute the redirect
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(1).setHeight(1), 'Opening...');
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Employee Attendance Form')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Get employee data from Settings sheet for dropdown
 */
function getEmployeeData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var settingsSheet = ss.getSheetByName('Settings');
    
    // Get Employee IDs from column D (D3 onwards)
    var employeeIds = settingsSheet.getRange('D3:D').getValues()
      .map(function(row) { return row[0]; })
      .filter(function(val) { return val !== ''; });
    
    // Get Employee Names from column E (E3 onwards)
    var employeeNames = settingsSheet.getRange('E3:E').getValues()
      .map(function(row) { return row[0]; })
      .filter(function(val) { return val !== ''; });
    
    // Create employee map for quick lookup
    var employeeMap = {};
    for (var i = 0; i < employeeIds.length; i++) {
      employeeMap[employeeIds[i]] = employeeNames[i];
    }
    
    return {
      ids: employeeIds,
      names: employeeNames,
      map: employeeMap
    };
  } catch (error) {
    throw new Error('Failed to load employee data: ' + error.toString());
  }
}

/**
 * Get standard working hours from Settings sheet (J1)
 */
function getStandardHours() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var settingsSheet = ss.getSheetByName('Settings');
    var standardHours = settingsSheet.getRange('J1').getValue();
    return parseFloat(standardHours) || 8; // Default to 8 hours
  } catch (error) {
    return 8;
  }
}

/**
 * Submit attendance records to the sheet
 */
/**
 * Submit attendance records to the sheet
 */
function submitAttendance(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dataSheet = ss.getSheetByName('Attendance log');
    
    if (!dataSheet) {
      dataSheet = ss.insertSheet('Attendance log');
      dataSheet.getRange('A1:H1').setValues([[
        'UNIQID', 'Date', 'Employee Id', 'Employee Name', 
        'Check-in', 'Check-out', 'Total Working Hour', 'Status'
      ]]);
    }
    
    // Parse date string correctly to avoid timezone issues
    var dateParts = data.date.split('-');
    var year = parseInt(dateParts[0]);
    var month = parseInt(dateParts[1]) - 1;
    var day = parseInt(dateParts[2]);
    var selectedDate = new Date(year, month, day);
    
    var standardHours = getStandardHours();
    var halfDayHours = standardHours / 2;
    var rows = [];
    
    // Calculate date serial using Google Sheets formula
    // Google Sheets epoch: December 30, 1899
    var epoch = new Date(Date.UTC(1899, 11, 30));
    var dateSerial = Math.round((selectedDate - epoch) / (1000 * 60 * 60 * 24));
    
    // Process each attendance record
    for (var i = 0; i < data.records.length; i++) {
      var record = data.records[i];
      
      if (!record.employeeId) continue;
      
      var employeeId = record.employeeId;
      var employeeName = record.employeeName;
      var checkIn = record.checkIn;
      var checkOut = record.checkOut;
      
      // Calculate UNIQID (Date Serial * Employee ID)
      var uniqId = dateSerial * parseInt(employeeId);
      
      // Calculate total working hours and status
      var totalHours = 0;
      var status = 'A';
      
      if (checkIn && checkOut) {
        totalHours = calculateHours(checkIn, checkOut);
        
        // Status logic based on working hours
        if (totalHours < 3) {
          status = 'A'; // Absent - less than 3 hours
        } else if (totalHours < halfDayHours) {
          status = 'HD'; // Half Day - less than half of full day
        } else if (totalHours >= standardHours) {
          status = 'E'; // Overtime/Extra Hours
        } else {
          status = 'P'; // Present/Full Day
        }
      } else if (checkIn && !checkOut) {
        status = 'A';
      }
      
      rows.push([
        uniqId,
        selectedDate,
        employeeId,
        employeeName,
        checkIn || '',
        checkOut || '',
        totalHours.toFixed(2),
        status
      ]);
    }
    
    if (rows.length > 0) {
      dataSheet.getRange(dataSheet.getLastRow() + 1, 1, rows.length, 8).setValues(rows);
      
      // Format the date column as MM/DD/YYYY
      var lastRow = dataSheet.getLastRow();
      dataSheet.getRange(lastRow - rows.length + 1, 2, rows.length, 1)
        .setNumberFormat('MM/dd/yyyy');
    }
    
    return {
      success: true,
      message: 'Successfully submitted ' + rows.length + ' attendance record(s)!'
    };
    
  } catch (error) {
    return {
      success: false,
      message: 'Error: ' + error.toString()
    };
  }
}


/**
 * Calculate hours between two times
 */
function calculateHours(checkIn, checkOut) {
  if (!checkIn || !checkOut) return 0;
  
  var inParts = checkIn.split(':');
  var outParts = checkOut.split(':');
  
  var inMinutes = parseInt(inParts[0]) * 60 + parseInt(inParts[1]);
  var outMinutes = parseInt(outParts[0]) * 60 + parseInt(outParts[1]);
  
  var diffMinutes = outMinutes - inMinutes;
  if (diffMinutes < 0) diffMinutes += 24 * 60; // Handle overnight shift
  
  return diffMinutes / 60;
}
