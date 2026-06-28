function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('School Management Pro V10')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function sortSheet(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).sort([
      {column: 5, ascending: true},
      {column: 6, ascending: true} 
    ]);
  }
}

// ==========================================
// THE MAGIC ENGINE: Recalculates all sheets perfectly
// ==========================================
function syncAllSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();

  var monthSheets = [];
  for(var s=0; s<sheets.length; s++) {
    var n = sheets[s].getName();
    if(n !== "Main_Sheet" && n !== "Left_Students") monthSheets.push(sheets[s]);
  }

  var studentState = {};

  for(var s=0; s<monthSheets.length; s++) {
    var sh = monthSheets[s];
    var data = sh.getDataRange().getValues();
    var isModified = false;

    for(var r=1; r<data.length; r++) {
      var regNo = data[r][0];
      if(!regNo) continue;

      if(!studentState[regNo]) { studentState[regNo] = { history: [], totalPaid: 0 }; }
      var state = studentState[regNo];

      // Step 1: Calculate pending arrears up to this month
      var cash = state.totalPaid;
      var pending = [];
      for(var i=0; i<state.history.length; i++) {
        if(cash >= state.history[i].amount) {
          cash -= state.history[i].amount;
        } else {
          var due = state.history[i].amount - cash;
          cash = 0;
          if(due > 0) pending.push(state.history[i].month + ":" + due);
        }
      }

      var prevArr = 0;
      for(var p=0; p<pending.length; p++) prevArr += parseFloat(pending[p].split(":")[1]);
      var logStr = pending.join("|");

      // Step 2: Apply to current row
      var curFee = parseFloat(data[r][4]) || 0;
      var totalPay = curFee + prevArr;
      var paid = parseFloat(data[r][7]) || 0;
      var rem = totalPay - paid;
      var status = rem <= 0 ? "Paid" : (paid > 0 ? "Pending" : "Unpaid");

      if(data[r][5] !== prevArr || data[r][6] !== totalPay || data[r][8] !== rem || data[r][9] !== status || data[r][10] !== logStr) {
        data[r][5] = prevArr; data[r][6] = totalPay; data[r][8] = rem; data[r][9] = status; data[r][10] = logStr;
        isModified = true;
      }

      // Step 3: Add this month to history for the NEXT sheet loop
      state.history.push({month: sh.getName(), amount: curFee});
      state.totalPaid += paid;
    }

    if(isModified) sh.getDataRange().setValues(data);
  }
}

// 1. ADD STUDENT
function addStudent(studentData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Main_Sheet");
  if (!sheet) {
    sheet = ss.insertSheet("Main_Sheet");
    sheet.appendRow(["Reg. No", "Name", "Father Name", "Roll No", "Class", "Section", "Admission Date", "Monthly Fee", "Status"]);
  }
  
  var lastRow = sheet.getLastRow();
  var nextRegNo = 1001;
  if (lastRow > 1) {
    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var maxId = 0;
    for (var i = 0; i < ids.length; i++) {
      if (Number(ids[i][0]) > maxId) maxId = Number(ids[i][0]);
    }
    nextRegNo = maxId + 1;
  }
  
  sheet.appendRow([
    nextRegNo, studentData.name, studentData.fatherName, studentData.rollNo, 
    studentData.studentClass, studentData.section.toUpperCase(), studentData.admissionDate, 
    parseFloat(studentData.monthlyFee), "Active"
  ]);
  sortSheet(sheet);
  return {success: true, regNo: nextRegNo};
}

// 2. SEARCH STUDENT
function searchStudent(regNo, studentClass, name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("Main_Sheet");
  var leftSheet = ss.getSheetByName("Left_Students");
  var results = [];

  var sReg = regNo ? String(regNo).trim().toLowerCase() : "";
  var sClass = studentClass ? String(studentClass).trim().toLowerCase() : "";
  var sName = name ? String(name).trim().toLowerCase() : "";

  function scanData(sheet, isLeft) {
    if (!sheet) return;
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue; 
      var cReg = String(data[i][0]).trim().toLowerCase();
      var cName = String(data[i][1]).trim().toLowerCase();
      var cClass = String(data[i][4]).trim().toLowerCase();

      var match = false;
      if (sReg !== "" && cReg === sReg) {
        match = true;
      } else if (sReg === "" && (sClass !== "" || sName !== "")) {
        var classMatch = sClass === "" ? true : cClass.includes(sClass);
        var nameMatch = sName === "" ? true : cName.includes(sName);
        if (classMatch && nameMatch) match = true;
      }

      if (match) {
        results.push({
          regNo: data[i][0], name: data[i][1], fatherName: data[i][2], rollNo: data[i][3],
          studentClass: data[i][4], section: data[i][5], admissionDate: data[i][6],
          monthlyFee: data[i][7], status: data[i][8], isLeftMode: isLeft
        });
      }
    }
  }
  
  scanData(mainSheet, false);
  scanData(leftSheet, true);
  return results;
}

// 3. EDIT STUDENT
function updateStudent(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Main_Sheet");
  var rows = sheet.getDataRange().getValues();
  for(var i=1; i < rows.length; i++) {
    if(String(rows[i][0]).trim() === String(data.regNo).trim()) {
      sheet.getRange(i+1, 2).setValue(data.name);
      sheet.getRange(i+1, 3).setValue(data.fatherName);
      sheet.getRange(i+1, 4).setValue(data.rollNo);
      sheet.getRange(i+1, 5).setValue(data.studentClass);
      sheet.getRange(i+1, 6).setValue(data.section.toUpperCase());
      sheet.getRange(i+1, 8).setValue(parseFloat(data.monthlyFee));
      sheet.getRange(i+1, 9).setValue(data.status);
      
      if(data.status === "Left") {
        var leftSheet = ss.getSheetByName("Left_Students");
        if(!leftSheet) {
          leftSheet = ss.insertSheet("Left_Students");
          leftSheet.appendRow(["Reg. No", "Name", "Father Name", "Roll No", "Class", "Section", "Admission Date", "Monthly Fee", "Status", "Left Date"]);
        }
        var rowData = sheet.getRange(i+1, 1, 1, 9).getValues()[0];
        rowData.push(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"));
        leftSheet.appendRow(rowData);
        sheet.deleteRow(i+1);
        sortSheet(leftSheet);
      } else {
        sortSheet(sheet);
      }
      return {success: true};
    }
  }
  return {success: false, message: "Record not found."};
}

// 4. GENERATE MONTH SHEET & RUN JS CALCULATIONS
function generateMonthSheet(newMonthName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if(ss.getSheetByName(newMonthName)) return {success: false, message: "Sheet already exists!"};

  var mainSheet = ss.getSheetByName("Main_Sheet");
  if(!mainSheet) return {success: false, message: "Main_Sheet missing!"};

  var newSheet = ss.insertSheet(newMonthName);
  newSheet.appendRow(["Reg. No", "Name", "Class", "Section", "Current Fee", "Previous Arrears", "Total Payable", "Amount Paid", "Remaining Balance", "Fee Status", "Pending Log"]);

  var activeStudents = mainSheet.getDataRange().getValues();
  var rows = [];

  for(var i=1; i<activeStudents.length; i++) {
    if(activeStudents[i][8] !== "Active") continue;
    var regNo = activeStudents[i][0];
    var curFee = parseFloat(activeStudents[i][7]) || 0;
    // We add empty arrears initially, then syncAllSheets will fill them with exact data!
    rows.push([regNo, activeStudents[i][1], activeStudents[i][4], activeStudents[i][5], curFee, 0, curFee, 0, curFee, "Unpaid", ""]);
  }
  
  if(rows.length > 0) newSheet.getRange(2, 1, rows.length, 11).setValues(rows);
  
  syncAllSheets(); // RUN JS ON CREATION TO CALCULATE EVERY PENDING FEE!
  return {success: true, message: newMonthName + " Generated & Previous Balances Successfully Carried Forward!"};
}

// 5. GET STUDENT LEDGER
function getStudentLedger(regNo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var ledger = [];
  var info = {name: "N/A", cls: "N/A", roll: "N/A"};
  
  var mainData = ss.getSheetByName("Main_Sheet").getDataRange().getValues();
  for(var i=1; i<mainData.length; i++) {
    if(String(mainData[i][0]).trim() === String(regNo).trim()) {
      info = {name: mainData[i][1], cls: mainData[i][4]+"-"+mainData[i][5], roll: mainData[i][3]}; break;
    }
  }

  for(var s=0; s<sheets.length; s++) {
    var shName = sheets[s].getName();
    if(shName === "Main_Sheet" || shName === "Left_Students") continue;
    var sheetData = sheets[s].getDataRange().getValues();
    for(var r=1; r<sheetData.length; r++) {
      if(String(sheetData[r][0]).trim() === String(regNo).trim()) {
        ledger.push({
          month: shName,
          curFee: parseFloat(sheetData[r][4]) || 0,
          prevArr: parseFloat(sheetData[r][5]) || 0,
          totalPay: parseFloat(sheetData[r][6]) || 0,
          paid: parseFloat(sheetData[r][7]) || 0,
          rem: parseFloat(sheetData[r][8]) || 0,
          status: sheetData[r][9],
          rowIdx: r + 1
        });
        break;
      }
    }
  }
  return {info: info, ledger: ledger};
}

// 6. UPDATE BULK FEES
function updateBulkFees(regNo, updatesArray) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  for(var i=0; i<updatesArray.length; i++) {
    var u = updatesArray[i];
    var sheet = ss.getSheetByName(u.month);
    if(sheet) {
      sheet.getRange(u.rowIdx, 8).setValue(parseFloat(u.paid) || 0); // Updates only paid amount
      sheet.getRange(u.rowIdx, 10).setValue(u.status);
    }
  }
  syncAllSheets(); // RIPPLE EVERYTHING FORWARD
  return {success: true, message: "Ledger updated and next months adjusted dynamically!"};
}

// 7. INVOICE GETTER
function getInvoiceData(regNo, targetMonth) {
  var ledgerData = getStudentLedger(regNo);
  if(!ledgerData || ledgerData.ledger.length === 0) return null;

  var targetRecord = null;
  if(targetMonth && targetMonth.trim() !== "") {
    var sTarget = String(targetMonth).trim().toLowerCase().replace(/[\s_]/g, '');
    for(var i=0; i<ledgerData.ledger.length; i++) {
      var sMonth = String(ledgerData.ledger[i].month).trim().toLowerCase().replace(/[\s_]/g, '');
      if(sMonth === sTarget) { targetRecord = ledgerData.ledger[i]; break; }
    }
  } else {
    targetRecord = ledgerData.ledger[ledgerData.ledger.length - 1];
  }

  if(!targetRecord) return null;

  // Find the log of the TARGET record directly from the sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(targetRecord.month);
  var logData = "";
  if(sheet) {
    var data = sheet.getDataRange().getValues();
    for(var r=1; r<data.length; r++) {
      if(data[r][0] == regNo) { logData = data[r][10]; break; }
    }
  }

  return {
    regNo: regNo,
    name: ledgerData.info.name,
    cls: ledgerData.info.cls,
    roll: ledgerData.info.roll,
    month: targetRecord.month,
    curFee: targetRecord.curFee,
    prevArr: targetRecord.prevArr,
    totalPay: targetRecord.totalPay,
    paid: targetRecord.paid,
    rem: targetRecord.rem,
    status: targetRecord.status,
    log: logData // This will contain 'Jan_2026:500|Feb_2026:1000'
  };
}

// 8. GOOGLE CONSOLE STYLE REPORTS
function getReports(monthName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  
  var aCol=0, aRec=0, aPaid=0, aPend=0, aUnp=0;
  var mName="", mExp=0, mCol=0, mRec=0, mPaid=0, mPend=0, mUnp=0;
  var monthFound = false;
  var latestSheet = null;
  var graphHistory = [];
  
  var sTarget = monthName ? String(monthName).trim().toLowerCase().replace(/[\s_]/g, '') : "";

  for(var s = 0; s < sheets.length; s++) { 
    var n = sheets[s].getName();
    if(n !== "Main_Sheet" && n !== "Left_Students") {
      latestSheet = sheets[s];
      var expectedTotal = 0;
      var collectedTotal = 0;
      var d = sheets[s].getDataRange().getValues();
      for(var r=1; r<d.length; r++) {
         aCol += (parseFloat(d[r][7]) || 0); 
         expectedTotal += (parseFloat(d[r][6]) || 0); 
         collectedTotal += (parseFloat(d[r][7]) || 0);
      }
      graphHistory.push({month: n, expected: expectedTotal, collected: collectedTotal});
      
      var sMonth = String(n).trim().toLowerCase().replace(/[\s_]/g, '');
      if(sTarget !== "" && sMonth === sTarget) {
         monthFound = true;
         mName = n;
         for(var r=1; r<d.length; r++) {
            mExp += (parseFloat(d[r][4]) || 0); 
            mCol += (parseFloat(d[r][7]) || 0); 
            mRec += (parseFloat(d[r][8]) || 0); 
            var st = d[r][9];
            if(st==="Paid") mPaid++; else if(st==="Pending") mPend++; else if(st==="Unpaid") mUnp++;
         }
      }
    }
  }
  
  if(graphHistory.length > 6) graphHistory = graphHistory.slice(graphHistory.length - 6);

  if(latestSheet) {
    var ld = latestSheet.getDataRange().getValues();
    for(var r=1; r<ld.length; r++) {
      aRec += (parseFloat(ld[r][8]) || 0); 
      var st = ld[r][9];
      if(st==="Paid") aPaid++; else if(st==="Pending") aPend++; else if(st==="Unpaid") aUnp++;
    }
  }

  var mData = monthFound ? {name: mName, exp: mExp, col: mCol, rec: mRec, paid: mPaid, pend: mPend, unp: mUnp} : null;
  return { a: {col:aCol, rec:aRec, paid:aPaid, pend:aPend, unp:aUnp}, graph: graphHistory, m: mData };
}
