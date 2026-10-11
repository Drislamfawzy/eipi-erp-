/**
 * نظام إدارة المرتبات (Payroll) - EIPI Factory ERP
 * النسخة النهائية المصححة (تم تعريف وتفعيل مكافأة الإنتاج fixedProductionBonus في بداية الحسابات)
 */

var SPREADSHEET_ID = "14YN47JzXBKLwbhu5h1gDAGYIjbme7ft16Y-m7DiqPiI";

function processPayroll(payload) {
  var lock = LockService.getScriptLock();
  
  try {
    lock.waitLock(30000);
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheetName = "Payroll";
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      var headers = [
        "Timestamp", "التاريخ", "الإسم", "عدد ايام الحضور", "خصومات وجزاءات", 
        "عدد غ بدون اذن", "عدد ساعات التاخير", "عدد ساعات اوفر تايم", "السلف", 
        "عدد ساعات العمل اليومية", "يستحق مكافاه انتاج", "حوافز", "أجر اليوم", 
        "أجر الساعة", "خصم التأخير", "خصم جزاءات", "خصم غياب بدون إذن", 
        "أجر الأوفر تايم", "إجمالي الخصم", "إجمالي الإضافات", "عمولة مدير الإنتاج", 
        "صافى المستحق", "اجمالى المرتب"
      ];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f3f3f3");
      sheet.setFrozenRows(1);
    }
    
    var date = payload.date;
    var name = payload.employeeName;
    var baseHours = parseFloat(payload.baseHours) || 0;
    var attendanceDays = parseFloat(payload.attendanceDays) || 0;
    var unexcusedAbsence = parseFloat(payload.unexcusedAbsence) || 0;
    var delayHours = parseFloat(payload.delayHours) || 0;
    var overtimeHours = parseFloat(payload.overtimeHours) || 0;
    var deductionsInput = parseFloat(payload.deductions) || 0;
    
    // --- جلب السلف أوتوماتيكياً من شيت Advances ---
    var advances = 0;
    try {
      advances = getEmployeeAdvancesForMonth(name, date);
      Logger.log("إجمالي السلف المحسوبة للعامل (" + name + ") عن تاريخ (" + date + ") = " + advances);
    } catch (err) {
      Logger.log("خطأ أثناء استدعاء السلف: " + err.message);
      advances = parseFloat(payload.advances) || 0;
    }
    
    var incentives = parseFloat(payload.incentives) || 0;
    var productionBonusStr = payload.productionBonus || "لا"; // "نعم" أو "لا"
    
    // --- 1. حساب أجر اليوم وأجر الساعة ---
    var dailyWage = 0;
    if (baseHours === 9) dailyWage = 161;
    else if (baseHours === 12) dailyWage = 207;
    else if (baseHours === 24) dailyWage = 550;
    
    var hourlyWage = baseHours > 0 ? (dailyWage / baseHours) : 0;
    
    // --- 2. حساب الخصومات ---
    var delayDeduction = delayHours * 2 * hourlyWage; // خصم التأخير
    var unexcusedAbsenceDeduction = unexcusedAbsence * 2 * dailyWage; // غياب بدون إذن يخصم يومين
    var totalDeductions = delayDeduction + deductionsInput + unexcusedAbsenceDeduction + advances;
    
    // --- 3. حساب الإضافات ومكافأة الإنتاج (تعريف صريح للمتغير) ---
    var overtimePay = overtimeHours * 2 * hourlyWage; 
    
    // التعريف الصريح لمتغير مكافأة الإنتاج الثابتة (500 جنيه عند نعم، و0 عند لا)
    var fixedProductionBonus = (productionBonusStr === "نعم") ? 500 : 0;
    
    // إجمالي الإضافات يضم الأوفر تايم ومكافأة الإنتاج
    var totalAdditions = overtimePay + fixedProductionBonus; 
    
    // --- 4. أيام الراحة المستحقة ---
    var restDays = 0;
    if (dailyWage !== 550) {
      if (attendanceDays >= 22 && attendanceDays <= 26) {
        restDays = 4;
      } else if (attendanceDays >= 15 && attendanceDays < 22) {
        restDays = 3;
      } else if (attendanceDays > 26) {
        restDays = 4;
      }
    }
    var baseSalary = dailyWage * (attendanceDays + restDays);
    
    // --- 5. عمولة مدير الإنتاج (لـ 24 ساعة) ---
    var managerCommission = (dailyWage === 550) ? 6666.5 : 0;
    
    // --- 6. حساب صافي المستحق وإجمالي المرتب ---
    var netSalary = baseSalary - totalDeductions + totalAdditions + incentives + managerCommission;
    var grossSalary = netSalary + totalDeductions;
    
    // --- 7. حفظ البيانات في الشيت ---
    var timestamp = new Date();
    var rowData = [
      timestamp, date, name, attendanceDays, deductionsInput, 
      unexcusedAbsence, delayHours, overtimeHours, advances, 
      baseHours, productionBonusStr, incentives, dailyWage.toFixed(2), 
      hourlyWage.toFixed(2), delayDeduction.toFixed(2), deductionsInput.toFixed(2), unexcusedAbsenceDeduction.toFixed(2), 
      overtimePay.toFixed(2), totalDeductions.toFixed(2), totalAdditions.toFixed(2), managerCommission.toFixed(2),
      netSalary.toFixed(2), grossSalary.toFixed(2)
    ];
    
    sheet.appendRow(rowData);
    
    return "تم تسجيل المرتب بنجاح! المكافأة المضافة: " + fixedProductionBonus + " ج.م | السلف المخصومة: " + advances + " ج.م | الصافي: " + netSalary.toFixed(2) + " ج.م";
    
  } catch (e) {
    throw new Error("حدث خطأ أثناء حفظ المرتب: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * دالة جلب سلف الشهر من شيت Advances المربوط بالملف الأصلي
 */
function getEmployeeAdvancesForMonth(employeeName, payrollDateStr) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("Advances");
    if (!sheet) return 0;
    
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return 0;
    
    var payrollDate = payrollDateStr ? new Date(payrollDateStr) : new Date();
    var targetYear = payrollDate.getFullYear();
    var targetMonth = payrollDate.getMonth(); 
    
    var totalAdvances = 0;
    var targetName = employeeName ? employeeName.toString().trim() : "";
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rowDateVal = row[1] || row[0]; 
      var rowName = row[2] ? row[2].toString().trim() : "";
      var rowType = row[3] ? row[3].toString().trim() : "";
      var rowAmount = parseFloat(row[4]) || 0;
      
      if (rowName === targetName && (rowType === "سلفة" || rowType.indexOf("سلف") !== -1)) {
        var rDate = new Date(rowDateVal);
        if (!isNaN(rDate.getTime())) {
          if (rDate.getFullYear() === targetYear && rDate.getMonth() === targetMonth) {
            totalAdvances += rowAmount;
          }
        }
      }
    }
    
    return totalAdvances;
  } catch (e) {
    Logger.log("خطأ داخل getEmployeeAdvancesForMonth: " + e.message);
    return 0;
  }
}

/**
 * دالة جلب أسماء العمال
 */
function getEmployeesList() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("Employees");
    
    if (!sheet) {
      sheet = ss.insertSheet("Employees");
      sheet.appendRow(["اسم العامل"]);
      sheet.appendRow(["محمد على"]);
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return ["محمد على"];
    
    var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var employees = [];
    for (var i = 0; i < data.length; i++) {
      if (data[i][0]) {
        employees.push(data[i][0].toString().trim());
      }
    }
    return employees;
  } catch (e) {
    return ["خطأ في جلب الأسماء"];
  }
}