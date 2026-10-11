/**
 * نظام إدارة السلف والمعاملات المالية - النسخة التشخيصية المؤكدة
 */
var SPREADSHEET_ID = "14YN47JzXBKLwbhu5h1gDAGYIjbme7ft16Y-m7DiqPiI";

function processAdvanceTransaction(payload) {
  var lock = LockService.getScriptLock();
  
  try {
    lock.waitLock(30000);
    
    // فتح ملف الإكسل الأصلي حصرياً برقم المعرف
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // البحث عن شيت Advances (إذا لم يكن موجوداً يتم إنشاؤه فوراً)
    var sheetName = "Advances"; 
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      var headers = ["Timestamp", "التاريخ", "اسم الموظف", "نوع المعاملة", "المبلغ", "ملاحظات"];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f3f3f3");
      sheet.setFrozenRows(1);
    }
    
    var timestamp = new Date();
    var date = payload.date;
    var employeeName = payload.employeeName;
    var transactionType = payload.transactionType; // "سلفة" أو "مرتب"
    var amount = parseFloat(payload.amount) || 0;
    var notes = payload.notes || "";
    
    var rowData = [
      timestamp,
      date,
      employeeName,
      transactionType,
      amount,
      notes
    ];
    
    // إدراج الصف في الملف الأصلي
    sheet.appendRow(rowData);
    
    // طباعة تتبع في السجلات للتأكيد
    Logger.log("تمت الكتابة بنجاح في ملف: " + ss.getName() + " الشيت: " + sheet.getName() + " المبلغ: " + amount);
    
    return "تم حفظ السلفة وتوثيقها في شيت Advances بملف المصنع الأساسي بنجاح!";
    
  } catch (e) {
    throw new Error("فشل الحفظ بسبب الصلاحيات أو المعرف: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

function getEmployeeAdvancesForMonth(employeeName, targetDateStr) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName("Advances");
    if (!sheet) return 0;
    
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return 0;
    
    var targetDate = targetDateStr ? new Date(targetDateStr) : new Date();
    var targetYear = targetDate.getFullYear();
    var targetMonth = targetDate.getMonth(); 
    
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
    console.error("خطأ في قراءة السلف: " + e.message);
    return 0;
  }
}