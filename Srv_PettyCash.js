/**
 * نظام إدارة العهدة (Petty Cash) - EIPI Factory ERP
 * دالة استقبال وحفظ حركات العهدة
 */
function savePettyCashTransaction(payload) {
  // 1. تفعيل نظام الحماية والتزامن (LockService)
  var lock = LockService.getScriptLock();
  
  try {
    // انتظار القفل لمدة أقصاها 30 ثانية
    lock.waitLock(30000);
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = "PettyCash"; // اسم الشيت في قاعدة البيانات
    var sheet = ss.getSheetByName(sheetName);
    
    // 2. إذا لم يكن الشيت موجوداً، قم بإنشائه وإضافة العناوين (Headers)
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      var headers = ["Timestamp", "التاريخ", "نوع الحركة", "البند", "المبلغ", "البيان/الملاحظات"];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f3f3f3");
      sheet.setFrozenRows(1);
    }
    
    // 3. تجهيز صف البيانات الجديد
    var timestamp = new Date();
    // تحويل المبلغ لرقم عشري للتأكد من صحة العمليات الحسابية لاحقاً
    var amount = parseFloat(payload.amount); 
    
    // يمكننا جعل المصروفات بالسالب والوارد بالموجب لتسهيل الجمع في الداشبورد
    var finalAmount = (payload.type === "صادر") ? -Math.abs(amount) : Math.abs(amount);
    
    var rowData = [
      timestamp,           // Timestamp
      payload.date,        // التاريخ
      payload.type,        // نوع الحركة
      payload.category,    // البند
      finalAmount,         // المبلغ (بالموجب للوارد، بالسالب للصادر)
      payload.notes        // البيان/الملاحظات
    ];
    
    // 4. إدراج البيانات في الشيت
    sheet.appendRow(rowData);
    
    // إرجاع رسالة نجاح للواجهة
    return "تم حفظ حركة العهدة بنجاح (بند: " + payload.category + ").";
    
  } catch (e) {
    // في حالة حدوث خطأ أو انتهاء وقت القفل
    throw new Error("عفواً، حدث خطأ أثناء الحفظ أو النظام مشغول: " + e.message);
    
  } finally {
    // 5. تحرير القفل دائماً في النهاية للسماح للعمليات الأخرى
    lock.releaseLock();
  }
}