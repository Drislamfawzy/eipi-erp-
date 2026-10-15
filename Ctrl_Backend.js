/**
 * نظام إدارة موارد مصنع EIPI Factory ERP - خادم المعالجة (Backend)
 * الإصدار المحدث - النسخة V3 مع حساب OEE الديناميكي ووقت التوقف
 */

// 1. دالة حفظ الإنتاج والشحن وقراءات العدادات وقيمة الشحن ووقت التوقف
function saveProductionData(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Production") || ss.insertSheet("Production");
    
    // ضبط عناوين الأعمدة بترتيب دقيق وصحيح لا يزيح الملاحظات
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "وقت التسجيل", 
        "تاريخ الوردية", 
        "نوع الوردية", 
        "الماكينة / الخط", 
        "الصنف المُنتج", 
        "الكمية المصنعة", 
        "الهالك (عدد)", 
        "وقت التوقف (دقيقة)", 
        "ملاحظات التشغيل", 
        "قراءة كهرباء 1", 
        "قراءة كهرباء 2", 
        "شحن كهرباء", 
        "قيمة الشحن (ج.م)", 
        "رابط إيصال الشحن"
      ]);
    }
    
    var timestamp = new Date();
    var shiftDate = payload.shiftDate;
    var shiftType = payload.shiftType;
    var notes = payload.notes || "";
    var elec1 = payload.elec1 || 0;
    var elec2 = payload.elec2 || 0;
    var hasRecharged = payload.hasRecharged || "لا";
    var rechargeAmount = payload.rechargeAmount || 0;
    var receiptUrl = payload.receiptUrl || "لا توجد صورة";
    
    if (payload.machines && Array.isArray(payload.machines)) {
      payload.machines.forEach(function(m) {
        sheet.appendRow([
          timestamp, 
          shiftDate, 
          shiftType, 
          m.machineName, 
          m.productName, 
          m.quantity, 
          m.waste, 
          m.downtimeMins || 0, // وقت التوقف في عمود H (8) منفصل تماماً
          notes,               // ملاحظات التشغيل في عمود I (9) لوحدها
          elec1, 
          elec2, 
          hasRecharged, 
          rechargeAmount, 
          receiptUrl
        ]);
      });
    }
    
    return { status: "success", message: "تم حفظ وتسيير بيانات الإنتاج والشحن بنجاح" };
  } catch (error) {
    throw new Error("خطأ أثناء حفظ الإنتاج: " + error.message);
  } finally {
    lock.releaseLock();
  }
}

// 2. استقبال وحفظ أذونات خروج البضاعة
function saveDispatchData(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Dispatch") || ss.insertSheet("Dispatch");
    
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["وقت التسجيل اللحظي", "تاريخ الإذن", "اسم العميل", "رقم الإذن", "الصنف المنصرف", "الكمية المصروفة", "المبلغ النقدي (ج.م)"]);
    }
    
    var timestamp = new Date();
    var dispatchDate = payload.date || new Date().toISOString().split('T')[0];
    var clientName = payload.clientName;
    var dispatchNo = payload.dispatchNo || "";
    var cashAmount = payload.cashAmount || 0;
    
    if (payload.items && Array.isArray(payload.items)) {
      payload.items.forEach(function(item) {
        sheet.appendRow([timestamp, dispatchDate, clientName, dispatchNo, item.itemName, item.quantity, cashAmount]);
      });
    }
    
    return { status: "success", message: "تم إصدار وتسجيل إذن خروج البضاعة بنجاح" };
  } catch (error) {
    throw new Error("خطأ أثناء إصدار إذن الخروج: " + error.message);
  } finally {
    lock.releaseLock();
  }
}

// 3. استقبال وحفظ أذونات استلام الخامات
function saveInboundData(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Inbound") || ss.insertSheet("Inbound");

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["وقت الاستلام", "اسم المورد", "الخامة / الصنف", "الكمية المستلمة", "صورة الإذن"]);
    }

    var timestamp = new Date();
    var supplierName = payload.supplierName;
    var receiptUrl = payload.receiptUrl || "لا توجد صورة";

    if (payload.items && Array.isArray(payload.items)) {
      payload.items.forEach(function (item) {
        sheet.appendRow([timestamp, supplierName, item.materialName, item.quantity, receiptUrl]);
      });
    }

    return { status: "success", message: "تم تسجيل استلام الخامات وإضافتها للمخزن بنجاح" };
  } catch (error) {
    throw new Error("خطأ أثناء حفظ الاستلام: " + error.message);
  } finally {
    lock.releaseLock();
  }
}

// 4. دالة جلب قاعدة البيانات المركزية (Master Data)
function getMasterData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = "قاعدة البيانات";
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(["العمال", "العملاء", "الموردين", "الخامات", "المنتجات", "الماكينات"]);
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { workers: [], clients: [], suppliers: [], materials: [], products: [], machines: [] };
    }

    var data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    var masterData = { workers: [], clients: [], suppliers: [], materials: [], products: [], machines: [] };

    for (var i = 0; i < data.length; i++) {
      if (data[i][0]) masterData.workers.push(data[i][0].toString().trim());   
      if (data[i][1]) masterData.clients.push(data[i][1].toString().trim());   
      if (data[i][2]) masterData.suppliers.push(data[i][2].toString().trim()); 
      if (data[i][3]) masterData.materials.push(data[i][3].toString().trim()); 
      if (data[i][4]) masterData.products.push(data[i][4].toString().trim());  
      if (data[i][5]) masterData.machines.push(data[i][5].toString().trim());  
    }

    return masterData;
  } catch (error) {
    throw new Error("خطأ في جلب قاعدة البيانات المركزية: " + error.message);
  }
}

// 5. رفع صورة الإذن (استلام)
function uploadReceiptImage(fileData, fileName) {
  try {
    var base64Data = fileData.split(',')[1];
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), "image/jpeg", fileName || "receipt.jpg");

    var mainFolderName = "Inbound_Receipts";
    var folders = DriveApp.getFoldersByName(mainFolderName);
    var mainFolder = folders.hasNext() ? folders.next() : DriveApp.createFolder(mainFolderName);

    var date = new Date();
    var subFolderName = (date.getMonth() + 1) + "-" + date.getFullYear();
    var subFolders = mainFolder.getFoldersByName(subFolderName);
    var subFolder = subFolders.hasNext() ? subFolders.next() : mainFolder.createFolder(subFolderName);

    var file = subFolder.createFile(blob);
    return file.getUrl();

  } catch (error) {
    throw new Error("خطأ أثناء رفع الصورة: " + error.message);
  }
}

// 6. رفع إيصال الكهرباء في المجلد المخصص
function uploadElecReceiptImage(fileData, fileName) {
  try {
    var base64Data = fileData.split(',')[1];
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), "image/jpeg", fileName || "elec_receipt.jpg");

    var mainFolderName = "elec_reciet";
    var folders = DriveApp.getFoldersByName(mainFolderName);
    var mainFolder = folders.hasNext() ? folders.next() : DriveApp.createFolder(mainFolderName);

    var date = new Date();
    var subFolderName = (date.getMonth() + 1) + "-" + date.getFullYear();
    var subFolders = mainFolder.getFoldersByName(subFolderName);
    var subFolder = subFolders.hasNext() ? subFolders.next() : mainFolder.createFolder(subFolderName);

    var file = subFolder.createFile(blob);
    return file.getUrl();

  } catch (error) {
    throw new Error("خطأ أثناء رفع إيصال الكهرباء: " + error.message);
  }
}

// 7. حفظ أو تحديث الجرد الفعلي
function saveInventoryData(payload) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(30000);
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        
        var invSheet = ss.getSheetByName("Inventory");
        if (!invSheet) {
            invSheet = ss.insertSheet("Inventory");
            invSheet.appendRow(["الصنف", "الرصيد الفعلي", "تاريخ آخر جرد"]);
            invSheet.getRange(1, 1, 1, 3).setFontWeight("bold");
        }

        var inventoryData = invSheet.getDataRange().getValues();
        var itemRowMap = {}; 

        for (var i = 1; i < inventoryData.length; i++) {
            var itemName = String(inventoryData[i][0]).trim();
            if (itemName) {
                itemRowMap[itemName] = i + 1; 
            }
        }

        payload.items.forEach(function(item) {
            var name = item.itemName.trim();
            var qty = item.quantity;
            var date = payload.date;

            if (itemRowMap[name]) {
                invSheet.getRange(itemRowMap[name], 2).setValue(qty);
                invSheet.getRange(itemRowMap[name], 3).setValue(date);
            } else {
                invSheet.appendRow([name, qty, date]);
                itemRowMap[name] = invSheet.getLastRow();
            }
        });

        var logSheet = ss.getSheetByName("Inventory_Log");
        var columnsOrder = [
            "تاريخ الجرد", "بريفورم 29", "بريفورم 24", "بريفورم 20", "بريفورم مصنعيه",
            "عبوات 500", "عبوات 750", "عبوات 850 مشرشر", "عبوات 900", "عبوات لتر 29", "عبوات لتر 20 جرام", "عبوات مصنعيه",
            "غطاء أزرق مارينا", "غطاء أزرق شكاير", "غطاء لبنى", "غطاء اصفر", "غطاء احمر", "غطاء اخضر",
            "أكياس", "أكياس خارجي", "كراتين سكوتش"
        ];

        if (!logSheet) {
            logSheet = ss.insertSheet("Inventory_Log");
            logSheet.appendRow(columnsOrder);
            logSheet.getRange(1, 1, 1, columnsOrder.length).setFontWeight("bold").setBackground("#f3f3f3");
            logSheet.setFrozenRows(1);
        }

        var logRow = new Array(columnsOrder.length).fill("");
        logRow[0] = payload.date;

        payload.items.forEach(function(item) {
            var name = item.itemName.trim();
            var colIndex = columnsOrder.indexOf(name);
            if (colIndex !== -1) {
                logRow[colIndex] = item.quantity;
            }
        });

        logSheet.appendRow(logRow);
        return { status: "success", message: "تم تحديث المخزون وحفظ سجل الجرد بنجاح" };
        
    } catch (e) {
        throw new Error("خطأ في تحديث الجرد: " + e.message);
    } finally {
        lock.releaseLock();
    }
}

// 8. جلب مؤشرات لوحة القيادة (Dashboard V2) وحساب OEE الديناميكي
function getDashboardMetrics_V3(filterType) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var filterDate = new Date();
    var today = new Date();
    
    if (filterType === 'يوم') { filterDate.setDate(today.getDate() - 1); } 
    else if (filterType === 'أسبوع') { filterDate.setDate(today.getDate() - 7); } 
    else if (filterType === 'شهر') { filterDate.setMonth(today.getMonth() - 1); } 
    else if (filterType === 'ربع سنوي') { filterDate.setMonth(today.getMonth() - 3); } 
    else if (filterType === 'سنوي') { filterDate.setFullYear(today.getFullYear() - 1); } 
    else { filterDate.setFullYear(2000); }

    var totalProduction = 0, totalWaste = 0, totalDispatchAbwa = 0, pettyCashBalance = 0;
    var maxElec1 = 0, maxElec2 = 0, totalShiftsCount = 0;
    var chartLabels = [], chartProdData = [], chartWasteData = [];
    
    // متغيرات حساب الـ OEE الديناميكي
    var totalAvailableMins = 0;
    var totalOperatingMins = 0;
    var totalActualQtyForOEE = 0;
    var totalTargetQtyForOEE = 0;
    var totalGoodQtyForOEE = 0;

    // 1. قراءة شيت الإنتاج
    var prodSheet = ss.getSheetByName("Production");
    if (prodSheet) {
      var prodData = prodSheet.getDataRange().getValues();
      if (prodData.length > 1) {
        for (var i = prodData.length - 1; i > 0; i--) {
          if (!prodData[i][0] && !prodData[i][1]) continue; 
          var e1 = Number(prodData[i][9]) || 0; // تم ضبط الفهرس ليطابق أعمدة الحفظ الجديدة
          var e2 = Number(prodData[i][10]) || 0;
          if (e1 > 0 || e2 > 0) {
            if(maxElec1 === 0) maxElec1 = e1;
            if(maxElec2 === 0) maxElec2 = e2;
            if(maxElec1 > 0 && maxElec2 > 0) break;
          }
        }

        for (var i = 1; i < prodData.length; i++) {
          var row = prodData[i];
          if (!row[0] && !row[1]) continue; 
          
          var rowDate = new Date(row[1]);
          if (isNaN(rowDate.getTime())) rowDate = new Date(2000, 0, 1);
          
          if (rowDate >= filterDate) {
            totalShiftsCount++;
            var machineName = String(row[3] || "").trim();
            var qty = Number(row[5]) || 0;
            var waste = Number(row[6]) || 0;
            var downtimeMins = Number(row[7]) || 0;

            totalProduction += qty;
            totalWaste += waste;

            // حساب عناصر الـ OEE لكل سطر وردية/ماكينة
            var shiftAvailableMins = 660; // 11 ساعة × 60 دقيقة
            var shiftOperatingMins = Math.max(0, shiftAvailableMins - downtimeMins);
            
            totalAvailableMins += shiftAvailableMins;
            totalOperatingMins += shiftOperatingMins;
            totalActualQtyForOEE += qty;
            
            var goodQty = Math.max(0, qty - waste);
            totalGoodQtyForOEE += goodQty;

            // تحديد المستهدف حسب نوع الماكينة (أوتوماتيك = 30000، العادية = 8000)
            var targetQty = (machineName.includes("اوتوماتيك") || machineName.includes("أوتوماتيك")) ? 35000 : 11000;
            totalTargetQtyForOEE += targetQty;

            var dateStr = rowDate.toISOString().split('T')[0];
            var dateIndex = chartLabels.indexOf(dateStr);
            if (dateIndex === -1) {
              chartLabels.push(dateStr);
              chartProdData.push(qty);
              chartWasteData.push(waste);
            } else {
              chartProdData[dateIndex] += qty;
              chartWasteData[dateIndex] += waste;
            }
          }
        }
      }
    }

    // حساب نسبة الـ OEE الحقيقية (الإتاحة × الأداء × الجودة)
    var finalOee = 100.00;
    if (totalAvailableMins > 0 && totalTargetQtyForOEE > 0 && totalProduction > 0) {
      var availability = totalOperatingMins / totalAvailableMins;
      var performance = totalActualQtyForOEE / totalTargetQtyForOEE;
      var quality = totalGoodQtyForOEE / totalProduction;
      
      finalOee = (availability * performance * quality) * 100;
      if (finalOee > 100) finalOee = 100; // الحد الأقصى 100%
      if (finalOee < 0) finalOee = 0;
    } else {
      finalOee = 0.00;
    }

    // 2. قراءة شيت الصرف
    var dispatchSheet = ss.getSheetByName("Dispatch");
    if (dispatchSheet) {
      var dispData = dispatchSheet.getDataRange().getValues();
      for (var i = 1; i < dispData.length; i++) {
        var row = dispData[i];
        if (!row[0] && !row[1]) continue; 

        var rowDate = new Date(row[1]);
        if (isNaN(rowDate.getTime())) rowDate = new Date(2000, 0, 1);

        if (rowDate >= filterDate) {
          var itemName = String(row[4] || "").trim();
          var qty = Number(row[5]) || 0;
          
          if (itemName.includes("عبو")) {
            totalDispatchAbwa += qty;
          }
        }
      }
    }

    // 3. قراءة شيت العهدة (PettyCash)
    var pettySheet = ss.getSheetByName("PettyCash");
    if (pettySheet) {
      var pettyData = pettySheet.getDataRange().getValues();
      for (var i = 1; i < pettyData.length; i++) {
        var row = pettyData[i];
        if (!row[0] && !row[1]) continue;

        var rowDate = new Date(row[1]);
        if (isNaN(rowDate.getTime())) rowDate = new Date(row[0]);
        if (isNaN(rowDate.getTime())) rowDate = new Date(2000, 0, 1);

        if (rowDate >= filterDate) {
          var transType = String(row[2] || "").trim();
          var amount = Math.abs(Number(row[4]) || 0);
          
          if (transType.includes("وارد") || transType.includes("إضافة")) {
              pettyCashBalance += amount;
          } else if (transType.includes("صادر") || transType.includes("خصم") || transType.includes("منصرف")) {
              pettyCashBalance -= amount;
          }
        }
      }
    }

    // 4. قراءة الأرصدة
    var invSheet = ss.getSheetByName("Inventory");
    var actualInventory = {};
    if (invSheet) {
      var invData = invSheet.getDataRange().getValues();
      for (var i = 1; i < invData.length; i++) {
        var row = invData[i];
        var name = String(row[0] || "").trim();
        var qty = Number(row[1]) || 0;
        var minLimit = Number(row[3]) || 500;  
        var safeLimit = Number(row[4]) || 2000; 

        if (name) {
            actualInventory[name] = { qty: qty, min: minLimit, safe: safeLimit };
        }
      }
    }

    return {
      status: "success",
      metrics: {
        totalProduction: totalProduction,
        totalWaste: totalWaste,
        totalDowntime: 0,
        totalDispatchAbwa: totalDispatchAbwa,
        oee: finalOee.toFixed(2), // إرجاع القيمة المحسوبة ديناميكياً
        targetProgress: 0,
        compressorRemaining: totalShiftsCount > 0 ? (1000 - (totalShiftsCount * 11)) : 1000,
        managerCommission: (totalDispatchAbwa * 0.011).toFixed(2),
        pettyCashBalance: pettyCashBalance.toFixed(2),
        electricityPlant1: maxElec1,
        electricityPlant2: maxElec2,
        chartLabels: chartLabels,
        chartProdData: chartProdData,
        chartWasteData: chartWasteData,
        actualInventory: actualInventory
      }
    };
  } catch (error) {
    return { status: "error", message: error.message };
  }
}

// 9. دالة اختبار الداشبورد 
function testDashboard_V3_Direct() {
  Logger.log("جاري سحب البيانات من الدالة المنظفة...");
  var res = getDashboardMetrics_V3('شهر');
  
  if (res.status === "error") {
    Logger.log("الخطأ هو: " + res.message);
  } else {
    Logger.log("تمت العملية بنجاح! الأرقام:");
    Logger.log("- الإنتاج: " + res.metrics.totalProduction);
    Logger.log("- OEE %: " + res.metrics.oee);
    Logger.log("- المبيعات (عبوات): " + res.metrics.totalDispatchAbwa);
  }
}

// 10. إرسال تقرير الداشبورد كملف PDF
function sendDashboardEmail() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var url = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/export?format=pdf&portrait=true&size=A4";
    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token } });
    var pdfBlob = response.getBlob().setName("Dashboard_Report.pdf");
    var emailBody = "السادة الإدارة،\n\nمرفق لسيادتكم تقرير لوحة القيادة التنفيذية الأحدث لمصنع EIPI.\n\nمرسل آلياً من نظام الـ ERP.";
    
    MailApp.sendEmail({
      to: "egyptianplast@gmail.com",
      subject: "تقرير لوحة القيادة التنفيذية (Dashboard) - EIPI Factory",
      body: emailBody,
      attachments: [pdfBlob]
    });
    
    return "تم إرسال التقرير بنجاح إلى الإدارة.";
  } catch(e) {
    throw new Error("خطأ في إرسال التقرير: " + e.message);
  }
}

// دوال نظام التخطيط والتحميل
function extractAddedValueFromFormula(formula) {
  if (!formula || typeof formula !== 'string' || !formula.startsWith('=')) {
    return 0;
  }
  var match = formula.match(/\+\s*(\d+)/);
  if (match) {
    return parseInt(match[1], 10);
  }
  return 0;
}

function getCardData() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("التخطيط");
    if (!sheet) throw new Error('شيت "التخطيط" غير موجود');
    
    var data = sheet.getDataRange().getValues();
    var formulasNahar = sheet.getRange(1, 3, sheet.getLastRow(), 1).getFormulas();
    var formulasLeil = sheet.getRange(1, 4, sheet.getLastRow(), 1).getFormulas();
    
    var result = [];
    var extractedNaharVal = 0;
    var extractedLeilVal = 0;
    
    var now = new Date();
    var todayDateOnly = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    
    var dayOfWeek = now.getDay(); 
    var daysToSubtract = (dayOfWeek + 1) % 7; 
    var lastSaturday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - daysToSubtract);
    lastSaturday.setHours(0, 0, 0, 0);

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rawDate = row[0];
      
      if (rawDate instanceof Date) {
        var rowDate = new Date(rawDate.getFullYear(), rawDate.getMonth(), rawDate.getDate());
        
        if (rowDate.getTime() === todayDateOnly.getTime()) {
          var formulaNaharStr = formulasNahar[i] ? formulasNahar[i][0] : "";
          var formulaLeilStr = formulasLeil[i] ? formulasLeil[i][0] : "";
          
          extractedNaharVal = extractAddedValueFromFormula(formulaNaharStr);
          extractedLeilVal = extractAddedValueFromFormula(formulaLeilStr);
        }

        if (rowDate >= lastSaturday) {
          var tahmeelVal = parseFloat(row[4]) || 0;
          
          if (tahmeelVal > 0) { 
            var naharVal = parseFloat(row[2]) || 0;
            var leilVal = parseFloat(row[3]) || 0;

            result.push({
              date: formatDate(rawDate),
              day: row[1] ? row[1].toString() : '---',
              nahar: naharVal,
              leil: leilVal,
              tahmeel: tahmeelVal,
              client: row[5] ? row[5].toString() : ' ',
              ezn: row[6] !== "" ? row[6] : '---',
              ghtaa: row[7] ? row[7].toString() : ' '
            });
          }
        }
      }
    }

    var todayFormatted = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy/MM/dd");

    return { 
      data: result, 
      extractedNahar: extractedNaharVal, 
      extractedLeil: extractedLeilVal, 
      activeDays: result.length,
      currentDate: todayFormatted
    };
  } catch(e) {
    return { data: [], extractedNahar: 0, extractedLeil: 0, activeDays: 0, currentDate: '---' };
  }
}

function formatDate(dateVal) {
  if (!dateVal) return '---';
  if (dateVal instanceof Date) {
    return Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy/MM/dd");
  }
  return dateVal.toString();
}
// جلب رصيد بريفورم 29 وعبوات لتر 29 من شيت Inventory للكارت التفاعلي في الرئيسية
function getFactoryStock() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Inventory");
    var stock = { preform: 0, bottle: 0 };

    if (!sheet || sheet.getLastRow() < 2) {
      return stock;
    }

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    for (var i = 0; i < data.length; i++) {
      var itemName = String(data[i][0] || "").trim();
      var qty = Number(data[i][1]) || 0;

      if (itemName === "بريفورم 29") {
        stock.preform = qty;
      } else if (itemName === "عبوات لتر 29") {
        stock.bottle = qty;
      }
    }

    return stock;
  } catch (e) {
    return { preform: 0, bottle: 0 };
  }
}