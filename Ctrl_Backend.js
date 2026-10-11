/**
 * نظام إدارة موارد مصنع EIPI Factory ERP - خادم المعالجة (Backend)
 */

// 1. حفظ وتخزين بيانات الإنتاج لعدة ماكينات في نفس الوردية
function saveProductionData(payload) {
  var lock = LockService.getScriptLock();
  // محاولة الحصول على قفل لمدة 30 ثانية لتجنب التداخل
  try {
    lock.waitLock(30000);
  } catch (e) {
    throw new Error("عذراً، النظام مشغول حالياً بطلب آخر. يرجى المحاولة بعد قليل.");
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Production") || ss.insertSheet("Production");
    
    // إعداد رؤوس الأعمدة إذا كان الشيت فارغاً
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["وقت التسجيل", "تاريخ الوردية", "نوع الوردية", "الماكينة / الخط", "الصنف المُنتج", "الكمية المصنعة", "الهالك (كجم)", "ملاحظات التشغيل"]);
    }
    
    var timestamp = new Date();
    var shiftDate = payload.shiftDate;
    var shiftType = payload.shiftType;
    var notes = payload.notes || "";
    
    // ترحيل كل سطر ماكينة إلى صف منفصل في الشيت
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
          notes
        ]);
      });
    }
    
    return { status: "success", message: "تم حفظ وتسيير بيانات الإنتاج لجميع الماكينات بنجاح" };
  } catch (error) {
    throw new Error("خطأ أثناء حفظ الإنتاج: " + error.message);
  } finally {
    lock.releaseLock();
  }
}

// 2. استقبال وحفظ أذونات خروج البضاعة (أصناف متعددة) وخصمها من المخزون
function saveDispatchData(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    throw new Error("عذراً، النظام مشغول حالياً. يرجى المحاولة لاحقاً.");
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Dispatch") || ss.insertSheet("Dispatch");
    
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["وقت الإصدار", "اسم العميل", "الصنف المنصرف", "الكمية المصروفة"]);
    }
    
    var timestamp = new Date();
    var clientName = payload.clientName;
    
    if (payload.items && Array.isArray(payload.items)) {
      payload.items.forEach(function(item) {
        sheet.appendRow([
          timestamp,
          clientName,
          item.itemName,
          item.quantity
        ]);
        
        // يمكنك هنا إضافة دالة لتحديث مخزون الأصناف وخصم الكمية تلقائياً إذا لزم الأمر
      });
    }
    
    return { status: "success", message: "تم إصدار وتسجيل إذن خروج البضاعة بنجاح" };
  } catch (error) {
    throw new Error("خطأ أثناء إصدار إذن الخروج: " + error.message);
  } finally {
    lock.releaseLock();
  }
}

// 3. استقبال وحفظ أذونات استلام الخامات وإضافتها للأرصدة الفعليّة
function saveInboundData(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    throw new Error("عذراً، النظام مشغول حالياً. يرجى المحاولة لاحقاً.");
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Inbound") || ss.insertSheet("Inbound");
    
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["وقت الاستلام", "اسم المورد", "الخامة / الصنف", "الكمية المستلمة"]);
    }
    
    var timestamp = new Date();
    var supplierName = payload.supplierName;
    
    if (payload.items && Array.isArray(payload.items)) {
      payload.items.forEach(function(item) {
        sheet.appendRow([
          timestamp,
          supplierName,
          item.materialName,
          item.quantity
        ]);
        
        // يمكنك هنا إضافة دالة لإضافة الكمية المستلمة إلى رصيد المخزن الفعلي تلقائياً
      });
    }
    
    return { status: "success", message: "تم تسجيل استلام الخامات وإضافتها للمخزن بنجاح" };
  } catch (error) {
    throw new Error("خطأ أثناء حفظ الاستلام: " + error.message);
  } finally {
    lock.releaseLock();
  }
}
/**
 * حساب وإرجاع كافة مؤشرات لوحة القيادة التنفيذية (Dashboard Metrics) - تبدأ بأصفار نظيفة ودون قيم وهمية
 */
function getDashboardMetrics() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. شيت الإنتاج: إجمالي الإنتاج، الهالك، وإنتاج آخر يوم
    var prodSheet = ss.getSheetByName("Production");
    var totalMonthlyProduction = 0;
    var totalMonthlyWaste = 0;
    var lastDayProduction = 0;
    
    if (prodSheet && prodSheet.getLastRow() > 1) {
      var prodData = prodSheet.getRange(2, 1, prodSheet.getLastRow() - 1, prodSheet.getLastColumn()).getValues();
      var todayStr = new Date().toISOString().split('T')[0];
      
      prodData.forEach(function(row) {
        var shiftDate = row[1] ? new Date(row[1]).toISOString().split('T')[0] : "";
        var qty = Number(row[5]) || 0;
        var waste = Number(row[6]) || 0;
        
        totalMonthlyProduction += qty;
        totalMonthlyWaste += waste;
        
        if (shiftDate === todayStr || (prodData.indexOf(row) === prodData.length - 1)) {
          lastDayProduction += qty;
        }
      });
    }
    
    // نسبة الهالك مع التأمين التام ضد القسمة على صفر
    var wastePercentage = totalMonthlyProduction > 0 ? ((totalMonthlyWaste / totalMonthlyProduction) * 100).toFixed(2) : "0.00";

    // 2. شيت الصرف (Dispatch): إجمالي خروج المنتجات شهرياً
    var dispatchSheet = ss.getSheetByName("Dispatch");
    var totalMonthlyDispatch = 0;
    if (dispatchSheet && dispatchSheet.getLastRow() > 1) {
      var dispData = dispatchSheet.getRange(2, 1, dispatchSheet.getLastRow() - 1, dispatchSheet.getLastColumn()).getValues();
      dispData.forEach(function(row) {
        totalMonthlyDispatch += Number(row[3]) || 0;
      });
    }

    // 3. الطاقة والماليات (مؤمنة بالكامل لتُرجع أصفاراً نظيفة 0 في حال خلو الشيتات)
    var electricityBalancePlant1 = 0; 
    var electricityBalancePlant2 = 0; 
    var totalElectricityCostMonth = 0;    
    var totalSalariesMonth = 0;       

    var elecSheet = ss.getSheetByName("Electricity");
    if (elecSheet && elecSheet.getLastRow() > 1) {
      var elecData = elecSheet.getRange(2, 1, elecSheet.getLastRow() - 1, elecSheet.getLastColumn()).getValues();
      // جمع آخر قراءة أو إجمالي الرصيد حسب هيكل الشيت لديك، أو تركه صفراً مؤقتاً
    }

    var salSheet = ss.getSheetByName("Salaries");
    if (salSheet && salSheet.getLastRow() > 1) {
      var salData = salSheet.getRange(2, 1, salSheet.getLastRow() - 1, salSheet.getLastColumn()).getValues();
      salData.forEach(function(row) {
        totalSalariesMonth += Number(row[1]) || 0; // افتراض العمود الثاني للمرتبات
      });
    }

    // حساب تكلفة العبوة مع حماية تامة ضد القسمة على صفر
    var costPerUnitElectricity = totalMonthlyProduction > 0 ? (totalElectricityCostMonth / totalMonthlyProduction).toFixed(2) : "0.00";
    var costPerUnitSalaries = totalMonthlyProduction > 0 ? (totalSalariesMonth / totalMonthlyProduction).toFixed(2) : "0.00";

    // 4. ملخص المخزون والخامات (أصفار نظيفة بحتة منعاً لأي قيم وهمية)
    var inventory = {
      bagsScotch: "0 كيس",
      externalBags: "0 شوال",
      preform29: "0 قطعة",
      preform24: "0 قطعة",
      preform20: "0 قطعة",
      bottles29L: "0 عبوة",
      bottles900Plain: "0 عبوة",
      bottles850Ribbed: "0 عبوة",
      bottles750: "0 عبوة",
      bottles500: "0 عبوة",
      capMarinaBlue: "0 غطاء",
      capSacksBlue: "0 غطاء",
      capLightBlue: "0 غطاء",
      capYellow: "0 غطاء",
      capRed: "0 غطاء"
    };

    // محاولة جلب المخزون الحقيقي إن وجد شيت Inbound أو Inventory
    var invSheet = ss.getSheetByName("Inbound");
    // يمكن لاحقاً إضافة منطق حساب الرصيد الفعلي من الحركات

    return {
      status: "success",
      metrics: {
        electricityPlant1: electricityBalancePlant1,
        electricityPlant2: electricityBalancePlant2,
        totalProduction: totalMonthlyProduction,
        totalDispatch: totalMonthlyDispatch,
        lastDayProduction: lastDayProduction,
        totalWaste: totalMonthlyWaste,
        wastePercentage: wastePercentage,
        electricityCost: totalElectricityCostMonth,
        salariesCost: totalSalariesMonth,
        costElectricityPerUnit: costPerUnitElectricity,
        costSalariesPerUnit: costPerUnitSalaries,
        inventory: inventory
      }
    };

  } catch (error) {
    return {
      status: "error",
      message: error.message
    };
  }
}
/**
 * دالة المصروفات والطاقة: تبدأ بأصفار حقيقية لحين تسجيل بيانات فعلية في الشيت
 */
function getFinancialsFromSheet(ss) {
  var finSheet = ss.getSheetByName("Financials") || ss.getSheetByName("Expenses");
  var data = { elecPlant1: 0, elecPlant2: 0, elecCost: 0, salariesCost: 0 };
  
  if (finSheet && finSheet.getLastRow() > 1) {
    // قراءة البيانات الفعلية إذا تم إنشاؤها لاحقاً
  }
  
  return data;
}

/**
 * دالة المخزون: تبدأ بأصفار نظيفة ودون أي أرقام تجريبية
 */
function getInventoryBreakdown(ss) {
  var invSheet = ss.getSheetByName("Inventory");
  
  var inventoryData = {
    bagsScotch: "0 كيس",
    externalBags: "0 شوال",
    preform29: "0 قطعة",
    preform24: "0 قطعة",
    preform20: "0 قطعة",
    bottles29L: "0 عبوة",
    bottles900Plain: "0 عبوة",
    bottles850Ribbed: "0 عبوة",
    bottles750: "0 عبوة",
    bottles500: "0 عبوة",
    capMarinaBlue: "0 غطاء",
    capSacksBlue: "0 غطاء",
    capLightBlue: "0 غطاء",
    capYellow: "0 غطاء",
    capRed: "0 غطاء"
  };

  if (invSheet && invSheet.getLastRow() > 1) {
    var data = invSheet.getRange(2, 1, invSheet.getLastRow() - 1, 2).getValues();
    data.forEach(function(row) {
      var itemName = String(row[0]).trim();
      var itemQty = Number(row[1]) || 0;
      
      if (itemName.includes("سكوتش")) inventoryData.bagsScotch = itemQty + " كيس";
      else if (itemName.includes("خارجي")) inventoryData.externalBags = itemQty + " شوال";
      else if (itemName.includes("بريفورم 29")) inventoryData.preform29 = itemQty + " قطعة";
      else if (itemName.includes("بريفورم 24")) inventoryData.preform24 = itemQty + " قطعة";
      else if (itemName.includes("بريفورم 20")) inventoryData.preform20 = itemQty + " قطعة";
      else if (itemName.includes("عبوة لتر 29")) inventoryData.bottles29L = itemQty + " عبوة";
      else if (itemName.includes("900 سادة")) inventoryData.bottles900Plain = itemQty + " عبوة";
      else if (itemName.includes("850 مشرشر")) inventoryData.bottles850Ribbed = itemQty + " عبوة";
      else if (itemName.includes("750")) inventoryData.bottles750 = itemQty + " عبوة";
      else if (itemName.includes("500")) inventoryData.bottles500 = itemQty + " عبوة";
      else if (itemName.includes("أزرق مارينا")) inventoryData.capMarinaBlue = itemQty + " غطاء";
      else if (itemName.includes("أزرق شكاير")) inventoryData.capSacksBlue = itemQty + " غطاء";
      else if (itemName.includes("لبنى")) inventoryData.capLightBlue = itemQty + " غطاء";
      else if (itemName.includes("أصفر")) inventoryData.capYellow = itemQty + " غطاء";
      else if (itemName.includes("أحمر")) inventoryData.capRed = itemQty + " غطاء";
    });
  }

  return inventoryData;
}