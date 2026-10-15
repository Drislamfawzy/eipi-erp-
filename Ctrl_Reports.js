/**
 * Controller Layer: API Endpoints for Dashboard Metrics & Reports Center
 */

function api_getDashboardMetrics() {
  try {
    return getDashboardMetrics();
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function savePettyCashTransaction(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("PettyCash");
    
    if (!sheet) {
      sheet = ss.insertSheet("PettyCash");
      sheet.appendRow(["Timestamp", "التاريخ", "نوع الحركة", "بند الحركة", "المبلغ", "البيان / ملاحظات"]);
      sheet.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#f3f3f3");
      sheet.setFrozenRows(1);
    }
    
    var timestamp = new Date();
    var date = payload.date;
    var type = payload.type; 
    var category = payload.category;
    var amount = parseFloat(payload.amount) || 0;
    var notes = payload.notes || "";
    
    sheet.appendRow([timestamp, date, type, category, amount, notes]);
    return "تم حفظ حركة العهدة بنجاح!";
  } catch (error) {
    throw new Error("خطأ في حفظ حركة العهدة: " + error.message);
  } finally {
    lock.releaseLock();
  }
}

function getComprehensiveReportsData(startDateStr, endDateStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var startDate = startDateStr ? new Date(startDateStr) : new Date(2000, 0, 1);
    startDate.setHours(0, 0, 0, 0);
    
    var endDate = endDateStr ? new Date(endDateStr) : new Date(2100, 11, 31);
    endDate.setHours(23, 59, 59, 999);

    // ==========================================
    // 1. تقرير العهدة (PettyCash)
    // ==========================================
    var pettyCashSummary = { totalInbound: 0, totalOutbound: 0, netBalance: 0 };
    var pettyCashItems = [];
    var pettySheet = ss.getSheetByName("PettyCash");
    if (pettySheet && pettySheet.getLastRow() > 1) {
      var data = pettySheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var rowDate = new Date(row[1] || row[0]);
        if (isNaN(rowDate.getTime())) rowDate = new Date(2000, 0, 1);
        if (rowDate >= startDate && rowDate <= endDate) {
          var type = String(row[2] || "").trim();
          var category = String(row[3] || "").trim();
          var amount = Math.abs(Number(row[4]) || 0);
          var notes = String(row[5] || "").trim();
          if (type.includes("وارد") || type.includes("إضافة")) {
            pettyCashSummary.totalInbound += amount;
          } else {
            pettyCashSummary.totalOutbound += amount;
          }
          pettyCashItems.push({
            date: Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy/MM/dd"),
            type: type, category: category, amount: amount, notes: notes
          });
        }
      }
      pettyCashSummary.netBalance = pettyCashSummary.totalInbound - pettyCashSummary.totalOutbound;
    }

    // ==========================================
    // 2. تقرير خروج منتج (Dispatch)
    // ==========================================
    var dispatchSummary = { totalQty: 0, totalCash: 0 };
    var dispatchGroupedMap = {};
    var dispatchItems = [];
    var dispSheet = ss.getSheetByName("Dispatch");
    if (dispSheet && dispSheet.getLastRow() > 1) {
      var data = dispSheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var rowDate = new Date(row[1] || row[0]);
        if (isNaN(rowDate.getTime())) rowDate = new Date(2000, 0, 1);
        if (rowDate >= startDate && rowDate <= endDate) {
          var clientName = String(row[2] || "").trim();
          var dispatchNo = String(row[3] || "").trim();
          var itemName = String(row[4] || "").trim();
          var qty = Number(row[5]) || 0;
          var cash = Number(row[6]) || 0;
          dispatchSummary.totalQty += qty;
          dispatchSummary.totalCash += cash;
          if (!dispatchGroupedMap[itemName]) dispatchGroupedMap[itemName] = 0;
          dispatchGroupedMap[itemName] += qty;
          dispatchItems.push({
            date: Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy/MM/dd"),
            clientName: clientName, dispatchNo: dispatchNo, itemName: itemName, qty: qty, cash: cash
          });
        }
      }
    }
    var dispatchGrouped = [];
    for (var item in dispatchGroupedMap) {
      dispatchGrouped.push({ itemName: item, totalQty: dispatchGroupedMap[item] });
    }

    // ==========================================
    // 3. تقرير السلف (Advances)
    // ==========================================
    var advancesSummary = { totalAmount: 0 };
    var advancesGroupedMap = {};
    var advancesItems = [];
    var advSheet = ss.getSheetByName("Advances");
    if (advSheet && advSheet.getLastRow() > 1) {
      var data = advSheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var rowDate = new Date(row[1] || row[0]);
        if (isNaN(rowDate.getTime())) rowDate = new Date(2000, 0, 1);
        if (rowDate >= startDate && rowDate <= endDate) {
          var empName = String(row[2] || "").trim();
          var transType = String(row[3] || "").trim();
          var amount = Number(row[4]) || 0;
          var notes = String(row[5] || "").trim();
          if (empName && (transType === "سلفة" || transType.indexOf("سلف") !== -1)) {
            advancesSummary.totalAmount += amount;
            var dateStr = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy/MM/dd");
            if (!advancesGroupedMap[empName]) {
              advancesGroupedMap[empName] = { count: 0, total: 0, lastDate: dateStr };
            }
            advancesGroupedMap[empName].count += 1;
            advancesGroupedMap[empName].total += amount;
            advancesGroupedMap[empName].lastDate = dateStr;
            advancesItems.push({ date: dateStr, employeeName: empName, amount: amount, notes: notes });
          }
        }
      }
    }
    var advancesGrouped = [];
    for (var emp in advancesGroupedMap) {
      advancesGrouped.push({
        employeeName: emp, lastDate: advancesGroupedMap[emp].lastDate,
        count: advancesGroupedMap[emp].count, total: advancesGroupedMap[emp].total
      });
    }

    // ==========================================
    // 4. تقرير المرتبات (Payroll) - مع تفاصيل إضافية للمدير والعمال
    // ==========================================
    var payrollSummary = { totalNet: 0, totalDeductions: 0, totalAdditions: 0 };
    var payrollGroupedMap = {};
    var payrollItems = [];
    var paySheet = ss.getSheetByName("Payroll");
    
    if (paySheet && paySheet.getLastRow() > 1) {
      var data = paySheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var rowDate = new Date(row[1] || row[0]);
        if (isNaN(rowDate.getTime())) rowDate = new Date(2000, 0, 1);
        
        if (rowDate >= startDate && rowDate <= endDate) {
          var name = String(row[2] || "").trim();
          var attendanceDays = Number(row[3]) || 0;
          var dailyWage = Number(row[12]) || 0;
          var managerComm = Number(row[20]) || 0;
          
          var deductions = Number(row[18]) || Number(row[4]) || 0;
          var additions = Number(row[19]) || 0;
          var netSalary = Number(row[21]) || 0;
          
          payrollSummary.totalNet += netSalary;
          payrollSummary.totalDeductions += deductions;
          payrollSummary.totalAdditions += additions;
          
          if (!payrollGroupedMap[name]) {
            payrollGroupedMap[name] = { days: 0, deductions: 0, additions: 0, net: 0 };
          }
          payrollGroupedMap[name].days += attendanceDays;
          payrollGroupedMap[name].deductions += deductions;
          payrollGroupedMap[name].additions += additions;
          payrollGroupedMap[name].net += netSalary;
          
          // حساب إجمالي عبوات الشهر لمدير الإنتاج إذا وجد
          var abwaQtyForMonth = 0;
          if (dailyWage === 550 || managerComm > 0) {
            abwaQtyForMonth = getAbwaQtyForMonth(ss, rowDate);
          }

          payrollItems.push({
            date: Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy/MM/dd"),
            employeeName: name,
            attendanceDays: attendanceDays,
            dailyWage: dailyWage,
            managerCommission: managerComm,
            abwaQty: abwaQtyForMonth,
            // تفصيل الإضافات
            overtimeHours: Number(row[7]) || 0,
            overtimePay: Number(row[17]) || 0,
            productionBonus: String(row[10] || "لا"),
            productionBonusAmount: (String(row[10]) === "نعم") ? 500 : 0,
            incentives: Number(row[11]) || 0,
            totalAdditions: additions,
            // تفصيل الخصومات
            delayHours: Number(row[6]) || 0,
            delayDeduction: Number(row[14]) || 0,
            penaltyDeduction: Number(row[15]) || Number(row[4]) || 0,
            unexcusedAbsence: Number(row[5]) || 0,
            absenceDeduction: Number(row[16]) || 0,
            advances: Number(row[8]) || 0,
            totalDeductions: deductions,
            netSalary: netSalary
          });
        }
      }
    }
    
    var payrollGrouped = [];
    for (var emp in payrollGroupedMap) {
      payrollGrouped.push({
        employeeName: emp,
        days: payrollGroupedMap[emp].days,
        deductions: payrollGroupedMap[emp].deductions,
        additions: payrollGroupedMap[emp].additions,
        net: payrollGroupedMap[emp].net
      });
    }

    return {
      status: "success",
      data: {
        pettyCash: { summary: pettyCashSummary, items: pettyCashItems },
        dispatch: { summary: dispatchSummary, grouped: dispatchGrouped, items: dispatchItems },
        advances: { summary: advancesSummary, grouped: advancesGrouped, items: advancesItems },
        payroll: { summary: payrollSummary, grouped: payrollGrouped, items: payrollItems }
      }
    };

  } catch (error) {
    return { status: "error", message: error.message };
  }
}

function getAbwaQtyForMonth(ss, targetDate) {
  try {
    var dispSheet = ss.getSheetByName("Dispatch");
    if (!dispSheet || dispSheet.getLastRow() < 2) return 0;
    var data = dispSheet.getDataRange().getValues();
    var targetYear = targetDate.getFullYear();
    var targetMonth = targetDate.getMonth();
    var totalQty = 0;
    for (var i = 1; i < data.length; i++) {
      var rDate = new Date(data[i][1] || data[i][0]);
      if (!isNaN(rDate.getTime()) && rDate.getFullYear() === targetYear && rDate.getMonth() === targetMonth) {
        var itemName = String(data[i][4] || "").trim();
        var qty = Number(data[i][5]) || 0;
        if (itemName.includes("عبو")) {
          totalQty += qty;
        }
      }
    }
    return totalQty;
  } catch(e) {
    return 0;
  }
}
/**
 * دالة لتوليد تفاصيل مرتب العامل كنص أو PDF وإرسالها عبر البريد الإلكتروني للمدير
 */
function sendWorkerPayrollEmail(workerDetailPayload) {
  try {
    var emailBody = "تقرير تفاصيل أجر العامل / الموظف\n" +
                    "==================================\n" +
                    "اسم الموظف: " + workerDetailPayload.employeeName + "\n" +
                    "تاريخ الاستحقاق: " + workerDetailPayload.date + "\n" +
                    "أيام الحضور: " + workerDetailPayload.attendanceDays + " يوم\n" +
                    "أجر اليوم: " + workerDetailPayload.dailyWage + " ج.م\n" +
                    "----------------------------------\n" +
                    "إجمالي الإضافات: " + workerDetailPayload.totalAdditions + " ج.م\n" +
                    "إجمالي الخصومات: " + workerDetailPayload.totalDeductions + " ج.م\n" +
                    "صافي المستحق: " + workerDetailPayload.netSalary + " ج.م\n" +
                    "==================================\n" +
                    "مرسل آلياً من نظام EIPI Factory ERP.";

    MailApp.sendEmail({
      to: "egyptianplast@gmail.com",
      subject: "تفاصيل مرتب العامل: " + workerDetailPayload.employeeName + " - " + workerDetailPayload.date,
      body: emailBody
    });

    return "تم إرسال تفاصيل المرتب بنجاح إلى البريد الإلكتروني للإدارة!";
  } catch (error) {
    throw new Error("خطأ في إرسال الإيميل: " + error.message);
  }
}
/**
 * دالة لتوليد ملف PDF يضم البيانات المفلترة فقط (الفترة المحددة) وإرساله كمرفق
 */
function sendComprehensivePayrollReportPDF(startDateStr, endDateStr) {
  try {
    var reportsRes = getComprehensiveReportsData(startDateStr, endDateStr);
    if (reportsRes.status !== "success") {
      throw new Error("فشل جلب بيانات المرتبات للفترة المحددة.");
    }

    var payrollData = reportsRes.data.payroll;
    var summary = payrollData.summary;
    var items = payrollData.items;

    // 1. بناء هيكل HTML مصغر ومنسق للبيانات المفلترة فقط
    var htmlContent = "<html dir='rtl'><head><meta charset='UTF-8'>" +
                      "<style>body { font-family: 'Cairo', Arial, sans-serif; padding: 20px; color: #333; }" +
                      "h2 { text-align: center; color: #0f172a; }" +
                      "p { text-align: center; font-size: 0.9rem; color: #64748b; }" +
                      "table { width: 100%; border-collapse: collapse; margin-top: 20px; text-align: right; font-size: 0.85rem; }" +
                      "th, td { padding: 8px 10px; border: 1px solid #cbd5e1; }" +
                      "th { background: #0f172a; color: #fff; }" +
                      ".summary { display: flex; justify-content: space-around; margin: 15px 0; font-weight: bold; text-align: center; }" +
                      "</style></head><body>" +
                      "<h2>تقرير أجور ومرتبات العاملين (المفلتر)</h2>" +
                      "<p>الفترة من: <b>" + (startDateStr || "البداية") + "</b> إلى: <b>" + (endDateStr || "اليوم") + "</b></p>" +
                      
                      "<div class='summary'>" +
                        "<div>إجمالي الصافي: " + summary.totalNet.toLocaleString() + " ج.م</div>" +
                        "<div>إجمالي الخصومات: " + summary.totalDeductions.toLocaleString() + " ج.م</div>" +
                        "<div>إجمالي الإضافات: " + summary.totalAdditions.toLocaleString() + " ج.م</div>" +
                      "</div>" +

                      "<table>" +
                      "<thead><tr><th>التاريخ</th><th>اسم الموظف</th><th>أيام الحضور</th><th>الخصومات</th><th>الإضافات</th><th>صافي المستحق</th></tr></thead>" +
                      "<tbody>";

    if (items.length === 0) {
      htmlContent += "<tr><td colspan='6' style='text-align: center;'>لا توجد سجلات مرتبات في هذه الفترة المحددة</td></tr>";
    } else {
      items.forEach(function(i) {
        htmlContent += "<tr>" +
                       "<td>" + i.date + "</td>" +
                       "<td><b>" + i.employeeName + "</b></td>" +
                       "<td>" + i.attendanceDays + " يوم</td>" +
                       "<td style='color: #dc2626;'>" + i.totalDeductions.toLocaleString() + " ج.م</td>" +
                       "<td style='color: #16a34a;'>" + i.totalAdditions.toLocaleString() + " ج.م</td>" +
                       "<td style='font-weight: bold;'>" + i.netSalary.toLocaleString() + " ج.م</td>" +
                       "</tr>";
      });
    }

    htmlContent += "</tbody></table></body></html>";

    // 2. تحويل محتوى الـ HTML المفلتر حصرياً إلى ملف PDF
    var blob = HtmlService.createHtmlOutput(htmlContent).getAs('application/pdf')
               .setName("Payroll_Filtered_Report_" + (startDateStr || "All") + ".pdf");

    var emailBody = "السادة الإدارة،\n\n" +
                    "مرفق لسيادتكم تقرير أجور ومرتبات العاملين (المفلتر حسب الفترة المحددة) بصيغة PDF.\n" +
                    "الفترة من: " + (startDateStr || "البداية") + " إلى: " + (endDateStr || "اليوم") + "\n\n" +
                    "مرسل آلياً من مركز التقارير الشامل - EIPI Factory ERP.";

    // 3. إرسال الإيميل مع إرفاق الـ PDF النظيف والصغير
    MailApp.sendEmail({
      to: "egyptianplast@gmail.com",
      subject: "تقرير المرتبات المفلتر (PDF) - الفترة (" + (startDateStr || "البداية") + " إلى " + (endDateStr || "اليوم") + ")",
      body: emailBody,
      attachments: [blob]
    });

    return "تم إنشاء وإرسال ملف الـ PDF المفلتر بنجاح إلى الإدارة كمرفق!";
  } catch (error) {
    throw new Error("خطأ في تصدير وإرسال الـ PDF المفلتر: " + error.message);
  }
}
/**
 * دالة لتصدير وإرسال "تقرير العهدة المفلتر" كملف PDF مرفق
 */
function sendPettyCashReportPDF(startDateStr, endDateStr) {
  try {
    var reportsRes = getComprehensiveReportsData(startDateStr, endDateStr);
    if (reportsRes.status !== "success") {
      throw new Error("فشل جلب بيانات العهدة للفترة المحددة.");
    }

    var pcData = reportsRes.data.pettyCash;
    var summary = pcData.summary;
    var items = pcData.items;

    var htmlContent = "<html dir='rtl'><head><meta charset='UTF-8'>" +
                      "<style>body { font-family: 'Cairo', Arial, sans-serif; padding: 20px; color: #333; }" +
                      "h2 { text-align: center; color: #0f172a; }" +
                      "p { text-align: center; font-size: 0.9rem; color: #64748b; }" +
                      "table { width: 100%; border-collapse: collapse; margin-top: 20px; text-align: right; font-size: 0.85rem; }" +
                      "th, td { padding: 8px 10px; border: 1px solid #cbd5e1; }" +
                      "th { background: #0f172a; color: #fff; }" +
                      ".summary { display: flex; justify-content: space-around; margin: 15px 0; font-weight: bold; text-align: center; }" +
                      "</style></head><body>" +
                      "<h2>تقرير حركة العهدة والنقدية (المفلتر)</h2>" +
                      "<p>الفترة من: <b>" + (startDateStr || "البداية") + "</b> إلى: <b>" + (endDateStr || "اليوم") + "</b></p>" +
                      
                      "<div class='summary'>" +
                        "<div>إجمالي الوارد: " + summary.totalInbound.toLocaleString() + " ج.م</div>" +
                        "<div>إجمالي الصادر: " + summary.totalOutbound.toLocaleString() + " ج.م</div>" +
                        "<div>صافي الرصيد: " + summary.netBalance.toLocaleString() + " ج.م</div>" +
                      "</div>" +

                      "<table>" +
                      "<thead><tr><th>التاريخ</th><th>نوع الحركة</th><th>البند</th><th>المبلغ</th><th>البيان / ملاحظات</th></tr></thead>" +
                      "<tbody>";

    if (items.length === 0) {
      htmlContent += "<tr><td colspan='5' style='text-align: center;'>لا توجد حركات عهدة في هذه الفترة المحددة</td></tr>";
    } else {
      items.forEach(function(i) {
        htmlContent += "<tr>" +
                       "<td>" + i.date + "</td>" +
                       "<td><b>" + i.type + "</b></td>" +
                       "<td>" + i.category + "</td>" +
                       "<td style='font-weight: bold;'>" + i.amount.toLocaleString() + " ج.م</td>" +
                       "<td>" + i.notes + "</td>" +
                       "</tr>";
      });
    }

    htmlContent += "</tbody></table></body></html>";

    var blob = HtmlService.createHtmlOutput(htmlContent).getAs('application/pdf')
               .setName("PettyCash_Filtered_Report_" + (startDateStr || "All") + ".pdf");

    var emailBody = "السادة الإدارة،\n\n" +
                    "مرفق لسيادتكم تقرير العهدة والنقدية (المفلتر حسب الفترة المحددة) بصيغة PDF.\n" +
                    "الفترة من: " + (startDateStr || "البداية") + " إلى: " + (endDateStr || "اليوم") + "\n\n" +
                    "مرسل آلياً من مركز التقارير الشامل - EIPI Factory ERP.";

    MailApp.sendEmail({
      to: "egyptianplast@gmail.com",
      subject: "تقرير العهدة والنقدية المفلتر (PDF) - الفترة (" + (startDateStr || "البداية") + " إلى " + (endDateStr || "اليوم") + ")",
      body: emailBody,
      attachments: [blob]
    });

    return "تم إنشاء وإرسال ملف PDF العهدة المفلتر بنجاح إلى الإدارة كمرفق!";
  } catch (error) {
    throw new Error("خطأ في إرسال تقرير العهدة: " + error.message);
  }
}

/**
 * دالة لتصدير وإرسال "تقرير السلف المفلتر" كملف PDF مرفق
 */
function sendAdvancesReportPDF(startDateStr, endDateStr) {
  try {
    var reportsRes = getComprehensiveReportsData(startDateStr, endDateStr);
    if (reportsRes.status !== "success") {
      throw new Error("فشل جلب بيانات السلف للفترة المحددة.");
    }

    var advData = reportsRes.data.advances;
    var summary = advData.summary;
    var grouped = advData.grouped;

    var htmlContent = "<html dir='rtl'><head><meta charset='UTF-8'>" +
                      "<style>body { font-family: 'Cairo', Arial, sans-serif; padding: 20px; color: #333; }" +
                      "h2 { text-align: center; color: #0f172a; }" +
                      "p { text-align: center; font-size: 0.9rem; color: #64748b; }" +
                      "table { width: 100%; border-collapse: collapse; margin-top: 20px; text-align: right; font-size: 0.85rem; }" +
                      "th, td { padding: 8px 10px; border: 1px solid #cbd5e1; }" +
                      "th { background: #0f172a; color: #fff; }" +
                      ".summary { text-align: center; margin: 15px 0; font-weight: bold; }" +
                      "</style></head><body>" +
                      "<h2>تقرير السلف الإجمالي للموظفين (المفلتر)</h2>" +
                      "<p>الفترة من: <b>" + (startDateStr || "البداية") + "</b> إلى: <b>" + (endDateStr || "اليوم") + "</b></p>" +
                      
                      "<div class='summary'>" +
                        "إجمالي السلف المنصرفة بالفترة: " + summary.totalAmount.toLocaleString() + " ج.م" +
                      "</div>" +

                      "<table>" +
                      "<thead><tr><th>اسم الموظف</th><th>تاريخ المعاملة (آخر سلفة)</th><th>عدد مرات السلف</th><th>إجمالي المبلغ المسحوب</th></tr></thead>" +
                      "<tbody>";

    if (grouped.length === 0) {
      htmlContent += "<tr><td colspan='4' style='text-align: center;'>لا توجد سلف مسجلة في هذه الفترة المحددة</td></tr>";
    } else {
      grouped.forEach(function(g) {
        htmlContent += "<tr>" +
                       "<td><b>" + g.employeeName + "</b></td>" +
                       "<td>" + g.lastDate + "</td>" +
                       "<td>" + g.count + " مرات</td>" +
                       "<td style='font-weight: bold; color: #d97706;'>" + g.total.toLocaleString() + " ج.م</td>" +
                       "</tr>";
      });
    }

    htmlContent += "</tbody></table></body></html>";

    var blob = HtmlService.createHtmlOutput(htmlContent).getAs('application/pdf')
               .setName("Advances_Filtered_Report_" + (startDateStr || "All") + ".pdf");

    var emailBody = "السادة الإدارة،\n\n" +
                    "مرفق لسيادتكم تقرير السلف الإجمالي للموظفين (المفلتر حسب الفترة المحددة) بصيغة PDF.\n" +
                    "الفترة من: " + (startDateStr || "البداية") + " إلى: " + (endDateStr || "اليوم") + "\n\n" +
                    "مرسل آلياً من مركز التقارير الشامل - EIPI Factory ERP.";

    MailApp.sendEmail({
      to: "egyptianplast@gmail.com",
      subject: "تقرير السلف المفلتر (PDF) - الفترة (" + (startDateStr || "البداية") + " إلى " + (endDateStr || "اليوم") + ")",
      body: emailBody,
      attachments: [blob]
    });

    return "تم إنشاء وإرسال ملف PDF السلف المفلتر بنجاح إلى الإدارة كمرفق!";
  } catch (error) {
    throw new Error("خطأ في إرسال تقرير السلف: " + error.message);
  }
}