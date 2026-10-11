function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page.trim() : 'Home';
  
  // تم إضافة 'Form_PettyCash' هنا لتفعيل التوجيه لشاشة العهدة
  const allowedPages = ['Home', 'Index', 'Form_Dispatch', 'Form_Inbound', 'Dashboard', 'Form_PettyCash','Form_Advances', 'Form_Payroll'];
  
  const targetPage = allowedPages.includes(page) ? page : 'Home';

  try {
    const template = HtmlService.createTemplateFromFile(targetPage);
    // تمرير الرابط الحقيقي للـ Web App إلى الـ Frontend بأمان
    template.scriptUrl = ScriptApp.getService().getUrl();
    
    return template.evaluate()
      .setTitle('EIPI Factory ERP - ' + targetPage)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    return HtmlService.createHtmlOutput("<div style='text-align:center; padding:50px; color:red;'>خطأ في تحميل الصفحة: " + err.message + "</div>");
  }
}

function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (e) {
    return "";
  }
}