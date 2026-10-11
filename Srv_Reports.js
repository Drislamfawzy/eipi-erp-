/**
 * Service Layer: Aggregates data for Executive Dashboard and Operational Reports.
 * Implements read-only logic without altering database state.
 */
function getDashboardMetrics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Total Production Quantity Sum
  const prodLineSheet = ss.getSheetByName('trn_prod_line');
  let totalProduction = 0;
  if (prodLineSheet && prodLineSheet.getLastRow() > 1) {
    const data = prodLineSheet.getRange(2, 6, prodLineSheet.getLastRow() - 1, 1).getValues(); // Produced_Qty is col 6
    totalProduction = data.flat().reduce((acc, val) => acc + (parseFloat(val) || 0), 0);
  }

  // 2. Total Scrap Quantity Sum
  const scrapSheet = ss.getSheetByName('trn_scrap_event');
  let totalScrap = 0;
  if (scrapSheet && scrapSheet.getLastRow() > 1) {
    const data = scrapSheet.getRange(2, 5, scrapSheet.getLastRow() - 1, 1).getValues(); // Quantity is col 5
    totalScrap = data.flat().reduce((acc, val) => acc + (parseFloat(val) || 0), 0);
  }

  // 3. Total Dispatched Quantity
  const dispLineSheet = ss.getSheetByName('trn_dispatch_line');
  let totalDispatched = 0;
  if (dispLineSheet && dispLineSheet.getLastRow() > 1) {
    const data = dispLineSheet.getRange(2, 4, dispLineSheet.getLastRow() - 1, 1).getValues(); // Quantity is col 4
    totalDispatched = data.flat().reduce((acc, val) => acc + (parseFloat(val) || 0), 0);
  }

  // 4. Calculate Scrap Percentage
  let scrapPercentage = totalProduction > 0 ? ((totalScrap / totalProduction) * 100).toFixed(2) : 0;

  return {
    success: true,
    data: {
      totalProduction: totalProduction,
      totalScrap: totalScrap,
      totalDispatched: totalDispatched,
      scrapPercentage: scrapPercentage + "%"
    }
  };
}