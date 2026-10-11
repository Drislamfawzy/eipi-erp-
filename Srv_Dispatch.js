/**
 * Service Layer: Processes Product Dispatch and generates Inventory Movements.
 */
function saveDispatch(payload) {
  if (!payload.lines || payload.lines.length === 0) {
    throw new Error("Validation Error: الإذن يجب أن يحتوي على صنف واحد على الأقل.");
  }
  
  const eventId = generateTransactionID('trn_dispatch_hdr', 'DSP', 4);
  const createdAt = new Date().toISOString();
  
  // 1. Insert Header
  insertRecord('trn_dispatch_hdr', {
    Dispatch_ID: eventId,
    Partner_ID: payload.partnerId,
    Trans_Date: payload.date,
    Doc_Reference: payload.docRef || "",
    Cash_Amount: payload.cashAmount || 0, // ADR-06 is OPEN: Just a record, no auto-treasury
    Status: "Posted"
  });

  // 2. Insert Lines & Inventory Movements
  payload.lines.forEach((line, index) => {
    // Dispatch Line
    const lineId = `${eventId}-L${String(index + 1).padStart(2, '0')}`;
    insertRecord('trn_dispatch_line', {
      DispLine_ID: lineId,
      Dispatch_ID: eventId,
      Item_ID: line.itemId,
      Quantity: line.qty
    });

    // Inventory Movement (Deduct from Stock)
    const movId = generateTransactionID('trn_inv_movement', 'INV', 6);
    insertRecord('trn_inv_movement', {
      Movement_ID: movId,
      Item_ID: line.itemId,
      Quantity: line.qty,
      UOM_ID: "", // Will be handled in DB defaults
      Direction: "Decrease",
      Movement_Type: "Dispatch",
      Warehouse_ID: "", // Default till ADR-04 is closed
      Source_Entity_Type: "trn_dispatch_hdr",
      Source_Entity_ID: eventId,
      Status: "Posted",
      Created_At: createdAt
    });
  });

  return eventId;
}