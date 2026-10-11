/**
 * Service Layer: Processes Inbound Material Receipts.
 */
function saveInbound(payload) {
  if (!payload.lines || payload.lines.length === 0) {
    throw new Error("Validation Error: إذن الاستلام يجب أن يحتوي على صنف واحد على الأقل.");
  }
  
  const eventId = generateTransactionID('trn_inbound_hdr', 'INB', 4);
  const createdAt = new Date().toISOString();
  
  // 1. Insert Header
  insertRecord('trn_inbound_hdr', {
    Inbound_ID: eventId,
    Partner_ID: payload.partnerId,
    Trans_Date: payload.date,
    Doc_Reference: payload.docRef || "",
    Status: "Posted"
  });

  // 2. Insert Lines & Inventory Movements (Increase)
  payload.lines.forEach((line, index) => {
    const lineId = `${eventId}-L${String(index + 1).padStart(2, '0')}`;
    insertRecord('trn_inbound_line', {
      InbLine_ID: lineId,
      Inbound_ID: eventId,
      Item_ID: line.itemId,
      Quantity: line.qty
    });

    // Inventory Movement (Add to Stock)
    const movId = generateTransactionID('trn_inv_movement', 'INV', 6);
    insertRecord('trn_inv_movement', {
      Movement_ID: movId,
      Item_ID: line.itemId,
      Quantity: line.qty,
      UOM_ID: "", 
      Direction: "Increase", // زيادة الرصيد
      Movement_Type: "Inbound",
      Warehouse_ID: "", 
      Source_Entity_Type: "trn_inbound_hdr",
      Source_Entity_ID: eventId,
      Status: "Posted",
      Created_At: createdAt
    });
  });

  return eventId;
}