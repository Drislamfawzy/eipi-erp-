/**
 * Controller Layer: API Endpoint for Inbound.
 */
function api_saveInbound(payload) {
  return withLock(() => {
    try {
      const eventId = saveInbound(payload);
      return { success: true, message: "تم تسجيل إذن الاستلام بنجاح", data: { eventId: eventId } };
    } catch (error) {
      return { success: false, error: error.message };
    }
  });
}