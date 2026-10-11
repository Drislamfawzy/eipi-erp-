/**
 * Controller Layer: API Endpoint for Frontend UI to save Dispatch.
 */
function api_saveDispatch(payload) {
  return withLock(() => {
    try {
      const eventId = saveDispatch(payload);
      return { success: true, message: "تم تسجيل إذن الصرف بنجاح", data: { eventId: eventId } };
    } catch (error) {
      return { success: false, error: error.message };
    }
  });
}