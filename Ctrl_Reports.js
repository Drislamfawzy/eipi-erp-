/**
 * Controller Layer: API Endpoint for Dashboard Metrics.
 */
function api_getDashboardMetrics() {
  try {
    return getDashboardMetrics();
  } catch (error) {
    return { success: false, error: error.message };
  }
}