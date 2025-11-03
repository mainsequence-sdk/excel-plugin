/**
 * Utility functions for Google Sheets add-on
 */
function isValidDate(dateString: string): boolean {
  const date = new Date(dateString);
  return date instanceof Date && !isNaN(date.getTime());
}

function dateRangeToQuery(startDate: string, endDate: string): string {
  if (!isValidDate(startDate) || !isValidDate(endDate)) {
    throw new Error('Invalid date format');
  }
  return `start=${encodeURIComponent(startDate)}&end=${encodeURIComponent(endDate)}`;
}

function logError(message: string, error: any): void {
  const timestamp = new Date().toISOString();
  console.error(`[${timestamp}] ${message}:`, error);
}

function getCurrentUserEmail(): string | null {
  try {
    return Session.getActiveUser().getEmail();
  } catch (error) {
    logError('Failed to get user email', error);
    return null;
  }
}
