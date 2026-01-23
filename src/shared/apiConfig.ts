export const DEFAULT_API_BASE_URL = "https://main-sequence.app";
export const API_BASE_URL_STORAGE_KEY = "api_base_url";

const normalizeBaseUrl = (value: string): string => {
  const trimmed = value.trim();
  if (!trimmed) return DEFAULT_API_BASE_URL;

  const withScheme = /^https?:\/\//i.test(trimmed) ? trimmed : `https://${trimmed}`;
  return withScheme.replace(/\/+$/, "");
};

const safeGetLocalStorage = (): string | null => {
  try {
    return localStorage.getItem(API_BASE_URL_STORAGE_KEY);
  } catch {
    return null;
  }
};

const safeSetLocalStorage = (value: string): void => {
  try {
    localStorage.setItem(API_BASE_URL_STORAGE_KEY, value);
  } catch {
    // Ignore storage failures (e.g., blocked in some environments).
  }
};

const safeRemoveLocalStorage = (): void => {
  try {
    localStorage.removeItem(API_BASE_URL_STORAGE_KEY);
  } catch {
    // Ignore storage failures.
  }
};

const safeGetOfficeStorage = async (): Promise<string | null> => {
  try {
    if (typeof OfficeRuntime === "undefined" || !OfficeRuntime.storage) return null;
    return await OfficeRuntime.storage.getItem(API_BASE_URL_STORAGE_KEY);
  } catch {
    return null;
  }
};

const safeSetOfficeStorage = async (value: string): Promise<void> => {
  try {
    if (typeof OfficeRuntime === "undefined" || !OfficeRuntime.storage) return;
    await OfficeRuntime.storage.setItem(API_BASE_URL_STORAGE_KEY, value);
  } catch {
    // Ignore storage failures.
  }
};

const safeRemoveOfficeStorage = async (): Promise<void> => {
  try {
    if (typeof OfficeRuntime === "undefined" || !OfficeRuntime.storage) return;
    await OfficeRuntime.storage.removeItem(API_BASE_URL_STORAGE_KEY);
  } catch {
    // Ignore storage failures.
  }
};

export const getApiBaseUrlSync = (): string => {
  const stored = safeGetLocalStorage();
  return stored ? normalizeBaseUrl(stored) : DEFAULT_API_BASE_URL;
};

export const getApiBaseUrl = async (): Promise<string> => {
  const stored = await safeGetOfficeStorage();
  if (stored) return normalizeBaseUrl(stored);

  const localStored = safeGetLocalStorage();
  return localStored ? normalizeBaseUrl(localStored) : DEFAULT_API_BASE_URL;
};

export const setApiBaseUrl = async (value: string): Promise<string> => {
  const normalized = normalizeBaseUrl(value);
  safeSetLocalStorage(normalized);
  await safeSetOfficeStorage(normalized);
  return normalized;
};

export const clearApiBaseUrl = async (): Promise<void> => {
  safeRemoveLocalStorage();
  await safeRemoveOfficeStorage();
};
