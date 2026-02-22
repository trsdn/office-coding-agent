/**
 * Zustand storage adapter backed by OfficeRuntime.storage.
 *
 * OfficeRuntime.storage is a persistent, async key-value store provided by
 * Office.js for add-ins using the SharedRuntime. Unlike localStorage, it
 * persists reliably across Excel sessions and is not tied to a specific
 * workbook — settings survive when the user opens a different file.
 *
 * This add-in requires SharedRuntime (manifest declares it), so
 * OfficeRuntime.storage is always available in production. Tests supply a
 * lightweight in-memory mock via tests/setup.ts.
 *
 * @see https://learn.microsoft.com/javascript/api/office-runtime/officeruntime.storage
 */
import type { StateStorage } from 'zustand/middleware';

interface OfficeRuntimeStorage {
  getItem(key: string): Promise<string | null>;
  setItem(key: string, value: string): Promise<void>;
  removeItem(key: string): Promise<void>;
}

function getStorage(): OfficeRuntimeStorage {
  if (typeof OfficeRuntime !== 'undefined' && OfficeRuntime?.storage) {
    return OfficeRuntime.storage as OfficeRuntimeStorage;
  }

  // Fallback to localStorage for hosts without SharedRuntime (e.g. Outlook)
  if (typeof localStorage !== 'undefined') {
    return {
      getItem: (key: string) => Promise.resolve(localStorage.getItem(key)),
      setItem: (key: string, value: string) => {
        localStorage.setItem(key, value);
        return Promise.resolve();
      },
      removeItem: (key: string) => {
        localStorage.removeItem(key);
        return Promise.resolve();
      },
    };
  }

  throw new Error(
    '[officeStorage] No storage backend available. ' +
      'OfficeRuntime.storage requires SharedRuntime and localStorage is unavailable.'
  );
}

export const officeStorage: StateStorage = {
  getItem: async (name: string) =>
    getStorage()
      .getItem(name)
      .then(v => v ?? null),
  setItem: async (name: string, value: string) => getStorage().setItem(name, value),
  removeItem: async (name: string) => getStorage().removeItem(name),
};
