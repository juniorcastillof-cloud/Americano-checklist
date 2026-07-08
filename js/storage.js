/**
 * storage.js
 * Local persistence layer using IndexedDB.
 *
 * This is the app's local "database". It is the source of truth the app
 * reads from to decide which checklist items are already Pass/locked for a
 * given Phase + Area/Location, and it is what gets rebuilt into the Excel
 * workbook on every submit.
 *
 * Two object stores:
 *   - "items"   : one row per (phase, area, checklistItem) -> latest status
 *   - "records" : one row per submitted inspection event (full audit trail)
 *   - "photos"  : blobs keyed by photoId, referenced from items/records
 */

const DB_NAME = 'americano_checklist_db';
const DB_VERSION = 1;

let dbPromise = null;

function openDB() {
  if (dbPromise) return dbPromise;
  dbPromise = new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = (event) => {
      const db = event.target.result;
      if (!db.objectStoreNames.contains('items')) {
        const items = db.createObjectStore('items', { keyPath: 'key' });
        items.createIndex('phaseArea', 'phaseArea', { unique: false });
      }
      if (!db.objectStoreNames.contains('records')) {
        const records = db.createObjectStore('records', { keyPath: 'id', autoIncrement: true });
        records.createIndex('phaseArea', 'phaseArea', { unique: false });
        records.createIndex('timestamp', 'timestamp', { unique: false });
      }
      if (!db.objectStoreNames.contains('photos')) {
        db.createObjectStore('photos', { keyPath: 'id' });
      }
      if (!db.objectStoreNames.contains('meta')) {
        db.createObjectStore('meta', { keyPath: 'key' });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
  return dbPromise;
}

function tx(storeName, mode) {
  return openDB().then((db) => db.transaction(storeName, mode).objectStore(storeName));
}

function itemKey(phaseId, area, itemId) {
  return `${phaseId}::${area}::${itemId}`;
}

const Storage = {
  /** Get the current stored state of a single checklist item for a phase+area. */
  async getItem(phaseId, area, itemId) {
    const store = await tx('items', 'readonly');
    return new Promise((resolve, reject) => {
      const req = store.get(itemKey(phaseId, area, itemId));
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => reject(req.error);
    });
  },

  /** Get all stored item states for a given phase + area/location. */
  async getItemsForPhaseArea(phaseId, area) {
    const store = await tx('items', 'readonly');
    return new Promise((resolve, reject) => {
      const idx = store.index('phaseArea');
      const req = idx.getAll(IDBKeyRange.only(`${phaseId}::${area}`));
      req.onsuccess = () => resolve(req.result || []);
      req.onerror = () => reject(req.error);
    });
  },

  /** Upsert a single item's status/notes. Locked items (Pass) are preserved unless explicitly reopened. */
  async putItem(phaseId, area, itemId, data) {
    const store = await tx('items', 'readwrite');
    const record = {
      key: itemKey(phaseId, area, itemId),
      phaseArea: `${phaseId}::${area}`,
      phaseId,
      area,
      itemId,
      ...data,
    };
    return new Promise((resolve, reject) => {
      const req = store.put(record);
      req.onsuccess = () => resolve(record);
      req.onerror = () => reject(req.error);
    });
  },

  /** Get every item row in the whole local database (used to rebuild the full workbook). */
  async getAllItems() {
    const store = await tx('items', 'readonly');
    return new Promise((resolve, reject) => {
      const req = store.getAll();
      req.onsuccess = () => resolve(req.result || []);
      req.onerror = () => reject(req.error);
    });
  },

  /** Append an inspection submission event to the audit trail. */
  async addRecord(record) {
    const store = await tx('records', 'readwrite');
    return new Promise((resolve, reject) => {
      const req = store.add(record);
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  },

  /** Get every submitted inspection record (used for the audit/log sheet). */
  async getAllRecords() {
    const store = await tx('records', 'readonly');
    return new Promise((resolve, reject) => {
      const req = store.getAll();
      req.onsuccess = () => resolve((req.result || []).sort((a, b) => a.timestamp.localeCompare(b.timestamp)));
      req.onerror = () => reject(req.error);
    });
  },

  /** Store a photo blob and return its generated id. */
  async putPhoto(id, blob, meta) {
    const store = await tx('photos', 'readwrite');
    return new Promise((resolve, reject) => {
      const req = store.put({ id, blob, ...meta });
      req.onsuccess = () => resolve(id);
      req.onerror = () => reject(req.error);
    });
  },

  async getPhoto(id) {
    const store = await tx('photos', 'readonly');
    return new Promise((resolve, reject) => {
      const req = store.get(id);
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => reject(req.error);
    });
  },

  async getAllPhotos() {
    const store = await tx('photos', 'readonly');
    return new Promise((resolve, reject) => {
      const req = store.getAll();
      req.onsuccess = () => resolve(req.result || []);
      req.onerror = () => reject(req.error);
    });
  },

  /** Small key/value store for app settings (e.g. last inspector name, Graph account cache flags). */
  async setMeta(key, value) {
    const store = await tx('meta', 'readwrite');
    return new Promise((resolve, reject) => {
      const req = store.put({ key, value });
      req.onsuccess = () => resolve(value);
      req.onerror = () => reject(req.error);
    });
  },

  async getMeta(key) {
    const store = await tx('meta', 'readonly');
    return new Promise((resolve, reject) => {
      const req = store.get(key);
      req.onsuccess = () => resolve(req.result ? req.result.value : null);
      req.onerror = () => reject(req.error);
    });
  },

  /** Wipe all local data (used by the "Reset local data" admin action). */
  async clearAll() {
    const db = await openDB();
    return new Promise((resolve, reject) => {
      const t = db.transaction(['items', 'records', 'photos'], 'readwrite');
      t.objectStore('items').clear();
      t.objectStore('records').clear();
      t.objectStore('photos').clear();
      t.oncomplete = () => resolve();
      t.onerror = () => reject(t.error);
    });
  },
};

window.Storage = Storage;
