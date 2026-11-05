import { StoredReport, StoredUserAccessReport } from '../models/models';

const DB_NAME = 'SmartPermissionsHistory';
const DB_VERSION = 2;
const STORE_NAME = 'reports';
const UA_STORE_NAME = 'userAccessReports';
const MAX_REPORTS = 10;

export class ReportHistoryService {
  // IndexedDB is partitioned by origin (the tenant), not by user or site, so every user of
  // this browser profile shares one database. Records are stamped with the signed-in user
  // and each user only reads, and can only clear, their own. Records saved before this
  // existed carry no owner and stay visible so no history is lost on upgrade.
  constructor(private readonly ownerKey = '') {}

  private isVisible(record: { ownerKey?: string }): boolean {
    return record.ownerKey === undefined || record.ownerKey === this.ownerKey;
  }

  private openDb(): Promise<IDBDatabase> {
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = (e) => {
        const db = (e.target as IDBOpenDBRequest).result;
        if (!db.objectStoreNames.contains(STORE_NAME)) {
          const store = db.createObjectStore(STORE_NAME, { keyPath: 'id' });
          store.createIndex('timestamp', 'timestamp', { unique: false });
        }
        if (!db.objectStoreNames.contains(UA_STORE_NAME)) {
          const uaStore = db.createObjectStore(UA_STORE_NAME, { keyPath: 'id' });
          uaStore.createIndex('timestamp', 'timestamp', { unique: false });
        }
      };
      req.onsuccess = (e) => resolve((e.target as IDBOpenDBRequest).result);
      req.onerror = (e) => reject((e.target as IDBOpenDBRequest).error);
    });
  }

  // Inserts `record`, evicting the oldest entry once the store holds MAX_REPORTS.
  // Every failure path must settle the promise and close the db: a quota-exceeded
  // write aborts the transaction (onabort, not onerror), which previously left the
  // caller hanging with the connection open.
  private async putWithLimit(storeName: string, record: StoredReport | StoredUserAccessReport): Promise<void> {
    const db = await this.openDb();
    return new Promise((resolve, reject) => {
      let settled = false;
      const finish = (error?: unknown): void => {
        if (settled) return;
        settled = true;
        db.close();
        if (error) reject(error); else resolve();
      };

      const tx = db.transaction(storeName, 'readwrite');
      tx.oncomplete = () => finish();
      tx.onerror = () => finish(tx.error ?? new Error('IndexedDB write failed'));
      tx.onabort = () => finish(tx.error ?? new Error('IndexedDB write aborted'));

      const store = tx.objectStore(storeName);
      const countReq = store.count();
      countReq.onerror = () => finish(countReq.error);
      countReq.onsuccess = () => {
        if (countReq.result < MAX_REPORTS) {
          store.put({ ...record, ownerKey: this.ownerKey });
          return;
        }
        // Delete the oldest report (lowest timestamp) to stay within the limit
        const cursorReq = store.index('timestamp').openCursor(null, 'next');
        cursorReq.onerror = () => finish(cursorReq.error);
        cursorReq.onsuccess = (e) => {
          const cursor = (e.target as IDBRequest<IDBCursorWithValue>).result;
          if (cursor) store.delete(cursor.primaryKey);
          store.put({ ...record, ownerKey: this.ownerKey });
        };
      };
    });
  }

  async add(report: StoredReport): Promise<void> {
    return this.putWithLimit(STORE_NAME, report);
  }

  async getAll(): Promise<StoredReport[]> {
    const db = await this.openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, 'readonly');
      const req = tx.objectStore(STORE_NAME).index('timestamp').getAll();
      req.onsuccess = () => {
        db.close();
        resolve((req.result as StoredReport[]).filter((r) => this.isVisible(r)).reverse()); // newest first
      };
      req.onerror = () => { db.close(); reject(req.error); };
    });
  }

  async delete(id: string): Promise<void> {
    const db = await this.openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, 'readwrite');
      tx.objectStore(STORE_NAME).delete(id);
      tx.oncomplete = () => { db.close(); resolve(); };
      tx.onerror = () => { db.close(); reject(tx.error); };
    });
  }

  private async clearStore(storeName: string): Promise<void> {
    const db = await this.openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(storeName, 'readwrite');
      const store = tx.objectStore(storeName);
      const cursorReq = store.openCursor();
      cursorReq.onsuccess = () => {
        const cursor = cursorReq.result;
        if (!cursor) return;
        // Only this user's records (and unowned legacy ones they can already see).
        if (this.isVisible(cursor.value as { ownerKey?: string })) cursor.delete();
        cursor.continue();
      };
      tx.oncomplete = () => { db.close(); resolve(); };
      tx.onerror = () => { db.close(); reject(tx.error); };
      tx.onabort = () => { db.close(); reject(tx.error); };
    });
  }

  clearReports(): Promise<void> { return this.clearStore(STORE_NAME); }

  clearUserAccessReports(): Promise<void> { return this.clearStore(UA_STORE_NAME); }

  // ── User Access report store ──────────────────────────────────────────────

  async addUserAccess(report: StoredUserAccessReport): Promise<void> {
    return this.putWithLimit(UA_STORE_NAME, report);
  }

  async getAllUserAccess(): Promise<StoredUserAccessReport[]> {
    const db = await this.openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(UA_STORE_NAME, 'readonly');
      const req = tx.objectStore(UA_STORE_NAME).index('timestamp').getAll();
      req.onsuccess = () => {
        db.close();
        resolve((req.result as StoredUserAccessReport[]).filter((r) => this.isVisible(r)).reverse());
      };
      req.onerror = () => { db.close(); reject(req.error); };
    });
  }

  async deleteUserAccess(id: string): Promise<void> {
    const db = await this.openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(UA_STORE_NAME, 'readwrite');
      tx.objectStore(UA_STORE_NAME).delete(id);
      tx.oncomplete = () => { db.close(); resolve(); };
      tx.onerror = () => { db.close(); reject(tx.error); };
    });
  }
}
