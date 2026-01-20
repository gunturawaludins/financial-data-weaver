// IndexedDB Database Manager for ETL System
import { openDB, IDBPDatabase } from 'idb';
import { DatabaseRecord } from './types';

const DB_NAME = 'etl_mkbd_database';
const DB_VERSION = 1;
const META_STORE = 'meta_tables';

interface TableMeta {
  tableName: string;
  headers: string[];
  recordCount: number;
  lastUpdated: string;
  createdAt: string;
}

let dbInstance: IDBPDatabase | null = null;

export async function getDatabase(): Promise<IDBPDatabase> {
  if (dbInstance) return dbInstance;

  dbInstance = await openDB(DB_NAME, DB_VERSION, {
    upgrade(db) {
      // Create meta store for table information
      if (!db.objectStoreNames.contains(META_STORE)) {
        db.createObjectStore(META_STORE, { keyPath: 'tableName' });
      }
    },
  });

  return dbInstance;
}

export async function createTableIfNotExists(tableName: string, headers: string[]): Promise<void> {
  const db = await getDatabase();
  
  // Check if table already exists
  if (!db.objectStoreNames.contains(tableName)) {
    // Close and reopen with new version to create new object store
    const currentVersion = db.version;
    db.close();
    dbInstance = null;

    dbInstance = await openDB(DB_NAME, currentVersion + 1, {
      upgrade(database) {
        if (!database.objectStoreNames.contains(tableName)) {
          database.createObjectStore(tableName, { keyPath: '_id', autoIncrement: true });
        }
        if (!database.objectStoreNames.contains(META_STORE)) {
          database.createObjectStore(META_STORE, { keyPath: 'tableName' });
        }
      },
    });
  }

  // Update or create meta info
  const meta: TableMeta = {
    tableName,
    headers,
    recordCount: 0,
    lastUpdated: new Date().toISOString(),
    createdAt: new Date().toISOString(),
  };

  const existingMeta = await dbInstance!.get(META_STORE, tableName);
  if (existingMeta) {
    meta.createdAt = existingMeta.createdAt;
  }

  await dbInstance!.put(META_STORE, meta);
}

export async function appendRecords(tableName: string, records: DatabaseRecord[]): Promise<number> {
  const db = await getDatabase();
  
  if (!db.objectStoreNames.contains(tableName)) {
    throw new Error(`Table ${tableName} does not exist`);
  }

  const tx = db.transaction(tableName, 'readwrite');
  const store = tx.objectStore(tableName);
  
  let addedCount = 0;
  for (const record of records) {
    const { _id, ...recordWithoutId } = record;
    await store.add(recordWithoutId);
    addedCount++;
  }
  
  await tx.done;

  // Update meta record count
  const meta = await db.get(META_STORE, tableName) as TableMeta;
  if (meta) {
    const allRecords = await db.getAll(tableName);
    meta.recordCount = allRecords.length;
    meta.lastUpdated = new Date().toISOString();
    await db.put(META_STORE, meta);
  }

  return addedCount;
}

export async function getAllTables(): Promise<TableMeta[]> {
  const db = await getDatabase();
  
  if (!db.objectStoreNames.contains(META_STORE)) {
    return [];
  }
  
  return await db.getAll(META_STORE);
}

export async function getTableData(tableName: string): Promise<DatabaseRecord[]> {
  const db = await getDatabase();
  
  if (!db.objectStoreNames.contains(tableName)) {
    return [];
  }
  
  return await db.getAll(tableName);
}

export async function clearTable(tableName: string): Promise<void> {
  const db = await getDatabase();
  
  if (!db.objectStoreNames.contains(tableName)) {
    return;
  }
  
  await db.clear(tableName);
  
  // Update meta
  const meta = await db.get(META_STORE, tableName) as TableMeta;
  if (meta) {
    meta.recordCount = 0;
    meta.lastUpdated = new Date().toISOString();
    await db.put(META_STORE, meta);
  }
}

export async function deleteTable(tableName: string): Promise<void> {
  const db = await getDatabase();
  
  // Delete meta entry
  await db.delete(META_STORE, tableName);
  
  // Clear the table data
  if (db.objectStoreNames.contains(tableName)) {
    await db.clear(tableName);
  }
}

export async function getTableStats(): Promise<{ totalTables: number; totalRecords: number }> {
  const tables = await getAllTables();
  const totalRecords = tables.reduce((sum, t) => sum + t.recordCount, 0);
  return { totalTables: tables.length, totalRecords };
}
