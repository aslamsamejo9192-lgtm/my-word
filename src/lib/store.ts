import { openDB, type IDBPDatabase } from 'idb'
import type { MyDocument } from './types'

const DB_NAME = 'my-word-db'
const DB_VERSION = 1

let dbPromise: Promise<IDBPDatabase> | null = null

function getDb(): Promise<IDBPDatabase> {
  if (!dbPromise) {
    dbPromise = openDB(DB_NAME, DB_VERSION, {
      upgrade(db) {
        if (!db.objectStoreNames.contains('documents')) {
          const store = db.createObjectStore('documents', { keyPath: 'id' })
          store.createIndex('updatedAt', 'updatedAt')
        }
        if (!db.objectStoreNames.contains('meta')) {
          db.createObjectStore('meta')
        }
      },
    })
  }
  return dbPromise
}

export async function listDocuments(): Promise<MyDocument[]> {
  const db = await getDb()
  const all = await db.getAll('documents')
  return all.sort((a, b) => b.updatedAt - a.updatedAt)
}

export async function getDocument(id: string): Promise<MyDocument | undefined> {
  const db = await getDb()
  return db.get('documents', id)
}

export async function putDocument(doc: MyDocument): Promise<void> {
  const db = await getDb()
  await db.put('documents', { ...doc, updatedAt: Date.now() })
}

export async function deleteDocument(id: string): Promise<void> {
  const db = await getDb()
  await db.delete('documents', id)
  const last = await getLastOpenId()
  if (last === id) await setLastOpenId(null)
}

export async function duplicateDocument(id: string): Promise<MyDocument | null> {
  const doc = await getDocument(id)
  if (!doc) return null
  const copy: MyDocument = {
    ...doc,
    id: `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 9)}`,
    title: `${doc.title} (Copy)`,
    createdAt: Date.now(),
    updatedAt: Date.now(),
  }
  await putDocument(copy)
  return copy
}

export async function getLastOpenId(): Promise<string | null> {
  const db = await getDb()
  return (await db.get('meta', 'lastOpenId')) ?? null
}

export async function setLastOpenId(id: string | null): Promise<void> {
  const db = await getDb()
  if (id) await db.put('meta', id, 'lastOpenId')
  else await db.delete('meta', 'lastOpenId')
}
