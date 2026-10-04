import type * as XLSX from 'xlsx'

export type TransactionType = 'income' | 'expense'

export type Transaction = {
  id: number
  title: string
  historyTitle?: string
  category: string
  amount: number
  type: TransactionType
  date: string
}

export type BrowserDirectoryHandle = FileSystemDirectoryHandle & {
  queryPermission: (descriptor: { mode: 'readwrite' }) => Promise<PermissionState>
}

type DirectoryPickerWindow = Window & {
  showDirectoryPicker?: (options: { mode: 'readwrite' }) => Promise<BrowserDirectoryHandle>
}

type WorkbookData = {
  transactions: Transaction[]
  history: Transaction[]
}

const databaseFileName = 'budget-database.xlsx'
const settingsDatabaseName = 'budget-tracker-settings'
const settingsStoreName = 'settings'
const starterTransactions: Transaction[] = [
  { id: 1, title: 'Monthly salary', category: 'Salary', amount: 3200, type: 'income', date: '2026-09-20' },
  { id: 2, title: 'Apartment rent', category: 'Housing', amount: 1150, type: 'expense', date: '2026-09-20' },
  { id: 3, title: 'Grocery run', category: 'Food', amount: 86.45, type: 'expense', date: '2026-09-19' },
  { id: 4, title: 'Freelance project', category: 'Side income', amount: 450, type: 'income', date: '2026-09-18' },
  { id: 5, title: 'Train pass', category: 'Transport', amount: 42, type: 'expense', date: '2026-09-18' },
]

export function supportsFolderSelection() {
  return typeof (window as DirectoryPickerWindow).showDirectoryPicker === 'function'
}

export async function chooseDatabaseDirectory() {
  const pickerWindow = window as DirectoryPickerWindow
  if (!pickerWindow.showDirectoryPicker) {
    throw new Error('Folder access requires Chrome or Edge on a secure connection.')
  }
  return pickerWindow.showDirectoryPicker({ mode: 'readwrite' })
}

function openSettingsDatabase() {
  return new Promise<IDBDatabase>((resolve, reject) => {
    const request = indexedDB.open(settingsDatabaseName, 1)
    request.onupgradeneeded = () => request.result.createObjectStore(settingsStoreName)
    request.onsuccess = () => resolve(request.result)
    request.onerror = () => reject(request.error ?? new Error('Unable to open browser storage.'))
  })
}

export async function getSavedDirectory() {
  const database = await openSettingsDatabase()
  return new Promise<BrowserDirectoryHandle | undefined>((resolve, reject) => {
    const transaction = database.transaction(settingsStoreName, 'readonly')
    const request = transaction.objectStore(settingsStoreName).get('directory') as IDBRequest<BrowserDirectoryHandle | undefined>
    request.onsuccess = () => {
      database.close()
      resolve(request.result)
    }
    request.onerror = () => {
      database.close()
      reject(request.error ?? new Error('Unable to restore the saved folder.'))
    }
  })
}

export async function saveDirectory(handle: BrowserDirectoryHandle) {
  const database = await openSettingsDatabase()
  return new Promise<void>((resolve, reject) => {
    const transaction = database.transaction(settingsStoreName, 'readwrite')
    transaction.objectStore(settingsStoreName).put(handle, 'directory')
    transaction.oncomplete = () => {
      database.close()
      resolve()
    }
    transaction.onerror = () => {
      database.close()
      reject(transaction.error ?? new Error('Unable to remember the selected folder.'))
    }
    transaction.onabort = () => {
      database.close()
      reject(transaction.error ?? new Error('Unable to remember the selected folder.'))
    }
  })
}

function workbookTransactions(XLSX: typeof import('xlsx'), sheet: XLSX.WorkSheet | undefined) {
  if (!sheet) return []
  return XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: '' }).map((row) => ({
    id: Number(row.ID),
    title: String(row.Description ?? ''),
    historyTitle: String(row.HistoryTitle ?? ''),
    category: String(row.Category ?? 'Other'),
    amount: Number(row.Amount ?? 0),
    type: row.Type === 'income' ? 'income' as const : 'expense' as const,
    date: String(row.Date ?? ''),
  }))
}

async function createWorkbook(transactions: Transaction[], history: Transaction[]) {
  const XLSX = await import('xlsx')
  const totals = transactions.reduce((summary, transaction) => {
    summary[transaction.type] += transaction.amount
    return summary
  }, { income: 0, expense: 0 })
  const workbook = XLSX.utils.book_new()
  const summarySheet = XLSX.utils.json_to_sheet([
    { Metric: 'Total income', Amount: totals.income },
    { Metric: 'Total expenses', Amount: totals.expense },
    { Metric: 'Available balance', Amount: totals.income - totals.expense },
  ])
  const transactionSheet = XLSX.utils.json_to_sheet(transactions.map((transaction) => ({
    ID: transaction.id,
    Date: transaction.date,
    Description: transaction.title,
    Category: transaction.category,
    Type: transaction.type,
    Amount: transaction.amount,
  })))
  const historySheet = XLSX.utils.json_to_sheet(history.map((transaction) => ({
    ID: transaction.id,
    Date: transaction.date,
    HistoryTitle: transaction.historyTitle ?? '',
    Description: transaction.title,
    Category: transaction.category,
    Type: transaction.type,
    Amount: transaction.amount,
  })))
  summarySheet['!cols'] = [{ wch: 22 }, { wch: 16 }]
  transactionSheet['!cols'] = [{ wch: 12 }, { wch: 14 }, { wch: 26 }, { wch: 18 }, { wch: 12 }, { wch: 14 }]
  historySheet['!cols'] = transactionSheet['!cols']
  XLSX.utils.book_append_sheet(workbook, summarySheet, 'Summary')
  XLSX.utils.book_append_sheet(workbook, transactionSheet, 'Transactions')
  XLSX.utils.book_append_sheet(workbook, historySheet, 'Expense History')
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'array' })
}

export async function writeWorkbook(handle: BrowserDirectoryHandle, data: WorkbookData) {
  const fileHandle = await handle.getFileHandle(databaseFileName, { create: true })
  const writable = await fileHandle.createWritable()
  await writable.write(await createWorkbook(data.transactions, data.history))
  await writable.close()
}

export async function loadWorkbook(handle: BrowserDirectoryHandle): Promise<WorkbookData> {
  const permission = await handle.queryPermission({ mode: 'readwrite' })
  if (permission !== 'granted') {
    throw new Error('Folder access needs permission again. Choose the folder to reconnect.')
  }

  const fileHandle = await handle.getFileHandle(databaseFileName, { create: true })
  const file = await fileHandle.getFile()
  if (!file.size) {
    const initialData = { transactions: starterTransactions, history: [] }
    await writeWorkbook(handle, initialData)
    return initialData
  }

  const XLSX = await import('xlsx')
  const workbook = XLSX.read(await file.arrayBuffer())
  return {
    transactions: workbookTransactions(XLSX, workbook.Sheets.Transactions),
    history: workbookTransactions(XLSX, workbook.Sheets['Expense History']),
  }
}
