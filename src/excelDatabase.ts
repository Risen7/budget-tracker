// Import the spreadsheet namespace as a type so it adds no runtime code.
import type * as XLSX from 'xlsx'

// Restrict transaction type values to the two supported ledger categories.
export type TransactionType = 'income' | 'expense'

// Describe the fields stored for each income or expense record.
export type Transaction = {
  // Keep an identifier that remains unique when records are displayed.
  id: number
  // Store the transaction's human-readable description.
  title: string
  // Optionally associate an archived row with its archive group.
  historyTitle?: string
  // Store the user-selected category for the record.
  category: string
  // Store the transaction amount as a number.
  amount: number
  // Distinguish income from expense so totals can be calculated correctly.
  type: TransactionType
  // Store the calendar date in ISO date format.
  date: string
// Close the current object, callback, or expression.
}

export type GroceryProduct = {
  id: string
  name: string
  store: string
  stock: number
  price: number
  toBuy: boolean
  bought: boolean
}

export type BillItem = {
  id: string
  name: string
  amount: number
  paid: boolean
}

export type BudgetListItem = {
  id: string
  name: string
  amount: number
}

// Extend the browser directory handle with the permission methods used by the app.
export type BrowserDirectoryHandle = FileSystemDirectoryHandle & {
  // Perform this step as part of the surrounding workbook operation.
  queryPermission: (descriptor: { mode: 'readwrite' }) => Promise<PermissionState>
  // Perform this step as part of the surrounding workbook operation.
  requestPermission: (descriptor: { mode: 'readwrite' }) => Promise<PermissionState>
// Close the current object, callback, or expression.
}

// Describe the optional browser directory-picker API without assuming every browser has it.
type DirectoryPickerWindow = Window & {
  // Perform this step as part of the surrounding workbook operation.
  showDirectoryPicker?: (options: { mode: 'readwrite' }) => Promise<BrowserDirectoryHandle>
// Close the current object, callback, or expression.
}

// Keep the active ledger and archived ledger together when reading or writing a workbook.
export type WorkbookData = {
  // Perform this step as part of the surrounding workbook operation.
  transactions: Transaction[]
  // Perform this step as part of the surrounding workbook operation.
  history: Transaction[]
  groceryProducts: GroceryProduct[]
  nonGroceryProducts: GroceryProduct[]
  firstBills: BillItem[]
  secondBills: BillItem[]
  familyItems: BudgetListItem[]
  allowanceItems: BudgetListItem[]
  seldomItems: BudgetListItem[]
// Close the current object, callback, or expression.
}

// Use one consistent workbook name inside the selected folder.
const databaseFileName = 'budget-database.xlsx'
// Name the IndexedDB database that remembers the chosen folder handle.
const settingsDatabaseName = 'budget-tracker-settings'
// Name the IndexedDB object store that holds browser settings.
const settingsStoreName = 'settings'
// Provide initial example records when a new empty workbook is first opened.
const starterTransactions: Transaction[] = [
  // Continue the surrounding list or object with this value.
  { id: 1, title: 'Monthly salary', category: 'Salary', amount: 3200, type: 'income', date: '2026-09-20' },
  // Continue the surrounding list or object with this value.
  { id: 2, title: 'Apartment rent', category: 'Housing', amount: 1150, type: 'expense', date: '2026-09-20' },
  // Continue the surrounding list or object with this value.
  { id: 3, title: 'Grocery run', category: 'Food', amount: 86.45, type: 'expense', date: '2026-09-19' },
  // Continue the surrounding list or object with this value.
  { id: 4, title: 'Freelance project', category: 'Side income', amount: 450, type: 'income', date: '2026-09-18' },
  // Continue the surrounding list or object with this value.
  { id: 5, title: 'Train pass', category: 'Transport', amount: 42, type: 'expense', date: '2026-09-18' },
// Perform this step as part of the surrounding workbook operation.
]

// Check whether this browser exposes the folder-selection API.
export function supportsFolderSelection() {
  // Return the value produced by this function or callback.
  return typeof (window as DirectoryPickerWindow).showDirectoryPicker === 'function'
// Close the current object, callback, or expression.
}

// Ask the browser to let the user select a folder for the Excel workbook.
export async function chooseDatabaseDirectory() {
  // Access the optional picker through the extended browser-window type.
  const pickerWindow = window as DirectoryPickerWindow
  // Fail clearly when the browser cannot provide the required folder access.
  if (!pickerWindow.showDirectoryPicker) {
    // Stop the operation and report this error to the caller.
    throw new Error('Folder access requires Chrome or Edge on a secure connection.')
  // Close the current object, callback, or expression.
  }
  // Open a folder chooser and request permission to read and write the folder.
  return pickerWindow.showDirectoryPicker({ mode: 'readwrite' })
// Close the current object, callback, or expression.
}

// Open the persistent browser database used to remember the last selected folder.
function openSettingsDatabase() {
  // Adapt IndexedDB's event-based API to a promise for async callers.
  return new Promise<IDBDatabase>((resolve, reject) => {
    // Create the settings store the first time the database is upgraded.
    const request = indexedDB.open(settingsDatabaseName, 1)
    // Handle this IndexedDB event and settle the pending operation.
    request.onupgradeneeded = () => request.result.createObjectStore(settingsStoreName)
    // Return the open database handle once IndexedDB has finished opening.
    request.onsuccess = () => resolve(request.result)
    // Reject with the browser's underlying error if opening fails.
    request.onerror = () => reject(request.error ?? new Error('Unable to open browser storage.'))
  // Close the current object, callback, or expression.
  })
// Close the current object, callback, or expression.
}

// Retrieve the saved directory handle from IndexedDB, if one exists.
export async function getSavedDirectory() {
  // Open the settings database before starting a read transaction.
  const database = await openSettingsDatabase()
  // Resolve the stored value or reject if the IndexedDB read fails.
  return new Promise<BrowserDirectoryHandle | undefined>((resolve, reject) => {
    // Compute transaction to hold the database operation and its access mode.
    const transaction = database.transaction(settingsStoreName, 'readonly')
    // Request the previously persisted directory handle.
    const request = transaction.objectStore(settingsStoreName).get('directory') as IDBRequest<BrowserDirectoryHandle | undefined>
    // Close the database and return the saved handle when the request succeeds.
    request.onsuccess = () => {
      // Release the IndexedDB connection after the operation finishes.
      database.close()
      // Resolve the surrounding promise with the completed result.
      resolve(request.result)
    // Close the current object, callback, or expression.
    }
    // Close the database and report any failure while reading the saved handle.
    request.onerror = () => {
      // Release the IndexedDB connection after the operation finishes.
      database.close()
      // Reject the surrounding promise with a useful failure reason.
      reject(request.error ?? new Error('Unable to restore the saved folder.'))
    // Close the current object, callback, or expression.
    }
  // Close the current object, callback, or expression.
  })
// Close the current object, callback, or expression.
}

// Store the selected folder handle so the app can offer to reconnect on a later visit.
export async function saveDirectory(handle: BrowserDirectoryHandle) {
  // Open IndexedDB before writing the handle.
  const database = await openSettingsDatabase()
  // Resolve only after the write transaction finishes successfully.
  return new Promise<void>((resolve, reject) => {
    // Compute transaction to hold the database operation and its access mode.
    const transaction = database.transaction(settingsStoreName, 'readwrite')
    // Save the browser directory handle under the stable "directory" setting key.
    transaction.objectStore(settingsStoreName).put(handle, 'directory')
    // Close the database after the write is committed.
    transaction.oncomplete = () => {
      // Release the IndexedDB connection after the operation finishes.
      database.close()
      // Resolve the surrounding promise with the completed result.
      resolve()
    // Close the current object, callback, or expression.
    }
    // Close the database and report errors raised by the write transaction.
    transaction.onerror = () => {
      // Release the IndexedDB connection after the operation finishes.
      database.close()
      // Reject the surrounding promise with a useful failure reason.
      reject(transaction.error ?? new Error('Unable to remember the selected folder.'))
    // Close the current object, callback, or expression.
    }
    // Also report an aborted transaction rather than silently losing the setting.
    transaction.onabort = () => {
      // Release the IndexedDB connection after the operation finishes.
      database.close()
      // Reject the surrounding promise with a useful failure reason.
      reject(transaction.error ?? new Error('Unable to remember the selected folder.'))
    // Close the current object, callback, or expression.
    }
  // Close the current object, callback, or expression.
  })
// Close the current object, callback, or expression.
}

// Convert rows from a workbook worksheet into the application's transaction shape.
function workbookTransactions(XLSX: typeof import('xlsx'), sheet: XLSX.WorkSheet | undefined) {
  // Treat a missing worksheet as an empty transaction list.
  if (!sheet) return []
  // Read every spreadsheet row and normalize the columns to typed application fields.
  return XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: '' }).map((row) => ({
    // Store the record identifier.
    id: Number(row.ID),
    // Store the transaction description.
    title: String(row.Description ?? ''),
    // Store the title that groups this archived record.
    historyTitle: String(row.HistoryTitle ?? ''),
    // Store the selected transaction category.
    category: String(row.Category ?? 'Other'),
    // Store the monetary amount.
    amount: Number(row.Amount ?? 0),
    // Store whether this row represents income or an expense.
    type: row.Type === 'income' ? 'income' as const : 'expense' as const,
    // Store the transaction or archive date.
    date: String(row.Date ?? ''),
  // Perform this step as part of the surrounding workbook operation.
  }))
// Close the current object, callback, or expression.
}

// Build an XLSX workbook containing totals, current transactions, and saved history.
async function createWorkbook(data: WorkbookData) {
  // Load SheetJS only when an Excel operation is actually needed.
  const XLSX = await import('xlsx')
  // Calculate income and expense totals from the current transaction list.
  const totals = data.transactions.reduce((summary, transaction) => {
    // Add this transaction amount to its matching running total.
    summary[transaction.type] += transaction.amount
    // Return the value produced by this function or callback.
    return summary
  // Perform this step as part of the surrounding workbook operation.
  }, { income: 0, expense: 0 })
  // Create the empty workbook that will hold all three worksheets.
  const workbook = XLSX.utils.book_new()
  // Build a summary sheet containing current totals and the remaining balance.
  const summarySheet = XLSX.utils.json_to_sheet([
    // Continue the surrounding list or object with this value.
    { Metric: 'Total income', Amount: totals.income },
    // Continue the surrounding list or object with this value.
    { Metric: 'Total expenses', Amount: totals.expense },
    // Continue the surrounding list or object with this value.
    { Metric: 'Available balance', Amount: totals.income - totals.expense },
  // Close the current object, callback, or expression.
  ])
  // Convert active transactions into column names suitable for an Excel sheet.
  const transactionSheet = XLSX.utils.json_to_sheet(data.transactions.map((transaction) => ({
    // Write the transaction identifier to the Excel column.
    ID: transaction.id,
    // Write the transaction date to the Excel column.
    Date: transaction.date,
    // Write the transaction description to the Excel column.
    Description: transaction.title,
    // Write the transaction category to the Excel column.
    Category: transaction.category,
    // Write the income-or-expense type to the Excel column.
    Type: transaction.type,
    // Write the monetary amount to the Excel column.
    Amount: transaction.amount,
  // Perform this step as part of the surrounding workbook operation.
  })))
  // Convert archived transactions into rows, including their history group title.
  const historySheet = XLSX.utils.json_to_sheet(data.history.map((transaction) => ({
    // Write the transaction identifier to the Excel column.
    ID: transaction.id,
    // Write the transaction date to the Excel column.
    Date: transaction.date,
    // Write the archive group title to the Excel column.
    HistoryTitle: transaction.historyTitle ?? '',
    // Write the transaction description to the Excel column.
    Description: transaction.title,
    // Write the transaction category to the Excel column.
    Category: transaction.category,
    // Write the income-or-expense type to the Excel column.
    Type: transaction.type,
    // Write the monetary amount to the Excel column.
    Amount: transaction.amount,
  // Perform this step as part of the surrounding workbook operation.
  })))
  const productRows = (products: GroceryProduct[]) => products.map((product) => ({
    ID: product.id,
    Name: product.name,
    Store: product.store,
    Stock: product.stock,
    Price: product.price,
    ToBuy: product.toBuy,
    Bought: product.bought,
  }))
  const billRows = (bills: BillItem[]) => bills.map((bill) => ({
    ID: bill.id,
    Name: bill.name,
    Amount: bill.amount,
    Paid: bill.paid,
  }))
  const budgetItemRows = (items: BudgetListItem[]) => items.map((item) => ({
    ID: item.id,
    Name: item.name,
    Amount: item.amount,
  }))
  // Set readable widths for summary-sheet columns.
  summarySheet['!cols'] = [{ wch: 22 }, { wch: 16 }]
  // Set readable widths for transaction-sheet columns.
  transactionSheet['!cols'] = [{ wch: 12 }, { wch: 14 }, { wch: 26 }, { wch: 18 }, { wch: 12 }, { wch: 14 }]
  // Reuse the transaction column layout for the structurally similar history sheet.
  historySheet['!cols'] = transactionSheet['!cols']
  // Add the summary sheet under its workbook tab name.
  XLSX.utils.book_append_sheet(workbook, summarySheet, 'Summary')
  // Add the active transaction rows under their workbook tab name.
  XLSX.utils.book_append_sheet(workbook, transactionSheet, 'Transactions')
  // Add archived entries under the expense-history tab name.
  XLSX.utils.book_append_sheet(workbook, historySheet, 'Expense History')
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(productRows(data.groceryProducts)), 'Grocery Products')
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(productRows(data.nonGroceryProducts)), 'Non-Grocery Products')
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(billRows(data.firstBills)), 'Day 1-15 Bills')
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(billRows(data.secondBills)), 'Day 16-30 Bills')
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(budgetItemRows(data.familyItems)), 'Family Bonding')
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(budgetItemRows(data.allowanceItems)), 'Cut-off Allowance')
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(budgetItemRows(data.seldomItems)), 'Seldom Buys')
  // Serialize the workbook as bytes that the browser file API can write.
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'array' })
// Close the current object, callback, or expression.
}

// Write current and archived records into the selected folder's workbook file.
export async function writeWorkbook(handle: BrowserDirectoryHandle, data: WorkbookData) {
  // Open or create the workbook file in the selected directory.
  const fileHandle = await handle.getFileHandle(databaseFileName, { create: true })
  // Open a writable stream for replacing the file contents.
  const writable = await fileHandle.createWritable()
  // Generate the workbook bytes and write them into the file.
  await writable.write(await createWorkbook(data))
  // Close the stream so the browser commits the completed workbook.
  await writable.close()
// Close the current object, callback, or expression.
}

// Load transactions and history from the selected folder's workbook.
export async function loadWorkbook(handle: BrowserDirectoryHandle): Promise<WorkbookData> {
  // Check that the browser still grants the app read/write access to this folder.
  const permission = await handle.queryPermission({ mode: 'readwrite' })
  // Stop with a useful message when permission must be granted again.
  if (permission !== 'granted') {
    // Stop the operation and report this error to the caller.
    throw new Error('Folder access needs permission again. Reconnect the saved folder or choose it again.')
  // Close the current object, callback, or expression.
  }

  // Open or create the workbook file within the selected folder.
  const fileHandle = await handle.getFileHandle(databaseFileName, { create: true })
  // Read the actual file object so its size and bytes can be inspected.
  const file = await fileHandle.getFile()
  // Initialize a blank file with sample records and return that initial data.
  if (!file.size) {
    // Compute initialData to hold the starter ledger and empty history.
    const initialData: WorkbookData = {
      transactions: starterTransactions,
      history: [],
      groceryProducts: [],
      nonGroceryProducts: [],
      firstBills: [],
      secondBills: [],
      familyItems: [],
      allowanceItems: [],
      seldomItems: [],
    }
    // Wait for this asynchronous operation to complete before continuing.
    await writeWorkbook(handle, initialData)
    // Return the value produced by this function or callback.
    return initialData
  // Close the current object, callback, or expression.
  }

  // Load SheetJS only when parsing a non-empty workbook.
  const XLSX = await import('xlsx')
  // Parse the workbook from the browser-provided file bytes.
  const workbook = XLSX.read(await file.arrayBuffer())
  const rowsFromSheet = (name: string) => workbook.Sheets[name]
    ? XLSX.utils.sheet_to_json<Record<string, unknown>>(workbook.Sheets[name], { defval: '' })
    : []
  const readProducts = (name: string): GroceryProduct[] => rowsFromSheet(name).map((row) => ({
    id: String(row.ID ?? ''),
    name: String(row.Name ?? ''),
    store: String(row.Store || 'Other'),
    stock: Number(row.Stock ?? 0),
    price: Number(row.Price ?? 0),
    toBuy: row.ToBuy === true || row.ToBuy === 'TRUE',
    bought: row.Bought === true || row.Bought === 'TRUE',
  }))
  const readBills = (name: string): BillItem[] => rowsFromSheet(name).map((row) => ({
    id: String(row.ID ?? ''),
    name: String(row.Name ?? ''),
    amount: Number(row.Amount ?? 0),
    paid: row.Paid === true || row.Paid === 'TRUE',
  }))
  const readBudgetItems = (name: string): BudgetListItem[] => rowsFromSheet(name).map((row) => ({
    id: String(row.ID ?? ''),
    name: String(row.Name ?? ''),
    amount: Number(row.Amount ?? 0),
  }))
  // Return normalized active and archived transaction rows from their sheets.
  return {
    // Continue the surrounding list or object with this value.
    transactions: workbookTransactions(XLSX, workbook.Sheets.Transactions),
    // Continue the surrounding list or object with this value.
    history: workbookTransactions(XLSX, workbook.Sheets['Expense History']),
    groceryProducts: readProducts('Grocery Products'),
    nonGroceryProducts: readProducts('Non-Grocery Products'),
    firstBills: readBills('Day 1-15 Bills'),
    secondBills: readBills('Day 16-30 Bills'),
    familyItems: readBudgetItems('Family Bonding'),
    allowanceItems: readBudgetItems('Cut-off Allowance'),
    seldomItems: readBudgetItems('Seldom Buys'),
  // Close the current object, callback, or expression.
  }
// Close the current object, callback, or expression.
}
