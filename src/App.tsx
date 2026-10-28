// Import React state and derived-state hooks.
import { useEffect, useMemo, useState } from 'react'
// Import the form event type without adding it to the runtime bundle.
import type { Dispatch, FormEvent, SetStateAction } from 'react'
// Open the callback or block that contains the following operation.
import {
  // Perform this step in the dashboard calculation or state update.
  chooseDatabaseDirectory as openDatabaseDirectoryPicker,
  // Begin loading the saved directory handle from browser storage.
  getSavedDirectory,
  // Perform this step in the dashboard calculation or state update.
  loadWorkbook,
  // Perform this step in the dashboard calculation or state update.
  saveDirectory,
  // Perform this step in the dashboard calculation or state update.
  supportsFolderSelection,
  // Perform this step in the dashboard calculation or state update.
  writeWorkbook,
// Perform this step in the dashboard calculation or state update.
} from './excelDatabase'
// Perform this step in the dashboard calculation or state update.
import type { BrowserDirectoryHandle, Transaction, TransactionType } from './excelDatabase'
// Load the dashboard stylesheet.
import './App.css'

// Format all displayed amounts consistently as Philippine pesos.
const currency = new Intl.NumberFormat('en-PH', { style: 'currency', currency: 'PHP' })

// Format a local calendar date as YYYY-MM-DD for transaction matching.
function toISODate(date: Date) {
  const year = date.getFullYear()
  const month = String(date.getMonth() + 1).padStart(2, '0')
  const day = String(date.getDate()).padStart(2, '0')
  return `${year}-${month}-${day}`
}

// Capture today's local date for the transaction date input.
const today = toISODate(new Date())

type TodoItem = {
  id: string
  title: string
  completed: boolean
}

type TodoStorageState = {
  items: TodoItem[]
  writable: boolean
  message: string
}

type GroceryProduct = {
  id: string
  name: string
  stock: number
  price: number
  toBuy: boolean
  bought: boolean
}

type GroceryStorageState = {
  products: GroceryProduct[]
  writable: boolean
  message: string
}

function readProductStorage(storageKey: string, listName: string): GroceryStorageState {
  try {
    const savedProducts = window.localStorage.getItem(storageKey)
    if (!savedProducts) return { products: [], writable: true, message: '' }
    const parsed: unknown = JSON.parse(savedProducts)
    if (!Array.isArray(parsed) || !parsed.every((product) => (
      typeof product.id === 'string'
      && typeof product.name === 'string'
      && typeof product.stock === 'number'
      && Number.isFinite(product.stock)
      && product.stock >= 0
      && (product.price === undefined || (
        typeof product.price === 'number'
        && Number.isFinite(product.price)
        && product.price >= 0
      ))
      && typeof product.toBuy === 'boolean'
      && (product.bought === undefined || typeof product.bought === 'boolean')
    ))) {
      throw new Error(`Saved ${listName} list has an invalid format.`)
    }
    return {
      products: parsed.map((product) => ({
        id: product.id,
        name: product.name,
        stock: product.stock,
        price: product.price ?? 0,
        toBuy: product.toBuy,
        bought: product.bought ?? false,
      })) as GroceryProduct[],
      writable: true,
      message: '',
    }
  } catch (error) {
    return {
      products: [],
      writable: false,
      message: error instanceof Error ? error.message : `Unable to load saved ${listName} products.`,
    }
  }
}

type ProductListCardProps = {
  id: string
  title: string
  description: string
  icon: string
  products: GroceryProduct[]
  writable: boolean
  message: string
  onSave: (products: GroceryProduct[]) => boolean
  onAdd: (name: string) => boolean
}

function ProductListCard({ id, title, description, icon, products, writable, message, onSave, onAdd }: ProductListCardProps) {
  const [name, setName] = useState('')
  const priorityProducts = [...products].sort(
    (first, second) => Number(first.bought) - Number(second.bought)
      || Number(second.toBuy) - Number(first.toBuy),
  )
  const isGroceryList = id === 'grocery-list'
  const total = products.reduce(
    (sum, product) => sum + (isGroceryList && (!product.toBuy || product.bought) ? 0 : product.stock * product.price),
    0,
  )

  function addProduct(event: FormEvent<HTMLFormElement>) {
    event.preventDefault()
    const trimmedName = name.trim()
    if (trimmedName && onAdd(trimmedName)) setName('')
  }

  return (
    <article className="budget-card product-list-card">
      <div className="grocery-card-header">
        <span className="budget-card-icon" aria-hidden="true">{icon}</span>
        <div><h2 id={`${id}-heading`}>{title}</h2><p>{description}</p></div>
        <div className="product-list-total"><span>{isGroceryList ? 'To buy total' : 'Stock value'}</span><div className="product-list-total-value"><strong>{currency.format(total)}</strong><span className="estimate-only">ESTIMATE ONLY</span></div></div>
        <span className="grocery-priority-count">{products.filter((product) => product.toBuy && !product.bought).length} priority to buy</span>
      </div>
      <form className="grocery-form" onSubmit={addProduct}>
        <input aria-label={`${title} product name`} value={name} onChange={(event) => setName(event.target.value)} placeholder="Add a product…" required disabled={!writable} />
        <button className="add-button grocery-add-button" type="submit" aria-label={`Add ${title} product`} disabled={!writable || !name.trim()}>+ Add</button>
      </form>
      {message && <p className="grocery-error" role="alert">{message}</p>}
      {!products.length && <p className="grocery-empty">No products yet. Add a product to get started.</p>}
      {!!products.length && <ul className="grocery-product-list" aria-labelledby={`${id}-heading`}>
        {priorityProducts.map((product) => (
          <li className={`grocery-product${product.toBuy && !product.bought ? ' priority' : ''}${product.bought ? ' bought' : ''}`} key={product.id}>
            <label className="grocery-bought-control">
              <input type="checkbox" checked={product.bought} disabled={!writable} onChange={() => onSave(products.map((item) => item.id === product.id ? { ...item, bought: !item.bought } : item))} />
              <span className="visually-hidden">Mark {product.name} as bought</span>
            </label>
            <div className="grocery-product-name"><strong>{product.name}</strong>{product.toBuy && !product.bought && <span className="grocery-priority-badge">To buy</span>}</div>
            <label className="grocery-row-field grocery-stock-field"><span className="visually-hidden">Stock</span><input aria-label={`Current stock for ${product.name}`} type="number" min="0" step="any" value={product.stock} disabled={!writable} onChange={(event) => {
              const stock = Number(event.target.value)
              if (Number.isFinite(stock) && stock >= 0) {
                onSave(products.map((item) => item.id === product.id ? { ...item, stock } : item))
              }
            }} /></label>
            <label className="grocery-row-field grocery-price-field"><span className="visually-hidden">Unit price</span><span className="currency-prefix" aria-hidden="true">₱</span><input aria-label={`Unit price for ${product.name}`} type="number" min="0" step="0.01" value={product.price.toFixed(2)} disabled={!writable} onChange={(event) => {
              const price = Number(event.target.value)
              if (Number.isFinite(price) && price >= 0) {
                onSave(products.map((item) => item.id === product.id ? { ...item, price } : item))
              }
            }} /></label>
            <label className="grocery-row-field grocery-buy-field"><span className="visually-hidden">To buy</span><select aria-label={`Mark ${product.name} to buy`} value={product.toBuy ? 'yes' : 'no'} disabled={!writable} onChange={(event) => onSave(products.map((item) => item.id === product.id ? { ...item, toBuy: event.target.value === 'yes' } : item))}><option value="yes">Buy</option><option value="no">Skip</option></select></label>
            <button className="grocery-delete" type="button" aria-label={`Remove ${product.name}`} disabled={!writable} onClick={() => onSave(products.filter((item) => item.id !== product.id))}>×</button>
          </li>
        ))}
      </ul>}
    </article>
  )
}

type BillItem = {
  id: string
  name: string
  amount: number
  paid: boolean
}

type BillStorageState = {
  bills: BillItem[]
  writable: boolean
  message: string
}

function readBillStorage(storageKey: string, period: string): BillStorageState {
  try {
    const savedBills = window.localStorage.getItem(storageKey)
    if (!savedBills) return { bills: [], writable: true, message: '' }
    const parsed: unknown = JSON.parse(savedBills)
    if (!Array.isArray(parsed) || !parsed.every((bill) => (
      typeof bill.id === 'string'
      && typeof bill.name === 'string'
      && typeof bill.amount === 'number'
      && Number.isFinite(bill.amount)
      && bill.amount >= 0
      && (bill.paid === undefined || typeof bill.paid === 'boolean')
    ))) {
      throw new Error(`Saved ${period} bills have an invalid format.`)
    }
    return {
      bills: parsed.map((bill) => ({ ...bill, paid: bill.paid ?? false })) as BillItem[],
      writable: true,
      message: '',
    }
  } catch (error) {
    return {
      bills: [],
      writable: false,
      message: error instanceof Error ? error.message : `Unable to load ${period} bills.`,
    }
  }
}

type BillsCardProps = {
  period: 'Day 1–15' | 'Day 16–30'
  bills: BillItem[]
  writable: boolean
  message: string
  onSave: (bills: BillItem[]) => boolean
}

function BillsCard({ period, bills, writable, message, onSave }: BillsCardProps) {
  const [name, setName] = useState('')
  const total = bills.reduce((sum, bill) => sum + bill.amount, 0)

  function addBill(event: FormEvent<HTMLFormElement>) {
    event.preventDefault()
    const trimmedName = name.trim()
    if (!trimmedName || !writable) return
    const saved = onSave([...bills, {
      id: `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      name: trimmedName,
      amount: 0,
      paid: false,
    }])
    if (saved) setName('')
  }

  return (
    <article className="budget-card bills bills-editor">
      <div className="bill-card-header">
        <span className="budget-card-icon" aria-hidden="true">🧾</span>
        <div><p className="budget-card-kicker">{period === 'Day 1–15' ? 'First half of the month' : 'Second half of the month'}</p><h2 id={`${period === 'Day 1–15' ? 'day-1-15' : 'day-16-30'}-heading`}>{period} bills</h2></div>
        <div className="bill-total"><span>Total bills</span><strong>{currency.format(total)}</strong></div>
      </div>
      <form className="bill-form" onSubmit={addBill}>
        <input aria-label={`${period} bill name`} value={name} onChange={(event) => setName(event.target.value)} placeholder="Add a bill…" required disabled={!writable} />
        <button className="add-button" type="submit" disabled={!writable || !name.trim()}>+ Add bill</button>
      </form>
      {message && <p className="grocery-error" role="alert">{message}</p>}
      {!bills.length && <p className="grocery-empty">No bills yet. Add a bill to get started.</p>}
      {!!bills.length && <ul className="bill-list" aria-labelledby={`${period === 'Day 1–15' ? 'day-1-15' : 'day-16-30'}-heading`}>
        {bills.map((bill) => (
          <li className={bill.paid ? 'bill-item paid' : 'bill-item'} key={bill.id}>
            <label className="bill-paid-control">
              <input type="checkbox" checked={bill.paid} disabled={!writable} onChange={() => onSave(bills.map((item) => item.id === bill.id ? { ...item, paid: !item.paid } : item))} />
              <span className="visually-hidden">Mark {bill.name} as paid</span>
            </label>
            <strong className="bill-name">{bill.name}</strong>
            <label className="bill-amount-field">
              <span className="visually-hidden">Amount for {bill.name}</span>
              <span aria-hidden="true">₱</span>
              <input aria-label={`Amount for ${bill.name}`} type="number" min="0" step="0.01" value={bill.amount.toFixed(2)} disabled={!writable} onChange={(event) => {
                const amount = Number(event.target.value)
                if (Number.isFinite(amount) && amount >= 0) {
                  onSave(bills.map((item) => item.id === bill.id ? { ...item, amount } : item))
                }
              }} />
            </label>
            <button className="grocery-delete" type="button" aria-label={`Remove ${bill.name}`} disabled={!writable} onClick={() => onSave(bills.filter((item) => item.id !== bill.id))}>×</button>
          </li>
        ))}
      </ul>}
    </article>
  )
}

function readTodoStorage(): TodoStorageState {
  try {
    const savedItems = window.localStorage.getItem('budget-tracker-todos')
    if (!savedItems) return { items: [], writable: true, message: '' }
    const parsed: unknown = JSON.parse(savedItems)
    if (!Array.isArray(parsed) || !parsed.every((item) => (
      typeof item.id === 'string'
      && typeof item.title === 'string'
      && typeof item.completed === 'boolean'
    ))) {
      throw new Error('Saved to-do list has an invalid format.')
    }
    const savedDate = window.localStorage.getItem('budget-tracker-todos-date')
    const items = parsed as TodoItem[]
    return {
      items: savedDate && savedDate !== toISODate(new Date())
        ? items.map((item) => ({ ...item, completed: false }))
        : items,
      writable: true,
      message: '',
    }
  } catch (error) {
    return {
      items: [],
      writable: false,
      message: error instanceof Error ? error.message : 'Unable to load saved to-do items.',
    }
  }
}

// Format the current local date for the dashboard greeting.
function formatCurrentDate() {
  return new Date().toLocaleDateString('en-US', {
    weekday: 'long',
    month: 'long',
    day: 'numeric',
    year: 'numeric',
  })
}

// Convert stored ISO dates into friendly labels for the activity list.
function formatDate(date: string) {
  // Label the current sample day as Today.
  if (date === '2026-09-20') return 'Today'
  // Label the previous sample day as Yesterday.
  if (date === '2026-09-19') return 'Yesterday'
  // Format all older dates with a short month and day.
  return new Date(`${date}T12:00:00`).toLocaleDateString('en-US', { month: 'short', day: 'numeric' })
// Close the current callback, block, or value collection.
}

// Render the complete budget dashboard.
function App() {
  const [activeTab, setActiveTab] = useState<'dashboard' | 'budgets'>('dashboard')
  // Restore the user's saved color theme, defaulting to light mode.
  const [theme, setTheme] = useState<'light' | 'dark'>(() => (
    window.localStorage.getItem('budget-tracker-theme') === 'dark' ? 'dark' : 'light'
  ))
  // Restore the saved text-size step, using the middle step as the default.
  const [fontSizeScale, setFontSizeScale] = useState(() => {
    const savedScale = Number(window.localStorage.getItem('budget-tracker-font-size'))
    return Number.isInteger(savedScale) && savedScale >= 0 && savedScale <= 2 ? savedScale : 1
  })
  // Hold the transactions loaded from the selected Excel workbook.
  const [transactions, setTransactions] = useState<Transaction[]>([])
  // Hold transactions that have been moved into expense history.
  const [history, setHistory] = useState<Transaction[]>([])
  // Track whether the new entry is income or an expense.
  const [type, setType] = useState<TransactionType>('expense')
  // Track the new transaction description field.
  const [title, setTitle] = useState('')
  // Track the selected transaction category.
  const [category, setCategory] = useState('Food')
  // Filter recent activity to one transaction date, or show every date.
  const [activityDate, setActivityDate] = useState('all')
  // Track date groups manually while defaulting today's activity to collapsed.
  const [collapsedActivityGroups, setCollapsedActivityGroups] = useState<Record<string, boolean>>({})
  // Track the new amount as text while the user types.
  const [amount, setAmount] = useState('')
  // Track the date selected for the new transaction.
  const [date, setDate] = useState(today)
  // Show the result of archiving the active transaction list.
  const [archiveMessage, setArchiveMessage] = useState('')
  // Track the title used to group an archived batch.
  const [historyTitle, setHistoryTitle] = useState('')
  // Hold dashboard to-do items and their form input.
  const [todoState, setTodoState] = useState(readTodoStorage)
  const { items: todoItems, writable: todoStorageWritable, message: todoMessage } = todoState
  const [todoDraft, setTodoDraft] = useState('')
  const [groceryState, setGroceryState] = useState(() => readProductStorage('budget-tracker-grocery-products', 'grocery'))
  const { products: groceryProducts, writable: groceryStorageWritable, message: groceryMessage } = groceryState
  const [nonGroceryState, setNonGroceryState] = useState(() => readProductStorage('budget-tracker-nongrocery-products', 'non-grocery'))
  const { products: nonGroceryProducts, writable: nonGroceryStorageWritable, message: nonGroceryMessage } = nonGroceryState
  const [firstBillState, setFirstBillState] = useState(() => readBillStorage('budget-tracker-day-1-15-bills', 'Day 1–15'))
  const { bills: firstBills, writable: firstBillsWritable, message: firstBillsMessage } = firstBillState
  const [secondBillState, setSecondBillState] = useState(() => readBillStorage('budget-tracker-day-16-30-bills', 'Day 16–30'))
  const { bills: secondBills, writable: secondBillsWritable, message: secondBillsMessage } = secondBillState
  // Track which saved history groups are expanded or collapsed.
  const [collapsedGroups, setCollapsedGroups] = useState<Record<string, boolean>>({})
  // Show the current folder used for the Excel workbook.
  const [databaseDirectory, setDatabaseDirectory] = useState('')
  // Show errors from selecting or loading a database folder.
  const [databaseMessage, setDatabaseMessage] = useState('')
  // Prevent opening multiple folder pickers at once.
  const [isConnectingDatabase, setIsConnectingDatabase] = useState(false)
  // Retain browser access to the selected folder for workbook reads and writes.
  const [directoryHandle, setDirectoryHandle] = useState<BrowserDirectoryHandle | null>(null)
  // Keep the saved handle available when browser permission must be renewed.
  const [savedDirectoryHandle, setSavedDirectoryHandle] = useState<BrowserDirectoryHandle | null>(null)

  // Apply the selected theme to the page and remember it for the next visit.
  useEffect(() => {
    document.documentElement.dataset.theme = theme
    window.localStorage.setItem('budget-tracker-theme', theme)
  }, [theme])

  // Apply and persist the selected text size so it is shared by both themes.
  useEffect(() => {
    document.documentElement.dataset.fontSizeScale = String(fontSizeScale)
    window.localStorage.setItem('budget-tracker-font-size', String(fontSizeScale))
  }, [fontSizeScale])

  function saveTodoItems(items: TodoItem[]) {
    try {
      window.localStorage.setItem('budget-tracker-todos', JSON.stringify(items))
      window.localStorage.setItem('budget-tracker-todos-date', toISODate(new Date()))
      setTodoState((current) => ({ ...current, items }))
      return true
    } catch (error) {
      setTodoState((current) => ({
        ...current,
        writable: false,
        message: error instanceof Error ? error.message : 'Unable to save to-do items.',
      }))
      return false
    }
  }

  // Keep the task list saved, and clear completion checks when a new day begins.
  useEffect(() => {
    if (!todoStorageWritable) return
    const dailyReset = window.setInterval(() => {
      const currentDate = toISODate(new Date())
      try {
        const savedDate = window.localStorage.getItem('budget-tracker-todos-date')
        if (!savedDate) {
          window.localStorage.setItem('budget-tracker-todos-date', currentDate)
        } else if (savedDate !== currentDate) {
          saveTodoItems(todoItems.map((item) => ({ ...item, completed: false })))
        }
      } catch (error) {
        setTodoState((current) => ({
          ...current,
          writable: false,
          message: error instanceof Error ? error.message : 'Unable to reset to-do items for the new day.',
        }))
      }
    }, 60_000)

    return () => window.clearInterval(dailyReset)
  }, [todoItems, todoStorageWritable])

  // Restore the previously selected folder when the browser still grants access.
  useEffect(() => {
    // Ignore asynchronous results after this component has been removed.
    let isCurrent = true
    // Retrieve the folder handle stored during the previous session.
    getSavedDirectory()
      // Load the workbook after a usable saved folder is found.
      .then(async (handle) => {
        // Stop if no folder was saved or this effect is no longer active.
        if (!handle || !isCurrent) return
        // Keep the handle so the UI can request permission again later.
        setSavedDirectoryHandle(handle)
        // Show the selected folder name in the database controls.
        setDatabaseDirectory(handle.name)
        // Read the active transaction and history sheets from the workbook.
        const data = await loadWorkbook(handle)
        // Avoid updating React state if the component unmounted while loading.
        if (!isCurrent) return
        // Grant the rest of the app the loaded directory handle.
        setDirectoryHandle(handle)
        // Populate the active transaction state from the workbook.
        setTransactions(data.transactions)
        // Populate the archived transaction state from the workbook.
        setHistory(data.history)
      // Close the current callback, block, or value collection.
      })
      // Display a readable message if restoring or loading the workbook fails.
      .catch((error: unknown) => {
        // Check this condition before continuing the dashboard operation.
        if (isCurrent) setDatabaseMessage(error instanceof Error ? error.message : 'Unable to restore the Excel database folder.')
      // Close the current callback, block, or value collection.
      })
    // Mark pending async work stale when the component unmounts.
    return () => { isCurrent = false }
  // Run this restore attempt once when the dashboard first mounts.
  }, [])

  // Calculate total income and expenses whenever the transaction list changes.
  const totals = useMemo(() => transactions.reduce((summary, transaction) => {
    // Add each amount to its matching income or expense total.
    summary[transaction.type] += transaction.amount
    // Return the running summary for the next transaction.
    return summary
  // Perform this step in the dashboard calculation or state update.
  }, { income: 0, expense: 0 }), [transactions])

  // Summarize the last 15 days so the dashboard can chart recent spending.
  const weeklyExpenditure = useMemo(() => {
    // Create one entry for each of the 15 calendar days shown on the chart.
    const days = Array.from({ length: 15 }, (_, index) => {
      // Start with the current local date for this chart position.
      const date = new Date()
      // Normalize the time to midday to avoid date shifts around midnight.
      date.setHours(12, 0, 0, 0)
      // Move the date back from today to the day represented by this index.
      date.setDate(date.getDate() - (14 - index))
      // Convert the date to the same ISO format stored on transactions.
      const isoDate = toISODate(date)
      // Add the amounts of expenses whose date matches this chart day.
      const total = transactions
        // Perform this step in the dashboard calculation or state update.
        .filter((transaction) => transaction.type === 'expense' && transaction.date === isoDate)
        // Perform this step in the dashboard calculation or state update.
        .reduce((sum, transaction) => sum + transaction.amount, 0)

      // Return the date key, day label, and expense total for one bar.
      return {
        // Perform this step in the dashboard calculation or state update.
        key: isoDate,
        // Perform this step in the dashboard calculation or state update.
        label: String(date.getDate()),
        // Perform this step in the dashboard calculation or state update.
        total,
      // Close the current callback, block, or value collection.
      }
    // Close the current callback, block, or value collection.
    })

    // Return all 15 date entries to the chart calculation.
    return days
  // Recalculate chart values whenever the active transaction list changes.
  }, [transactions])
  const todayExpense = weeklyExpenditure[weeklyExpenditure.length - 1]?.total ?? 0

  // Ensure bars have a nonzero scale even when there are no recent expenses.
  const chartMax = Math.max(...weeklyExpenditure.map((item) => item.total), 1)
  // Convert daily totals into positioned rectangles for the SVG chart.
  const chartBars = weeklyExpenditure.map((day, index) => {
    // Space each bar evenly from the left side of the chart.
    const x = 20 + index * 22
    // Keep every bar the same width.
    const width = 12
    // Scale bar height in proportion to its day's expense total.
    const height = (day.total / chartMax) * 88
    // Position the rectangle so taller bars start higher in the chart.
    const y = 130 - height

    // Carry the day details together with its rendered rectangle geometry.
    return { ...day, x, y, width, height }
  // Close the current callback, block, or value collection.
  })

  function showActivityForDate(date: string) {
    setActivityDate('all')
    setCollapsedActivityGroups((current) => ({ ...current, [date]: false }))
    window.setTimeout(() => {
      const target = document.getElementById(`activity-day-${date}`)
        ?? document.getElementById('recent-activity-heading')
      target?.scrollIntoView({ behavior: 'smooth', block: 'center' })
    }, 0)
  }

  // Organize entries by date so the activity list can show daily groups.
  // Group each transaction under its ISO date for the recent-activity list.
  const groupedTransactions = useMemo(() => transactions
    .filter((transaction) => activityDate === 'all' || transaction.date === activityDate)
    .reduce<Record<string, Transaction[]>>((groups, transaction) => {
      // Append the current transaction to its date group.
      groups[transaction.date] = [...(groups[transaction.date] ?? []), transaction]
      // Return the running date groups for the next transaction.
      return groups
    // Perform this step in the dashboard calculation or state update.
    }, {}), [transactions, activityDate])
  const activityDates = useMemo(() => [...new Set(transactions.map((transaction) => transaction.date))]
    .sort((first, second) => second.localeCompare(first)), [transactions])

  // Compute summaryCards to hold a value used by the surrounding dashboard logic.
  const summaryCards = useMemo(() => {
    // Isolate saved summary rows from individual archived transactions.
    const summaries = history.filter((transaction) => transaction.category === 'Summary')
    // Create one accumulator per archive title with totals and archived rows.
    const grouped = new Map<string, { income: number, expense: number, balance: number, entries: Transaction[] }>()

    // Put each saved summary amount into its matching archive accumulator.
    summaries.forEach((transaction) => {
      // Use the archive title as the key, with a fallback for untitled summaries.
      const key = transaction.historyTitle ?? 'Summary'
      // Initialize totals and an empty row list the first time this title appears.
      const current = grouped.get(key) ?? { income: 0, expense: 0, balance: 0, entries: [] }

      // Store the saved total in the income slot.
      if (transaction.title === 'Total income') current.income = transaction.amount
      // Store the saved total in the expense slot.
      if (transaction.title === 'Total expenses') current.expense = transaction.amount
      // Store the saved remaining balance.
      if (transaction.title === 'Available balance') current.balance = transaction.amount

      // Save the updated accumulator under its archive title.
      grouped.set(key, current)
    // Close the current callback, block, or value collection.
    })

    // Attach each non-summary archived transaction to the matching group.
    history.filter((transaction) => transaction.category !== 'Summary').forEach((transaction) => {
      // Use the explicit history title, or make a date-based group for older rows.
      const key = transaction.historyTitle || `Saved ${formatDate(transaction.date)}`
      // Reuse an existing group or create a new one for this archive.
      const current = grouped.get(key) ?? { income: 0, expense: 0, balance: 0, entries: [] }
      // Append this archived row without mutating the previous list.
      current.entries = [...current.entries, transaction]
      // Save the updated group back into the map.
      grouped.set(key, current)
    // Close the current callback, block, or value collection.
    })

    // Convert the map to renderable objects, preserving each archive title.
    return Array.from(grouped.entries()).map(([title, values]) => ({ title, ...values }))
  // Rebuild archive groups whenever the saved history changes.
  }, [history])

  // Update the screen and persist the new list in the Excel database.
  async function save(nextTransactions: Transaction[]) {
    // Refuse to persist edits when no writable workbook folder is connected.
    if (!directoryHandle) {
      // Update React state or form behavior for the operation in progress.
      setDatabaseMessage('Choose an Excel database folder before making changes.')
      // Return the result of this function, callback, or calculation.
      return false
    // Close the current callback, block, or value collection.
    }
    // Write the next transaction list while preserving the current history.
    try {
      // Wait until this asynchronous save or browser operation finishes.
      await writeWorkbook(directoryHandle, { transactions: nextTransactions, history })
      // Reflect the saved transaction list in the dashboard.
      setTransactions(nextTransactions)
      // Clear a previous folder or workbook error after a successful save.
      setDatabaseMessage('')
      // Tell the caller that persistence completed.
      return true
    // Handle any failure raised by the preceding asynchronous operation.
    } catch (error) {
      // Surface the reason the workbook could not be updated.
      setDatabaseMessage(error instanceof Error ? error.message : 'Unable to save the Excel workbook.')
      // Tell the caller not to clear its form after a failed save.
      return false
    // Close the current callback, block, or value collection.
    }
  // Close the current callback, block, or value collection.
  }

  // Move the active list into the workbook history and clear it after success.
  async function archiveTransactions() {
    // Clear any previous success or error message before starting.
    setArchiveMessage('')
    // Require folder access, a title, and active records before archiving.
    if (!directoryHandle || !historyTitle.trim() || !transactions.length) {
      // Update React state or form behavior for the operation in progress.
      setArchiveMessage('Choose a folder, enter a history title, and add transactions before saving.')
      // Perform this step in the dashboard calculation or state update.
      return
    // Close the current callback, block, or value collection.
    }
    // Remove surrounding whitespace from the group title.
    const title = historyTitle.trim()
    // Compute the totals that will be stored with this archived group.
    const totals = transactions.reduce((summary, transaction) => {
      // Perform this step in the dashboard calculation or state update.
      summary[transaction.type] += transaction.amount
      // Return the result of this function, callback, or calculation.
      return summary
    // Perform this step in the dashboard calculation or state update.
    }, { income: 0, expense: 0 })
    // Record today's date on the summary records for this group.
    const archiveDate = new Date().toISOString().slice(0, 10)
    // Copy active transactions with the archive title attached.
    const archivedTransactions = transactions.map((transaction) => ({ ...transaction, historyTitle: title }))
    // Build three summary rows that summarize the group being archived.
    const summaryTransactions: Transaction[] = [
      // Perform this step in the dashboard calculation or state update.
      { id: Date.now() + 1, title: 'Total income', historyTitle: title, category: 'Summary', amount: totals.income, type: 'income', date: archiveDate },
      // Perform this step in the dashboard calculation or state update.
      { id: Date.now() + 2, title: 'Total expenses', historyTitle: title, category: 'Summary', amount: totals.expense, type: 'expense', date: archiveDate },
      // Perform this step in the dashboard calculation or state update.
      { id: Date.now() + 3, title: 'Available balance', historyTitle: title, category: 'Summary', amount: totals.income - totals.expense, type: 'income', date: archiveDate },
    // Perform this step in the dashboard calculation or state update.
    ]
    // Combine previously archived records with this group's rows and totals.
    const nextHistory = [...history, ...archivedTransactions, ...summaryTransactions]
    // Save the new workbook before clearing the active transaction list.
    try {
      // Wait until this asynchronous save or browser operation finishes.
      await writeWorkbook(directoryHandle, { transactions: [], history: nextHistory })
      // Clear the current ledger after the archive workbook write succeeds.
      setTransactions([])
      // Show the newly archived group in the history panel.
      setHistory(nextHistory)
      // Clear the archive title field after a successful save.
      setHistoryTitle('')
      // Clear any earlier database warning after the write succeeds.
      setDatabaseMessage('')
      // Report how many active transactions were moved into history.
      setArchiveMessage(`${transactions.length} transaction${transactions.length === 1 ? '' : 's'} saved to expense history.`)
    // Handle any failure raised by the preceding asynchronous operation.
    } catch (error) {
      // Keep the active list intact and show the workbook-write failure.
      setArchiveMessage(error instanceof Error ? error.message : 'Unable to save the expense history.')
    // Close the current callback, block, or value collection.
    }
  // Close the current callback, block, or value collection.
  }

  // Open the browser folder picker and load or create the selected workbook.
  async function chooseDatabaseDirectory() {
    // Clear any error from an earlier folder selection.
    setDatabaseMessage('')
    // Disable folder controls while the browser picker and file loading run.
    setIsConnectingDatabase(true)
    // Ask for a folder, load its workbook, and remember it for later.
    try {
      // Compute handle to hold a value used by the surrounding dashboard logic.
      const handle = await openDatabaseDirectoryPicker()
      // Compute data to hold the loaded workbook contents.
      const data = await loadWorkbook(handle)
      // Wait until this asynchronous save or browser operation finishes.
      await saveDirectory(handle)
      // Update React state or form behavior for the operation in progress.
      setSavedDirectoryHandle(handle)
      // Update React state or form behavior for the operation in progress.
      setDirectoryHandle(handle)
      // Update React state or form behavior for the operation in progress.
      setDatabaseDirectory(handle.name)
      // Update React state or form behavior for the operation in progress.
      setTransactions(data.transactions)
      // Update React state or form behavior for the operation in progress.
      setHistory(data.history)
      // Treat a cancelled picker as normal and report other failures.
    } catch (error) {
      // Check this condition before continuing the dashboard operation.
      if (error instanceof DOMException && error.name === 'AbortError') return
      // Update React state or form behavior for the operation in progress.
      setDatabaseMessage(error instanceof Error ? error.message : 'Unable to select the Excel database folder.')
      // Re-enable folder controls after success, cancellation, or failure.
    } finally {
      // Update React state or form behavior for the operation in progress.
      setIsConnectingDatabase(false)
    // Close the current callback, block, or value collection.
    }
  // Close the current callback, block, or value collection.
  }

  // Renew permission for the previously saved folder without opening the picker.
  async function reconnectSavedDirectory() {
    // Do nothing when there is no previously saved directory handle.
    if (!savedDirectoryHandle) return
    // Clear old status and mark the permission renewal as in progress.
    setDatabaseMessage('')
    // Update React state or form behavior for the operation in progress.
    setIsConnectingDatabase(true)
    // Request access again, then reload the workbook using the saved handle.
    try {
      // Compute permission to hold the renewed browser folder permission.
      const permission = await savedDirectoryHandle.requestPermission({ mode: 'readwrite' })
      // Check this condition before continuing the dashboard operation.
      if (permission !== 'granted') {
        // Perform this step in the dashboard calculation or state update.
        throw new Error('Folder access was not granted. Choose the folder to reconnect.')
      // Close the current callback, block, or value collection.
      }
      // Compute data to hold the loaded workbook contents.
      const data = await loadWorkbook(savedDirectoryHandle)
      // Update React state or form behavior for the operation in progress.
      setDirectoryHandle(savedDirectoryHandle)
      // Update React state or form behavior for the operation in progress.
      setTransactions(data.transactions)
      // Update React state or form behavior for the operation in progress.
      setHistory(data.history)
      // Report permission or file errors so the user can recover.
    } catch (error) {
      // Update React state or form behavior for the operation in progress.
      setDatabaseMessage(error instanceof Error ? error.message : 'Unable to reconnect the saved Excel database folder.')
      // Re-enable folder controls regardless of the reconnect result.
    } finally {
      // Update React state or form behavior for the operation in progress.
      setIsConnectingDatabase(false)
    // Close the current callback, block, or value collection.
    }
  // Close the current callback, block, or value collection.
  }

  // Validate and add a transaction submitted from the form.
  async function addTransaction(event: FormEvent<HTMLFormElement>) {
    // Prevent the browser from reloading the page on submit.
    event.preventDefault()
    // Convert the input amount from text into a number.
    const numericAmount = Number(amount)
    // Ignore incomplete or invalid entries.
    if (!title.trim() || !numericAmount || numericAmount < 0) return
    // Add the newest transaction to the selected date group.
    const saved = await save([{ id: Date.now(), title: title.trim(), category, amount: numericAmount, type, date }, ...transactions])
    // Check this condition before continuing the dashboard operation.
    if (!saved) return
    // Clear the text fields after a successful submission.
    setTitle('')
    // Update React state or form behavior for the operation in progress.
    setAmount('')
  // Close the current callback, block, or value collection.
  }

  // Add a new to-do item to the list.
  function addTodo(event: FormEvent<HTMLFormElement>) {
    event.preventDefault()
    const trimmedTitle = todoDraft.trim()
    if (!trimmedTitle || !todoStorageWritable) return
    const saved = saveTodoItems([...todoItems, {
      id: `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      title: trimmedTitle,
      completed: false,
    }])
    if (saved) setTodoDraft('')
  }

  function saveGroceryProducts(products: GroceryProduct[]) {
    try {
      window.localStorage.setItem('budget-tracker-grocery-products', JSON.stringify(products))
      setGroceryState((current) => ({ ...current, products }))
      return true
    } catch (error) {
      setGroceryState((current) => ({
        ...current,
        writable: false,
        message: error instanceof Error ? error.message : 'Unable to save grocery products.',
      }))
      return false
    }
  }

  function saveNonGroceryProducts(products: GroceryProduct[]) {
    try {
      window.localStorage.setItem('budget-tracker-nongrocery-products', JSON.stringify(products))
      setNonGroceryState((current) => ({ ...current, products }))
      return true
    } catch (error) {
      setNonGroceryState((current) => ({
        ...current,
        writable: false,
        message: error instanceof Error ? error.message : 'Unable to save non-grocery products.',
      }))
      return false
    }
  }

  function saveBills(nextBills: BillItem[], storageKey: string, setState: Dispatch<SetStateAction<BillStorageState>>, period: string) {
    try {
      window.localStorage.setItem(storageKey, JSON.stringify(nextBills))
      setState((current) => ({ ...current, bills: nextBills }))
      return true
    } catch (error) {
      setState((current) => ({
        ...current,
        writable: false,
        message: error instanceof Error ? error.message : `Unable to save ${period} bills.`,
      }))
      return false
    }
  }

  function addProduct(products: GroceryProduct[], saveProducts: (products: GroceryProduct[]) => boolean, writable: boolean, name: string) {
    if (!name || !writable) return false
    return saveProducts([...products, {
      id: `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      name,
      stock: 0,
      price: 0,
      toBuy: true,
      bought: false,
    }])
  }

  // Return the dashboard layout and its interactive controls.
  return (
    // Wrap the complete dashboard in its centered page container.
    <main className="app-shell">
      {/* Show the brand, subtitle, and account shortcut. */}
      <header className="topbar"><div className="brand-mark">₱</div><div><strong>Risen7</strong><span>Personal finances</span></div><div className="text-size-controls" role="group" aria-label="Text size"><button type="button" aria-label="Decrease text size" onClick={() => setFontSizeScale((size) => Math.max(0, size - 1))} disabled={fontSizeScale === 0}>A−</button><span aria-live="polite">{fontSizeScale === 0 ? '90%' : fontSizeScale === 1 ? '100%' : '110%'}</span><button type="button" aria-label="Increase text size" onClick={() => setFontSizeScale((size) => Math.min(2, size + 1))} disabled={fontSizeScale === 2}>A+</button></div><button className="theme-toggle" type="button" aria-label={`Switch to ${theme === 'light' ? 'dark' : 'light'} mode`} aria-pressed={theme === 'dark'} onClick={() => setTheme(theme === 'light' ? 'dark' : 'light')}>
        {theme === 'light'
          ? <svg aria-hidden="true" focusable="false" viewBox="0 0 24 24"><path d="M20.2 15.4A8.5 8.5 0 0 1 8.6 3.8 8.5 8.5 0 1 0 20.2 15.4Z" /></svg>
          : <svg aria-hidden="true" focusable="false" viewBox="0 0 24 24"><circle cx="12" cy="12" r="4" /><path d="M12 2v2m0 16v2M4.93 4.93l1.42 1.42m11.3 11.3 1.42 1.42M2 12h2m16 0h2M4.93 19.07l1.42-1.42m11.3-11.3 1.42-1.42" /></svg>}
      </button><button className="avatar" type="button" aria-label="Account menu">JD</button></header>
      <nav className="main-tabs" role="tablist" aria-label="Main sections">
        <button id="dashboard-tab" className={activeTab === 'dashboard' ? 'main-tab active' : 'main-tab'} type="button" role="tab" aria-selected={activeTab === 'dashboard'} aria-controls="dashboard-content" onClick={() => setActiveTab('dashboard')}>Dashboard</button>
        <button id="budgets-tab" className={activeTab === 'budgets' ? 'main-tab active' : 'main-tab'} type="button" role="tab" aria-selected={activeTab === 'budgets'} aria-controls="budgets-content" onClick={() => setActiveTab('budgets')}>Budgets</button>
      </nav>
      {activeTab === 'dashboard' ? <div id="dashboard-content" role="tabpanel" aria-labelledby="dashboard-tab">
      {/* Show the current day and selected reporting month. */}
      <section className="welcome-row"><p className="dashboard-date">{formatCurrentDate()}</p><div className="welcome-copy"><h1>Good morning</h1><p className="muted">Here is your financial snapshot for this month.</p></div><div className="month-chip">September 2026 <span>⌄</span></div></section>
      {/* Display the calculated balance, income, and expense totals. */}
      <section className="stats-grid" aria-label="Financial summary"><article className="stat-card balance"><div className="stat-label">Available balance <span className="info">i</span></div><strong>{currency.format(totals.income - totals.expense)}</strong><div className="trend positive">↗ 8.4% <small>vs last month</small></div><div className="balance-bar"><span style={{ width: `${Math.min(100, (totals.expense / totals.income) * 100)}%` }} /></div></article><article className="stat-card"><div className="stat-label"><span className="dot income-dot" />Total income</div><strong>{currency.format(totals.income)}</strong><div className="trend positive">↗ 12.6% <small>vs last month</small></div></article><article className="stat-card"><div className="stat-label"><span className="dot expense-dot" />Total expenses</div><strong>{currency.format(totals.expense)}</strong><div className="trend negative">↘ 3.2% <small>vs last month</small></div></article></section>
      {/* Place the daily activity list beside the entry form. */}
      <div className="content-grid">
        <section className="transactions-panel">
          <div className="weekly-panel">
            <div className="section-heading">
              <div><h2>15-day expenditure</h2><p className="muted">Your spend across the last 15 days</p></div>
              <strong className="weekly-total" aria-label={`Today's expenses: ${currency.format(todayExpense)}`}>{currency.format(todayExpense)}</strong>
            </div>
            <div className="weekly-chart" aria-label="15-day expenditure chart">
              <svg viewBox="0 0 370 160" preserveAspectRatio="none" role="img" aria-label="15-day expenditure bar chart">
                <g>{chartBars.map((bar) => <g key={bar.key} className="chart-bar-group" role="button" tabIndex={0} aria-label={`${new Date(`${bar.key}T12:00:00`).toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' })}: ${currency.format(bar.total)} expenses. Show recent activity.`} onClick={() => showActivityForDate(bar.key)} onKeyDown={(event) => {
                  if (event.key === 'Enter' || event.key === ' ') {
                    event.preventDefault()
                    showActivityForDate(bar.key)
                  }
                }}><rect className="chart-bar" x={bar.x} y={bar.y} width={bar.width} height={bar.height} rx="4" /><text className="chart-value" x={bar.x + bar.width / 2} y={bar.y - 8} textAnchor="middle">{currency.format(bar.total)}</text><text className="chart-label" x={bar.x + bar.width / 2} y="150" textAnchor="middle">{bar.label}</text></g>)}</g>
              </svg>
            </div>
          </div>
          <div className="section-heading">
            <div id="recent-activity-heading"><h2>Recent activity</h2><p className="muted">Your latest income and expenses</p></div>
            <select className="filter-button activity-date-filter" aria-label="Filter recent activity by date" value={activityDate} onChange={(event) => setActivityDate(event.target.value)}>
              <option value="all">All dates</option>
              {activityDates.map((date) => <option key={date} value={date}>{new Date(`${date}T12:00:00`).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' })}</option>)}
            </select>
          </div>
          <div className="transaction-list">
            {Object.entries(groupedTransactions).map(([date, entries]) => {
              const isCollapsed = collapsedActivityGroups[date] ?? date !== today
              const activityId = `activity-${date}`
              return (
                <div className="date-group" id={`activity-day-${date}`} key={date}>
                  <div className="activity-day-heading">
                    <div className="date-label">{formatDate(date)} <span>{new Date(`${date}T12:00:00`).toLocaleDateString('en-US', { weekday: 'short' })}</span></div>
                    <button className="collapse-toggle activity-day-toggle" type="button" aria-label={`${isCollapsed ? 'Expand' : 'Collapse'} activity for ${date}`} aria-expanded={!isCollapsed} aria-controls={activityId} onClick={() => setCollapsedActivityGroups((current) => ({ ...current, [date]: !isCollapsed }))}>
                      {isCollapsed ? 'Expand' : 'Collapse'}
                    </button>
                  </div>
                  <div id={activityId} hidden={isCollapsed}>{entries.map((transaction) => <div className="transaction" key={transaction.id}><div className={`transaction-icon ${transaction.type}`}>{transaction.type === 'income' ? '↙' : '↗'}</div><div className="transaction-info"><strong>{transaction.title}</strong><span>{transaction.category}</span></div><div className={transaction.type === 'income' ? 'amount income-amount' : 'amount'}>{transaction.type === 'income' ? '+' : '-'}{currency.format(transaction.amount)}</div><button className="delete-button" type="button" aria-label={`Delete ${transaction.title}`} onClick={() => save(transactions.filter((item) => item.id !== transaction.id))}>×</button></div>)}</div>
                </div>
              )
            })}
          </div>
          {Object.keys(groupedTransactions).length === 0 && <p className="muted">No transactions for this date.</p>}
          <label className="history-title-field">History title<input value={historyTitle} onChange={(event) => setHistoryTitle(event.target.value)} placeholder="e.g. September expenses" required /></label>
          <button className="export-button" type="button" onClick={archiveTransactions} disabled={!transactions.length || !historyTitle.trim()}>Save to expense history</button>
          {archiveMessage && <p className="database-note">{archiveMessage}</p>}
        </section>
        {/* Provide controls for adding income and expense entries. */}
        <div className="right-column-stack">
        <aside className="add-panel"><div className="section-heading"><div><h2>Add transaction</h2><p className="muted">Keep your ledger up to date</p></div><span className="plus-icon">+</span></div><form onSubmit={addTransaction}><div className="type-toggle"><button type="button" className={type === 'expense' ? 'selected expense-selected' : ''} onClick={() => setType('expense')}>Expense</button><button type="button" className={type === 'income' ? 'selected income-selected' : ''} onClick={() => setType('income')}>Income</button></div><label>Description<input value={title} onChange={(event) => setTitle(event.target.value)} placeholder="e.g. Coffee with friends" required /></label><div className="form-row"><label>Category<select value={category} onChange={(event) => setCategory(event.target.value)}><option>Food</option><option>Housing</option><option>Bills</option><option>Transport</option><option>Gala/Mall</option><option>Salary</option><option>Side income</option><option>Other</option></select></label><label>Amount<input type="number" min="0.01" step="0.01" value={amount} onChange={(event) => setAmount(event.target.value)} placeholder="₱ 0.00" required /></label></div><label>Date<input type="date" value={date} onChange={(event) => setDate(event.target.value)} required /></label><button className="add-button" type="submit">Add {type}</button></form><div className="database-location"><p>Excel database folder</p><code title={databaseDirectory}>{databaseDirectory || 'No folder selected'}</code>{supportsFolderSelection() ? <>{savedDirectoryHandle && !directoryHandle && <button className="folder-button" type="button" onClick={reconnectSavedDirectory} disabled={isConnectingDatabase}>{isConnectingDatabase ? 'Reconnecting…' : 'Reconnect saved folder'}</button>}<button className="folder-button" type="button" onClick={chooseDatabaseDirectory} disabled={isConnectingDatabase}>{isConnectingDatabase ? 'Connecting…' : 'Choose folder'}</button></> : <p className="database-error" role="alert">Folder access needs Chrome or Edge on a secure connection.</p>}<p className="database-note">The browser remembers the folder, but may ask you to reconnect if permission expires.</p>{databaseMessage && <p className="database-error" role="alert">{databaseMessage}</p>}</div></aside>
        <section className="todo-panel" aria-labelledby="todo-heading">
        <div className="section-heading">
          <div><h2 id="todo-heading">To-do list</h2><p className="muted">{todoItems.filter((item) => !item.completed).length} tasks remaining</p></div>
        </div>
        <form className="todo-form" onSubmit={addTodo}>
          <input aria-label="New to-do item" value={todoDraft} onChange={(event) => setTodoDraft(event.target.value)} placeholder="Add a task…" required disabled={!todoStorageWritable} />
          <button className="add-button" type="submit" disabled={!todoStorageWritable || !todoDraft.trim()}>Add task</button>
        </form>
        {todoMessage && <p className="todo-error" role="alert">{todoMessage}</p>}
        {!todoItems.length && <p className="todo-empty muted">No tasks yet. Add one above.</p>}
        <ul className="todo-list">
          {todoItems.map((item) => (
            <li className={item.completed ? 'todo-item completed' : 'todo-item'} key={item.id}>
              <label><input type="checkbox" checked={item.completed} disabled={!todoStorageWritable} onChange={() => saveTodoItems(todoItems.map((todo) => todo.id === item.id ? { ...todo, completed: !todo.completed } : todo))} /><span>{item.title}</span></label>
              <button className="todo-delete" type="button" aria-label={`Delete ${item.title}`} disabled={!todoStorageWritable} onClick={() => saveTodoItems(todoItems.filter((todo) => todo.id !== item.id))}>×</button>
            </li>
          ))}
        </ul>
        </section>
        </div></div>
      {/* Show each archived group, its saved totals, and its expandable transaction list. */}
      <section className="history-panel"><div className="section-heading"><div><h2>Expense history</h2><p className="muted">Transactions saved from recent activity</p></div><span className="history-count">{history.filter((transaction) => transaction.category !== 'Summary').length}</span></div>{summaryCards.length ? summaryCards.map((summary) => { const isCollapsed = collapsedGroups[summary.title] ?? true; return <div className="weekly-panel" key={summary.title}><div className="section-heading collapsible-header"><div><h2>{summary.title}</h2><p className="muted">Saved totals for this archive</p></div><button className="collapse-toggle" type="button" onClick={() => setCollapsedGroups((current) => ({ ...current, [summary.title]: !isCollapsed }))}>{isCollapsed ? 'Expand' : 'Collapse'}</button></div><div className="stats-grid"><article className="stat-card"><div className="stat-label"><span className="dot income-dot" />Total income</div><strong>{currency.format(summary.income)}</strong></article><article className="stat-card"><div className="stat-label"><span className="dot expense-dot" />Total expenses</div><strong>{currency.format(summary.expense)}</strong></article><article className="stat-card balance"><div className="stat-label">Available balance <span className="info">i</span></div><strong>{currency.format(summary.balance)}</strong></article></div>{!isCollapsed && summary.entries?.length ? <div className="history-list"><div className="history-group"><h3>Saved transactions</h3>{summary.entries.map((transaction) => <div className="transaction" key={`${transaction.id}-${transaction.date}-${transaction.historyTitle}`}><div className={`transaction-icon ${transaction.type}`}>{transaction.type === 'income' ? '↙' : '↗'}</div><div className="transaction-info"><strong>{transaction.title}</strong><span>{formatDate(transaction.date)} · {transaction.category}</span></div><div className={transaction.type === 'income' ? 'amount income-amount' : 'amount'}>{transaction.type === 'income' ? '+' : '-'}{currency.format(transaction.amount)}</div></div>)}</div></div> : null}</div> }) : <p className="muted">No saved transactions yet.</p>}</section>
      </div> : <section id="budgets-content" className="budgets-page" role="tabpanel" aria-labelledby="budgets-tab">
        <div className="budgets-heading">
          <div><p className="eyebrow">Plan your spending</p><h1>Budgets</h1><p className="muted">Keep everyday spending and upcoming plans organized.</p></div>
          <span className="budgets-total-label">7 budget categories</span>
        </div>
        <div className="budget-cards">
          <ProductListCard
            id="grocery-list"
            title="Grocery list"
            description="Stock and shopping priorities"
            icon="🛒"
            products={groceryProducts}
            writable={groceryStorageWritable}
            message={groceryMessage}
            onSave={saveGroceryProducts}
            onAdd={(name) => addProduct(groceryProducts, saveGroceryProducts, groceryStorageWritable, name)}
          />
          <ProductListCard
            id="non-grocery-list"
            title="Not-grocery list"
            description="Everyday spending"
            icon="🧺"
            products={nonGroceryProducts}
            writable={nonGroceryStorageWritable}
            message={nonGroceryMessage}
            onSave={saveNonGroceryProducts}
            onAdd={(name) => addProduct(nonGroceryProducts, saveNonGroceryProducts, nonGroceryStorageWritable, name)}
          />
          <BillsCard period="Day 1–15" bills={firstBills} writable={firstBillsWritable} message={firstBillsMessage} onSave={(nextBills) => saveBills(nextBills, 'budget-tracker-day-1-15-bills', setFirstBillState, 'Day 1–15')} />
          <BillsCard period="Day 16–30" bills={secondBills} writable={secondBillsWritable} message={secondBillsMessage} onSave={(nextBills) => saveBills(nextBills, 'budget-tracker-day-16-30-bills', setSecondBillState, 'Day 16–30')} />
          <article className="budget-card family"><span className="budget-card-icon" aria-hidden="true">💛</span><div><p className="budget-card-kicker">Make room for memories</p><h2>Family bonding</h2><p>Set aside room in your plan for activities with family.</p></div><span className="budget-card-arrow" aria-hidden="true">↗</span></article>
          <article className="budget-card allowance"><span className="budget-card-icon" aria-hidden="true">💵</span><div><p className="budget-card-kicker">Payday planning</p><h2>Cut-off allowance</h2><p>Organize allowance for each pay period or cut-off.</p></div><span className="budget-card-arrow" aria-hidden="true">↗</span></article>
          <article className="budget-card seldom"><span className="budget-card-icon" aria-hidden="true">✨</span><div><p className="budget-card-kicker">Occasional purchases</p><h2>Seldom buys</h2><p>Track less frequent buys before they catch you by surprise.</p></div><span className="budget-card-arrow" aria-hidden="true">↗</span></article>
        </div>
      </section>}
    </main>
  )
}

// Export the dashboard so the application entry point can render it.
export default App
