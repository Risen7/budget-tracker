// Import React state and derived-state hooks.
import { useEffect, useMemo, useState } from 'react'
// Import the form event type without adding it to the runtime bundle.
import type { FormEvent } from 'react'
// Load the dashboard stylesheet.
import './App.css'

// Restrict transactions to the two supported financial directions.
type TransactionType = 'income' | 'expense'
// Describe every value stored for a transaction.
type Transaction = {
  id: number
  title: string
  historyTitle?: string
  category: string
  amount: number
  type: TransactionType
  date: string
}

// Provide useful sample data when the tracker has no saved entries yet.
const initialTransactions: Transaction[] = [
  { id: 1, title: 'Monthly salary', category: 'Salary', amount: 3200, type: 'income', date: '2026-09-20' },
  { id: 2, title: 'Apartment rent', category: 'Housing', amount: 1150, type: 'expense', date: '2026-09-20' },
  { id: 3, title: 'Grocery run', category: 'Food', amount: 86.45, type: 'expense', date: '2026-09-19' },
  { id: 4, title: 'Freelance project', category: 'Side income', amount: 450, type: 'income', date: '2026-09-18' },
  { id: 5, title: 'Train pass', category: 'Transport', amount: 42, type: 'expense', date: '2026-09-18' },
]

// Format all displayed amounts consistently as Philippine pesos.
const currency = new Intl.NumberFormat('en-PH', { style: 'currency', currency: 'PHP' })

const today = new Date().toISOString().slice(0, 10)

// Convert stored ISO dates into friendly labels for the activity list.
function formatDate(date: string) {
  // Label the current sample day as Today.
  if (date === '2026-09-20') return 'Today'
  // Label the previous sample day as Yesterday.
  if (date === '2026-09-19') return 'Yesterday'
  // Format all older dates with a short month and day.
  return new Date(`${date}T12:00:00`).toLocaleDateString('en-US', { month: 'short', day: 'numeric' })
}

// Render the complete budget dashboard.
function App() {
  // Hold the transactions loaded from the Excel database API.
  const [transactions, setTransactions] = useState<Transaction[]>(initialTransactions)
  // Hold transactions that have been moved into expense history.
  const [history, setHistory] = useState<Transaction[]>([])
  // Track whether the new entry is income or an expense.
  const [type, setType] = useState<TransactionType>('expense')
  // Track the new transaction description field.
  const [title, setTitle] = useState('')
  // Track the selected transaction category.
  const [category, setCategory] = useState('Food')
  // Track the new amount as text while the user types.
  const [amount, setAmount] = useState('')
  // Track the date selected for the new transaction.
  const [date, setDate] = useState(today)
  // Show the result of archiving the active transaction list.
  const [archiveMessage, setArchiveMessage] = useState('')
  // Track the title used to group an archived batch.
  const [historyTitle, setHistoryTitle] = useState('')
  // Track which saved history groups are expanded or collapsed.
  const [collapsedGroups, setCollapsedGroups] = useState<Record<string, boolean>>({})

  // Load the current transaction table from the Excel database when the app starts.
  useEffect(() => {
    // Request the workbook data through the local API.
    fetch('/api/transactions')
      .then((response) => response.json())
      .then((savedTransactions: Transaction[]) => setTransactions(savedTransactions))
      .catch(() => setTransactions(initialTransactions))
    fetch('/api/transactions/history')
      .then((response) => response.json())
      .then((savedHistory: Transaction[]) => setHistory(savedHistory))
  }, [])

  // Calculate total income and expenses whenever the transaction list changes.
  const totals = useMemo(() => transactions.reduce((summary, transaction) => {
    // Add each amount to its matching income or expense total.
    summary[transaction.type] += transaction.amount
    // Return the running summary for the next transaction.
    return summary
  }, { income: 0, expense: 0 }), [transactions])

  // Summarize the last 15 days so the dashboard can chart recent spending.
  const weeklyExpenditure = useMemo(() => {
    const days = Array.from({ length: 15 }, (_, index) => {
      const date = new Date()
      date.setHours(12, 0, 0, 0)
      date.setDate(date.getDate() - (14 - index))
      const isoDate = date.toISOString().slice(0, 10)
      const total = transactions
        .filter((transaction) => transaction.type === 'expense' && transaction.date === isoDate)
        .reduce((sum, transaction) => sum + transaction.amount, 0)

      return {
        key: isoDate,
        label: String(date.getDate()),
        total,
      }
    })

    return days
  }, [transactions])

  const chartMax = Math.max(...weeklyExpenditure.map((item) => item.total), 1)
  const chartBars = weeklyExpenditure.map((day, index) => {
    const x = 20 + index * 22
    const width = 12
    const height = (day.total / chartMax) * 88
    const y = 130 - height

    return { ...day, x, y, width, height }
  })

  // Organize entries by date so the activity list can show daily groups.
  const groupedTransactions = useMemo(() => transactions.reduce<Record<string, Transaction[]>>((groups, transaction) => {
    // Append the current transaction to its date group.
    groups[transaction.date] = [...(groups[transaction.date] ?? []), transaction]
    // Return the running date groups for the next transaction.
    return groups
  }, {}), [transactions])

  const summaryCards = useMemo(() => {
    const summaries = history.filter((transaction) => transaction.category === 'Summary')
    const grouped = new Map<string, { income: number, expense: number, balance: number, entries: Transaction[] }>()

    summaries.forEach((transaction) => {
      const key = transaction.historyTitle ?? 'Summary'
      const current = grouped.get(key) ?? { income: 0, expense: 0, balance: 0, entries: [] }

      if (transaction.title === 'Total income') current.income = transaction.amount
      if (transaction.title === 'Total expenses') current.expense = transaction.amount
      if (transaction.title === 'Available balance') current.balance = transaction.amount

      grouped.set(key, current)
    })

    history.filter((transaction) => transaction.category !== 'Summary').forEach((transaction) => {
      const key = transaction.historyTitle || `Saved ${formatDate(transaction.date)}`
      const current = grouped.get(key) ?? { income: 0, expense: 0, balance: 0, entries: [] }
      current.entries = [...current.entries, transaction]
      grouped.set(key, current)
    })

    return Array.from(grouped.entries()).map(([title, values]) => ({ title, ...values }))
  }, [history])

  // Update the screen and persist the new list in the Excel database.
  async function save(nextTransactions: Transaction[]) {
    // Refresh the visible dashboard immediately.
    setTransactions(nextTransactions)
    // Send the complete transaction table to the local Excel database API.
    await fetch('/api/transactions', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(nextTransactions) })
  }

  // Move the active list into the workbook history and clear it after success.
  async function archiveTransactions() {
    setArchiveMessage('')
    const response = await fetch('/api/transactions/archive', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ title: historyTitle }) })
    const result = await response.json()
    if (!response.ok) {
      setArchiveMessage(result.error ?? 'Unable to save the expense history.')
      return
    }
    setTransactions([])
    setHistory((currentHistory) => [...currentHistory, ...transactions])
    setHistoryTitle('')
    setArchiveMessage(`${result.saved} transaction${result.saved === 1 ? '' : 's'} saved to expense history.`)
  }

  // Validate and add a transaction submitted from the form.
  function addTransaction(event: FormEvent<HTMLFormElement>) {
    // Prevent the browser from reloading the page on submit.
    event.preventDefault()
    // Convert the input amount from text into a number.
    const numericAmount = Number(amount)
    // Ignore incomplete or invalid entries.
    if (!title.trim() || !numericAmount || numericAmount < 0) return
    // Add the newest transaction to the selected date group.
    save([{ id: Date.now(), title: title.trim(), category, amount: numericAmount, type, date }, ...transactions])
    // Clear the text fields after a successful submission.
    setTitle('')
    setAmount('')
  }

  // Return the dashboard layout and its interactive controls.
  return (
    <main className="app-shell">
      {/* Show the brand, subtitle, and account shortcut. */}
      <header className="topbar"><div className="brand-mark">₱</div><div><strong>Risen7</strong><span>Personal finances</span></div><button className="avatar" type="button" aria-label="Account menu">JD</button></header>
      {/* Show the current day and selected reporting month. */}
      <section className="welcome-row"><div><p className="eyebrow">Sunday, September 20, 2026</p><h1>Good morning</h1><p className="muted">Here is your financial snapshot for this month.</p></div><div className="month-chip">September 2026 <span>⌄</span></div></section>
      {/* Display the calculated balance, income, and expense totals. */}
      <section className="stats-grid" aria-label="Financial summary"><article className="stat-card balance"><div className="stat-label">Available balance <span className="info">i</span></div><strong>{currency.format(totals.income - totals.expense)}</strong><div className="trend positive">↗ 8.4% <small>vs last month</small></div><div className="balance-bar"><span style={{ width: `${Math.min(100, (totals.expense / totals.income) * 100)}%` }} /></div></article><article className="stat-card"><div className="stat-label"><span className="dot income-dot" />Total income</div><strong>{currency.format(totals.income)}</strong><div className="trend positive">↗ 12.6% <small>vs last month</small></div></article><article className="stat-card"><div className="stat-label"><span className="dot expense-dot" />Total expenses</div><strong>{currency.format(totals.expense)}</strong><div className="trend negative">↘ 3.2% <small>vs last month</small></div></article></section>
      {/* Place the daily activity list beside the entry form. */}
      <div className="content-grid"><section className="transactions-panel"><div className="weekly-panel"><div className="section-heading"><div><h2>15-day expenditure</h2><p className="muted">Your spend across the last 15 days</p></div><strong className="weekly-total">{currency.format(Math.max(...weeklyExpenditure.map((day) => day.total), 0))}</strong></div><div className="weekly-chart" aria-label="15-day expenditure chart"><svg viewBox="0 0 370 160" preserveAspectRatio="none" role="img" aria-label="15-day expenditure bar chart"><g>{chartBars.map((bar) => <g key={bar.key} className="chart-bar-group"><rect className="chart-bar" x={bar.x} y={bar.y} width={bar.width} height={bar.height} rx="4" /><text className="chart-value" x={bar.x + bar.width / 2} y={bar.y - 8} textAnchor="middle">{currency.format(bar.total)}</text><text className="chart-label" x={bar.x + bar.width / 2} y="150" textAnchor="middle">{bar.label}</text></g>)}</g></svg></div></div><div className="section-heading"><div><h2>Recent activity</h2><p className="muted">Your latest income and expenses</p></div><button className="filter-button" type="button">All activity <span>⌄</span></button></div><div className="transaction-list">{Object.entries(groupedTransactions).map(([date, entries]) => <div className="date-group" key={date}><div className="date-label">{formatDate(date)} <span>{new Date(`${date}T12:00:00`).toLocaleDateString('en-US', { weekday: 'short' })}</span></div>{entries.map((transaction) => <div className="transaction" key={transaction.id}><div className={`transaction-icon ${transaction.type}`}>{transaction.type === 'income' ? '↙' : '↗'}</div><div className="transaction-info"><strong>{transaction.title}</strong><span>{transaction.category}</span></div><div className={transaction.type === 'income' ? 'amount income-amount' : 'amount'}>{transaction.type === 'income' ? '+' : '-'}{currency.format(transaction.amount)}</div><button className="delete-button" type="button" aria-label={`Delete ${transaction.title}`} onClick={() => save(transactions.filter((item) => item.id !== transaction.id))}>×</button></div>)}</div>)}</div><label className="history-title-field">History title<input value={historyTitle} onChange={(event) => setHistoryTitle(event.target.value)} placeholder="e.g. September expenses" required /></label><button className="export-button" type="button" onClick={archiveTransactions} disabled={!transactions.length || !historyTitle.trim()}>Save to expense history</button>{archiveMessage && <p className="database-note">{archiveMessage}</p>}</section>
        {/* Provide controls for adding income and expense entries. */}
        <aside className="add-panel"><div className="section-heading"><div><h2>Add transaction</h2><p className="muted">Keep your ledger up to date</p></div><span className="plus-icon">+</span></div><form onSubmit={addTransaction}><div className="type-toggle"><button type="button" className={type === 'expense' ? 'selected expense-selected' : ''} onClick={() => setType('expense')}>Expense</button><button type="button" className={type === 'income' ? 'selected income-selected' : ''} onClick={() => setType('income')}>Income</button></div><label>Description<input value={title} onChange={(event) => setTitle(event.target.value)} placeholder="e.g. Coffee with friends" required /></label><div className="form-row"><label>Category<select value={category} onChange={(event) => setCategory(event.target.value)}><option>Food</option><option>Housing</option><option>Transport</option><option>Shopping</option><option>Salary</option><option>Side income</option><option>Other</option></select></label><label>Amount<input type="number" min="0.01" step="0.01" value={amount} onChange={(event) => setAmount(event.target.value)} placeholder="₱ 0.00" required /></label></div><label>Date<input type="date" value={date} onChange={(event) => setDate(event.target.value)} required /></label><button className="add-button" type="submit">Add {type}</button></form><p className="database-note">Transactions are stored in budget-database.xlsx</p></aside></div>
      <section className="history-panel"><div className="section-heading"><div><h2>Expense history</h2><p className="muted">Transactions saved from recent activity</p></div><span className="history-count">{history.filter((transaction) => transaction.category !== 'Summary').length}</span></div>{summaryCards.length ? summaryCards.map((summary) => { const isCollapsed = collapsedGroups[summary.title] ?? true; return <div className="weekly-panel" key={summary.title}><div className="section-heading collapsible-header"><div><h2>{summary.title}</h2><p className="muted">Saved totals for this archive</p></div><button className="collapse-toggle" type="button" onClick={() => setCollapsedGroups((current) => ({ ...current, [summary.title]: !isCollapsed }))}>{isCollapsed ? 'Expand' : 'Collapse'}</button></div><div className="stats-grid"><article className="stat-card"><div className="stat-label"><span className="dot income-dot" />Total income</div><strong>{currency.format(summary.income)}</strong></article><article className="stat-card"><div className="stat-label"><span className="dot expense-dot" />Total expenses</div><strong>{currency.format(summary.expense)}</strong></article><article className="stat-card balance"><div className="stat-label">Available balance <span className="info">i</span></div><strong>{currency.format(summary.balance)}</strong></article></div>{!isCollapsed && summary.entries?.length ? <div className="history-list"><div className="history-group"><h3>Saved transactions</h3>{summary.entries.map((transaction) => <div className="transaction" key={`${transaction.id}-${transaction.date}-${transaction.historyTitle}`}><div className={`transaction-icon ${transaction.type}`}>{transaction.type === 'income' ? '↙' : '↗'}</div><div className="transaction-info"><strong>{transaction.title}</strong><span>{formatDate(transaction.date)} · {transaction.category}</span></div><div className={transaction.type === 'income' ? 'amount income-amount' : 'amount'}>{transaction.type === 'income' ? '+' : '-'}{currency.format(transaction.amount)}</div></div>)}</div></div> : null}</div> }) : <p className="muted">No saved transactions yet.</p>}</section>
    </main>
  )
}

// Export the dashboard so the application entry point can render it.
export default App
