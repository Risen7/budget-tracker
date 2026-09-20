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

// Format all displayed amounts consistently as US dollars.
const currency = new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' })

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
  // Track whether the new entry is income or an expense.
  const [type, setType] = useState<TransactionType>('expense')
  // Track the new transaction description field.
  const [title, setTitle] = useState('')
  // Track the selected transaction category.
  const [category, setCategory] = useState('Food')
  // Track the new amount as text while the user types.
  const [amount, setAmount] = useState('')

  // Load the current transaction table from the Excel database when the app starts.
  useEffect(() => {
    // Request the workbook data through the local API.
    fetch('/api/transactions')
      .then((response) => response.json())
      .then((savedTransactions: Transaction[]) => setTransactions(savedTransactions))
      .catch(() => setTransactions(initialTransactions))
  }, [])

  // Calculate total income and expenses whenever the transaction list changes.
  const totals = useMemo(() => transactions.reduce((summary, transaction) => {
    // Add each amount to its matching income or expense total.
    summary[transaction.type] += transaction.amount
    // Return the running summary for the next transaction.
    return summary
  }, { income: 0, expense: 0 }), [transactions])

  // Organize entries by date so the activity list can show daily groups.
  const groupedTransactions = useMemo(() => transactions.reduce<Record<string, Transaction[]>>((groups, transaction) => {
    // Append the current transaction to its date group.
    groups[transaction.date] = [...(groups[transaction.date] ?? []), transaction]
    // Return the running date groups for the next transaction.
    return groups
  }, {}), [transactions])

  // Update the screen and persist the new list in the Excel database.
  async function save(nextTransactions: Transaction[]) {
    // Refresh the visible dashboard immediately.
    setTransactions(nextTransactions)
    // Send the complete transaction table to the local Excel database API.
    await fetch('/api/transactions', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(nextTransactions) })
  }

  // Validate and add a transaction submitted from the form.
  function addTransaction(event: FormEvent<HTMLFormElement>) {
    // Prevent the browser from reloading the page on submit.
    event.preventDefault()
    // Convert the input amount from text into a number.
    const numericAmount = Number(amount)
    // Ignore incomplete or invalid entries.
    if (!title.trim() || !numericAmount || numericAmount < 0) return
    // Add the newest transaction at the top of today's entries.
    save([{ id: Date.now(), title: title.trim(), category, amount: numericAmount, type, date: '2026-09-20' }, ...transactions])
    // Clear the text fields after a successful submission.
    setTitle('')
    setAmount('')
  }

  // Return the dashboard layout and its interactive controls.
  return (
    <main className="app-shell">
      {/* Show the brand, subtitle, and account shortcut. */}
      <header className="topbar"><div className="brand-mark">$</div><div><strong>Risen7</strong><span>Personal finances</span></div><button className="avatar" type="button" aria-label="Account menu">JD</button></header>
      {/* Show the current day and selected reporting month. */}
      <section className="welcome-row"><div><p className="eyebrow">Sunday, September 20, 2026</p><h1>Good morning</h1><p className="muted">Here is your financial snapshot for this month.</p></div><div className="month-chip">September 2026 <span>⌄</span></div></section>
      {/* Display the calculated balance, income, and expense totals. */}
      <section className="stats-grid" aria-label="Financial summary"><article className="stat-card balance"><div className="stat-label">Available balance <span className="info">i</span></div><strong>{currency.format(totals.income - totals.expense)}</strong><div className="trend positive">↗ 8.4% <small>vs last month</small></div><div className="balance-bar"><span style={{ width: `${Math.min(100, (totals.expense / totals.income) * 100)}%` }} /></div></article><article className="stat-card"><div className="stat-label"><span className="dot income-dot" />Total income</div><strong>{currency.format(totals.income)}</strong><div className="trend positive">↗ 12.6% <small>vs last month</small></div></article><article className="stat-card"><div className="stat-label"><span className="dot expense-dot" />Total expenses</div><strong>{currency.format(totals.expense)}</strong><div className="trend negative">↘ 3.2% <small>vs last month</small></div></article></section>
      {/* Place the daily activity list beside the entry form. */}
      <div className="content-grid"><section className="transactions-panel"><div className="section-heading"><div><h2>Recent activity</h2><p className="muted">Your latest income and expenses</p></div><button className="filter-button" type="button">All activity <span>⌄</span></button></div><div className="transaction-list">{Object.entries(groupedTransactions).map(([date, entries]) => <div className="date-group" key={date}><div className="date-label">{formatDate(date)} <span>{new Date(`${date}T12:00:00`).toLocaleDateString('en-US', { weekday: 'short' })}</span></div>{entries.map((transaction) => <div className="transaction" key={transaction.id}><div className={`transaction-icon ${transaction.type}`}>{transaction.type === 'income' ? '↙' : '↗'}</div><div className="transaction-info"><strong>{transaction.title}</strong><span>{transaction.category}</span></div><div className={transaction.type === 'income' ? 'amount income-amount' : 'amount'}>{transaction.type === 'income' ? '+' : '-'}{currency.format(transaction.amount)}</div><button className="delete-button" type="button" aria-label={`Delete ${transaction.title}`} onClick={() => save(transactions.filter((item) => item.id !== transaction.id))}>×</button></div>)}</div>)}</div></section>
        {/* Provide controls for adding income and expense entries. */}
        <aside className="add-panel"><div className="section-heading"><div><h2>Add transaction</h2><p className="muted">Keep your ledger up to date</p></div><span className="plus-icon">+</span></div><form onSubmit={addTransaction}><div className="type-toggle"><button type="button" className={type === 'expense' ? 'selected expense-selected' : ''} onClick={() => setType('expense')}>Expense</button><button type="button" className={type === 'income' ? 'selected income-selected' : ''} onClick={() => setType('income')}>Income</button></div><label>Description<input value={title} onChange={(event) => setTitle(event.target.value)} placeholder="e.g. Coffee with friends" required /></label><div className="form-row"><label>Category<select value={category} onChange={(event) => setCategory(event.target.value)}><option>Food</option><option>Housing</option><option>Transport</option><option>Shopping</option><option>Salary</option><option>Side income</option><option>Other</option></select></label><label>Amount<input type="number" min="0.01" step="0.01" value={amount} onChange={(event) => setAmount(event.target.value)} placeholder="$ 0.00" required /></label></div><button className="add-button" type="submit">Add {type}</button></form><p className="database-note">Transactions are stored in budget-database.xlsx</p></aside></div>
    </main>
  )
}

// Export the dashboard so the application entry point can render it.
export default App
