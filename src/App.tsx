import { useMemo, useState } from 'react'
import type { FormEvent } from 'react'
import './App.css'

type TransactionType = 'income' | 'expense'
type Transaction = { id: number; title: string; category: string; amount: number; type: TransactionType; date: string }
const initialTransactions: Transaction[] = [
  { id: 1, title: 'Monthly salary', category: 'Salary', amount: 3200, type: 'income', date: '2026-09-20' },
  { id: 2, title: 'Apartment rent', category: 'Housing', amount: 1150, type: 'expense', date: '2026-09-20' },
  { id: 3, title: 'Grocery run', category: 'Food', amount: 86.45, type: 'expense', date: '2026-09-19' },
  { id: 4, title: 'Freelance project', category: 'Side income', amount: 450, type: 'income', date: '2026-09-18' },
  { id: 5, title: 'Train pass', category: 'Transport', amount: 42, type: 'expense', date: '2026-09-18' },
]
const currency = new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' })
function formatDate(date: string) { if (date === '2026-09-20') return 'Today'; if (date === '2026-09-19') return 'Yesterday'; return new Date(`${date}T12:00:00`).toLocaleDateString('en-US', { month: 'short', day: 'numeric' }) }

function App() {
  const [transactions, setTransactions] = useState<Transaction[]>(() => { const saved = localStorage.getItem('budget-transactions'); return saved ? JSON.parse(saved) : initialTransactions })
  const [type, setType] = useState<TransactionType>('expense')
  const [title, setTitle] = useState('')
  const [category, setCategory] = useState('Food')
  const [amount, setAmount] = useState('')
  const totals = useMemo(() => transactions.reduce((summary, transaction) => { summary[transaction.type] += transaction.amount; return summary }, { income: 0, expense: 0 }), [transactions])
  const groupedTransactions = useMemo(() => transactions.reduce<Record<string, Transaction[]>>((groups, transaction) => { groups[transaction.date] = [...(groups[transaction.date] ?? []), transaction]; return groups }, {}), [transactions])
  function save(nextTransactions: Transaction[]) { setTransactions(nextTransactions); localStorage.setItem('budget-transactions', JSON.stringify(nextTransactions)) }
  function addTransaction(event: FormEvent<HTMLFormElement>) { event.preventDefault(); const numericAmount = Number(amount); if (!title.trim() || !numericAmount || numericAmount < 0) return; save([{ id: Date.now(), title: title.trim(), category, amount: numericAmount, type, date: '2026-09-20' }, ...transactions]); setTitle(''); setAmount('') }

  return (
    <main className="app-shell"><header className="topbar"><div className="brand-mark">$</div><div><strong>Ledgerly</strong><span>Personal finances</span></div><button className="avatar" type="button" aria-label="Account menu">JD</button></header>
      <section className="welcome-row"><div><p className="eyebrow">Sunday, September 20, 2026</p><h1>Good morning, Jamie</h1><p className="muted">Here is your financial snapshot for this month.</p></div><div className="month-chip">September 2026 <span>⌄</span></div></section>
      <section className="stats-grid" aria-label="Financial summary"><article className="stat-card balance"><div className="stat-label">Available balance <span className="info">i</span></div><strong>{currency.format(totals.income - totals.expense)}</strong><div className="trend positive">↗ 8.4% <small>vs last month</small></div><div className="balance-bar"><span style={{ width: `${Math.min(100, (totals.expense / totals.income) * 100)}%` }} /></div></article><article className="stat-card"><div className="stat-label"><span className="dot income-dot" />Total income</div><strong>{currency.format(totals.income)}</strong><div className="trend positive">↗ 12.6% <small>vs last month</small></div></article><article className="stat-card"><div className="stat-label"><span className="dot expense-dot" />Total expenses</div><strong>{currency.format(totals.expense)}</strong><div className="trend negative">↘ 3.2% <small>vs last month</small></div></article></section>
      <div className="content-grid"><section className="transactions-panel"><div className="section-heading"><div><h2>Recent activity</h2><p className="muted">Your latest income and expenses</p></div><button className="filter-button" type="button">All activity <span>⌄</span></button></div><div className="transaction-list">{Object.entries(groupedTransactions).map(([date, entries]) => <div className="date-group" key={date}><div className="date-label">{formatDate(date)} <span>{new Date(`${date}T12:00:00`).toLocaleDateString('en-US', { weekday: 'short' })}</span></div>{entries.map((transaction) => <div className="transaction" key={transaction.id}><div className={`transaction-icon ${transaction.type}`}>{transaction.type === 'income' ? '↙' : '↗'}</div><div className="transaction-info"><strong>{transaction.title}</strong><span>{transaction.category}</span></div><div className={transaction.type === 'income' ? 'amount income-amount' : 'amount'}>{transaction.type === 'income' ? '+' : '-'}{currency.format(transaction.amount)}</div><button className="delete-button" type="button" aria-label={`Delete ${transaction.title}`} onClick={() => save(transactions.filter((item) => item.id !== transaction.id))}>×</button></div>)}</div>)}</div></section>
        <aside className="add-panel"><div className="section-heading"><div><h2>Add transaction</h2><p className="muted">Keep your ledger up to date</p></div><span className="plus-icon">+</span></div><form onSubmit={addTransaction}><div className="type-toggle"><button type="button" className={type === 'expense' ? 'selected expense-selected' : ''} onClick={() => setType('expense')}>Expense</button><button type="button" className={type === 'income' ? 'selected income-selected' : ''} onClick={() => setType('income')}>Income</button></div><label>Description<input value={title} onChange={(event) => setTitle(event.target.value)} placeholder="e.g. Coffee with friends" required /></label><div className="form-row"><label>Category<select value={category} onChange={(event) => setCategory(event.target.value)}><option>Food</option><option>Housing</option><option>Transport</option><option>Shopping</option><option>Salary</option><option>Side income</option><option>Other</option></select></label><label>Amount<input type="number" min="0.01" step="0.01" value={amount} onChange={(event) => setAmount(event.target.value)} placeholder="$ 0.00" required /></label></div><button className="add-button" type="submit">Add {type}</button></form></aside></div>
    </main>
  )
}

export default App
