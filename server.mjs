import { createServer } from 'node:http'
import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import XLSX from 'xlsx'

const port = Number(process.env.PORT ?? 3001)
const host = process.env.HOST ?? '0.0.0.0'
const distPath = join(process.cwd(), 'dist')
const databasePath = join(process.cwd(), 'budget-database.xlsx')
const starterTransactions = [
  { id: 1, title: 'Monthly salary', category: 'Salary', amount: 3200, type: 'income', date: '2026-09-20' },
  { id: 2, title: 'Apartment rent', category: 'Housing', amount: 1150, type: 'expense', date: '2026-09-20' },
  { id: 3, title: 'Grocery run', category: 'Food', amount: 86.45, type: 'expense', date: '2026-09-19' },
  { id: 4, title: 'Freelance project', category: 'Side income', amount: 450, type: 'income', date: '2026-09-18' },
  { id: 5, title: 'Train pass', category: 'Transport', amount: 42, type: 'expense', date: '2026-09-18' },
]

function readTransactions() {
  if (!existsSync(databasePath)) return starterTransactions
  const workbook = XLSX.readFile(databasePath)
  const sheet = workbook.Sheets.Transactions
  if (!sheet) return starterTransactions
  return XLSX.utils.sheet_to_json(sheet).map((row) => ({
    id: Number(row.ID),
    title: String(row.Description ?? ''),
    category: String(row.Category ?? 'Other'),
    amount: Number(row.Amount ?? 0),
    type: row.Type === 'income' ? 'income' : 'expense',
    date: String(row.Date ?? ''),
  }))
}

function writeDatabase(transactions) {
  const transactionRows = transactions.map((transaction) => ({
    ID: transaction.id,
    Date: transaction.date,
    Description: transaction.title,
    Category: transaction.category,
    Type: transaction.type,
    Amount: transaction.amount,
  }))
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
  const transactionSheet = XLSX.utils.json_to_sheet(transactionRows)
  summarySheet['!cols'] = [{ wch: 22 }, { wch: 16 }]
  transactionSheet['!cols'] = [{ wch: 12 }, { wch: 14 }, { wch: 26 }, { wch: 18 }, { wch: 12 }, { wch: 14 }]
  XLSX.utils.book_append_sheet(workbook, summarySheet, 'Summary')
  XLSX.utils.book_append_sheet(workbook, transactionSheet, 'Transactions')
  XLSX.writeFile(workbook, databasePath)
}

function sendJson(response, statusCode, data) {
  response.writeHead(statusCode, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' })
  response.end(JSON.stringify(data))
}

function serveFrontend(request, response) {
  if (!existsSync(join(distPath, 'index.html'))) {
    sendJson(response, 404, { error: 'Frontend build not found. Run npm run build first.' })
    return
  }

  const requestedPath = new URL(request.url, `http://${host}`).pathname
  const filePath = join(distPath, requestedPath === '/' ? 'index.html' : requestedPath)
  const assetPath = existsSync(filePath) ? filePath : join(distPath, 'index.html')
  const contentType = assetPath.endsWith('.js')
    ? 'text/javascript'
    : assetPath.endsWith('.css')
      ? 'text/css'
      : assetPath.endsWith('.svg')
        ? 'image/svg+xml'
        : 'text/html'

  response.writeHead(200, { 'Content-Type': contentType })
  response.end(readFileSync(assetPath))
}

const server = createServer((request, response) => {
  if (request.method === 'OPTIONS') {
    response.writeHead(204, { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Methods': 'GET, POST', 'Access-Control-Allow-Headers': 'Content-Type' })
    response.end()
    return
  }
  const requestPath = new URL(request.url, `http://${host}`).pathname
  if (request.method === 'GET' && requestPath === '/api/transactions') {
    sendJson(response, 200, readTransactions())
    return
  }
  if (request.method === 'POST' && requestPath === '/api/transactions') {
    let body = ''
    request.on('data', (chunk) => { body += chunk })
    request.on('end', () => {
      try {
        const transactions = JSON.parse(body)
        if (!Array.isArray(transactions)) throw new Error('Transactions must be an array')
        writeDatabase(transactions)
        sendJson(response, 200, { saved: true })
      } catch (error) {
        sendJson(response, 400, { error: error.message })
      }
    })
    return
  }
  if (request.method === 'GET') {
    serveFrontend(request, response)
    return
  }
  sendJson(response, 404, { error: 'Not found' })
})

server.listen(port, host, () => {
  writeDatabase(readTransactions())
  console.log(`Budget tracker running on ${host}:${port}`)
})
