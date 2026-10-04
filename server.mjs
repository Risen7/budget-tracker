import { createServer } from 'node:http'
import { execFile } from 'node:child_process'
import { copyFileSync, existsSync, readFileSync, statSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { promisify } from 'node:util'
import XLSX from 'xlsx'

const execFileAsync = promisify(execFile)
const port = Number(process.env.PORT ?? 3001)
const host = process.env.HOST ?? '0.0.0.0'
const distPath = join(process.cwd(), 'dist')
const databaseSettingsPath = join(process.cwd(), '.budget-database-location.json')
let databaseDirectory = process.cwd()
if (existsSync(databaseSettingsPath)) {
  const settings = JSON.parse(readFileSync(databaseSettingsPath, 'utf8'))
  if (typeof settings.directory !== 'string') throw new Error('The saved Excel database directory is invalid.')
  databaseDirectory = resolve(settings.directory)
  if (!existsSync(databaseDirectory) || !statSync(databaseDirectory).isDirectory()) {
    throw new Error(`The saved Excel database directory does not exist: ${databaseDirectory}`)
  }
}

function getDatabasePath() {
  return join(databaseDirectory, 'budget-database.xlsx')
}

const starterTransactions = [
  { id: 1, title: 'Monthly salary', category: 'Salary', amount: 3200, type: 'income', date: '2026-09-20' },
  { id: 2, title: 'Apartment rent', category: 'Housing', amount: 1150, type: 'expense', date: '2026-09-20' },
  { id: 3, title: 'Grocery run', category: 'Food', amount: 86.45, type: 'expense', date: '2026-09-19' },
  { id: 4, title: 'Freelance project', category: 'Side income', amount: 450, type: 'income', date: '2026-09-18' },
  { id: 5, title: 'Train pass', category: 'Transport', amount: 42, type: 'expense', date: '2026-09-18' },
]

function readTransactions(databasePath = getDatabasePath()) {
  if (!existsSync(databasePath)) return starterTransactions
  const workbook = XLSX.readFile(databasePath)
  const sheet = workbook.Sheets.Transactions
  if (!sheet) return starterTransactions
  return XLSX.utils.sheet_to_json(sheet).map((row) => ({
    id: Number(row.ID),
    title: String(row.Description ?? ''),
    historyTitle: String(row.HistoryTitle ?? ''),
    category: String(row.Category ?? 'Other'),
    amount: Number(row.Amount ?? 0),
    type: row.Type === 'income' ? 'income' : 'expense',
    date: String(row.Date ?? ''),
  }))
}

function readHistory(databasePath = getDatabasePath()) {
  if (!existsSync(databasePath)) return []
  const workbook = XLSX.readFile(databasePath)
  const sheet = workbook.Sheets['Expense History']
  if (!sheet) return []
  return XLSX.utils.sheet_to_json(sheet).map((row) => ({
    id: Number(row.ID),
    title: String(row.Description ?? ''),
    historyTitle: String(row.HistoryTitle ?? ''),
    category: String(row.Category ?? 'Other'),
    amount: Number(row.Amount ?? 0),
    type: row.Type === 'income' ? 'income' : 'expense',
    date: String(row.Date ?? ''),
  }))
}

function buildHistoryTotals(transactions, historyTitle) {
  const totals = transactions.reduce((summary, transaction) => {
    summary[transaction.type] += transaction.amount
    return summary
  }, { income: 0, expense: 0 })
  const archiveDate = new Date().toISOString().slice(0, 10)

  return [
    {
      id: Date.now() + 1,
      title: 'Total income',
      historyTitle,
      category: 'Summary',
      amount: totals.income,
      type: 'income',
      date: archiveDate,
    },
    {
      id: Date.now() + 2,
      title: 'Total expenses',
      historyTitle,
      category: 'Summary',
      amount: totals.expense,
      type: 'expense',
      date: archiveDate,
    },
    {
      id: Date.now() + 3,
      title: 'Available balance',
      historyTitle,
      category: 'Summary',
      amount: totals.income - totals.expense,
      type: 'income',
      date: archiveDate,
    },
  ]
}

function writeDatabase(transactions, history = readHistory()) {
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
  XLSX.writeFile(workbook, getDatabasePath())
}

function sendJson(response, statusCode, data) {
  response.writeHead(statusCode, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' })
  response.end(JSON.stringify(data))
}

function isLocalRequest(request) {
  const address = request.socket.remoteAddress?.replace(/^::ffff:/, '')
  return address === '127.0.0.1' || address === '::1'
}

async function showDatabaseDirectoryPicker() {
  if (process.platform !== 'win32') {
    throw new Error('The folder picker is currently available on Windows only.')
  }

  const script = "Add-Type -AssemblyName System.Windows.Forms; $dialog = New-Object System.Windows.Forms.FolderBrowserDialog; $dialog.Description = 'Select the folder for budget-database.xlsx'; $dialog.ShowNewFolderButton = $true; if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { [Console]::Out.Write($dialog.SelectedPath) }"
  const { stdout } = await execFileAsync('powershell.exe', ['-NoProfile', '-STA', '-Command', script], { windowsHide: false })
  return stdout.trim()
}

function setDatabaseDirectory(directory) {
  const selectedDirectory = resolve(directory)
  if (!existsSync(selectedDirectory) || !statSync(selectedDirectory).isDirectory()) {
    throw new Error('The selected folder is no longer available.')
  }

  const currentPath = getDatabasePath()
  const nextPath = join(selectedDirectory, 'budget-database.xlsx')
  if (nextPath !== currentPath) {
    if (existsSync(nextPath)) {
      XLSX.readFile(nextPath)
    } else {
      copyFileSync(currentPath, nextPath)
    }

    writeFileSync(databaseSettingsPath, JSON.stringify({ directory: selectedDirectory }, null, 2))
    databaseDirectory = selectedDirectory
  }

  return { directory: databaseDirectory, fileName: 'budget-database.xlsx' }
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
  if (request.method === 'GET' && requestPath === '/api/database-directory') {
    if (!isLocalRequest(request)) {
      sendJson(response, 403, { error: 'The database folder can only be viewed from this computer.' })
      return
    }
    sendJson(response, 200, { directory: databaseDirectory, fileName: 'budget-database.xlsx' })
    return
  }
  if (request.method === 'POST' && requestPath === '/api/database-directory/pick') {
    if (!isLocalRequest(request)) {
      sendJson(response, 403, { error: 'The folder picker can only be used from this computer.' })
      return
    }
    showDatabaseDirectoryPicker()
      .then((directory) => {
        if (!directory) {
          sendJson(response, 200, { canceled: true })
          return
        }
        sendJson(response, 200, setDatabaseDirectory(directory))
      })
      .catch((error) => sendJson(response, 500, { error: error.message }))
    return
  }
  if (request.method === 'GET' && requestPath === '/api/transactions') {
    sendJson(response, 200, readTransactions())
    return
  }
  if (request.method === 'GET' && requestPath === '/api/transactions/history') {
    sendJson(response, 200, readHistory())
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
  if (request.method === 'POST' && requestPath === '/api/transactions/archive') {
    let body = ''
    request.on('data', (chunk) => { body += chunk })
    request.on('end', () => {
      try {
        const { title } = JSON.parse(body || '{}')
        const historyTitle = String(title ?? '').trim()
        const transactions = readTransactions()
        if (!historyTitle) throw new Error('A history title is required.')
        if (transactions.length === 0) throw new Error('There are no transactions to save.')
        const titledTransactions = transactions.map((transaction) => ({ ...transaction, historyTitle }))
        const totals = buildHistoryTotals(transactions, historyTitle)
        writeDatabase([], [...readHistory(), ...titledTransactions, ...totals])
        sendJson(response, 200, { saved: transactions.length })
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
