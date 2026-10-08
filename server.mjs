// Import Node's HTTP server factory for API routes and static-file delivery.
import { createServer } from 'node:http'
// Import filesystem helpers for checking and reading built frontend files.
import { existsSync, readFileSync } from 'node:fs'
// Import a path helper for constructing OS-safe paths.
import { join } from 'node:path'
// Import SheetJS for reading and writing the Excel workbook.
import XLSX from 'xlsx'

// Read the configured server port or use the local development default.
const port = Number(process.env.PORT ?? 3001)
// Read the configured bind address or listen on all interfaces.
const host = process.env.HOST ?? '0.0.0.0'
// Resolve the production frontend directory from the current working directory.
const distPath = join(process.cwd(), 'dist')
// Resolve the persistent workbook file used by this server.
const databasePath = join(process.cwd(), 'budget-database.xlsx')

// Define example records used when the workbook has not been created yet.
const starterTransactions = [
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
// Perform this step as part of the surrounding operation.
]

// Load active transaction rows from the workbook, or provide starter records.
function readTransactions() {
  // Use starter records when no database file exists yet.
  if (!existsSync(databasePath)) return starterTransactions
  // Parse the Excel workbook from disk.
  const workbook = XLSX.readFile(databasePath)
  // Select the sheet containing active transactions.
  const sheet = workbook.Sheets.Transactions
  // Preserve a usable initial view if the sheet is missing.
  if (!sheet) return starterTransactions
  // Convert spreadsheet columns into the app's transaction objects.
  return XLSX.utils.sheet_to_json(sheet).map((row) => ({
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
    type: row.Type === 'income' ? 'income' : 'expense',
    // Store the transaction or archive date.
    date: String(row.Date ?? ''),
  // Perform this step as part of the surrounding operation.
  }))
// Close the current object, callback, or expression.
}

// Load archived transaction rows from the workbook's history sheet.
function readHistory() {
  // A missing database means there cannot be any saved history yet.
  if (!existsSync(databasePath)) return []
  // Parse the workbook so its history worksheet can be read.
  const workbook = XLSX.readFile(databasePath)
  // Select the worksheet where archived entries are stored.
  const sheet = workbook.Sheets['Expense History']
  // Return an empty archive when the workbook has no history sheet.
  if (!sheet) return []
  // Convert the worksheet rows to transaction objects for the UI.
  return XLSX.utils.sheet_to_json(sheet).map((row) => ({
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
    type: row.Type === 'income' ? 'income' : 'expense',
    // Store the transaction or archive date.
    date: String(row.Date ?? ''),
  // Perform this step as part of the surrounding operation.
  }))
// Close the current object, callback, or expression.
}

// Calculate summary rows that accompany a titled archived transaction group.
function buildHistoryTotals(transactions, historyTitle) {
  // Add each transaction amount to the bucket matching its type.
  const totals = transactions.reduce((summary, transaction) => {
    // Add this transaction amount to its matching running total.
    summary[transaction.type] += transaction.amount
    // Return the value produced by this function or callback.
    return summary
  // Perform this step as part of the surrounding operation.
  }, { income: 0, expense: 0 })
  // Record the current date for the archive summary rows.
  const archiveDate = new Date().toISOString().slice(0, 10)

  // Return the group's income, expenses, and resulting balance as rows.
  return [
    // Open the block or collection used by the surrounding operation.
    {
      // Store the record identifier.
      id: Date.now() + 1,
      // Store the transaction description.
      title: 'Total income',
      // Continue the surrounding list or object with this value.
      historyTitle,
      // Store the selected transaction category.
      category: 'Summary',
      // Store the monetary amount.
      amount: totals.income,
      // Store whether this row represents income or an expense.
      type: 'income',
      // Store the transaction or archive date.
      date: archiveDate,
    // Close the current object, callback, or expression.
    },
    // Open the block or collection used by the surrounding operation.
    {
      // Store the record identifier.
      id: Date.now() + 2,
      // Store the transaction description.
      title: 'Total expenses',
      // Continue the surrounding list or object with this value.
      historyTitle,
      // Store the selected transaction category.
      category: 'Summary',
      // Store the monetary amount.
      amount: totals.expense,
      // Store whether this row represents income or an expense.
      type: 'expense',
      // Store the transaction or archive date.
      date: archiveDate,
    // Close the current object, callback, or expression.
    },
    // Open the block or collection used by the surrounding operation.
    {
      // Store the record identifier.
      id: Date.now() + 3,
      // Store the transaction description.
      title: 'Available balance',
      // Continue the surrounding list or object with this value.
      historyTitle,
      // Store the selected transaction category.
      category: 'Summary',
      // Store the monetary amount.
      amount: totals.income - totals.expense,
      // Store whether this row represents income or an expense.
      type: 'income',
      // Store the transaction or archive date.
      date: archiveDate,
    // Close the current object, callback, or expression.
    },
  // Perform this step as part of the surrounding operation.
  ]
// Close the current object, callback, or expression.
}

// Serialize current transactions and history into the database workbook.
function writeDatabase(transactions, history = readHistory()) {
  // Rename transaction fields to the column headings stored in Excel.
  const transactionRows = transactions.map((transaction) => ({
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
  // Perform this step as part of the surrounding operation.
  }))
  // Compute totals to populate the workbook's summary worksheet.
  const totals = transactions.reduce((summary, transaction) => {
    // Add this transaction amount to its matching running total.
    summary[transaction.type] += transaction.amount
    // Return the value produced by this function or callback.
    return summary
  // Perform this step as part of the surrounding operation.
  }, { income: 0, expense: 0 })
  // Start an empty workbook before adding its sheets.
  const workbook = XLSX.utils.book_new()
  // Create a summary sheet with income, expense, and available balance.
  const summarySheet = XLSX.utils.json_to_sheet([
    // Continue the surrounding list or object with this value.
    { Metric: 'Total income', Amount: totals.income },
    // Continue the surrounding list or object with this value.
    { Metric: 'Total expenses', Amount: totals.expense },
    // Continue the surrounding list or object with this value.
    { Metric: 'Available balance', Amount: totals.income - totals.expense },
  // Close the current object, callback, or expression.
  ])
  // Convert active transaction rows into a worksheet.
  const transactionSheet = XLSX.utils.json_to_sheet(transactionRows)
  // Convert archived history rows into a worksheet.
  const historySheet = XLSX.utils.json_to_sheet(history.map((transaction) => ({
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
  // Perform this step as part of the surrounding operation.
  })))
  // Give summary columns readable widths in spreadsheet applications.
  summarySheet['!cols'] = [{ wch: 22 }, { wch: 16 }]
  // Give transaction columns widths that fit their content.
  transactionSheet['!cols'] = [{ wch: 12 }, { wch: 14 }, { wch: 26 }, { wch: 18 }, { wch: 12 }, { wch: 14 }]
  // Reuse the same column widths for history rows.
  historySheet['!cols'] = transactionSheet['!cols']
  // Add the totals worksheet to the workbook.
  XLSX.utils.book_append_sheet(workbook, summarySheet, 'Summary')
  // Add the active-ledger worksheet to the workbook.
  XLSX.utils.book_append_sheet(workbook, transactionSheet, 'Transactions')
  // Add the archived-ledger worksheet to the workbook.
  XLSX.utils.book_append_sheet(workbook, historySheet, 'Expense History')
  // Save the constructed workbook at the configured database path.
  XLSX.writeFile(workbook, databasePath)
// Close the current object, callback, or expression.
}

// Return an HTTP response whose payload is encoded as JSON.
function sendJson(response, statusCode, data) {
  // Set the status and JSON content type, plus browser cross-origin access.
  response.writeHead(statusCode, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' })
  // Serialize the response object and finish the request.
  response.end(JSON.stringify(data))
// Close the current object, callback, or expression.
}

// Serve built frontend files and support client-side routes.
function serveFrontend(request, response) {
  // Report a clear build error if the frontend output is missing.
  if (!existsSync(join(distPath, 'index.html'))) {
    // Perform this step as part of the surrounding operation.
    sendJson(response, 404, { error: 'Frontend build not found. Run npm run build first.' })
    // Perform this step as part of the surrounding operation.
    return
  // Close the current object, callback, or expression.
  }

  // Parse the request path using the server address as the URL base.
  const requestedPath = new URL(request.url, `http://${host}`).pathname
  // Resolve the requested asset, mapping the root URL to the HTML entry point.
  const filePath = join(distPath, requestedPath === '/' ? 'index.html' : requestedPath)
  // Fall back to index.html for paths handled by the React router.
  const assetPath = existsSync(filePath) ? filePath : join(distPath, 'index.html')
  // Select a MIME type based on the asset extension.
  const contentType = assetPath.endsWith('.js')
    // Perform this step as part of the surrounding operation.
    ? 'text/javascript'
    // Perform this step as part of the surrounding operation.
    : assetPath.endsWith('.css')
      // Perform this step as part of the surrounding operation.
      ? 'text/css'
      // Perform this step as part of the surrounding operation.
      : assetPath.endsWith('.svg')
        // Perform this step as part of the surrounding operation.
        ? 'image/svg+xml'
        // Perform this step as part of the surrounding operation.
        : 'text/html'

  // Send an HTTP success response with the detected content type.
  response.writeHead(200, { 'Content-Type': contentType })
  // Read and send the requested frontend asset.
  response.end(readFileSync(assetPath))
// Close the current object, callback, or expression.
}

// Route API requests and send all other GET requests to the frontend.
const server = createServer((request, response) => {
  // Answer CORS preflight requests before processing the requested route.
  if (request.method === 'OPTIONS') {
    // Handle this event or write the corresponding HTTP response.
    response.writeHead(204, { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Methods': 'GET, POST', 'Access-Control-Allow-Headers': 'Content-Type' })
    // Handle this event or write the corresponding HTTP response.
    response.end()
    // Perform this step as part of the surrounding operation.
    return
  // Close the current object, callback, or expression.
  }
  // Extract the URL path so route handlers can match it.
  const requestPath = new URL(request.url, `http://${host}`).pathname
  // Return the active transactions as JSON.
  if (request.method === 'GET' && requestPath === '/api/transactions') {
    // Perform this step as part of the surrounding operation.
    sendJson(response, 200, readTransactions())
    // Perform this step as part of the surrounding operation.
    return
  // Close the current object, callback, or expression.
  }
  // Return archived transactions as JSON.
  if (request.method === 'GET' && requestPath === '/api/transactions/history') {
    // Perform this step as part of the surrounding operation.
    sendJson(response, 200, readHistory())
    // Perform this step as part of the surrounding operation.
    return
  // Close the current object, callback, or expression.
  }
  // Accept and persist replacement data for the active transaction list.
  if (request.method === 'POST' && requestPath === '/api/transactions') {
    // Perform this step as part of the surrounding operation.
    let body = ''
    // Collect incoming request chunks into one JSON string.
    request.on('data', (chunk) => { body += chunk })
    // Parse and save the completed request, returning validation errors to the caller.
    request.on('end', () => {
      // Begin the block that contains the following related logic.
      try {
        // Compute transactions to hold the decoded list of transaction records.
        const transactions = JSON.parse(body)
        // Check this condition before continuing with the operation.
        if (!Array.isArray(transactions)) throw new Error('Transactions must be an array')
        // Perform this step as part of the surrounding operation.
        writeDatabase(transactions)
        // Perform this step as part of the surrounding operation.
        sendJson(response, 200, { saved: true })
      // Begin the block that contains the following related logic.
      } catch (error) {
        // Perform this step as part of the surrounding operation.
        sendJson(response, 400, { error: error.message })
      // Close the current object, callback, or expression.
      }
    // Close the current object, callback, or expression.
    })
    // Perform this step as part of the surrounding operation.
    return
  // Close the current object, callback, or expression.
  }
  // Archive the active list under a user-provided title and clear the active sheet.
  if (request.method === 'POST' && requestPath === '/api/transactions/archive') {
    // Perform this step as part of the surrounding operation.
    let body = ''
    // Collect the archive request body before attempting to parse it.
    request.on('data', (chunk) => { body += chunk })
    // Validate the title and available rows, then store the archived records and totals.
    request.on('end', () => {
      // Begin the block that contains the following related logic.
      try {
        // Read the requested fields from the parsed input object.
        const { title } = JSON.parse(body || '{}')
        // Compute historyTitle to hold the normalized title for the archive group.
        const historyTitle = String(title ?? '').trim()
        // Compute transactions to hold the decoded list of transaction records.
        const transactions = readTransactions()
        // Check this condition before continuing with the operation.
        if (!historyTitle) throw new Error('A history title is required.')
        // Check this condition before continuing with the operation.
        if (transactions.length === 0) throw new Error('There are no transactions to save.')
        // Compute titledTransactions to hold active rows tagged with the archive group title.
        const titledTransactions = transactions.map((transaction) => ({ ...transaction, historyTitle }))
        // Compute totals to hold the current income and expense sums.
        const totals = buildHistoryTotals(transactions, historyTitle)
        // Perform this step as part of the surrounding operation.
        writeDatabase([], [...readHistory(), ...titledTransactions, ...totals])
        // Perform this step as part of the surrounding operation.
        sendJson(response, 200, { saved: transactions.length })
      // Begin the block that contains the following related logic.
      } catch (error) {
        // Perform this step as part of the surrounding operation.
        sendJson(response, 400, { error: error.message })
      // Close the current object, callback, or expression.
      }
    // Close the current object, callback, or expression.
    })
    // Perform this step as part of the surrounding operation.
    return
  // Close the current object, callback, or expression.
  }
  // Serve the frontend for any remaining GET route.
  if (request.method === 'GET') {
    // Perform this step as part of the surrounding operation.
    serveFrontend(request, response)
    // Perform this step as part of the surrounding operation.
    return
  // Close the current object, callback, or expression.
  }
  // Respond with not found when no method and route combination matched.
  sendJson(response, 404, { error: 'Not found' })
// Close the current object, callback, or expression.
})

// Start the server and initialize its workbook before announcing readiness.
server.listen(port, host, () => {
  // Ensure the database workbook exists and has the latest active rows.
  writeDatabase(readTransactions())
  // Print the configured bind address and port for diagnostics.
  console.log(`Budget tracker running on ${host}:${port}`)
// Close the current object, callback, or expression.
})
