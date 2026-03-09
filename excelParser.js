const XLSX = require("xlsx")

const HEADER_ALIASES = {
  dept: ["單位", "部門", "單位/部門", "部門別"],
  name: ["姓名", "員工姓名"],
  cert: ["證照", "證照名稱", "證照項目"],
  certNo: ["證號", "合格證字號", "證照字號", "證照編號"],
  issueDate: ["發證日", "發照日", "發證日期"],
  expiry: ["複訓期限", "到期日", "有效期限", "複訓到期日"],
  training: ["在職訓練要求", "在職訓練"],
  retrain: ["已複訓日", "複訓日", "最近複訓日"],
}

const ALL_ALIASES = Object.values(HEADER_ALIASES).flat().map(normalizeHeader)

function normalizeHeader(value) {
  return String(value || "")
    .replace(/\s+/g, "")
    .replace(/[：:]/g, "")
    .trim()
}

function excelDateToISO(value) {
  if (value === "" || value == null) return ""

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10)
  }

  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value)
    if (!parsed) return ""
    const mm = String(parsed.m).padStart(2, "0")
    const dd = String(parsed.d).padStart(2, "0")
    return `${parsed.y}-${mm}-${dd}`
  }

  return String(value).trim()
}

function findHeaderRow(rows) {
  let bestIndex = 0
  let bestScore = -1

  rows.slice(0, 20).forEach((row, idx) => {
    const normalizedCells = row.map(normalizeHeader).filter(Boolean)
    const score = normalizedCells.filter((cell) => ALL_ALIASES.includes(cell)).length
    if (score > bestScore) {
      bestScore = score
      bestIndex = idx
    }
  })

  return bestIndex
}

function findColumnIndex(headerCells, aliases) {
  const aliasSet = new Set(aliases.map(normalizeHeader))
  return headerCells.findIndex((cell) => aliasSet.has(normalizeHeader(cell)))
}

function getCell(row, idx) {
  if (idx < 0) return ""
  const value = row[idx]
  return value == null ? "" : value
}

function parseExcel(path) {
  const workbook = XLSX.readFile(path, { cellDates: true })
  const sheet = workbook.Sheets[workbook.SheetNames[0]]
  const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" })

  if (!aoa.length) return []

  const headerRowIndex = findHeaderRow(aoa)
  const header = aoa[headerRowIndex] || []
  const dataRows = aoa.slice(headerRowIndex + 1)

  const columnIndex = {
    dept: findColumnIndex(header, HEADER_ALIASES.dept),
    name: findColumnIndex(header, HEADER_ALIASES.name),
    cert: findColumnIndex(header, HEADER_ALIASES.cert),
    certNo: findColumnIndex(header, HEADER_ALIASES.certNo),
    issueDate: findColumnIndex(header, HEADER_ALIASES.issueDate),
    expiry: findColumnIndex(header, HEADER_ALIASES.expiry),
    training: findColumnIndex(header, HEADER_ALIASES.training),
    retrain: findColumnIndex(header, HEADER_ALIASES.retrain),
  }

  return dataRows
    .map((row) => ({
      factory: "台南工廠",
      dept: String(getCell(row, columnIndex.dept)).trim(),
      name: String(getCell(row, columnIndex.name)).trim(),
      cert: String(getCell(row, columnIndex.cert)).trim(),
      certNo: String(getCell(row, columnIndex.certNo)).trim(),
      issueDate: excelDateToISO(getCell(row, columnIndex.issueDate)),
      expiry: excelDateToISO(getCell(row, columnIndex.expiry)),
      training: String(getCell(row, columnIndex.training)).trim(),
      retrain: excelDateToISO(getCell(row, columnIndex.retrain)),
    }))
    .filter((row) => row.name || row.certNo || row.cert || row.dept)
}

module.exports = parseExcel
