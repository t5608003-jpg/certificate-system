const XLSX = require("xlsx")

const HEADER_ALIASES = {
  dept: ["單位", "部門", "單位/部門", "部門別"],
  name: ["姓名", "員工姓名"],
  cert: ["證照", "證照名稱", "證照項目"],
  certNo: ["證號", "合格證字號", "證照字號", "證照編號"],
  issueDate: ["發證日", "發照日", "發證日期"],
  expiry: ["複訓期限", "複訊期限", "到期日", "有效期限", "複訓到期日"],
  training: ["在職訓練要求", "在職訓練"],
  retrain: ["已複訓日", "複訓日", "最近複訓日"],
}

const CERT_ABBR = {
  乙種職業安全衛生業務主管: "乙安",
  有機溶劑作業主管: "有機溶劑",
  特定化學物質作業主管: "特化",
  缺氧作業主管: "缺氧",
  乙級鍋爐操作人員: "乙鍋",
  高壓氣體特定設備操作人員: "高壓氣體",
  荷重在一公噸以上堆高機操作人員: "1噸以下堆高機",
  小型鍋爐操作人員: "小鍋",
  防爆電氣配線施工人員: "防爆配線",
  高空工作車操作人員: "高空作業車",
  急救人員: "急救",
  防火管理人: "防火",
  保安監督人: "保安監督人",
  保安檢查員: "保安檢查員",
}

const ALLOWED_DEPTS = new Set([
  "副廠長",
  "倉管課",
  "倉管課A",
  "倉管課B",
  "職安課",
  "管理課",
  "技術課",
  "溶劑業務課",
  "生產管理課",
  "製造課",
  "製造課一班",
  "製造課二班",
  "製造課三班",
])

const ALL_ALIASES = Object.values(HEADER_ALIASES).flat().map(normalizeHeader)

function normalizeHeader(value) {
  return String(value || "")
    .replace(/\s+/g, "")
    .replace(/[：:]/g, "")
    .trim()
}

function normalizeCert(cert) {
  const text = String(cert || "").trim()
  return CERT_ABBR[text] || text
}

function normalizeDept(dept) {
  const text = String(dept || "").trim()
  return ALLOWED_DEPTS.has(text) ? text : ""
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
      dept: normalizeDept(getCell(row, columnIndex.dept)),
      name: String(getCell(row, columnIndex.name)).trim(),
      cert: normalizeCert(getCell(row, columnIndex.cert)),
      certNo: String(getCell(row, columnIndex.certNo)).trim(),
      issueDate: excelDateToISO(getCell(row, columnIndex.issueDate)),
      expiry: excelDateToISO(getCell(row, columnIndex.expiry)),
      training: String(getCell(row, columnIndex.training)).trim(),
      retrain: excelDateToISO(getCell(row, columnIndex.retrain)),
    }))
    .filter((row) => row.dept)
    .filter((row) => row.name || row.certNo || row.cert)
}

module.exports = parseExcel
