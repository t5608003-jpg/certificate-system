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

function normalizeHeader(value) {
  return String(value || "").replace(/\s+/g, "").trim()
}

function getValueByAliases(row, aliases) {
  const entries = Object.entries(row)
  for (const alias of aliases) {
    const key = normalizeHeader(alias)
    const found = entries.find(([rawKey]) => normalizeHeader(rawKey) === key)
    if (found && found[1] !== "") return found[1]
  }
  return ""
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

function parseExcel(path) {
  const workbook = XLSX.readFile(path, { cellDates: true })
  const sheet = workbook.Sheets[workbook.SheetNames[0]]
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" })

  return rows
    .map((r) => ({
      factory: "台南工廠",
      dept: getValueByAliases(r, HEADER_ALIASES.dept),
      name: getValueByAliases(r, HEADER_ALIASES.name),
      cert: getValueByAliases(r, HEADER_ALIASES.cert),
      certNo: getValueByAliases(r, HEADER_ALIASES.certNo),
      issueDate: excelDateToISO(getValueByAliases(r, HEADER_ALIASES.issueDate)),
      expiry: excelDateToISO(getValueByAliases(r, HEADER_ALIASES.expiry)),
      training: getValueByAliases(r, HEADER_ALIASES.training),
      retrain: excelDateToISO(getValueByAliases(r, HEADER_ALIASES.retrain)),
    }))
    .filter((row) => row.name || row.certNo || row.cert)
}

module.exports = parseExcel
