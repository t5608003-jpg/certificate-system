const XLSX = require("xlsx")

const HEADER_ALIASES = {
  dept: ["單位", "部門", "單位/部門", "部門別"],
  name: ["姓名", "員工姓名"],
  cert: ["證照", "證照名稱", "證照項目", "證照別", "資格", "證照全名"],
  certNo: ["證號", "合格證字號", "證照字號", "證照編號"],
  issueDate: ["發證日", "發照日", "發證日期"],
  expiry: ["複訓期限", "複訊期限", "到期日", "有效期限", "複訓到期日"],
  training: ["在職訓練要求", "在職訓練"],
  retrain: ["已複訓日", "複訓日", "最近複訓日"],
}

const CERT_ABBR = {
  乙種職業安全衛生業務主管: "乙業主管",
  有機溶劑作業主管: "有機溶劑",
  特定化學物質作業主管: "特化物質",
  缺氧作業主管: "缺氧作業",
  乙級鍋爐操作人員: "乙級鍋爐",
  高壓氣體特定設備操作人員: "高壓氣體",
  荷重在一公噸以上堆高機操作人員: "堆高機",
  小型鍋爐操作人員: "小型鍋爐",
  防爆電氣配線施工人員: "防爆配線",
  高空工作車操作人員: "高空車",
  急救人員: "急救人員",
  防火管理人: "防火管理",
  保安監督人: "保安監督",
  保安檢查員: "保安檢查",
  使用起重機具從事吊掛作業人員: "吊掛作業",
}

const CERT_KEYS = Object.keys(CERT_ABBR)

const CERT_PATTERNS = {
  乙種職業安全衛生業務主管: [/乙種.*職業安全衛生.*(業務)?主.?管/, /職業安全衛生.*乙種/, /乙.*職.*安.*主.?管/, /職安.*乙種/],
  有機溶劑作業主管: [/有機溶劑.*作業主管/, /作業主管.*有機溶劑/],
  特定化學物質作業主管: [/特定化學物質.*作業主管/, /特化(物質)?/],
  缺氧作業主管: [/缺氧.*作業主管/, /缺氧作業/],
  乙級鍋爐操作人員: [/乙級鍋爐.*操作人員/, /乙級鍋爐/],
  高壓氣體特定設備操作人員: [/高壓氣體.*特定設備.*操作人員/, /高壓氣體/],
  荷重在一公噸以上堆高機操作人員: [/堆高機.*操作人員/, /荷重.*堆高機/, /堆高機/],
  小型鍋爐操作人員: [/小型鍋爐.*操作人員/, /小型鍋爐/],
  防爆電氣配線施工人員: [/防爆.*配線/, /防爆配線/],
  高空工作車操作人員: [/高空工作車.*操作人員/, /高空車/],
  急救人員: [/急救人員/, /急救/],
  防火管理人: [/防火管理人/, /防火管理/],
  保安監督人: [/保安監督人/, /保安監督/],
  保安檢查員: [/保安檢查員/, /保安檢查/],
  使用起重機具從事吊掛作業人員: [/使用起重機具從事吊掛作業人員/, /起重機具.*吊掛作業/, /吊掛作業/],
}


const DEPT_ALIASES = {
  廠副廠長: "副廠長",
}

const ALLOWED_DEPTS = new Set([
  "副廠長",
  "倉管課",
  "倉管課A",
  "倉管課B",
  "職安課",
  "廠務課",
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

function compactText(value) {
  return String(value || "").replace(/\s+/g, "").trim()
}


function cleanText(value) {
  return String(value || "")
    .replace(/_x000D_/gi, " ")
    .replace(/&#10;|&#13;/gi, " ")
    .replace(/\r\n/g, " ")
    .replace(/[\r\n\t]/g, " ")
    .replace(/\\r|\\n|\\t/g, " ")
    .replace(/\s*LF\s*/gi, " ")
    .replace(/\s+/g, " ")
    .trim()
}

function normalizeCertFull(cert) {
  return cleanText(cert)
}

function findCertByPattern(cert) {
  const source = cleanText(cert)
  if (!source) return ""
  for (const [fullName, patterns] of Object.entries(CERT_PATTERNS)) {
    if (patterns.some((p) => p.test(source))) return fullName
  }
  return ""
}

function findCertCanonicalName(cert) {
  const source = compactText(cert)
  if (!source) return ""
  for (const fullName of CERT_KEYS) {
    const normalizedName = compactText(fullName)
    if (source === normalizedName) return fullName
    if (source.includes(normalizedName) || normalizedName.includes(source)) return fullName
  }
  return findCertByPattern(cert)
}

function inferCertByCertNo(certNo) {
  const text = cleanText(certNo)
  if (!text) return ""

  const certNoRules = [
    { pattern: /中職甲訓\d+字第\d+號(?:\(補\))?/, cert: "乙種職業安全衛生業務主管" },
    { pattern: /勞安管勞員字第\d+號/, cert: "乙種職業安全衛生業務主管" },
    { pattern: /嘉市工業字第\d+-\d+號?/, cert: "乙種職業安全衛生業務主管" },
    { pattern: /^110S\d+$/, cert: "乙種職業安全衛生業務主管" },
    { pattern: /安基小鍋證字第\d+號?/, cert: "小型鍋爐操作人員" },
  ]

  const matched = certNoRules.find((r) => r.pattern.test(text))
  return matched ? matched.cert : ""
}

function normalizeCert(cert) {
  const text = normalizeCertFull(cert)
  const canonical = findCertCanonicalName(text)
  if (!canonical) return text
  return CERT_ABBR[canonical] || text
}

function normalizeDept(dept) {
  const text = cleanText(dept)
  const normalized = DEPT_ALIASES[text] || text
  return ALLOWED_DEPTS.has(normalized) ? normalized : ""
}

function formatISODate(year, month, day) {
  return `${String(year).padStart(4, "0")}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`
}

function parseDateText(value) {
  const text = cleanText(value)
  if (!text) return ""

  const rocMatch = text.match(/^(?:民國)?\s*(\d{2,3})[\/.\-年]\s*(\d{1,2})[\/.\-月]\s*(\d{1,2})/)
  if (rocMatch) {
    const year = Number(rocMatch[1]) + 1911
    const month = Number(rocMatch[2])
    const day = Number(rocMatch[3])
    return formatISODate(year, month, day)
  }

  const adMatch = text.match(/^(\d{4})[\/.\-年]\s*(\d{1,2})[\/.\-月]\s*(\d{1,2})/)
  if (adMatch) {
    const year = Number(adMatch[1])
    const month = Number(adMatch[2])
    const day = Number(adMatch[3])
    return formatISODate(year, month, day)
  }

  return text
}

function excelDateToISO(value) {
  if (value === "" || value == null) return ""
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10)
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value)
    if (!parsed) return ""
    return formatISODate(parsed.y, parsed.m, parsed.d)
  }
  return parseDateText(value)
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

function inferCertColumnIndex(dataRows, startAt, excludedIdx = new Set()) {
  if (!dataRows.length) return -1
  const width = Math.max(...dataRows.map((row) => row.length), 0)
  let bestIdx = -1
  let bestScore = 0

  for (let i = startAt; i < width; i++) {
    if (excludedIdx.has(i)) continue
    let score = 0
    for (const row of dataRows.slice(0, 200)) {
      const cell = cleanText(row[i])
      if (!cell) continue
      if (findCertCanonicalName(cell)) score += 2
      if (/作業主管|操作人員|管理人|監督人|檢查員/.test(cell)) score += 1
    }
    if (score > bestScore) {
      bestScore = score
      bestIdx = i
    }
  }
  return bestScore > 0 ? bestIdx : -1
}

function getCell(row, idx) {
  if (idx < 0) return ""
  const value = row[idx]
  return value == null ? "" : value
}



function findCertFromRow(row) {
  const cells = row.map((v) => cleanText(v)).filter(Boolean)
  const joined = compactText(cells.join(" "))
  if (!joined) return { certFull: "", cert: "" }

  for (const fullName of CERT_KEYS) {
    const normalizedName = compactText(fullName)
    if (joined.includes(normalizedName)) {
      return { certFull: fullName, cert: CERT_ABBR[fullName] || fullName }
    }
  }

  const byPattern = findCertByPattern(cells.join(" "))
  if (byPattern) {
    return { certFull: byPattern, cert: CERT_ABBR[byPattern] || byPattern }
  }

  for (const [fullName, abbr] of Object.entries(CERT_ABBR)) {
    const normalizedAbbr = compactText(abbr)
    if (normalizedAbbr && joined.includes(normalizedAbbr)) {
      return { certFull: fullName, cert: abbr }
    }
  }

  return { certFull: "", cert: "" }
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

  if (columnIndex.cert < 0) {
    const startAt = Math.max(columnIndex.name + 1, columnIndex.dept + 1, 0)
    columnIndex.cert = inferCertColumnIndex(dataRows, startAt, new Set([columnIndex.certNo]))
  }

  const parsedRows = dataRows.map((row) => {
    const certNoValue = cleanText(getCell(row, columnIndex.certNo))
    let certFromColumn = normalizeCertFull(getCell(row, columnIndex.cert))

    if (certFromColumn && certNoValue && certFromColumn === certNoValue) {
      certFromColumn = ""
    }

    const certFallback = findCertFromRow(row)
    const hintedCert = inferCertByCertNo(certNoValue)
    const certSource = certFromColumn || certFallback.certFull || hintedCert
    const certFull = certSource
    const cert = normalizeCert(certSource) || certFallback.cert

    return {
      factory: "台南工廠",
      dept: normalizeDept(getCell(row, columnIndex.dept)),
      name: cleanText(getCell(row, columnIndex.name)),
      certFull,
      cert,
      certNo: certNoValue,
      issueDate: excelDateToISO(getCell(row, columnIndex.issueDate)),
      expiry: excelDateToISO(getCell(row, columnIndex.expiry)),
      training: cleanText(getCell(row, columnIndex.training)),
      retrain: excelDateToISO(getCell(row, columnIndex.retrain)),
    }
  })

  return parsedRows
    .filter((row) => row.dept)
    .filter((row) => row.name || row.certNo || row.cert)
}

module.exports = parseExcel
