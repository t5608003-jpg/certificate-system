const fs = require("fs")

const FILE = "./data/employees.json"

function loadEmployees() {
  if (!fs.existsSync(FILE)) return []
  try {
    const raw = fs.readFileSync(FILE, "utf8")
    const data = JSON.parse(raw)
    return Array.isArray(data) ? data : []
  } catch (_) {
    return []
  }
}

function saveEmployees(rows) {
  fs.writeFileSync(FILE, JSON.stringify(rows, null, 2))
}

module.exports = {
  loadEmployees,
  saveEmployees,
}