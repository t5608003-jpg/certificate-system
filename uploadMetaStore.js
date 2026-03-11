const fs = require("fs")

const FILE = "./data/uploadMeta.json"

function loadUploadMeta() {
  if (!fs.existsSync(FILE)) return { fileName: "" }
  try {
    const raw = fs.readFileSync(FILE, "utf8")
    const data = JSON.parse(raw)
    return { fileName: String(data.fileName || "") }
  } catch (_) {
    return { fileName: "" }
  }
}

function saveUploadMeta(meta) {
  const payload = { fileName: String((meta && meta.fileName) || "") }
  fs.writeFileSync(FILE, JSON.stringify(payload, null, 2))
}

module.exports = {
  loadUploadMeta,
  saveUploadMeta,
}