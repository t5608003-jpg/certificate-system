const XLSX = require("xlsx")

function parseExcel(path){

 const workbook = XLSX.readFile(path)

 const sheet = workbook.Sheets[workbook.SheetNames[0]]

 const rows = XLSX.utils.sheet_to_json(sheet,{defval:""})

 return rows.map(r=>({

  factory:"台南工廠",

  dept:r["單位"] || r["部門"] || "",

  name:r["姓名"],

  cert:r["證照"] || r["證照名稱"],

  certNo:r["證號"] || r["合格證字號"],

  issueDate:r["發證日"],

  expiry:r["複訓期限"] || r["到期日"],

  training:r["在職訓練要求"],

  retrain:r["已複訓日"]

 }))

}

module.exports=parseExcel