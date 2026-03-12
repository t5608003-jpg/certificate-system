const express=require("express")
const multer=require("multer")
const cors=require("cors")

const parseExcel=require("./excelParser")
const {setCourse,getCourses}=require("./courseStore")
const {loadEmployees,saveEmployees}=require("./employeeStore")
const {loadUploadMeta,saveUploadMeta}=require("./uploadMetaStore")

const app=express()

app.use(cors())
app.use(express.json())
app.use((req,res,next)=>{
 res.set("Cache-Control","no-store")
 next()
})
app.use(express.static("public",{etag:false,lastModified:false,maxAge:0}))

const upload=multer({dest:"uploads/"})

let employees=loadEmployees()
let uploadMeta=loadUploadMeta()

function parseDate(str){
 if(!str) return null
 const d=new Date(str)
 return Number.isNaN(d.getTime()) ? null : d
}

function addMonths(date,months){
 return new Date(date.getFullYear(),date.getMonth()+months,1)
}

function endOfMonth(date){
 return new Date(date.getFullYear(),date.getMonth()+1,0,23,59,59,999)
}

function startOfCurrentMonth(){
 const now=new Date()
 return new Date(now.getFullYear(),now.getMonth(),1)
}

function isPlanned(course){
 return !!(course && (course.date || course.place))
}

function cleanDisplayText(value){
 return String(value || "")
  .replace(/_x000D_/gi," ")
  .replace(/&#10;|&#13;/gi," ")
  .replace(/\r\n/g," ")
  .replace(/[\r\n\t]/g," ")
  .replace(/\\r|\\n|\\t/g," ")
  .replace(/\s*LF\s*/gi," ")
  .replace(/\s+/g," ")
  .trim()
}


function decodeUploadFileName(name){
 const raw=String(name || "")
 try{
  const decoded=Buffer.from(raw,"latin1").toString("utf8")
  if(decoded.includes("�")) return cleanDisplayText(raw)
  return cleanDisplayText(decoded)
 }catch(_){
  return cleanDisplayText(raw)
 }
}

function sanitizeEmployeeForResponse(row){
 return {
  ...row,
  dept:cleanDisplayText(row.dept),
  name:cleanDisplayText(row.name),
  certFull:cleanDisplayText(row.certFull),
  cert:cleanDisplayText(row.cert),
  certNo:cleanDisplayText(row.certNo),
  issueDate:cleanDisplayText(row.issueDate),
  expiry:cleanDisplayText(row.expiry),
  training:cleanDisplayText(row.training),
  retrain:cleanDisplayText(row.retrain),
 }
}

function courseStatus(expiry,course){
 if(isPlanned(course)) return "已規劃"

 const today=new Date()
 const exp=parseDate(expiry)

 if(!exp) return "合格"
 if(exp < today) return "已過期"

 const limit=endOfMonth(addMonths(startOfCurrentMonth(),3))
 if(exp <= limit) return "待規劃"

 return "合格"
}

app.post("/upload",upload.single("file"),(req,res)=>{
 if(!req.file){
  return res.status(400).json({error:"請先選擇檔案"})
 }

 employees=parseExcel(req.file.path)
 saveEmployees(employees)

 uploadMeta={fileName:decodeUploadFileName(req.file.originalname)}
 saveUploadMeta(uploadMeta)

 res.json({ count:employees.length, fileName:uploadMeta.fileName })
})


app.get("/upload-meta",(req,res)=>{
 res.json({fileName:cleanDisplayText(uploadMeta.fileName)})
})

app.get("/search",(req,res)=>{
 const keyword=req.query.keyword || ""
 const expiry=req.query.expiry || ""
 const keywordText=keyword.toLowerCase()

 const monthOffset=Number(expiry)
 const useMonthFilter=expiry!=="" && !Number.isNaN(monthOffset)
 const rangeStart=startOfCurrentMonth()
 const rangeEnd=useMonthFilter ? endOfMonth(addMonths(rangeStart,monthOffset)) : null

 const courses=getCourses()

 let result=employees
  .filter(e=>{
   const cleaned=sanitizeEmployeeForResponse(e)
   const text=(cleaned.factory+cleaned.dept+cleaned.name+(cleaned.cert||"")+(cleaned.certFull||"")).toLowerCase()
   if(!text.includes(keywordText)) return false

   if(!useMonthFilter) return true
   if(!cleaned.name || !cleaned.expiry) return false

   const exp=parseDate(cleaned.expiry)
   if(!exp) return false

   return exp>=rangeStart && exp<=rangeEnd
  })
  .map(r=>{
   const cleaned=sanitizeEmployeeForResponse(r)
   const course=courses[cleaned.certNo] || courses[r.certNo]
   const status=courseStatus(cleaned.expiry,course)
   return{
    ...cleaned,
    cert:cleaned.cert || cleaned.certFull || "",
    status,
    course
   }
  })

 if(useMonthFilter){
  result=result.filter(r=>r.status!=="合格")
 }

 res.json(result)
})

app.get("/detail/:certNo",(req,res)=>{
 const certNo=req.params.certNo
 const emp=employees.find(e=>e.certNo==certNo)
 const cleaned=emp ? sanitizeEmployeeForResponse(emp) : undefined
 const allCourses=getCourses()
 const course=allCourses[certNo] || allCourses[cleaned?.certNo]
 const status=cleaned ? courseStatus(cleaned.expiry,course) : "合格"
 res.json({...cleaned,course,status})
})

app.post("/course",(req,res)=>{
 const {certNo,date,place}=req.body
 setCourse(certNo,{date,place})
 res.json({ok:true})
})

app.listen(8080,()=>{
 console.log("Server running on port 8080")
})
