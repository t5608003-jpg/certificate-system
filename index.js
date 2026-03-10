const express=require("express")
const multer=require("multer")
const cors=require("cors")

const parseExcel=require("./excelParser")
const {setCourse,getCourses}=require("./courseStore")
const {loadEmployees,saveEmployees}=require("./employeeStore")

const app=express()

app.use(cors())
app.use(express.json())
app.use(express.static("public"))

const upload=multer({dest:"uploads/"})

let employees=loadEmployees()

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

 res.json({ count:employees.length })
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
   const text=(e.factory+e.dept+e.name+(e.cert||"")+(e.certFull||"")).toLowerCase()
   if(!text.includes(keywordText)) return false

   if(!useMonthFilter) return true
   if(!e.name || !e.expiry) return false

   const exp=parseDate(e.expiry)
   if(!exp) return false

   return exp>=rangeStart && exp<=rangeEnd
  })
  .map(r=>{
   const course=courses[r.certNo]
   const status=courseStatus(r.expiry,course)
   return{
    ...r,
    cert:r.cert || r.certFull || "",
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
 const course=getCourses()[certNo]
 const status=emp ? courseStatus(emp.expiry,course) : "合格"
 res.json({...emp,course,status})
})

app.post("/course",(req,res)=>{
 const {certNo,date,place}=req.body
 setCourse(certNo,{date,place})
 res.json({ok:true})
})

app.listen(8080,()=>{
 console.log("Server running on port 8080")
})
