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
 return new Date(str)
}

function isValidDate(date){
 return date instanceof Date && !Number.isNaN(date.getTime())
}

function courseStatus(expiry){

 const today=new Date()

 const exp=parseDate(expiry)

 if(!isValidDate(exp)) return "合格"

 if(exp < today) return "已過期"

 const diff=(exp-today)/(1000*60*60*24)

 if(diff <= 90) return "待規劃"

 return "合格"
}

app.post("/upload",upload.single("file"),(req,res)=>{

 if(!req.file){
  return res.status(400).json({error:"請先選擇檔案"})
 }

 employees=parseExcel(req.file.path)
 saveEmployees(employees)

 res.json({
   count:employees.length
 })

})

app.get("/search",(req,res)=>{

 const keyword=req.query.keyword || ""
 const expiry=req.query.expiry || ""

 const today=new Date()

 let result=employees.filter(e=>{

 const text=(e.factory+e.dept+e.name+e.cert).toLowerCase()

 if(!text.includes(keyword.toLowerCase())) return false

 if(!expiry) return true

 if(!e.name || !e.expiry) return false

 const exp=parseDate(e.expiry)

 if(!isValidDate(exp)) return false

 const diff=(exp-today)/(1000*60*60*24)

 if(exp < today) return false

 if(expiry=="month")
   return exp.getMonth()==today.getMonth() && exp.getFullYear()==today.getFullYear()

 if(expiry=="1")
   return diff<=30

 if(expiry=="2")
   return diff<=60

 if(expiry=="3")
   return diff<=90

 return true

 })

 const courses=getCourses()

 result=result.map(r=>{

 const status=courseStatus(r.expiry)

 const course=courses[r.certNo]

 return{
   ...r,
   status,
   course
 }

 })

 res.json(result)

})

app.get("/detail/:certNo",(req,res)=>{

 const certNo=req.params.certNo

 const emp=employees.find(e=>e.certNo==certNo)

 const course=getCourses()[certNo]

 res.json({...emp,course})

})

app.post("/course",(req,res)=>{

 const {certNo,date,place}=req.body

 setCourse(certNo,{date,place})

 res.json({ok:true})

})

app.listen(8080,()=>{

 console.log("Server running on port 8080")

})
