const fs = require("fs")

const FILE = "./data/courses.json"

function loadCourses(){
  if(!fs.existsSync(FILE)) return {}
  return JSON.parse(fs.readFileSync(FILE))
}

function saveCourses(data){
  fs.writeFileSync(FILE, JSON.stringify(data,null,2))
}

function setCourse(certNo,course){
  const db = loadCourses()
  db[certNo] = course
  saveCourses(db)
}

function getCourses(){
  return loadCourses()
}

module.exports={
  setCourse,
  getCourses
}