const api=""

function cleanText(value){
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

function statusColor(status){
 if(status==="合格") return "#111111"
 if(status==="已過期") return "#dc2626"
 if(status==="待規劃") return "#f59e0b"
 if(status==="已規劃") return "#2563eb"
 return "inherit"
}

function setAttachmentName(name){
 const el=document.getElementById("attachment")
 if(!el) return
 const cleaned=cleanText(name)
 el.innerText=cleaned ? `附件：${cleaned}` : "附件：尚未上傳檔案"
}

async function loadUploadMeta(){
 try{
  const r=await fetch(api+"/upload-meta")
  const j=await r.json()
  setAttachmentName(j.fileName)
 }catch(_){
  setAttachmentName("")
 }
}

function triggerGlobalCleanup(){
 if(typeof window!=="undefined" && typeof window.__cleanupLFNodes==="function"){
  window.__cleanupLFNodes()
 }
}

function removePageLFNoise(){
 const walker=document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT)
 const targets=[]
 while(walker.nextNode()){
  const node=walker.currentNode
  const cleaned=cleanText(node.nodeValue)
  if(!cleaned) targets.push(node)
  else if(cleaned!==node.nodeValue) node.nodeValue=cleaned
 }
 targets.forEach(n=>n.parentNode && n.parentNode.removeChild(n))
 triggerGlobalCleanup()
}

function getCurrentFilters(){
 return {
  keyword: document.getElementById("keyword")?.value || "",
  expiry: document.getElementById("expiry")?.value || "",
 }
}

function buildDetailUrl(certNo){
 const {keyword,expiry}=getCurrentFilters()
 const q=new URLSearchParams({
  certNo: cleanText(certNo),
  keyword,
  expiry,
 })
 return `detail.html?${q.toString()}`
}

async function upload(){
 const file=document.getElementById("file").files[0]
 const form=new FormData()
 form.append("file",file)

 const r=await fetch(api+"/upload",{ method:"POST", body:form })
 const j=await r.json()

 if(!r.ok){
  alert(j.error || "上傳失敗")
  return
 }

 document.getElementById("count").innerText="目前資料筆數："+j.count
 setAttachmentName(j.fileName)
}

async function search(){
 const {keyword,expiry}=getCurrentFilters()

 const q=new URLSearchParams({keyword,expiry})
 const r=await fetch(api+"/search?"+q.toString())
 const data=await r.json()

 const tbody=document.getElementById("result")
 tbody.innerHTML=""

 data.forEach(d=>{
  const tr=document.createElement("tr")
  tr.innerHTML=`
 <td>${cleanText(d.factory)}</td>
 <td>${cleanText(d.dept)}</td>
 <td>${cleanText(d.name)}</td>
 <td>${cleanText(d.cert)}</td>
 <td>${cleanText(d.expiry)}</td>
 <td>
 <a href="${buildDetailUrl(d.certNo)}">
 <span style="color:${statusColor(cleanText(d.status))};font-weight:600;">${cleanText(d.status)}</span>
 </a>
 </td>
 `
  tbody.appendChild(tr)
 })
 triggerGlobalCleanup()
}

async function loadDetail(){
 const params=new URLSearchParams(location.search)
 const certNo=params.get("certNo")

 const r=await fetch("/detail/"+encodeURIComponent(certNo || ""))
 const d=await r.json()

 const table=document.getElementById("detail")
 table.innerHTML=`
 <tr><td>廠別</td><td>${cleanText(d.factory)}</td></tr>
 <tr><td>單位</td><td>${cleanText(d.dept)}</td></tr>
 <tr><td>姓名</td><td>${cleanText(d.name)}</td></tr>
 <tr><td>證照</td><td>${cleanText(d.cert)}</td></tr>
 <tr><td>合格證字號</td><td>${cleanText(d.certNo)}</td></tr>
 <tr><td>發證日</td><td>${cleanText(d.issueDate)}</td></tr>
 <tr><td>已複訓日</td><td>${cleanText(d.retrain)}</td></tr>
 <tr><td>複訓期限</td><td>${cleanText(d.expiry)}</td></tr>
 <tr><td>課程安排日期</td><td>${cleanText(d.course?.date)}</td></tr>
 <tr><td>課程安排地點</td><td>${cleanText(d.course?.place)}</td></tr>
 `
 triggerGlobalCleanup()
}

function backToResult(){
 const params=new URLSearchParams(location.search)
 const keyword=params.get("keyword") || ""
 const expiry=params.get("expiry") || ""
 const q=new URLSearchParams({keyword,expiry,runSearch:"1"})
 location.href=`index.html?${q.toString()}`
}

function resetToSearch(){
 location.href="index.html"
}

async function saveCourse(){
 const params=new URLSearchParams(location.search)
 const certNo=params.get("certNo")
 const date=document.getElementById("date").value
 const place=document.getElementById("place").value

 await fetch("/course",{
  method:"POST",
  headers:{"Content-Type":"application/json"},
  body:JSON.stringify({certNo,date,place})
 })

 alert("儲存成功")
}

function restoreSearchFromQuery(){
 const params=new URLSearchParams(location.search)
 const keyword=params.get("keyword")
 const expiry=params.get("expiry")
 const runSearch=params.get("runSearch")

 if(keyword!==null && document.getElementById("keyword")){
  document.getElementById("keyword").value=keyword
 }
 if(expiry!==null && document.getElementById("expiry")){
  document.getElementById("expiry").value=expiry
 }
 if(runSearch==="1"){
  search()
 }
}

if (typeof window !== "undefined") {
 window.addEventListener("DOMContentLoaded", ()=>{
  removePageLFNoise()
  loadUploadMeta()
  restoreSearchFromQuery()
 })
}
