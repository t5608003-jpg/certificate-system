const api=""

async function upload(){

 const file=document.getElementById("file").files[0]

 const form=new FormData()

 form.append("file",file)

 const r=await fetch(api+"/upload",{
  method:"POST",
  body:form
 })

 const j=await r.json()

 if(!r.ok){
  alert(j.error || "上傳失敗")
  return
 }

 document.getElementById("count").innerText=
 "目前資料筆數："+j.count

}

async function search(){

 const keyword=document.getElementById("keyword").value

 const expiry=document.getElementById("expiry").value

 const q=new URLSearchParams({keyword,expiry})

 const r=await fetch(api+"/search?"+q.toString())

 const data=await r.json()

 const tbody=document.getElementById("result")

 tbody.innerHTML=""

 data.forEach(d=>{

 const tr=document.createElement("tr")

 tr.innerHTML=`

 <td>${d.factory}</td>
 <td>${d.dept}</td>
 <td>${d.name}</td>
 <td>${d.cert}</td>
 <td>${d.expiry}</td>
 <td>
 <a href="detail.html?certNo=${d.certNo}">
 ${d.status}
 </a>
 </td>
 `

 tbody.appendChild(tr)

 })

}

async function loadDetail(){

 const params=new URLSearchParams(location.search)

 const certNo=params.get("certNo")

 const r=await fetch("/detail/"+certNo)

 const d=await r.json()

 const table=document.getElementById("detail")

 table.innerHTML=`

 <tr><td>廠別</td><td>${d.factory}</td></tr>
 <tr><td>單位</td><td>${d.dept}</td></tr>
 <tr><td>姓名</td><td>${d.name}</td></tr>
 <tr><td>證照</td><td>${d.cert}</td></tr>
 <tr><td>合格證字號</td><td>${d.certNo}</td></tr>
 <tr><td>發證日</td><td>${d.issueDate}</td></tr>
 <tr><td>已複訓日</td><td>${d.retrain}</td></tr>
 <tr><td>複訓期限</td><td>${d.expiry}</td></tr>
 `

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
