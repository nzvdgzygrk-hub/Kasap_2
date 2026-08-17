let lang="de",customers=[],selected=0,company="Omnia Meat GmbH",baseFile="Fatih Kasap Bestellung";
function normalizeCustomer(c){return {name:String(c?.name||""),orders:(c&&typeof c.orders==="object"&&c.orders)||{},page:Number.isInteger(c?.page)&&c.page>=1&&c.page<=10?c.page:null}}
(function load(){
  try{
    let raw=localStorage.getItem(STORAGE_KEY),s=raw?JSON.parse(raw):null;
    if(!s){const old=localStorage.getItem(OLD_STORAGE_KEY);if(old)s=JSON.parse(old)}
    if(s){customers=Array.isArray(s.customers)?s.customers.map(normalizeCustomer):[];selected=Number.isInteger(s.selected)?s.selected:0;lang=s.lang==="tr"?"tr":"de";company=s.company||company;baseFile=s.baseFile||baseFile}
  }catch(e){}
  if(!customers.length)customers=DEFAULT_CUSTOMERS.map(name=>({name,orders:{},page:null}));
  if(selected<0||selected>=customers.length)selected=0;
})();

function t(){return T[lang]}
function save(){company=companyName.value||company;baseFile=fileName.value||baseFile;localStorage.setItem(STORAGE_KEY,JSON.stringify({customers,selected,lang,company,baseFile}))}
function orderAt(v){if(typeof v==="number")return {kg:v,count:""};if(v&&typeof v==="object")return {kg:v.kg??"",count:v.count??""};return {kg:"",count:""}}
function hasOrder(o){return Number(o.kg)>0||Number(o.count)>0}
function hasCustomerOrder(c){return ARTICLES.some((_,i)=>hasOrder(orderAt(c.orders?.[i])))}
function activeCustomers(){return customers.filter(hasCustomerOrder)}
function usedPages(){return [...new Set(activeCustomers().map(c=>c.page).filter(p=>Number.isInteger(p)&&p>=1&&p<=10))].sort((a,b)=>a-b)}
function cleanFileName(n){return String(n||"Bestellung").replace(/[\\/:*?"<>|]/g,"").trim()||"Bestellung"}
function fmtNum(n){if(n===""||n==null||!Number.isFinite(Number(n)))return "";return String(Number(n)).replace(".",",")}
function amountText(o){const p=[];if(Number(o.kg)>0)p.push(fmtNum(o.kg)+" kg");if(Number(o.count)>0)p.push(fmtNum(o.count)+"x");return p.join(" / ")}
function today(){return new Date().toLocaleDateString("de-DE")}
function articleName(i){return ARTICLES[i][lang==="de"?0:1]}
function articleClass(i){
  const s=(ARTICLES[i][0]+" "+ARTICLES[i][1]).toLowerCase();
  if(/hähn|tavuk|pute|hindi|geflügel/.test(s))return "tavuk";
  if(/rind|dana|tosun|beef|roastbeef|sucuk|pastirma|pastırma/.test(s))return "tosun";
  return "other";
}
function currentOrders(c,filter=null){const rows=[];ARTICLES.forEach((a,i)=>{const o=orderAt(c.orders?.[i]);if(hasOrder(o)&&(!filter||articleClass(i)===filter))rows.push({i,name:articleName(i),de:a[0],tr:a[1],o})});return rows}

function applyLanguage(){
 const x=t();document.documentElement.lang=lang;exportTitle.textContent=x.exportTitle;companyLabel.textContent=x.company;fileLabel.textContent=x.file;excelBtn.textContent=x.excel;pdfBtn.textContent=x.pdf;exportHint.textContent=x.exportHint;
 manageTitle.textContent=x.manage;customerName.placeholder=x.customerName;password.placeholder=x.password;addBtn.textContent=x.add;deleteBtn.textContent=x.del;assignTitle.textContent=x.assign;assignHint.textContent=x.assignHint;selectTitle.textContent=x.select;itemsTitle.textContent=x.items;articleSearch.placeholder=x.search;
 deBtn.classList.toggle("active",lang==="de");trBtn.classList.toggle("active",lang==="tr");renderCustomers();renderAssignments();renderItems();
}
deBtn.onclick=()=>{lang="de";save();applyLanguage()};trBtn.onclick=()=>{lang="tr";save();applyLanguage()};companyName.value=company;fileName.value=baseFile;companyName.onchange=save;fileName.onchange=save;

function checkPassword(){if(password.value!==PASSWORD){alert(t().wrong);return false}return true}
function addCustomer(){if(!checkPassword())return;const name=customerName.value.trim();if(!name)return alert(t().enter);customers.push({name,orders:{},page:null});selected=customers.length-1;customerName.value="";password.value="";save();renderCustomers();renderAssignments();renderItems()}
function deleteCustomer(){if(!checkPassword())return;if(!customers.length)return;if(!confirm(t().confirm))return;customers.splice(selected,1);selected=customers.length?Math.min(selected,customers.length-1):-1;password.value="";save();renderCustomers();renderAssignments();renderItems()}
function renderCustomers(){customerSelect.innerHTML="";if(!customers.length){customerSelect.innerHTML=`<option>${t().none}</option>`;return}customers.forEach((c,i)=>{const o=document.createElement("option");o.value=i;o.textContent=c.name+(c.page&&hasCustomerOrder(c)?` · Sayfa ${c.page}`:"");if(i===selected)o.selected=true;customerSelect.appendChild(o)})}
function selectCustomer(){selected=Number(customerSelect.value);save();renderItems()}

function renderAssignments(){
 const active=activeCustomers();
 if(!active.length){assignmentGrid.innerHTML=`<div class="empty-assign">${escapeHtml(t().noActiveOrders)}</div>`;assignmentStatus.className='status ok';assignmentStatus.textContent=t().noActiveOrders;return}
 let html='<div class="assign-list">';
 active.forEach(c=>{
   const i=customers.indexOf(c);
   html+=`<div class="assign-row"><div class="assign-customer">${escapeHtml(c.name)}</div><select class="assign-select" aria-label="${escapeHtml(c.name)} Sayfa" onchange="setPage(${i},Number(this.value))"><option value="">${lang==='de'?'Sayfa wählen':'Sayfa seç'}</option>`;
   for(let p=1;p<=10;p++)html+=`<option value="${p}" ${c.page===p?'selected':''}>Sayfa ${p}</option>`;
   html+='</select></div>';
 });
 assignmentGrid.innerHTML=html+'</div>';
 const missing=active.filter(c=>!c.page).length;assignmentStatus.className='status '+(missing?'warn':'ok');assignmentStatus.textContent=missing?`${missing} ${t().unassignedActive}`:t().assignedActive;
}
function setPage(i,p){customers[i].page=(Number.isInteger(p)&&p>=1&&p<=10)?p:null;save();renderCustomers();renderAssignments()}

function renderItems(){
 if(selected<0||!customers[selected]){items.textContent=t().first;return}const q=(articleSearch.value||"").trim().toLowerCase();let html=`<table class="items-table"><thead><tr><th class="item">${t().ware}</th><th class="num">kg</th><th class="num">${t().count}</th></tr></thead><tbody>`;let lastGroup="";
 ARTICLES.forEach((a,i)=>{const name=articleName(i),cls=articleClass(i);if(q&&!name.toLowerCase().includes(q)&&!a[0].toLowerCase().includes(q)&&!a[1].toLowerCase().includes(q))return;const group=cls==='tavuk'?t().tavuk:(cls==='tosun'?t().tosun:(lang==='de'?'Weitere Artikel':'Diğer ürünler'));if(group!==lastGroup){html+=`<tr class="group-row"><td colspan="3">${escapeHtml(group)}</td></tr>`;lastGroup=group}const o=orderAt(customers[selected].orders?.[i]);html+=`<tr><td>${escapeHtml(name)}</td><td><input type="number" min="0" step="0.01" inputmode="decimal" value="${escapeHtml(o.kg)}" oninput="saveOrder(${i},'kg',this.value)"></td><td><input type="number" min="0" step="1" inputmode="numeric" value="${escapeHtml(o.count)}" oninput="saveOrder(${i},'count',this.value)"></td></tr>`});items.innerHTML=html+'</tbody></table>';
}
function saveOrder(i,field,value){if(selected<0)return;customers[selected].orders ||= {};const o=orderAt(customers[selected].orders[i]);if(value==="")o[field]="";else{const n=Number(String(value).replace(",","."));o[field]=Number.isFinite(n)&&n>=0?n:""}if(!hasOrder(o))delete customers[selected].orders[i];else customers[selected].orders[i]=o;save();renderAssignments();renderCustomers()}
function escapeHtml(v){return String(v??"").replace(/[&<>'"]/g,ch=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[ch]))}
function validateAssignments(){const missing=activeCustomers().filter(c=>!c.page);if(missing.length){alert(t().mustAssignActive+'\n\n'+missing.map(c=>'• '+c.name).join('\n'));return false}return true}
