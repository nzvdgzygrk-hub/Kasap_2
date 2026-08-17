// Finale Auswertungsregeln nach den gelieferten Tavuk-/Tosun-Vorlagen.
// Grundregel: Nur tatsächlich bestellte Positionen erscheinen in den Übersichten.
function idxByDe(name){return ARTICLES.findIndex(a=>a[0]===name)}
function valueForOverview(o){o=orderAt(o);if(Number(o.kg)>0)return fmtNum(o.kg)+" kg";if(Number(o.count)>0)return fmtNum(o.count);return ""}
function sumOrders(list){return list.reduce((s,o)=>{o=orderAt(o);s.kg+=Number(o.kg)||0;s.count+=Number(o.count)||0;return s},{kg:0,count:0})}

const FINAL_TAVUK_CONFIG=[
 ["BUT","Hähnchen - Keule"],
 ["KANAT","Hähnchen - Flügel"],
 ["FILE","Hähnchen - Brustfilet o. Kn"],
 ["FILE MK","Hähnchen - Brustfilet m. Kn"],
 ["LEBER","Hähnchen - Leber"],
 ["MAGEN","Hähnchen - Magen"],
 ["HERZ","Hähnchen - Herz"],
 ["800-1000","Hähnchen - Ganz 800-1000"],
 ["1100-1300","Hähnchen - Ganz 1100-1300"],
 ["1400-1500","Hähnchen - Ganz 1400-1500"],
 ["2000","Hähnchen - Ganz 2000"],
 ["BUT OK","Hähnchen - Keule o. Kn"],
 ["PIRZOLA","Hähnchen - Oberkeule"],
 ["INCIK","Hähnchen - Unterkeule"],
 ["HINDI MK","Puten - Oberkeule mit Kn., mit Haut"],
 ["HINDI OK","Puten - Oberkeule ohne Kn., ohne Haut"],
 ["HINDI BOYUN","Puten - Hals"],
 ["HINDI FILE","Puten - Brustfilet"],
 ["KANAT PLAT","Flügel plat"]
];
function activeTavukColumns(){return FINAL_TAVUK_CONFIG.map(([label,name])=>({label,name,i:idxByDe(name)})).filter(x=>x.i>=0&&activeCustomers().some(c=>hasOrder(orderAt(c.orders?.[x.i]))))}
function tavukIndices(){return activeTavukColumns().map(x=>x.i)}
function makeTavukSheet(){
 const cols=activeTavukColumns();
 const rows=[[companyName.value||company,...Array(Math.max(0,cols.length-1)).fill("")],[t().customer,...cols.map(x=>x.label)]];
 const totals=cols.map(()=>({kg:0,count:0}));
 activeCustomers().forEach(c=>{
  const vals=cols.map((x,k)=>{const o=orderAt(c.orders?.[x.i]);totals[k].kg+=Number(o.kg)||0;totals[k].count+=Number(o.count)||0;return valueForOverview(o)});
  if(vals.some(Boolean))rows.push([c.name,...vals]);
 });
 rows.push([t().total,...totals.map(valueForOverview)]);
 const ws=XLSX.utils.aoa_to_sheet(rows);
 ws['!cols']=[{wch:24},...cols.map(()=>({wch:13}))];
 if(cols.length){ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:Math.max(0,cols.length)}}]}
 styleCell(ws.A1,{font:{bold:true,sz:13},alignment:{horizontal:'center'}});
 const range=XLSX.utils.decode_range(ws['!ref']);
 for(let c=0;c<=range.e.c;c++)styleCell(ws[XLSX.utils.encode_cell({r:1,c})],S.header);
 for(let r=2;r<range.e.r;r++)for(let c=0;c<=range.e.c;c++)styleCell(ws[XLSX.utils.encode_cell({r,c})],S.normal);
 for(let c=0;c<=range.e.c;c++)styleCell(ws[XLSX.utils.encode_cell({r:range.e.r,c})],S.total);
 ws['!pageSetup']={orientation:'landscape',fitToWidth:1,fitToHeight:1,paperSize:9};
 return ws;
}
function sheetHTMLTavuk(){
 const cols=activeTavukColumns(),totals=cols.map(()=>({kg:0,count:0}));
 let h=`<div class="pdfsheet" style="width:1500px"><div class="topline"><div>${escapeHtml(companyName.value||company)}</div><div>${escapeHtml(t().tavuk)}</div><div>${today()}</div></div><table style="width:100%;table-layout:fixed"><tr><th style="width:170px">${escapeHtml(t().customer)}</th>`;
 cols.forEach(x=>h+=`<th style="font-size:10px;word-break:break-word">${escapeHtml(x.label)}</th>`);h+='</tr>';
 activeCustomers().forEach(c=>{const vals=cols.map((x,k)=>{const o=orderAt(c.orders?.[x.i]);totals[k].kg+=Number(o.kg)||0;totals[k].count+=Number(o.count)||0;return valueForOverview(o)});if(vals.some(Boolean)){h+=`<tr><td style="font-weight:700">${escapeHtml(c.name)}</td>`;vals.forEach(v=>h+=`<td style="text-align:center;font-size:10px">${escapeHtml(v)}</td>`);h+='</tr>'}});
 h+=`<tr><td style="font-weight:700">${escapeHtml(t().total)}</td>`;totals.forEach(s=>h+=`<td style="font-weight:700;text-align:center;font-size:10px">${escapeHtml(valueForOverview(s))}</td>`);return h+'</tr></table></div>';
}

// Nur Bezeichnungen, die in der gelieferten Tosun-Liste vorkommen, werden dort ausgewertet.
const FINAL_TOSUN_MAIN=[
 ["tosun boyun","Bullen - Hals 1/2"],
 ["tosun pistole","Bullen - Pistole"],
 ["tosun but","Bullen - Keule"],
 ["tosun vv. Ml.","Bullen - W. mit Lappen"],
 ["tosun vv. Ol.","Bullen - W. ohne Lappen"],
 ["tosun kol","JB - Schulter"],
 ["tosun lappen ok","Bullen - Lappen o. Kn"],
 ["tosun lappen mk","Bullen - Lappen m. Kn"],
 ["tosun oberschale","Bullen - Oberschale"]
];
const FINAL_DANA=[
 ["dana pistole","Kalbs - Pistole"],
 ["dana but","Kalbskeule mit Knochen"],
 ["dana vv. Ol.","Kalbs - W. ohne Lappen"],
 ["dana vv. Ml","Kalbs - W."],
 ["dana incik","Kalbs - Scheibe"]
];
const FINAL_KUZU=[
 ["kuzu","Lamm"],
 ["kuzu butsuz","Lamm (ohne Keule)"],
 ["rumpf","Rumpf"],
 ["kuzu rest.","Lammverarbeitungsfleisch"],
 ["kuzu pirzola","Lammkotelett"]
];
function resolveConfig(cfg){return cfg.map(([label,name])=>({label,name,i:idxByDe(name)})).filter(x=>x.i>=0)}
function activeConfig(cfg){return resolveConfig(cfg).filter(x=>activeCustomers().some(c=>hasOrder(orderAt(c.orders?.[x.i]))))}
function configTotals(cfg){return activeConfig(cfg).map(x=>({label:x.label,o:sumOrders(activeCustomers().map(c=>c.orders?.[x.i]))}))}
function mainTosunCustomerRows(){const cfg=activeConfig(FINAL_TOSUN_MAIN),rows=[];activeCustomers().forEach(c=>{const parts=[];cfg.forEach(x=>{const o=orderAt(c.orders?.[x.i]);if(hasOrder(o))parts.push(`${valueForOverview(o)} ${x.label}`)});if(parts.length)rows.push([c.name,parts.join(' / ')])});return rows}
function tosunIndices(){return [...resolveConfig(FINAL_TOSUN_MAIN),...resolveConfig(FINAL_DANA),...resolveConfig(FINAL_KUZU)].map(x=>x.i)}
function makeTosunSheet(){
 const left=mainTosunCustomerRows(),mainTotals=configTotals(FINAL_TOSUN_MAIN),dana=configTotals(FINAL_DANA),kuzu=configTotals(FINAL_KUZU);
 const rows=[[companyName.value||company,'',t().tosun,today()],[t().customer,t().product,t().product,t().summary]];
 const max=Math.max(left.length,mainTotals.length,1);
 for(let r=0;r<max;r++){const l=left[r]||['',''],tr=mainTotals[r];rows.push([l[0],l[1],tr?tr.label:'',tr?valueForOverview(tr.o):''])}
 function addSection(title,data){if(!data.length)return;rows.push([]);rows.push(['','',title,'']);data.forEach(x=>rows.push(['','',x.label,valueForOverview(x.o)]))}
 addSection('DANA',dana);addSection('KUZU',kuzu);
 const ws=XLSX.utils.aoa_to_sheet(rows);ws['!cols']=[{wch:22},{wch:70},{wch:35},{wch:18}];
 const range=XLSX.utils.decode_range(ws['!ref']);for(let c=0;c<4;c++)styleCell(ws[XLSX.utils.encode_cell({r:1,c})],S.header);for(let r=2;r<=range.e.r;r++)for(let c=0;c<4;c++)if(ws[XLSX.utils.encode_cell({r,c})])styleCell(ws[XLSX.utils.encode_cell({r,c})],S.normal);styleCell(ws.A1,{font:{bold:true}});styleCell(ws.C1,{font:{bold:true,sz:14},alignment:{horizontal:'center'}});ws['!pageSetup']={orientation:'landscape',fitToWidth:1,fitToHeight:1,paperSize:9};return ws;
}
function sheetHTMLTosun(){
 const left=mainTosunCustomerRows(),mainTotals=configTotals(FINAL_TOSUN_MAIN),dana=configTotals(FINAL_DANA),kuzu=configTotals(FINAL_KUZU);
 let h=`<div class="pdfsheet" style="width:1300px"><div class="topline"><div>${escapeHtml(companyName.value||company)}</div><div>${escapeHtml(t().tosun)}</div><div>${today()}</div></div><div class="tosun-grid"><table><tr><th style="width:210px">${escapeHtml(t().customer)}</th><th>${escapeHtml(t().product)}</th></tr>`;
 left.forEach(r=>h+=`<tr><td style="font-weight:700">${escapeHtml(r[0])}</td><td>${escapeHtml(r[1])}</td></tr>`);h+='</table><div><table><tr><th>'+escapeHtml(t().product)+'</th><th style="width:130px">'+escapeHtml(t().summary)+'</th></tr>';mainTotals.forEach(x=>h+=`<tr><td>${escapeHtml(x.label)}</td><td>${escapeHtml(valueForOverview(x.o))}</td></tr>`);h+='</table>';
 function section(title,data){if(!data.length)return;h+=`<table style="margin-top:14px"><tr><th colspan="2">${escapeHtml(title)}</th></tr>`;data.forEach(x=>h+=`<tr><td>${escapeHtml(x.label)}</td><td style="width:130px">${escapeHtml(valueForOverview(x.o))}</td></tr>`);h+='</table>'}
 section('DANA',dana);section('KUZU',kuzu);return h+'</div></div></div>';
}

// Ursprüngliche Besteller einmalig sicher ergänzen, ohne vorhandene Daten zu überschreiben.
(function restoreOriginalCustomersFinal(){const flag='fatihKasap2DefaultCustomersRestored_20260817_v2';if(localStorage.getItem(flag)==='1')return;const existing=new Set(customers.map(c=>String(c.name||'').trim().toLocaleLowerCase('de-DE')));DEFAULT_CUSTOMERS.forEach(name=>{const key=name.trim().toLocaleLowerCase('de-DE');if(!existing.has(key)){customers.push({name,orders:{},page:null});existing.add(key)}});localStorage.setItem(flag,'1');save();renderCustomers();renderAssignments()})();
