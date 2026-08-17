// Kasap_2 Korrektur: Tosun Listesi nur nach der vorgegebenen Tosun-PDF-Liste.
// Außerdem werden die ursprünglich hinterlegten Standard-Besteller einmalig wiederhergestellt,
// ohne vorhandene Bestellungen oder zusätzlich angelegte Kunden zu überschreiben.

const TOSUN_PDF_EXPORT_MAP = new Map([
  ["Bullen - Hals 1/2", "tosun boyun"],
  ["Bullen - Pistole", "tosun pistole"],
  ["Bullen - Keule", "tosun but"],
  ["Bullen - Keule, gefroren", "tosun but"],
  ["Bullen - W. mit Lappen", "tosun vv. Ml."],
  ["Bullen - W. ohne Lappen", "tosun vv. Ol."],
  ["Bullen - Oberschale", "tosun oberschale"],
  ["Bullen - Lappen o. Kn", "tosun lappen ok"],
  ["Bullen - Lappen m. Kn", "tosun lappen mk"],
  ["JB - Schulter", "tosun kol"],
  ["Kalbskeule mit Knochen", "dana but"],
  ["Kalbs - Scheibe", "dana incik"],
  ["Kalbs - Pistole", "dana pistole"]
]);

function tosunPdfLabel(i){
  return TOSUN_PDF_EXPORT_MAP.get(ARTICLES[i]?.[0]) || null;
}

function tosunIndices(){
  return ARTICLES.map((_,i)=>i).filter(i=>!!tosunPdfLabel(i));
}

function tosunGroups(){
  const groups=[];
  const byLabel=new Map();
  tosunIndices().forEach(i=>{
    const label=tosunPdfLabel(i);
    if(!byLabel.has(label)){
      const g={label,indices:[]};
      byLabel.set(label,g);
      groups.push(g);
    }
    byLabel.get(label).indices.push(i);
  });
  return groups;
}

function summedOrderForIndices(c,indices){
  const out={kg:0,count:0};
  indices.forEach(i=>{
    const o=orderAt(c.orders?.[i]);
    out.kg += Number(o.kg)||0;
    out.count += Number(o.count)||0;
  });
  return out;
}

// Excel-Tosunliste: nur Positionen aus der Tosun-PDF-Zuordnung.
makeTosunSheet = function(){
  const groups=tosunGroups(), left=[];
  const totals=groups.map(g=>({label:g.label,kg:0,count:0}));
  activeCustomers().forEach(c=>{
    const parts=[];
    groups.forEach((g,k)=>{
      const o=summedOrderForIndices(c,g.indices);
      if(hasOrder(o)){
        parts.push(amountText(o)+' '+g.label);
        totals[k].kg+=Number(o.kg)||0;
        totals[k].count+=Number(o.count)||0;
      }
    });
    if(parts.length)left.push([c.name,parts.join(' / ')]);
  });
  const activeTotals=totals.filter(x=>x.kg>0||x.count>0);
  const max=Math.max(left.length,activeTotals.length,1);
  const rows=[[companyName.value||company,'',t().tosun,today()],[t().customer,t().product,t().product,t().summary]];
  for(let r=0;r<max;r++){
    const l=left[r]||['',''];
    const tr=activeTotals[r];
    rows.push([l[0],l[1],tr?tr.label:'',tr?cellSummary(tr):'']);
  }
  const ws=XLSX.utils.aoa_to_sheet(rows);
  ws['!cols']=[{wch:22},{wch:70},{wch:35},{wch:18}];
  for(let c=0;c<4;c++)styleCell(ws[XLSX.utils.encode_cell({r:1,c})],S.header);
  const range=XLSX.utils.decode_range(ws['!ref']);
  for(let r=2;r<=range.e.r;r++)for(let c=0;c<4;c++)styleCell(ws[XLSX.utils.encode_cell({r,c})],S.normal);
  styleCell(ws.A1,{font:{bold:true}});
  styleCell(ws.C1,{font:{bold:true,sz:14},alignment:{horizontal:'center'}});
  ws['!pageSetup']={orientation:'landscape',fitToWidth:1,fitToHeight:1,paperSize:9};
  return ws;
};

// PDF-Tosunliste: dieselbe strikte PDF-Zuordnung wie im Excel-Export.
sheetHTMLTosun = function(){
  const groups=tosunGroups(), left=[];
  const totals=groups.map(g=>({label:g.label,kg:0,count:0}));
  activeCustomers().forEach(c=>{
    const parts=[];
    groups.forEach((g,k)=>{
      const o=summedOrderForIndices(c,g.indices);
      if(hasOrder(o)){
        parts.push(amountText(o)+' '+g.label);
        totals[k].kg+=Number(o.kg)||0;
        totals[k].count+=Number(o.count)||0;
      }
    });
    if(parts.length)left.push([c.name,parts.join(' / ')]);
  });
  let h=`<div class="pdfsheet" style="width:1300px"><div class="topline"><div>${escapeHtml(companyName.value||company)}</div><div>${escapeHtml(t().tosun)}</div><div>${today()}</div></div><div class="tosun-grid"><table><tr><th style="width:210px">${escapeHtml(t().customer)}</th><th>${escapeHtml(t().product)}</th></tr>`;
  left.forEach(r=>h+=`<tr><td style="font-weight:700">${escapeHtml(r[0])}</td><td>${escapeHtml(r[1])}</td></tr>`);
  h+='</table><table><tr><th>'+escapeHtml(t().product)+'</th><th style="width:130px">'+escapeHtml(t().summary)+'</th></tr>';
  totals.filter(x=>x.kg>0||x.count>0).forEach(x=>h+=`<tr><td>${escapeHtml(x.label)}</td><td>${escapeHtml(cellSummary(x))}</td></tr>`);
  return h+'</table></div></div>';
};

// Die 27 ursprünglich hinterlegten Besteller einmalig ergänzen.
// Danach bleiben normale Löschungen dauerhaft möglich.
(function restoreOriginalCustomersOnce(){
  const flag='fatihKasap2DefaultCustomersRestored_20260817_v1';
  if(localStorage.getItem(flag)==='1') return;
  const existing=new Set(customers.map(c=>String(c.name||'').trim().toLocaleLowerCase('de-DE')));
  DEFAULT_CUSTOMERS.forEach(name=>{
    const key=String(name).trim().toLocaleLowerCase('de-DE');
    if(!existing.has(key)){
      customers.push({name,orders:{},page:null});
      existing.add(key);
    }
  });
  localStorage.setItem(flag,'1');
  save();
  renderCustomers();
  renderAssignments();
})();
