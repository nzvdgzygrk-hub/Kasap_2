// Letzter Excel-Feinschliff: Firma links, Tavuk Listesi mittig, Datum rechts – wie in der Vorlage.
makeTavukSheet=function(){
 const cols=activeTavukColumns(),width=Math.max(3,cols.length+1),mid=Math.floor((width-1)/2);
 const top=Array(width).fill('');top[0]=companyName.value||company;top[mid]=t().tavuk;top[width-1]=today();
 const header=[t().customer,...cols.map(x=>x.label)];while(header.length<width)header.push('');
 const rows=[top,header];
 const totals=cols.map(()=>({kg:0,count:0}));
 activeCustomers().forEach(c=>{
  const vals=cols.map((x,k)=>{const o=orderAt(c.orders?.[x.i]);totals[k].kg+=Number(o.kg)||0;totals[k].count+=Number(o.count)||0;return valueForOverview(o)});
  if(vals.some(Boolean)){const row=[c.name,...vals];while(row.length<width)row.push('');rows.push(row)}
 });
 const totalRow=[t().total,...totals.map(valueForOverview)];while(totalRow.length<width)totalRow.push('');rows.push(totalRow);
 const ws=XLSX.utils.aoa_to_sheet(rows);ws['!cols']=[{wch:24},...Array(width-1).fill(0).map(()=>({wch:13}))];
 styleCell(ws.A1,{font:{bold:true},alignment:{horizontal:'left'}});
 styleCell(ws[XLSX.utils.encode_cell({r:0,c:mid})],{font:{bold:true,sz:14},alignment:{horizontal:'center'}});
 styleCell(ws[XLSX.utils.encode_cell({r:0,c:width-1})],{font:{bold:true},alignment:{horizontal:'right'}});
 const range=XLSX.utils.decode_range(ws['!ref']);
 for(let c=0;c<=range.e.c;c++)styleCell(ws[XLSX.utils.encode_cell({r:1,c})],S.header);
 for(let r=2;r<range.e.r;r++)for(let c=0;c<=range.e.c;c++)styleCell(ws[XLSX.utils.encode_cell({r,c})],S.normal);
 for(let c=0;c<=range.e.c;c++)styleCell(ws[XLSX.utils.encode_cell({r:range.e.r,c})],S.total);
 ws['!pageSetup']={orientation:'landscape',fitToWidth:1,fitToHeight:1,paperSize:9};return ws;
};
