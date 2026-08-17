// Sayfa-Zuordnung: immer alle Kunden anzeigen, inklusive später neu hinzugefügter Kunden.
renderAssignments = function(){
  if(!customers.length){
    assignmentGrid.innerHTML=`<div class="empty-assign">${escapeHtml(t().none)}</div>`;
    assignmentStatus.className='status';
    assignmentStatus.textContent=lang==='de'?'Sayfa ist optional.':'Sayfa isteğe bağlıdır.';
    return;
  }
  let html='<div class="assign-list">';
  customers.forEach((c,i)=>{
    html+=`<div class="assign-row"><div class="assign-customer">${escapeHtml(c.name)}</div><select class="assign-select" aria-label="${escapeHtml(c.name)} Sayfa" onchange="setPage(${i},Number(this.value))"><option value="">${lang==='de'?'Keine Sayfa':'Sayfa yok'}</option>`;
    for(let p=1;p<=10;p++) html+=`<option value="${p}" ${c.page===p?'selected':''}>Sayfa ${p}</option>`;
    html+='</select></div>';
  });
  assignmentGrid.innerHTML=html+'</div>';
  const assigned=customers.filter(c=>c.page).length;
  assignmentStatus.className='status';
  assignmentStatus.textContent=lang==='de'?`${assigned} von ${customers.length} Kunden haben eine Sayfa-Zuordnung. Sayfa ist optional.`:`${customers.length} müşteriden ${assigned} tanesine Sayfa atandı. Sayfa isteğe bağlıdır.`;
};
renderAssignments();
