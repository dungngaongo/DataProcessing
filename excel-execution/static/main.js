document.addEventListener('click', function(e){
  if(e.target.classList.contains('sheet-tab')){
    document.querySelectorAll('.sheet-tab').forEach(btn => btn.classList.remove('active'));
    e.target.classList.add('active');
    const sheetId = e.target.getAttribute('data-sheet');
    document.querySelectorAll('.sheet').forEach(s => s.classList.remove('active'));
    const activeSheet = document.getElementById(sheetId);
    if(activeSheet){ activeSheet.classList.add('active'); }
  }
});

// Cell update logic
document.addEventListener('blur', function(e){
  if(e.target.classList && e.target.classList.contains('cell')){
    const sheet = e.target.dataset.sheet;
    const row = parseInt(e.target.dataset.row, 10);
    const col = e.target.dataset.col;
    const value = e.target.textContent.trim();
    fetch('/update-cell', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ sheet, row, col, value })
    });
  }
}, true);

// Project mapping
document.addEventListener('click', function(e){
  if(e.target.classList && e.target.classList.contains('project-link')){
    const projectName = (e.target.dataset.project || '').trim();
    if(!projectName){ return; }
    const capTab = document.querySelector('.sheet-tab[data-sheet="cap-phat-sheet"]');
    if(capTab){ capTab.click(); }
    setTimeout(() => {
      const rows = document.querySelectorAll('#cap-phat-sheet .data-row');
      let targetRow = null;
      const lowerName = projectName.toLowerCase();
      rows.forEach(r => {
        if(targetRow) return;
        const projCell = r.querySelector('td[data-col="Dự án"]');
        if(projCell){
          const text = projCell.textContent.trim().toLowerCase();
          if(text){
            if(text === lowerName || text.includes(lowerName)){
              targetRow = r;
            }
          }
        }
      });
      if(targetRow){
        targetRow.scrollIntoView({ behavior: 'smooth', block: 'center' });
        targetRow.classList.add('mapped-row');
        targetRow.querySelectorAll('td').forEach(td => td.classList.add('mapped-highlight'));
        setTimeout(() => {
          targetRow.querySelectorAll('td').forEach(td => td.classList.remove('mapped-highlight'));
          targetRow.classList.remove('mapped-row');
        }, 2500);
      }
    }, 80);
  }
});

// CapPhat -> ChiTiet mapping on project click
document.addEventListener('click', function(e){
  if(e.target.classList && e.target.classList.contains('cap-project-link')){
    const projectName = ((e.target.dataset.capProject || e.target.textContent) || '').trim()
    if(!projectName){ return; }
    const chiTietTab = document.querySelector('.sheet-tab[data-sheet="chi-tiet-sheet"]');
    if(chiTietTab){ chiTietTab.click(); }
    setTimeout(() => {
      const rows = document.querySelectorAll('#chi-tiet-sheet .data-row');
      let targetRow = null;
      const lowerName = projectName.toLowerCase();
      rows.forEach(r => {
        if(targetRow) return;
        const projCell = r.querySelector('td[data-col="Dự án"]');
        if(projCell){
          const text = (projCell.textContent || '').trim().toLowerCase();
          if(text){
            if(text === lowerName || text.includes(lowerName)){
              targetRow = r;
            }
          }
        }
      });
      if(targetRow){
        targetRow.scrollIntoView({ behavior: 'smooth', block: 'center' });
        targetRow.classList.add('mapped-row');
        targetRow.querySelectorAll('td').forEach(td => td.classList.add('mapped-highlight'));
        setTimeout(() => {
          targetRow.querySelectorAll('td').forEach(td => td.classList.remove('mapped-highlight'));
          targetRow.classList.remove('mapped-row');
        }, 2500);
      }
    }, 80);
  }
});

// Column name update logic
document.addEventListener('blur', function(e){
  if(e.target.classList && e.target.classList.contains('col-header-name')){
    const sheet = e.target.dataset.sheet;
    const oldCol = e.target.dataset.oldColName;
    const newCol = e.target.textContent.trim();

    if(!newCol || newCol === oldCol) {
        e.target.textContent = oldCol;
        return;
    }

    fetch('/update-col-name', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ sheet, oldCol, newCol })
    })
    .then(response => {
        if(response.ok) {
            e.target.dataset.oldColName = newCol;
        } else {
            alert('Lỗi cập nhật tên cột!');
            e.target.textContent = oldCol;
        }
    })
    .catch(() => {
        alert('Lỗi mạng khi cập nhật tên cột.');
        e.target.textContent = oldCol;
    });
  }
}, true);