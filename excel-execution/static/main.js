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
    const rowId = e.target.dataset.rowId || (e.target.closest('tr[data-row-id]')?.dataset.rowId);
    const col = e.target.dataset.col;
    const value = e.target.textContent.trim();

    // Debounce quick successive blurs to avoid out-of-order updates
    if(!window.__cellUpdateTimers){ window.__cellUpdateTimers = new Map(); }
    const key = `${sheet}:${rowId || row}:${col}`;
    const prev = window.__cellUpdateTimers.get(key);
    if(prev){ clearTimeout(prev); }
    window.__cellUpdateTimers.set(key, setTimeout(() => {
      fetch('/update-cell', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ sheet, row, rowId, col, value })
      }).catch(() => {});
      window.__cellUpdateTimers.delete(key);
    }, 200));

    // Nếu là Sizing và cột Thời điểm đẩy yêu cầu thì tự động tính KPI
    if(sheet === 'Sizing' && col === 'Thời điểm đẩy yêu cầu'){
      // Tính ngày làm việc thứ 3 tiếp theo
      const kpiCol = 'Thời gian hoàn thành theo KPI';
      let date = parseDateVN(value);
      if(date){
        let daysAdded = 0;
        while(daysAdded < 3){
          date.setDate(date.getDate() + 1);
          const day = date.getDay();
          if(day !== 0 && day !== 6){ // 0: CN, 6: T7
            daysAdded++;
          }
        }
        const kpiDate = formatDateVN(date);
        // Gửi lên server cập nhật KPI
        // Update KPI using stable rowId
        fetch('/update-cell', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ sheet, row, rowId, col: kpiCol, value: kpiDate })
        }).catch(() => {});
        // Cập nhật trực tiếp giao diện
        const table = document.getElementById('sizing-sheet');
        if(table){
          const selector = rowId ? `.data-row[data-row-id="${rowId}"] td[data-col="${kpiCol}"]` : `.data-row[data-row="${row}"] td[data-col="${kpiCol}"]`;
          const kpiCell = table.querySelector(selector);
          if(kpiCell){
            kpiCell.textContent = kpiDate;
          }
        }
      }
    }
  }
}, true);

function parseDateVN(str){
  // Chấp nhận dd/mm/yyyy hoặc yyyy-mm-dd
  if(!str) return null;
  let parts = str.split('/');
  if(parts.length === 3){
    let d = parseInt(parts[0],10), m = parseInt(parts[1],10)-1, y = parseInt(parts[2],10);
    if(!isNaN(d) && !isNaN(m) && !isNaN(y)) return new Date(y,m,d);
  }
  parts = str.split('-');
  if(parts.length === 3){
    let y = parseInt(parts[0],10), m = parseInt(parts[1],10)-1, d = parseInt(parts[2],10);
    if(!isNaN(d) && !isNaN(m) && !isNaN(y)) return new Date(y,m,d);
  }
  return null;
}

function formatDateVN(date){
  if(!(date instanceof Date)) return '';
  let d = date.getDate().toString().padStart(2,'0');
  let m = (date.getMonth()+1).toString().padStart(2,'0');
  let y = date.getFullYear();
  return `${d}/${m}/${y}`;
}

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