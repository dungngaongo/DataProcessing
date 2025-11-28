//
// Giao diện JS tổng hợp cho tương tác bảng:
// - Chuyển tab giữa các sheet
// - Lưu ô khi blur (debounce nhẹ)
// - Tự tính KPI + cập nhật ô liên quan từ "Thời điểm đẩy yêu cầu"
// - Điều hướng/mapping dự án giữa các sheet và highlight dòng mục tiêu
// - Đổi tên cột ngay trên header
//

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

// Lưu ô khi blur, và xử lý lan truyền KPI cho Sizing
document.addEventListener('blur', function(e){
  if(e.target.classList && e.target.classList.contains('cell')){
    const sheet = e.target.dataset.sheet;
    const row = parseInt(e.target.dataset.row, 10);
    const rowId = e.target.dataset.rowId || (e.target.closest('tr[data-row-id]')?.dataset.rowId);
    const col = e.target.dataset.col;
    const value = e.target.textContent.trim();

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

    if(sheet === 'Sizing' && col === 'Thời điểm đẩy yêu cầu'){
      const kpiCol = 'Thời gian hoàn thành theo KPI';
      let date = parseDateVN(value);
      if(date){
        let daysAdded = 0;
        while(daysAdded < 3){
          date.setDate(date.getDate() + 1);
          const day = date.getDay();
          if(day !== 0 && day !== 6){
            daysAdded++;
          }
        }
        const kpiDate = formatDateVN(date);
        fetch('/update-cell', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ sheet, row, rowId, col: kpiCol, value: kpiDate })
        }).catch(() => {});
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

// Tiện ích parse/format ngày theo dạng Việt Nam (dd/mm/yyyy)
function parseDateVN(str){
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

// (ĐÃ BỎ) Logic mapping giữa các bảng được gỡ bỏ theo yêu cầu.

// Đổi tên cột trực tiếp trên header, đồng bộ lên server
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