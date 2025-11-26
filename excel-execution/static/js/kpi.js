document.addEventListener('DOMContentLoaded', () => {
    const sizingSheet = document.getElementById('sizing-sheet');
    if (!sizingSheet) return;        
   
    sizingSheet.addEventListener('blur', evt => {
        const cell = evt.target.closest('[data-sheet="Sizing"][data-col]');
        if (!cell) return;          

        if (cell.dataset.col !== 'Thời điểm đẩy yêu cầu') return;

        const row = cell.closest('tr[data-row]');
        const kpiCell = row?.querySelector('[data-col="Thời gian hoàn thành theo KPI"]');
        if (!kpiCell) return;

        const raw = cell.textContent.trim();
        if (!raw) {                    
            kpiCell.textContent = '';
            kpiCell.dispatchEvent(new Event('input', {bubbles:true}));
            return;
        }

        // Hỗ trợ định dạng ISO (YYYY‑MM‑DD) hoặc dd/mm/yyyy
        let startDate = new Date(raw);
        if (isNaN(startDate)) {
            const parts = raw.split('/');
            if (parts.length === 3) {
                const d = parseInt(parts[0], 10);
                const m = parseInt(parts[1], 10) - 1;
                const y = parseInt(parts[2], 10);
                startDate = new Date(y, m, d);
            }
        }
        if (isNaN(startDate)) {
            console.warn('Định dạng ngày không hợp lệ:', raw);
            return;
        }

        const kpiDate = addWorkingDays(startDate, 3);

        const dd = String(kpiDate.getDate()).padStart(2, '0');
        const mm = String(kpiDate.getMonth() + 1).padStart(2, '0');
        const yy = kpiDate.getFullYear();
        kpiCell.textContent = `${dd}/${mm}/${yy}`;

        if (progressCell) {
            const status = computeProgressStatus(raw, kpiCell.textContent);
            progressCell.textContent = status;
            const classMap = {
                'Quá hạn': 'status-overdue',
                'Đến hạn': 'status-due',
                'Còn 1 ngày': 'status-1day',
                'Còn 2 ngày': 'status-2day',
                'Còn 3 ngày': 'status-3day',
            };
            progressCell.className = 'cell';
            if (classMap[status]) {
                progressCell.classList.add(classMap[status]);
            }
        }

        kpiCell.dispatchEvent(new Event('input', {bubbles:true}));

        if (progressCell) {
            const status = computeProgressStatus(raw, kpiCell.textContent);
            progressCell.textContent = status;
            const classMap = {
                'Quá hạn': 'status-overdue',
                'Đến hạn': 'status-due',
                'Còn 1 ngày': 'status-1day',
                'Còn 2 ngày': 'status-2day',
                'Còn 3 ngày': 'status-3day'
            };
            progressCell.className = 'cell';
            if (classMap[status]) {
                progressCell.classList.add(classMap[status]);
            }
        }

        if (row) {
            const sheetName = 'Sizing';
            const kpiCol = 'Thời gian hoàn thành theo KPI';
            const rowIndex = parseInt(row.dataset.row, 10);
            fetch('/update-cell', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    sheet: sheetName,
                    row:   rowIndex,
                    col:   kpiCol,
                    value: kpiCell.textContent.trim()
                })
            }).catch(() => console.warn('Không thể lưu KPI lên server'));
        }

        kpiCell.dispatchEvent(new Event('input', {bubbles:true}));
    }, true);  

    function addWorkingDays(date, days) {
        const result = new Date(date);
        let added = 0;
        while (added < days) {
            result.setDate(result.getDate() + 1);
            const dw = result.getDay();
            if (dw !== 0 && dw !== 6) {
                added++;
            }
        }
        return result;
    }
});