document.addEventListener('DOMContentLoaded', () => {
    const sizingSheet = document.getElementById('sizing-sheet');
    if (!sizingSheet) return;        
   
    sizingSheet.addEventListener('blur', evt => {
        const cell = evt.target.closest('[data-sheet="Sizing"][data-col]');
        if (!cell) return;          

        if (cell.dataset.col !== 'Thời điểm đẩy yêu cầu') return;

        const row = cell.closest('tr[data-row]');
        const kpiCell = row?.querySelector('[data-col="Thời gian hoàn thành theo KPI"]');
        const progressCell = row?.querySelector('[data-col="Tiến độ"]');
        if (!kpiCell) return;

        const raw = cell.textContent.trim();
        if (!raw) {                    
            kpiCell.textContent = '';
            kpiCell.dispatchEvent(new Event('input', {bubbles:true}));
            return;
        }

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
            const status = computeProgressStatus(kpiCell.textContent);
            const classMap = {
                'Quá hạn': 'status-overdue',
                'Đến hạn': 'status-due',
                'Còn 1 ngày': 'status-1day',
                'Còn 2 ngày': 'status-2day',
                'Còn 3 ngày': 'status-3day'
            };
            progressCell.textContent = status;
            progressCell.className = 'cell';
            if (classMap[status]) progressCell.classList.add(classMap[status]);
        }

        kpiCell.dispatchEvent(new Event('input', {bubbles:true}));

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

    function computeProgressStatus(kpiText){
        const kpi = parseVNDate(kpiText);
        if(!kpi) return '';
        const today = new Date();
        today.setHours(0,0,0,0);
        const k = new Date(kpi.getFullYear(), kpi.getMonth(), kpi.getDate());
        if(k < today) return 'Quá hạn';
        const wdays = businessDaysBetween(today, k); 
        if(wdays === 0) return 'Đến hạn';
        if(wdays === 1) return 'Còn 1 ngày';
        if(wdays === 2) return 'Còn 2 ngày';
        if(wdays === 3) return 'Còn 3 ngày';
        return '';
    }

    function businessDaysBetween(start, end){
        let cnt = 0;
        const d = new Date(start);
        while(d < end){
            const dw = d.getDay();
            if(dw !== 0 && dw !== 6){ cnt++; }
            d.setDate(d.getDate() + 1);
        }
        return cnt;
    }

    function parseVNDate(text){
        if(!text) return null;
        let parts = text.split('/');
        if(parts.length === 3){
            const d = parseInt(parts[0],10), m = parseInt(parts[1],10)-1, y = parseInt(parts[2],10);
            const dt = new Date(y,m,d);
            return isNaN(dt) ? null : dt;
        }
        const t = new Date(text);
        return isNaN(t) ? null : t;
    }

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