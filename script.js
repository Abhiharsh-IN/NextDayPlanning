document.querySelectorAll('#add-row').forEach(button => {
    button.addEventListener('click', function() {
        const tableId = this.getAttribute('data-table');
        const table = document.getElementById(tableId).getElementsByTagName('tbody')[0];
        const newRow = table.insertRow(table.rows.length);
        
        const snCell = newRow.insertCell(0);
        snCell.textContent = table.rows.length;
        
        for (let i = 1; i < 5; i++) {
            const cell = newRow.insertCell(i);
            const input = document.createElement('input');
            input.type = 'text';
            input.name = `${tableId}-row${table.rows.length}[col${i}]`;
            cell.appendChild(input);
        }
    });
});

document.querySelectorAll('#delete-row').forEach(button => {
    button.addEventListener('click', function() {
        const tableId = this.getAttribute('data-table');
        const table = document.getElementById(tableId).getElementsByTagName('tbody')[0];
        
        if (table.rows.length > 1) {
            table.deleteRow(table.rows.length - 1);
        }
    });
});

document.getElementById('generate-excel').addEventListener('click', function() {
    const workbook = XLSX.utils.book_new();
    const sheetData = [];
    
    // Loop through each table
    const tableIds = ['today-table', 'nextDay-table', 'pileWork-table'];
    tableIds.forEach(tableId => {
        const table = document.getElementById(tableId);
        const tableData = [];

        for (let i = 1; i < table.rows.length; i++) {
            const rowData = [];
            for (let j = 0; j < table.rows[i].cells.length; j++) {
                const input = table.rows[i].cells[j].querySelector('input');
                rowData.push(input.value);
            }
            tableData.push(rowData);
        }
        
        sheetData.push([], [], ...tableData);
    });

    const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    
    const blob = XLSX.write(workbook, { bookType: 'xlsx', type: 'blob' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'output.xlsx';
    a.click();
});
