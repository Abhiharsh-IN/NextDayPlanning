// Get the input form
const inputForm = document.getElementById('input-form');

// Get the today table
const todayTable = document.getElementById('today-table');

// Get the nextDay table
const nextDayTable = document.getElementById('nextDay-table');

// Get the pileWork table
const pileWorkTable = document.getElementById('pileWork-table');

// Get the add row button
const addRowButton = document.getElementById('add-row');

// Get the delete row button
const deleteRowButton = document.getElementById('delete-row');

// Get the generate excel button
const generateExcelButton = document.getElementById('generate-excel');

// Add row event handler
addRowButton.addEventListener('click', function() {
    const numberOfRows = todayTable.rows.length;
    const newRow = todayTable.insertRow(numberOfRows);

    const snCell = newRow.insertCell(0);
    snCell.textContent = numberOfRows + 1;

    for (let i = 1; i < 5; i++) {
        const cell = newRow.insertCell(i);
        const input = document.createElement('input');
        input.type = 'text';
        input.name = `description[${numberOfRows}]`;
        cell.appendChild(input);
    }
});

// Delete row event handler
deleteRowButton.addEventListener('click', function() {
    const selectedRow = todayTable.querySelector('.selected');
    
    if (selectedRow) {
        selectedRow.remove();
    }
});

// Generate excel event handler
generateExcelButton.addEventListener('click', function() {
    const workbook = XLSX.utils.book_new();
    const sheetData = [];

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
