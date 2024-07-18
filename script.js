document.getElementById('upload').addEventListener('change', handleFile, false);
let workbook;

function handleFile(e) {
    const files = e.target.files;
    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        workbook = XLSX.read(data, { type: 'array' });
        showFeedback('File uploaded successfully. Ready to convert.', 'success');
    };
    reader.readAsArrayBuffer(files[0]);
}

function showFeedback(message, type) {
    const feedback = document.getElementById('feedback');
    feedback.className = `alert alert-${type}`;
    feedback.textContent = message;
    feedback.classList.remove('d-none');
}

// Function to convert Excel date serial number to JavaScript Date object
function convertExcelDate(excelDate) {
    const date = new Date((excelDate - (25567 + 1)) * 86400 * 1000);
    return date;
}

// Function to format Date object to Israeli date format
function formatDateToIsraeli(date) {
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0'); // Months are zero-based
    const year = String(date.getFullYear()).slice(-2); // Last two digits of the year
    return `${day}/${month}/${year}`;
}

// Function to format time to Israeli format
function formatTimeToIsraeli(date) {
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    const seconds = String(date.getSeconds()).padStart(2, '0');
    return `${hours}:${minutes}:${seconds}`;
}

function convertTime() {
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    const timeColumnIndex = 20; // Column U (zero-indexed)

    // Handle header row
    const headerRow = rows[0];
    headerRow.splice(timeColumnIndex + 1, 0, 'Review Time');
    for (let col = headerRow.length - 2; col > timeColumnIndex + 1; col--) {
        headerRow[col + 1] = headerRow[col];
    }
    headerRow[timeColumnIndex + 1] = 'Review Time';

    rows.forEach((row, index) => {
        if (index > 0) { // Skip header row
            const excelTime = row[timeColumnIndex];
            if (typeof excelTime === 'number') { // Check if the cell contains a number (Excel date serial)
                const date = convertExcelDate(excelTime);
                const israeliDate = formatDateToIsraeli(date);
                const israeliTime = formatTimeToIsraeli(date);

                // Shift all columns to the right from timeColumnIndex + 1 onwards
                for (let col = row.length - 1; col >= timeColumnIndex + 1; col--) {
                    row[col + 1] = row[col];
                }

                row[timeColumnIndex] = israeliDate; // Place the date in column 20
                row[timeColumnIndex + 1] = israeliTime; // Place the time in column 21
            } else {
                row[timeColumnIndex] = 'Invalid Date';
                row[timeColumnIndex + 1] = 'Invalid Time';
            }
        }
    });

    const newWorksheet = XLSX.utils.aoa_to_sheet(rows);
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, sheetName);

    XLSX.writeFile(newWorkbook, 'converted_times.xlsx');
}
