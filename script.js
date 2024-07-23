document.getElementById('upload').addEventListener('change', handleFile, false);
document.getElementById('drop-zone').addEventListener('click', () => document.getElementById('upload').click());
document.getElementById('drop-zone').addEventListener('dragover', (e) => {
    e.preventDefault();
    document.getElementById('drop-zone').classList.add('dragover');
});
document.getElementById('drop-zone').addEventListener('dragleave', (e) => {
    e.preventDefault();
    document.getElementById('drop-zone').classList.remove('dragover');
});
document.getElementById('drop-zone').addEventListener('drop', handleDrop, false);

let workbook;

function handleFile(e) {
    const files = e.target.files;
    processFile(files[0]);
    showFilePreview(files[0]);
}

function handleDrop(e) {
    e.preventDefault();
    document.getElementById('drop-zone').classList.remove('dragover');
    const files = e.dataTransfer.files;
    processFile(files[0]);
    showFilePreview(files[0]);
}

function processFile(file) {
    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        workbook = XLSX.read(data, { type: 'array' });
        showFeedback('File uploaded successfully. Ready to convert.', 'success');
    };
    reader.readAsArrayBuffer(file);
}

function showFeedback(message, type) {
    const feedback = document.getElementById('feedback');
    feedback.className = `alert alert-${type}`;
    feedback.textContent = message;
    feedback.classList.remove('d-none');
}

function showProgressBar() {
    const progressBar = document.getElementById('progress-bar');
    progressBar.classList.remove('d-none');
}

function updateProgressBar(percent) {
    const progressBarInner = document.getElementById('progress-bar-inner');
    progressBarInner.style.width = `${percent}%`;
}

function showLoadingSpinner() {
    const spinner = document.getElementById('loading-spinner');
    spinner.classList.remove('d-none');
}

function hideLoadingSpinner() {
    const spinner = document.getElementById('loading-spinner');
    spinner.classList.add('d-none');
}

function toggleDarkMode() {
    document.body.classList.toggle('dark-mode');
}

function showFilePreview(file) {
    const previewArea = document.getElementById('drop-zone');
    const fileDetails = `
        <p><strong>File name:</strong> ${file.name}</p>
        <p><strong>File size:</strong> ${(file.size / 1024).toFixed(2)} KB</p>
    `;
    previewArea.innerHTML = `
        <div class="drop-zone-inner">
            <i class="fas fa-file-excel fa-3x"></i>
            ${fileDetails}
        </div>
    `;
}

// Function to convert Excel date serial number to JavaScript Date object
function convertExcelDate(excelDate) {
    const date = new Date((excelDate - (25567 + 1)) * 86400 * 1000);
    return date;
}

// Function to format Date object to Israeli date format
function formatDateToIsraeli(date) {
    const day = String(date.getDate()).padStart(2, '0') - 1;
    const month = String(date.getMonth() + 1).padStart(2, '0'); // Months are zero-based
    const year = String(date.getFullYear()).slice(-4); // Full year
    return `${day}/${month}/${year}`;
}

// Function to format time to Israeli format
function formatTimeToIsraeli(date) {
    const hours = String(date.getHours()).padStart(2, '0') - 3;
    const minutes = String(date.getMinutes()).padStart(2, '0');
    const seconds = String(date.getSeconds()).padStart(2, '0');
    return `${hours}:${minutes}:${seconds}`;
}

function replaceDateWithTime(header) {
    return header.replace(/Date$/gi, 'Time');
}

function isExcelDate(value) {
    return typeof value === 'number' && value > 25567 && value < 2958465; // Valid Excel date range
}


function convertTime() {
    if (!workbook) {
        showFeedback('Please upload a file first.', 'warning');
        return;
    }

    showProgressBar();
    showLoadingSpinner();
    let progress = 0;
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    // Identify columns
    const headerRow = rows[0];

    // Identify all columns with "Date" in the header
    const dateColumns = [];
    headerRow.forEach((header, index) => {
        if (header.toLowerCase().includes('date')) {
            dateColumns.push(index);
        }
    });

    // Insert new headers for time columns and adjust data rows
    let offset = 0;
    dateColumns.forEach(colIndex => {
        const adjustedColIndex = colIndex + offset;
        const newHeader = replaceDateWithTime(headerRow[adjustedColIndex]);
        headerRow.splice(adjustedColIndex + 1, 0, newHeader);
        offset++;

        // Adjust each row for the new time column
        rows.forEach((row, rowIndex) => {
            if (rowIndex > 0) { // Skip header row
                const excelTime = row[adjustedColIndex];
                if (isExcelDate(excelTime)) { // Check if the cell contains a valid Excel date
                    const date = convertExcelDate(excelTime);
                    const israeliDate = formatDateToIsraeli(date);
                    const israeliTime = formatTimeToIsraeli(date);

                    // Insert the time value next to the date value
                    row.splice(adjustedColIndex + 1, 0, israeliTime);
                    row[adjustedColIndex] = israeliDate; // Place the date in the selected column
                } else {
                    // Ensure empty string in the new column if no date is present
                    row.splice(adjustedColIndex + 1, 0, '');
                }
            }
        });
    });

    // Ensure all rows have the correct number of columns
    const maxColumns = Math.max(...rows.map(row => row.length));
    rows.forEach(row => {
        while (row.length < maxColumns) {
            row.push('');
        }
    });

    const newWorksheet = XLSX.utils.aoa_to_sheet(processedRows);
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, sheetName);

    XLSX.writeFile(newWorkbook, 'converted_times.xlsx');
    showFeedback('Time conversion completed. The file is ready for download.', 'success');
    updateProgressBar(100); // Ensure the progress bar reaches 100%
    hideLoadingSpinner();
}
