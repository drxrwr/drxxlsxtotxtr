// Initialize variables
let workbook = null;
let fileInput = document.getElementById('fileInput');
let convertBtn = document.getElementById('convertBtn');
let sheetSelector = document.getElementById('sheetSelector');
let resultSection = document.getElementById('resultSection');
let loading = document.getElementById('loading');

// Handle file upload
fileInput.addEventListener('change', function (e) {
    let file = e.target.files[0];
    if (file && file.name.endsWith('.xlsx')) {
        // Show loading indicator while processing
        loading.style.display = 'block';

        // Read the uploaded file with SheetJS
        let reader = new FileReader();
        reader.onload = function (event) {
            let data = new Uint8Array(event.target.result);
            workbook = XLSX.read(data, { type: 'array' });
            
            // Hide loading indicator
            loading.style.display = 'none';

            // Populate sheet selector
            populateSheetSelector();
        };
        reader.readAsArrayBuffer(file);
    } else {
        alert('Harap unggah file dengan format XLSX!');
    }
});

// Populate sheet selector dropdown
function populateSheetSelector() {
    sheetSelector.innerHTML = '';  // Clear previous sheet options
    let sheets = workbook.SheetNames;

    if (sheets.length > 1) {
        let select = document.createElement('select');
        select.id = 'sheetSelect';

        sheets.forEach((sheet, index) => {
            let option = document.createElement('option');
            option.value = index;
            option.innerText = sheet;
            select.appendChild(option);
        });

        // Add select dropdown to the sheet selector
        sheetSelector.appendChild(select);
        select.addEventListener('change', displayColumns);
        displayColumns(); // Display columns for the first sheet
    } else {
        // Directly process if only one sheet
        displayColumns();
    }
}

// Display columns based on the selected sheet
function displayColumns() {
    let sheetIndex = document.getElementById('sheetSelect') ? document.getElementById('sheetSelect').value : 0;
    let sheet = workbook.Sheets[workbook.SheetNames[sheetIndex]];

    let rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    resultSection.innerHTML = '';  // Clear previous results

    // Display columns and generate download links
    rows[0].forEach((colName, colIndex) => {
        let colData = rows.map(row => row[colIndex] || '').filter(cell => cell !== '');  // Get non-empty rows
        
        let colDiv = document.createElement('div');
        let fileName = `${workbook.SheetNames[sheetIndex]}_${String.fromCharCode(65 + colIndex)}_${workbook.SheetNames[sheetIndex]}.txt`;
        let rowCount = colData.length;
        
        colDiv.innerHTML = `
            <h3>${fileName}</h3>
            <p>${rowCount} baris terisi</p>
            <button onclick="downloadTXT('${fileName}', ${colIndex}, ${sheetIndex})">Unduh</button>
        `;
        
        resultSection.appendChild(colDiv);
    });
}

// Generate and download TXT file for a column
function downloadTXT(fileName, colIndex, sheetIndex) {
    let sheet = workbook.Sheets[workbook.SheetNames[sheetIndex]];
    let rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    let colData = rows.map(row => row[colIndex] || '');

    let textContent = colData.join('\n');
    let blob = new Blob([textContent], { type: 'text/plain' });
    let url = URL.createObjectURL(blob);
    
    let link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    link.click();
    URL.revokeObjectURL(url);  // Clean up the URL
}
