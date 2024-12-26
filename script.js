document.getElementById('upload').addEventListener('change', handleFile, false);
document.getElementById('save').addEventListener('click', saveFile, false);

let hot; // Declare Handsontable instance variable

// Function to handle file upload
function handleFile(e) {
    const file = e.target.files[0];
    const reader = new FileReader();
    reader.onload = function(event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        // Display the data using Handsontable
        displayTable(jsonData);
    };
    reader.readAsArrayBuffer(file);
}

// Function to display the data in Handsontable
function displayTable(data) {
    const container = document.getElementById('excelTable');
    hot = new Handsontable(container, {
        data: data,
        rowHeaders: true,
        colHeaders: true,
        stretchH: 'all', // Stretch columns to fill the available width
        autoColumnSize: true, // Enable automatic column width adjustment
        manualColumnResize: true, // Allow manual column resizing
        manualRowResize: true, // Allow manual row resizing
        wordWrap: true, // Enable word wrapping
        formulas: {
            engine: HyperFormula // Enable formulas using HyperFormula
        },
        filters: true,
        dropdownMenu: true,
        contextMenu: true,
        licenseKey: 'non-commercial-and-evaluation' // for non-commercial use only
    });
}

// Function to save the edited data and download as an Excel file
function saveFile() {
    const editedData = hot.getData();
    const worksheet = XLSX.utils.aoa_to_sheet(editedData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    XLSX.writeFile(workbook, 'EditedData.xlsx');
}