// Get references to the HTML elements
const fileInput = document.getElementById('xlsxFile');
const convertBtn = document.getElementById('convertBtn');
const downloadBtn = document.getElementById('downloadBtn');
const selectAllBtn = document.getElementById('selectAllBtn');
const clearSelectionBtn = document.getElementById('clearSelectionBtn');
const sheetListContainer = document.getElementById('sheetListContainer');

let selectedSheets = [];

// Add event listener to the convert button
convertBtn.addEventListener('click', convertToJSON);

// Add event listener to the select all button
selectAllBtn.addEventListener('click', selectAllSheets);

// Add event listener to the clear selection button
clearSelectionBtn.addEventListener('click', clearSheetSelection);

// Update sheet selection when file input changes
fileInput.addEventListener('change', updateSheetList);

function updateSheetList() {
  const file = fileInput.files[0];
  const reader = new FileReader();

  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });

    // Clear previous sheet list
    sheetListContainer.innerHTML = '';

    // Add checkboxes for each sheet
    workbook.SheetNames.forEach(sheetName => {
      const sheetItem = document.createElement('div');
      sheetItem.classList.add('sheet-item');

      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.value = sheetName;
      checkbox.checked = true;
      checkbox.addEventListener('change', handleSheetSelection);

      const label = document.createElement('label');
      label.innerText = sheetName;

      sheetItem.appendChild(checkbox);
      sheetItem.appendChild(label);

      sheetListContainer.appendChild(sheetItem);

      // Add sheet to selected sheets
      selectedSheets.push(sheetName);
    });
  };

  reader.readAsArrayBuffer(file);
}

function handleSheetSelection(e) {
  const sheetName = e.target.value;
  if (e.target.checked) {
    selectedSheets.push(sheetName);
  } else {
    const index = selectedSheets.indexOf(sheetName);
    if (index > -1) {
      selectedSheets.splice(index, 1);
    }
  }
}

function selectAllSheets() {
  const checkboxes = sheetListContainer.querySelectorAll('input[type="checkbox"]');
  checkboxes.forEach(checkbox => {
    checkbox.checked = true;
    selectedSheets.push(checkbox.value);
  });
}

function clearSheetSelection() {
  const checkboxes = sheetListContainer.querySelectorAll('input[type="checkbox"]');
  checkboxes.forEach(checkbox => {
    checkbox.checked = false;
  });
  selectedSheets = [];
}

function convertToJSON() {
  // Get the uploaded file
  const file = fileInput.files[0];

  // Read the file data
  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });

    // Convert the selected sheets to JSON
    let jsonData = [];
    selectedSheets.forEach(sheetName => {
      const worksheet = workbook.Sheets[sheetName];
      const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      // Define the list of allowed exceptions
      const allowedExceptions = ['-', '_', ' ', '.'];

      // Check for special characters, cell length, and math formulas
      const invalidChars = new RegExp(`[^\\w${allowedExceptions.join('\\\\')}]`, 'g');
      const mathFormulaRegex = /^=/; // Regex to match math formulas
      let invalidCells = [];

      sheetData.forEach((row, rowIndex) => {
        row.forEach((cell, columnIndex) => {
          if (typeof cell === 'string') {
            const invalidCharFlag = invalidChars.test(cell);
            if (invalidCharFlag || cell.length > 128 || mathFormulaRegex.test(cell)) {
              const invalidChar = invalidCharFlag ? cell.match(invalidChars) : null;
              const columnHeader = sheetData[0][columnIndex]; // Get the header from the first row
              invalidCells.push({
                sheet: sheetName,
                row: rowIndex + 1,
                column: columnHeader,
                character: invalidChar,
                lengthExceeded: cell.length > 128,
                hasMathFormula: mathFormulaRegex.test(cell),
              });
            }
          }
        });
      });

      if (invalidCells.length > 0) {
        let errorMessage = `Error: Invalid data found in the following cells (Sheet: ${sheetName}):\n`;

        invalidCells.forEach((cell) => {
          errorMessage += `Row: ${cell.row}, Column: ${cell.column}`;

          if (cell.character) {
            errorMessage += `, Character: "${cell.character}"`;
          }

          if (cell.lengthExceeded) {
            errorMessage += `, Length Exceeded`;
          }

          if (cell.hasMathFormula) {
            errorMessage += `, Math Formula Detected`;
          }

          errorMessage += '\n';
        });

        // Show user-friendly error message
        alert('Invalid data found in the uploaded file. Please check the error log for more details.');

        // Write detailed error to error log file
        const errorLog = `logs/error_${new Date().toISOString()}.txt`;
        const errorContent = `Error Log - ${new Date().toISOString()}\n\n${errorMessage}`;

        const blob = new Blob([errorContent], { type: 'text/plain;charset=utf-8' });
        saveAs(blob, errorLog);

        return;
      }

      // Merge the current sheet data with the existing JSON data
      jsonData = jsonData.concat(sheetData);
    });

    if (jsonData.length === 0) {
      alert('No valid data found in the selected sheets.');
      return;
    }

    // Enable the download button
    downloadBtn.disabled = false;

    // Handle the download button click
    downloadBtn.addEventListener('click', function () {
      const jsonContent = JSON.stringify(jsonData, null, 2);
      const blob = new Blob([jsonContent], { type: 'application/json' });
      saveAs(blob, 'data.json');
    });
  };
  reader.readAsArrayBuffer(file);
}
