// Get references to the HTML elements
const fileInput = document.getElementById('xlsxFile');
const convertBtn = document.getElementById('convertBtn');
const downloadBtn = document.getElementById('downloadBtn');
const sheetList = document.getElementById('sheetList');
const selectAllBtn = document.getElementById('selectAllBtn');
const clearSelectionBtn = document.getElementById('clearSelectionBtn');

// Add event listener to the convert button
convertBtn.addEventListener('click', convertToJSON);

// Handle file selection
fileInput.addEventListener('change', handleFileSelect);

// Handle select all button
selectAllBtn.addEventListener('click', selectAllSheets);

// Handle clear selection button
clearSelectionBtn.addEventListener('click', clearSheetSelection);

function handleFileSelect() {
  // Get the uploaded file
  const file = fileInput.files[0];

  // Read the file data
  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });

    // Display sheet list
    const sheetNames = workbook.SheetNames;
    sheetList.innerHTML = '';
    sheetNames.forEach((sheetName) => {
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.name = 'selectedSheets';
      checkbox.value = sheetName;
      checkbox.checked = true;

      const label = document.createElement('label');
      label.appendChild(checkbox);
      label.appendChild(document.createTextNode(sheetName));

      sheetList.appendChild(label);
    });
  };
  reader.readAsArrayBuffer(file);
}

function selectAllSheets() {
  const checkboxes = document.querySelectorAll('input[name="selectedSheets"]');
  checkboxes.forEach((checkbox) => {
    checkbox.checked = true;
  });
}

function clearSheetSelection() {
  const checkboxes = document.querySelectorAll('input[name="selectedSheets"]');
  checkboxes.forEach((checkbox) => {
    checkbox.checked = false;
  });
}

function convertToJSON() {
  // Get the uploaded file
  const file = fileInput.files[0];

  // Read the file data
  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });

    // Get the selected sheets
    const selectedSheets = Array.from(document.querySelectorAll('input[name="selectedSheets"]:checked')).map(
      (checkbox) => checkbox.value
    );

    // Define the list of allowed exceptions
    const allowedExceptions = ['-', '_', ' ', '.'];

    // Check for special characters, cell length, and math formulas
    const invalidChars = new RegExp(`[^\\w${allowedExceptions.join('\\\\')}]`, 'g');
    const mathFormulaRegex = /^=/; // Regex to match math formulas
    let invalidCells = [];

    // Convert the selected sheets to JSON
    const jsonData = selectedSheets.reduce((result, sheetName) => {
      const worksheet = workbook.Sheets[sheetName];
      const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      sheetData.forEach((row, rowIndex) => {
        row.forEach((cell, columnIndex) => {
          if (typeof cell === 'string') {
            const invalidCharFlag = invalidChars.test(cell);
            if (invalidCharFlag || cell.length > 128 || mathFormulaRegex.test(cell)) {
              const invalidChar = invalidCharFlag ? cell.match(invalidChars) : null;
              const columnHeader = sheetData[0][columnIndex]; // Get the header from the
              invalidCells.push({
                sheet: sheetName,
                row: rowIndex + 1,
                column: columnHeader,
                character: invalidChar ? invalidChar[0] : null,
                lengthExceeded: cell.length > 128,
                hasMathFormula: mathFormulaRegex.test(cell),
              });
            }
          }
        });
      });

      result[sheetName] = sheetData;
      return result;
    }, {});

    if (invalidCells.length > 0) {
      let errorMessage = 'Error: Invalid data found in the following cells:\n';

      invalidCells.forEach((cell) => {
        errorMessage += `Sheet: ${cell.sheet}, Row: ${cell.row}, Column: ${cell.column}`;

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
