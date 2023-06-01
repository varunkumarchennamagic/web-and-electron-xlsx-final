// Get references to the HTML elements
const fileInput = document.getElementById('xlsxFile');
const convertBtn = document.getElementById('convertBtn');
const downloadBtn = document.getElementById('downloadBtn');
const sheetList = document.getElementById('sheetList');
const selectAllBtn = document.getElementById('selectAllBtn');
const clearSelectionBtn = document.getElementById('clearSelectionBtn');
const progressBar = document.getElementById('progressBar');
const progressLog = document.getElementById('progressLog');

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
    const allowedExceptions = ['-', '_', ' ', '.', '\í', '\ó', '\é'];

    // Check for special characters, cell length, and math formulas
    const invalidChars = new RegExp(`[^\\w${allowedExceptions.join('\\\\')}]`, 'g');
    const mathFormulaRegex = /^=/; // Regex to match math formulas
    let invalidCells = [];

    // Convert the selected sheets to JSON
    selectedSheets.forEach((sheetName, index) => {
      const worksheet = workbook.Sheets[sheetName];
      const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      sheetData.forEach((row, rowIndex) => {
        row.forEach((cell, columnIndex) => {
          if (typeof cell === 'string') {
            var invalidCharFlag = invalidChars.test(cell);
            if(invalidCharFlag) {
              var tempFlag = false
              console.log(cell.match(invalidChars))
              var invalidArr = cell.match(invalidChars)
              invalidArr.forEach(char => {
                console.log(allowedExceptions.includes(char))
                if(!allowedExceptions.includes(char)){
                  tempFlag = true
                }
              });
              invalidCharFlag = tempFlag
            }
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

      const jsonContent = JSON.stringify(sheetData, null, 2);
      const blob = new Blob([jsonContent], { type: 'application/json' });
      const fileName = `${sheetName}.json`;

      // Create a download link for each sheet
      const downloadLink = document.createElement('a');
      downloadLink.href = URL.createObjectURL(blob);
      downloadLink.download = fileName;
      downloadLink.textContent = fileName;

      // Append the download link to the sheet list
      sheetList.appendChild(downloadLink);

      // Append a line break after each download link
      sheetList.appendChild(document.createElement('br'));

      // Update progress bar
      const progress = ((index + 1) / selectedSheets.length) * 100;
      progressBar.style.width = `${progress}%`;

      // Update progress log
      progressLog.textContent = `Converting sheet ${index + 1} of ${selectedSheets.length}...`;
    });

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

    // Show the download button
    downloadBtn.style.display = 'block';
  };
  reader.readAsArrayBuffer(file);
}
