// Get references to the HTML elements
const fileInput = document.getElementById('xlsxFile');
const convertBtn = document.getElementById('convertBtn');
const downloadBtn = document.getElementById('downloadBtn');
const sheetList = document.getElementById('sheetList');
const downloadLinks = document.getElementById('downloadLinks');
const selectAllBtn = document.getElementById('selectAllBtn');
const clearSelectionBtn = document.getElementById('clearSelectionBtn');
const progressBar = document.getElementById('progressBar');
const progressLog = document.getElementById('progressLog');
const progressPercentage = document.getElementById('progressPercentage');
let xlsxFileName = '';
let invalidCells = []; // Declare invalidCells as a global variable

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
  xlsxFileName = file.name;
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
      label.htmlFor = sheetName;
      label.appendChild(checkbox);
      label.appendChild(document.createTextNode(sheetName));

      const sheetStatus = document.createElement('span');
      sheetStatus.className = 'sheet-status';

      label.appendChild(sheetStatus);

      sheetList.appendChild(label);
      sheetList.appendChild(document.createElement('br'));
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
  progressBar.querySelector('span').style.width = '0px';
  progressLog.innerHTML = '';
  progressPercentage.textContent = '0%';
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
    const allowedExceptions = ['-', '_', ' ', '.', '\u00ED', '\u00F3', '\u00E9']; // Include í, ó, and é as exceptions

    // Check for special characters, cell length, and math formulas
    const invalidChars = new RegExp(`[^a-zA-Z0-9${allowedExceptions.join('\\')}]`, 'g');
    const mathFormulaRegex = /^=/;

    // Initialize progress tracking variables
    let currentSheetIndex = 0;
    let totalSheets = selectedSheets.length;

    // Start conversion process
    convertSheetToJson(workbook, selectedSheets, currentSheetIndex, totalSheets);

    function convertSheetToJson(workbook, selectedSheets, currentSheetIndex, totalSheets) {
      const sheetName = selectedSheets[currentSheetIndex];
      const sheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(sheet);

      // Check for invalid cells
      const cellsWithErrors = [];
      jsonData.forEach((row, rowIndex) => {
        for (const [key, value] of Object.entries(row)) {
          if (typeof value === 'string' && value.match(invalidChars)) {
            cellsWithErrors.push(`${XLSX.utils.encode_col(key)}${rowIndex + 2}`);
          } else if (typeof value === 'string' && value.match(mathFormulaRegex)) {
            cellsWithErrors.push(`${XLSX.utils.encode_col(key)}${rowIndex + 2}`);
          }
        }
      });

      if (cellsWithErrors.length > 0) {
        invalidCells = invalidCells.concat(cellsWithErrors);
        const sheetStatus = createSheetStatusElement('Failed');
        updateSheetStatus(sheetName, sheetStatus);
      } else {
        const sheetStatus = createSheetStatusElement('Success');
        updateSheetStatus(sheetName, sheetStatus);

        const jsonDataStr = JSON.stringify(jsonData, null, 2);
        const blob = new Blob([jsonDataStr], { type: 'application/json' });
        const downloadLink = document.createElement('a');
        downloadLink.href = URL.createObjectURL(blob);
        downloadLink.download = `${xlsxFileName}_${sheetName}.json`;
        downloadLink.textContent = `${xlsxFileName}_${sheetName}.json`;
        downloadLink.classList.add('download-link');
        downloadLinks.appendChild(downloadLink);
        downloadLinks.appendChild(document.createElement('br'));
      }

      // Update progress bar and log
      const progressPercentageValue = ((currentSheetIndex + 1) / totalSheets) * 100;
      progressBar.querySelector('span').style.width = `${progressPercentageValue}%`;
      progressPercentage.textContent = `${Math.round(progressPercentageValue)}%`;
      progressLog.innerHTML += `Converted sheet: ${sheetName}<br>`;

      currentSheetIndex++;

      if (currentSheetIndex < totalSheets) {
        setTimeout(() => {
          convertSheetToJson(workbook, selectedSheets, currentSheetIndex, totalSheets);
        }, 100);
      } else {
        // Conversion completed
        if (invalidCells.length > 0) {
          const errorLogBtn = document.createElement('button');
          errorLogBtn.textContent = 'Download Error Log';
          errorLogBtn.addEventListener('click', downloadErrorLog);
          downloadLinks.appendChild(errorLogBtn);
        }
      }
    }
  };
  reader.readAsArrayBuffer(file);
}

function createSheetStatusElement(status) {
  const sheetStatus = document.createElement('span');
  sheetStatus.className = 'sheet-status';

  if (status === 'Failed') {
    sheetStatus.textContent = 'Failed';
    sheetStatus.classList.add('failed');
  } else if (status === 'Success') {
    sheetStatus.textContent = 'Success';
    sheetStatus.classList.add('success');
  }

  return sheetStatus;
}

function updateSheetStatus(sheetName, sheetStatusElement) {
  const label = sheetList.querySelector(`label[for="${sheetName}"]`);
  const existingSheetStatus = label.querySelector('.sheet-status');
  if (existingSheetStatus) {
    label.removeChild(existingSheetStatus);
  }
  label.appendChild(sheetStatusElement);
}

function downloadErrorLog() {
  const errorLogContent = invalidCells.join('\n');
  const blob = new Blob([errorLogContent], { type: 'text/plain' });
  const downloadLink = document.createElement('a');
  downloadLink.href = URL.createObjectURL(blob);
  downloadLink.download = 'error_log.txt';
  downloadLink.textContent = 'error_log.txt';
  downloadLink.classList.add('download-link');
  downloadLinks.appendChild(downloadLink);
}
