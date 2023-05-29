// Get references to the HTML elements
const fileInput = document.getElementById('xlsxFile');
const convertBtn = document.getElementById('convertBtn');
const downloadBtn = document.getElementById('downloadBtn');

// Add event listener to the convert button
convertBtn.addEventListener('click', convertToJSON);

function convertToJSON() {
  // Get the uploaded file
  const file = fileInput.files[0];

  // Read the file data
  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });

    // Convert the first sheet to JSON
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    // Define the list of allowed exceptions
    const allowedExceptions = ['-', '_', ' ', '.'];

    // Check for special characters, cell length, and math formulas
    const invalidChars = new RegExp(`[^\\w${allowedExceptions.join('\\\\')}]`, 'g');
    const mathFormulaRegex = /^=/; // Regex to match math formulas
    let invalidCells = [];

    jsonData.forEach((row, rowIndex) => {
      row.forEach((cell, columnIndex) => {
        if (typeof cell === 'string') {
          if (invalidChars.test(cell) || cell.length > 128 || mathFormulaRegex.test(cell)) {
            const invalidChar = invalidChars.test(cell) ? cell.match(invalidChars)[0] : null;
            const columnHeader = jsonData[0][columnIndex]; // Get the header from the first row
            invalidCells.push({
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
      let errorMessage = 'Error: Invalid data found in the following cells:\n';

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

