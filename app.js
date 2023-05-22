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
    const jsonData = XLSX.utils.sheet_to_json(worksheet);

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

