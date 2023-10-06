const xlsx = require("xlsx");

const inputExcel = document.getElementById("input-excel");
const processFilesButton = document.getElementById("process-files");
const saveToExcelButton = document.getElementById("save-to-excel");
const columnNamesInput = document.getElementById("column-names");
const refreshButton = document.getElementById('refresh');

let collectedData = [];
let uniqueData = []; // store this at the top scope to access in both event handlers

refreshButton.addEventListener('click', () => {
    // Reset global variables
    collectedData = [];
    uniqueData = [];

    // Clear the input field and table
    columnNamesInput.value = '';
    inputExcel.value = '';  // Clear file input selection
    const table = document.getElementById('preview-table');
    table.querySelector('thead').innerHTML = '';
    table.querySelector('tbody').innerHTML = '';

    alert('Refreshed and ready for new input!');
});

processFilesButton.addEventListener("click", () => {
  collectedData = []; // Reset the data on every click

  const columns = columnNamesInput.value
    .split(",")
    .map((column) => column.trim());
  if (!columns.length) {
    alert("Please specify at least one column name for filtering.");
    return;
  }

  collectedData = []; // Reset the data on every click

  const files = inputExcel.files;
  if (files.length) {
    Array.from(files).forEach((file) => {
      const reader = new FileReader();
      reader.onload = function (e) {
        const data = e.target.result;
        const workbook = xlsx.read(data, { type: "binary" });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = xlsx.utils.sheet_to_json(worksheet);
        collectedData = collectedData.concat(jsonData);
        // Filter logic moved here:
        uniqueData = filterUniqueRows(collectedData, columns);

        // Show a preview in the table (for example, first 5 rows)
        displayPreview(uniqueData.slice(0, 5));
      };
      reader.readAsBinaryString(file);
    });
  }
});

saveToExcelButton.addEventListener("click", () => {
  if (!collectedData.length) {
    alert("No data to save!");
    return;
  }

  const columns = columnNamesInput.value
    .split(",")
    .map((column) => column.trim());
  if (!columns.length) {
    alert("Please specify at least one column name for filtering.");
    return;
  }

  const uniqueData = filterUniqueRows(collectedData, columns);

  // New alert to display the numbers:
  alert(
    `total rows: ${collectedData.length}\n filtered out: ${uniqueData.length}`
  );

  const newWorkbook = xlsx.utils.book_new();
  const newWorksheet = xlsx.utils.json_to_sheet(uniqueData);
  xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, "Sheet1");
  xlsx.writeFile(newWorkbook, "filtered_data.xlsx");
});

function filterUniqueRows(data, columns) {
  const seen = new Set();
  return data.filter((row) => {
    const key = columns.map((col) => row[col]).join("|");
    if (seen.has(key)) {
      return false;
    }
    seen.add(key);
    return true;
  });
}

function displayPreview(data) {
    const table = document.getElementById('preview-table');
    const thead = table.querySelector('thead');
    const tbody = table.querySelector('tbody');

    // Clear existing rows
    thead.innerHTML = '';
    tbody.innerHTML = '';

    // Display headers
    const headers = Object.keys(data[0]);
    const headerRow = document.createElement('tr');
    headers.forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);

    // Display data rows
    data.forEach(row => {
        const tr = document.createElement('tr');
        headers.forEach(header => {
            const td = document.createElement('td');
            td.textContent = row[header] || '';
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });
}
