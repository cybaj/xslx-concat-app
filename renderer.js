const { ipcRenderer } = require("electron");
const xlsx = require("xlsx");

const inputExcel = document.getElementById("input-excel");
const processFilesButton = document.getElementById("process-files");
const saveToExcelButton = document.getElementById("save-to-excel");
const saveDuplicatedToExcelButton = document.getElementById(
  "save-duplicated-to-excel"
);
const saveUniqueToExcelButton = document.getElementById("save-unique-to-excel");
const columnNamesInput = document.getElementById("column-names");
const refreshButton = document.getElementById("refresh");

let collectedData = [];
let filteredData = [];
let duplicated = new Set();
let duplicatedKeys = new Set();
let colmunNames = [];

refreshButton.addEventListener("click", () => {
  // Reset global variables
  collectedData = [];
  filteredData = [];
  duplicated = new Set();
  duplicatedKeys = new Set();
  colmunNames = [];

  // Clear the input field and table
  columnNamesInput.value = "";
  inputExcel.value = ""; // Clear file input selection
  const table = document.getElementById("preview-table");
  table.querySelector("thead").innerHTML = "";
  table.querySelector("tbody").innerHTML = "";

  alert("Refreshed and ready for new input!");
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

        // Concat the column names of this file
        colmunNames = colmunNames.concat(Object.keys(jsonData[0]));

        // Filter logic moved here:
        const filteringResult = filterUniqueRows(
          collectedData,
          columns,
          new Set(colmunNames)
        );
        filteredData = filteringResult["filteredData"];
        duplicated = filteringResult["duplicatedData"];
        duplicatedKeys = filteringResult["duplicatedKeys"];

        // Show a preview in the table (for example, first 5 rows)
        displayPreview(filteredData.slice(0, 5));
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

  // const filteredData = filterUniqueRows(collectedData, columns);

  // New alert to display the numbers:
  alert(
    `total rows: ${collectedData.length}\n filtered out: ${filteredData.length}`
  );

  // Ask user for file location and name
  ipcRenderer
    .invoke("show-save-dialog")
    .then((filePath) => {
      if (filePath) {
        const newWorkbook = xlsx.utils.book_new();
        const newWorksheet = xlsx.utils.json_to_sheet(filteredData);
        xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, "Sheet1");
        xlsx.writeFile(newWorkbook, filePath);
      } else {
        console.log("No file path selected or provided.");
      }
    })
    .catch((err) => {
      console.error("Error saving the file:", err);
    });
});

saveUniqueToExcelButton.addEventListener("click", () => {
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

  ipcRenderer
    .invoke("show-save-unique-dialog")
    .then((filePath) => {
      if (filePath) {
        const uniqueData = filteredData.filter((row) => {
          const key = columns.map((col) => row[col]).join("|");
          if (duplicatedKeys.has(key)) {
            return false;
          }
          return true;
        });

        console.log("unique");
        console.dir(uniqueData);

        const newWorkbook = xlsx.utils.book_new();
        const newWorksheet = xlsx.utils.json_to_sheet(uniqueData);
        xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, "Sheet1");
        xlsx.writeFile(newWorkbook, filePath);
      } else {
        console.log("No file path selected or provided.");
      }
    })
    .catch((err) => {
      console.error("Error saving the file:", err);
    });
});

saveDuplicatedToExcelButton.addEventListener("click", () => {
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

  ipcRenderer
    .invoke("show-save-duplicated-dialog")
    .then((filePath) => {
      if (filePath) {
        const newWorkbook = xlsx.utils.book_new();
        const newWorksheet = xlsx.utils.json_to_sheet(Array.from(duplicated));
        xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, "Sheet1");
        xlsx.writeFile(newWorkbook, filePath);
      } else {
        console.log("No file path selected or provided.");
      }
    })
    .catch((err) => {
      console.error("Error saving the file:", err);
    });
});

function filterUniqueRows(data, columns, columnNames) {
  const seen = new Set();
  const duplicated = new Set();
  const duplicatedKeys = new Set();
  const filteredData = data.filter((row) => {
    const key = columns.map((col) => row[col]).join("|");
    if (seen.has(key)) {
      columnNames.forEach((col) => {
        if (row[col] === undefined) {
          row[col] = "";
        }
      });
      duplicated.add(row);
      duplicatedKeys.add(key);
      return false;
    }
    seen.add(key);
    return true;
  });
  const mergedData = filteredData.map((row) => {
    const key = columns.map((col) => row[col]).join("|");
    if (duplicatedKeys.has(key)) {
      columnNames.forEach((col) => {
        if (row[col] === undefined) {
          row[col] = "";
        }
      });
    }
    return row;
  });
  return {
    filteredData: mergedData,
    duplicatedData: duplicated,
    seen: seen,
    duplicatedKeys: duplicatedKeys,
  };
}

function displayPreview(data) {
  const table = document.getElementById("preview-table");
  const thead = table.querySelector("thead");
  const tbody = table.querySelector("tbody");

  // Clear existing rows
  thead.innerHTML = "";
  tbody.innerHTML = "";

  // Display headers
  const headers = Object.keys(data[0]);
  const headerRow = document.createElement("tr");
  headers.forEach((header) => {
    const th = document.createElement("th");
    th.textContent = header;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

  // Display data rows
  data.forEach((row) => {
    const tr = document.createElement("tr");
    headers.forEach((header) => {
      const td = document.createElement("td");
      td.textContent = row[header] || "";
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
}
