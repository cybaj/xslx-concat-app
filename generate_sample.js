const XLSX = require('xlsx');

function generateSampleExcel(filename, data) {
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  XLSX.writeFile(wb, filename);
}

// Sample Data for File 1
const data1 = [
  { Name: "John", Age: 25, Occupation: "Engineer" },
  { Name: "Doe", Age: 28, Occupation: "Doctor" },
  { Name: "Jane", Age: 29, Occupation: "Lawyer" },
];

// Sample Data for File 2 (with some duplicate rows from File 1)
const data2 = [
  { Name: "John", Age: 25, Occupation: "Engineer" },
  { Name: "Anna", Age: 27, Occupation: "Artist" },
  { Name: "Jane", Age: 29, Occupation: "Lawyer" },
];

generateSampleExcel('sample1.xlsx', data1);
generateSampleExcel('sample2.xlsx', data2);

console.log("Sample Excel files generated!");
