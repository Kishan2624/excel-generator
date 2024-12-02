const excelFileInput = document.getElementById("excelFile");
const dynamicInputs = document.getElementById("dynamicInputs");
const checkboxContainer = document.getElementById("checkboxContainer");

// Read Excel File
excelFileInput.addEventListener("change", (event) => {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });

    // Assume first sheet
    const firstSheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[firstSheetName];

    // Extract headings
    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    const headings = jsonData[0]; // Get first row as headings

    // Populate checkboxes
    checkboxContainer.innerHTML = ""; // Clear existing checkboxes
    headings.forEach((heading) => {
      const checkboxWrapper = document.createElement("div"); // Create a wrapper for each checkbox and label
      checkboxWrapper.style.display = "flex"; // Flex display for each checkbox-label pair
      checkboxWrapper.style.alignItems = "center"; // Vertically center the items

      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.value = heading;
      checkbox.id = `field-${heading}`;

      const label = document.createElement("label");
      label.htmlFor = `field-${heading}`;
      label.textContent = heading;
      label.style.margin = 0;

      // Append checkbox and label to the wrapper
      checkboxWrapper.appendChild(checkbox);
      checkboxWrapper.appendChild(label);

      // Append the wrapper to the checkbox container
      checkboxContainer.appendChild(checkboxWrapper);
    });

    // Show hidden inputs
    dynamicInputs.classList.remove("hidden");
  };

  reader.readAsArrayBuffer(file);
});
document
  .getElementById("excelForm")
  .addEventListener("submit", async (event) => {
    event.preventDefault();

    const fileInput = document.getElementById("excelFile");
    const studentNamesInput = document.getElementById("studentName").value;
    const studentNames = studentNamesInput
      .split(",")
      .map((name) => name.trim().toLowerCase()); // Normalize input names

    const selectedFields = Array.from(
      document.querySelectorAll(
        "#checkboxContainer input[type=checkbox]:checked"
      )
    ).map((checkbox) => checkbox.value);

    if (
      !fileInput.files.length ||
      !studentNames.length ||
      !selectedFields.length
    ) {
      alert("Please upload a file, enter student names, and select fields.");
      return;
    }

    const file = fileInput.files[0];
    const fileData = await file.arrayBuffer();
    const workbook = XLSX.read(fileData, { type: "array" });

    // Assume data is in the first sheet
    const sheetName = workbook.SheetNames[0];
    const sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

    // Track matched and unmatched students
    const unmatchedStudents = [];
    const filteredData = [];

    studentNames.forEach((inputName) => {
      const matchingRow = sheetData.find((row) => {
        const studentName = row["Name of the student"]
          ? row["Name of the student"].trim().toLowerCase()
          : "";
        return studentName === inputName;
      });

      if (matchingRow) {
        const newStudent = {};
        selectedFields.forEach((field) => {
          newStudent[field] = matchingRow[field] || ""; // Handle missing fields gracefully
        });
        filteredData.push(newStudent);
      } else {
        unmatchedStudents.push(inputName); // Add unmatched name to the list
      }
    });

    if (unmatchedStudents.length) {
      alert(
        `The following students were not found: ${unmatchedStudents.join(", ")}`
      );
      return; // Stop further processing if unmatched students exist
    }

    if (!filteredData.length) {
      alert("No matching students found!");
      return;
    }

   
    // Create a new Excel file with filtered data
    const newSheet = XLSX.utils.json_to_sheet(filteredData);
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, newSheet, "Filtered Data");

    XLSX.writeFile(newWorkbook, "Filtered_Student_Data.xlsx");
  });
