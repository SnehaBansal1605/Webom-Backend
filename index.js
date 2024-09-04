const express = require("express");
const bodyParser = require("body-parser");
const xlsx = require("xlsx");
const path = require("path");
const cors = require("cors");
const fs = require("fs");

const app = express();
const port = 3001;

// Enable CORS for all routes
app.use(cors());

// Middleware to parse JSON data
app.use(bodyParser.json());

// Route to handle form submissions
app.post("/submit-form", (req, res) => {
  const formData = req.body;
  const filePath = path.join(__dirname, "form_data.xlsx");

  try {
    let workbook;
    let worksheet;

    // Check if the file exists
    if (fs.existsSync(filePath)) {
      console.log("Excel file exists, appending data...");

      // Read the existing workbook
      workbook = xlsx.readFile(filePath);
      worksheet = workbook.Sheets["FormData"];

      // Check if worksheet is found
      if (!worksheet) {
        console.error("Worksheet 'FormData' not found!");
        return res.status(500).json({ message: "Worksheet not found!" });
      }

      // Convert the worksheet to JSON and append the new form data
      const jsonData = xlsx.utils.sheet_to_json(worksheet);
      jsonData.push(formData);

      // Convert updated JSON back to worksheet
      worksheet = xlsx.utils.json_to_sheet(jsonData);
      workbook.Sheets["FormData"] = worksheet; // Re-assign updated worksheet
    } else {
      console.log("Creating new Excel file and sheet...");

      // Create a new workbook and sheet if the file doesn't exist
      workbook = xlsx.utils.book_new();
      worksheet = xlsx.utils.json_to_sheet([formData]);

      // Append the worksheet to the workbook with the name 'FormData'
      xlsx.utils.book_append_sheet(workbook, worksheet, "FormData");
    }

    // Write the updated workbook to the file
    xlsx.writeFile(workbook, filePath);

    // Send success response
    res.json({ message: "Form data has been saved to Excel file!" });
  } catch (error) {
    // Log the error to the console for debugging
    console.error("Error processing form data:", error);

    // Send an error response to the client
    res
      .status(500)
      .json({ message: "Internal Server Error", error: error.message });
  }
});

app.listen(port, () => {
  console.log(`Server running on http://localhost:${port}`);
});
