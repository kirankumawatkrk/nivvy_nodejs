// main.js
const express = require("express");
const app = express();
const port = 3000;
const fs = require("fs");
const cors = require("cors");
const XlsxPopulate = require("xlsx-populate");

app.use(
  cors({
    origin: "*",
  })
);

// Define a simple API endpoint for GET request
app.get("/api/data", async (req, res) => {
  const jsonData = [
    { name: "John", age: 25 },
    { name: "Jane", age: 30 },
  ];

  XlsxPopulate.fromBlankAsync("", {password: 'kiran'}).then((workbook) => {
    const sheet = workbook.sheet(0);
    sheet.cell("A1").value(["Name", "Age"]);

    jsonData.forEach((data, index) => {
      sheet.cell(`A${index + 2}`).value([data.name, data.age]);
    });

    const filePath = "./data_with_password.xlsx";

    workbook
      .toFileAsync(filePath, { password: 'password' })
      .then(() => {
        res.sendFile(filePath, () => {
          console.log("Excel file exported as response");
          // Remove the temporary file after sending
          fs.unlinkSync(filePath);
        });
      })
      .catch((error) => {
        console.error("Error creating Excel file:", error);
        res.status(500).send("Internal Server Error");
      });
  });

  
});

// Start the server
app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});
