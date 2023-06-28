const express = require("express");
const randomstring = require("randomstring");
const XLSX = require("xlsx");
const fs = require("fs");

const app = express();

app.listen(3000, () => {
    console.log("Server started on port 3000");
  });

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

app.get("/", (req, res) => {
  res.sendFile(__dirname + "/index.html");
});

app.post("/generate-codes", (req, res) => {
  const prefix = req.body.prefix || "Pi";
  const count = parseInt(req.body.count) || 3000;

  const _data = [];
  let i = 0;

  while (i < count) {
    const _ran = prefix + randomstring.generate(9);
    const _code = { code: _ran };

    if (!_data.includes(_code)) {
      _data.push(_code);
      i++;
    }
  }

  const workSheet = XLSX.utils.json_to_sheet(_data);
  const workBook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workBook, workSheet, "Sheet 1");
  const filePath = `${prefix}_${count}.xlsx`;

  XLSX.writeFile(workBook, filePath, { bookType: "xlsx" });

  res.download(filePath, (err) => {
    if (err) {
      console.error("Error downloading file:", err);
      res.status(500).send("Internal Server Error");
    } else {
      // Clean up the file after it has been downloaded
      fs.unlinkSync(filePath);
    }
  });
});

module.exports = app;
