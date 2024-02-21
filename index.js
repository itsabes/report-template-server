const path = require("path");
const fs = require("fs");
const stream = require("stream");
const cors = require("cors");
const bodyParser = require("body-parser");

const XlsxPopulate = require("xlsx-populate");
const Docxtemplater = require("docxtemplater");
const ChartModule = require("./docxtemplater-chart-module");
const chartModule = new ChartModule();
const expressions = require("angular-expressions");
const angularParser = function(tag) {
  expr = expressions.compile(tag);
  return { get: expr };
};
const jsonata = require("jsonata");

const express = require("express");
const app = express();
const port = 3000;

app.use(cors());
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

app.get("/", (req, res) =>
  res.send(
    "Report Server, use POST /xlsx/<reportName> or POST /docx/<reportName>"
  )
);

app.post("/xlsx/:filename", (req, res, next) => {
  const fileIn = req.params.filename + ".xlsx";
  const fileOut = req.params.filename + ".xlsx";
  const data = req.body;

  const reportMapping = [
    {
      regex: new RegExp("{{(.*?)}}", "g"),
      replacer: match => {
        const query = match.substring(2, match.length - 2);
        let result = jsonata(query).evaluate(data);
        if (result === undefined || result === null) result = "";
        return result;
      }
    }
  ];

  XlsxPopulate.fromFileAsync(path.join(__dirname, "xlsx-template", fileIn))
    .then(workbook => {
      for (var i = 0; i < reportMapping.length; i += 1) {
        workbook.find(reportMapping[i].regex, reportMapping[i].replacer);
      }
      return workbook.outputAsync();
    })
    .then(data => {
      res.attachment(fileOut);
      res.send(data);
    })
    .catch(next);
});

app.post("/docx/:filename", function(req, res) {
  const fileIn = req.params.filename + ".docx";
  const fileOut = req.params.filename + ".docx";
  const data = req.body;
  const content = fs.readFileSync(
    path.join(__dirname, "docx-template", fileIn),
    "binary"
  );

  const docx = new Docxtemplater()
    .attachModule(chartModule)
    .load(content)
    .setOptions({
      delimiters: { start: "{{", end: "}}" },
      parser: angularParser
    })
    .setData(data)
    .render();

  const fileContents = docx.getZip().generate({ type: "nodebuffer" });

  var readStream = new stream.PassThrough();
  readStream.end(fileContents);

  res.set("Content-disposition", "attachment; filename=" + fileOut);
  res.set("Content-Type", "text/plain");

  readStream.pipe(res);
});

app.listen(port, () => console.log(`Example app listening on port ${port}!`));
