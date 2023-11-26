const express = require("express");
const bodyParser = require("body-parser");
const path = require("path");

const { createId } = require("@paralleldrive/cuid2");

const toPdf = require("office-to-pdf");

const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

const app = express();
const port = 3200;
const fs = require("fs");

const content = fs.readFileSync(
  path.resolve(__dirname, "templates", "template1.docx"),
  "binary"
);

app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());

app.get("/preview", async (req, res) => {
  const zip = new PizZip(content);

  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
  });

  doc.render(req.body);

  const buf = doc.getZip().generate({
    type: "nodebuffer",
    compression: "DEFLATE",
  });

  const pdf = await toPdf(buf).then(
    (pdfBuffer) => {
      return pdfBuffer;
    },
    (err) => {
      console.log(err);
    }
  );

  res.contentType("application/pdf");

  res.send(pdf);
});

app.post("/", async (req, res) => {
  const zip = new PizZip(content);

  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
  });

  doc.render(req.body);

  const buf = doc.getZip().generate({
    type: "nodebuffer",
    compression: "DEFLATE",
  });

  const pdf = await toPdf(buf).then(
    (pdfBuffer) => {
      return pdfBuffer;
    },
    (err) => {
      console.log(err);
    }
  );

  const id = `${createId()}.pdf`;

  const outputName = path.resolve(__dirname, "outputs", id);

  fs.writeFileSync(outputName, pdf);

  res.json({
    url: "/pdf/" + id,
  });
});

app.listen(port, () => {
  console.log(`Server berjalan di http://localhost:${port}`);
});
