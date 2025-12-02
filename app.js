// app.js
const express = require("express");
const bodyParser = require("body-parser");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const ImageModule = require("docxtemplater-image-module-free");
const { Buffer } = require("buffer");

const app = express();
app.use(bodyParser.json({ limit: "200mb" })); // Increase limit for screenshots

// Image processing
function getImage(tagValue) {
  return Buffer.from(tagValue.data, "base64");
}

function getSize(img, tagValue) {
  const width = tagValue.width || 550;     // Default width
  return [width, null];                    // Auto height
}

const imageOpts = {
  centered: false,
  getImage,
  getSize
};

app.post("/fill", (req, res) => {
  try {
    const { templateBase64, data, apiKey } = req.body;

    // Optional API key protection
    if (process.env.API_KEY && apiKey !== process.env.API_KEY) {
      return res.status(403).json({ error: "Invalid API key" });
    }

    if (!templateBase64 || !data) {
      return res.status(400).json({ error: "Missing templateBase64 or data" });
    }

    const templateBuffer = Buffer.from(templateBase64, "base64");
    const zip = new PizZip(templateBuffer);

    const doc = new Docxtemplater(zip, {
      modules: [new ImageModule(imageOpts)],
      paragraphLoop: true,
      linebreaks: true
    });

    doc.setData(data);

    try {
      doc.render();
    } catch (e) {
      console.error("Render error", e);
      return res.status(500).json({ error: "Render failed", details: e.message });
    }

    const buffer = doc.getZip().generate({ type: "nodebuffer" });
    const docxBase64 = buffer.toString("base64");

    return res.json({ docxBase64 });

  } catch (e) {
    console.error("Fatal server error", e);
    return res.status(500).json({ error: "Server crash", details: e.message });
  }
});

app.get("/", (req, res) => {
  res.send("DOCX Engine is running ✔");
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log("DOCX engine running on port", PORT));
