// ===============================
// DOCX ENGINE - FINAL v4+ VERSION
// ===============================

const express = require("express");
const bodyParser = require("body-parser");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const ImageModule = require("docxtemplater-image-module-free");
const { Buffer } = require("buffer");

const app = express();
app.use(bodyParser.json({ limit: "200mb" }));

// ===============================
// IMAGE MODULE HELPERS
// ===============================
const imageOpts = {
  centered: false,

  getImage(tagValue) {
    return Buffer.from(tagValue.data, "base64");
  },

  getSize(img, tagValue) {
    const width = tagValue.width || 550;
    return [width, null];
  }
};

// ===============================
// MAIN ROUTE: /fill
// ===============================
app.post("/fill", (req, res) => {
  try {
    console.log("👉 Request received");

    const { templateBase64, data, apiKey } = req.body;

    if (process.env.API_KEY && apiKey !== process.env.API_KEY) {
      return res.status(403).json({ error: "Invalid API key" });
    }
    if (!templateBase64 || !data) {
      return res.status(400).json({ error: "Missing templateBase64 or data" });
    }

    // 1. Load ZIP
    const templateBuffer = Buffer.from(templateBase64, "base64");
    const zip = new PizZip(templateBuffer);

    // 2. Initialize Docxtemplater (v4 syntax)
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      delimiters: { start: "%%", end: "%%" },
      modules: [
        new ImageModule(imageOpts)  // ✔ ATTACH MODULES HERE ONLY
      ]
    });

    // 3. Render with new API
    doc.render(data);

    // 4. Output file
    const buffer = doc.getZip().generate({ type: "nodebuffer" });
    const docxBase64 = buffer.toString("base64");

    console.log("✅ Document generated successfully.");
    return res.json({ docxBase64 });

  } catch (error) {
    console.error("❌ Template Processing Failed!");

    let detail = error.message;
    if (error.properties?.errors) {
      detail = error.properties.errors
        .map(e => e.properties.explanation)
        .join("\n");
    }

    return res.status(500).json({
      error: "DOCX Engine Crash",
      details: detail
    });
  }
});

// ===============================
app.get("/", (req, res) => {
  res.send("DOCX Engine is running ✔ (v4-safe)");
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log("🚀 DOCX engine running on port", PORT));


