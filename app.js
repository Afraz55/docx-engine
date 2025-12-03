// ===============================
// DOCX ENGINE - FINAL VERSION
// ===============================

const express = require("express");
const bodyParser = require("body-parser");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const ImageModule = require("docxtemplater-image-module-free");
const { Buffer } = require("buffer");

const app = express();

// Increase limit for screenshots + large payloads
app.use(bodyParser.json({ limit: "200mb" }));

// ===============================
// IMAGE MODULE HELPERS
// ===============================
const imageOpts = {
  centered: false,

  // Expecting: { data: "<base64>", width: <number> }
  getImage(tagValue) {
    try {
      return Buffer.from(tagValue.data, "base64");
    } catch (err) {
      console.error("❌ Image conversion failed:", err);
      return null;
    }
  },

  getSize(img, tagValue) {
    const width = tagValue.width || 550;
    return [width, null]; // auto height
  }
};

// ===============================
// MAIN ROUTE: /fill
// ===============================
app.post("/fill", (req, res) => {
  try {
    console.log("👉 Request received.");

    const { templateBase64, data, apiKey } = req.body;

    // Optional API key protection
    if (process.env.API_KEY && apiKey !== process.env.API_KEY) {
      return res.status(403).json({ error: "Invalid API key" });
    }

    if (!templateBase64 || !data) {
      return res
        .status(400)
        .json({ error: "Missing templateBase64 or data fields" });
    }

    // -------------------------------
    // 1. Convert base64 into zip
    // -------------------------------
    const templateBuffer = Buffer.from(templateBase64, "base64");
    const zip = new PizZip(templateBuffer);

    // -------------------------------
    // 2. Initialize Docxtemplater
    // -------------------------------
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      delimiters: { start: "%%", end: "%%" }
      doc.attachModule(new ImageModule(imageOpts)); // 💡 Safe delimiters
    });
    // -------------------------------
    // 3. Render with NEW API
    // -------------------------------
    doc.render(data);

    // -------------------------------
    // 4. Generate Output DOCX
    // -------------------------------
    const buffer = doc.getZip().generate({ type: "nodebuffer" });
    const docxBase64 = buffer.toString("base64");

    console.log("✅ Document generated successfully.");

    return res.json({ docxBase64 });

  } catch (error) {
    console.error("❌ Template Processing Failed!");

    let detail = error.message;

    if (error.properties && Array.isArray(error.properties.errors)) {
      detail = error.properties.errors
        .map(e => e.properties.explanation)
        .join("\n");
      console.error("🔍 DocxTemplater Errors Found:\n", detail);
    }

    return res.status(500).json({
      error: "DOCX Engine Crash",
      details: detail
    });
  }
});

// ===============================
// HEALTH CHECK
// ===============================
app.get("/", (req, res) => {
  res.send("DOCX Engine is running ✔ (Final Version)");
});

// ===============================
// START SERVER
// ===============================
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log("🚀 DOCX engine running on port", PORT));

