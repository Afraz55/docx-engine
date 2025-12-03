// app.js - FINAL DEBUG VERSION (Catches Initialization Errors)
const express = require("express");
const bodyParser = require("body-parser");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const ImageModule = require("docxtemplater-image-module-free");
const { Buffer } = require("buffer");

const app = express();
// Increase limit for screenshots
app.use(bodyParser.json({ limit: "200mb" })); 

// Image processing helper functions
function getImage(tagValue) {
  return Buffer.from(tagValue.data, "base64");
}

function getSize(img, tagValue) {
  const width = tagValue.width || 550; 
  return [width, null]; 
}

const imageOpts = {
  centered: false,
  getImage,
  getSize
};

app.post("/fill", (req, res) => {
  // 1. GLOBAL TRY BLOCK - Catches errors anywhere in the process
  try {
    console.log("👉 Request received.");

    const { templateBase64, data, apiKey } = req.body;

    if (process.env.API_KEY && apiKey !== process.env.API_KEY) {
      return res.status(403).json({ error: "Invalid API key" });
    }

    if (!templateBase64 || !data) {
      return res.status(400).json({ error: "Missing templateBase64 or data" });
    }

    // 2. Load Zip
    const templateBuffer = Buffer.from(templateBase64, "base64");
    const zip = new PizZip(templateBuffer);

    // 3. Initialize Docxtemplater (Parsing happens here too!)
    const doc = new Docxtemplater(zip, {
      modules: [new ImageModule(imageOpts)],
      paragraphLoop: true,
      linebreaks: true,
    });

    // 4. Set Data
    doc.setData(data);

    // 5. Render
    doc.render();

    // 6. Generate Output
    const buffer = doc.getZip().generate({ type: "nodebuffer" });
    const docxBase64 = buffer.toString("base64");

    console.log("✅ Document generated successfully.");
    return res.json({ docxBase64 });

  } catch (error) {
    // 7. INTELLIGENT ERROR HANDLING
    console.error("❌ Process Failed!");

    let errorDetails = error.message;

    // Check if this is a Docxtemplater Multi-Error
    if (error.properties && error.properties.errors instanceof Array) {
      const errorMessages = error.properties.errors.map(function (err) {
        return err.properties.explanation;
      }).join("\n");
      
      console.error("🔍 DETAILED ERRORS FOUND:", errorMessages);
      errorDetails = "TEMPLATE ERROR: " + errorMessages;
    }

    // Return the detailed error to Apps Script so you can read it in the logs
    return res.status(500).json({ 
      error: "Detailed Crash Report", 
      details: errorDetails 
    });
  }
});

app.get("/", (req, res) => {
  res.send("DOCX Engine is running (v3 - Catch All Mode) ✔");
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log("DOCX engine running on port", PORT));
