// app.js - DEBUG VERSION
const express = require("express");
const bodyParser = require("body-parser");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const ImageModule = require("docxtemplater-image-module-free");
const { Buffer } = require("buffer");

const app = express();
app.use(bodyParser.json({ limit: "200mb" })); 

// 1. Image processing settings
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
  try {
    console.log("👉 Request received. Data keys:", Object.keys(req.body.data || {}));

    const { templateBase64, data, apiKey } = req.body;

    if (process.env.API_KEY && apiKey !== process.env.API_KEY) {
      return res.status(403).json({ error: "Invalid API key" });
    }

    if (!templateBase64 || !data) {
      return res.status(400).json({ error: "Missing templateBase64 or data" });
    }

    const templateBuffer = Buffer.from(templateBase64, "base64");
    const zip = new PizZip(templateBuffer);

    // 2. Initialize Docxtemplater with better error handling settings
    const doc = new Docxtemplater(zip, {
      modules: [new ImageModule(imageOpts)],
      paragraphLoop: true,
      linebreaks: true,
    });

    doc.setData(data);

    try {
      doc.render();
    } catch (error) {
      // 3. THIS IS THE CRITICAL FIX FOR "MULTI ERROR"
      console.error("❌ Render Error Caught!");
      
      let errorDetails = error.message;
      
      // Extract the real reason from the properties
      if (error.properties && error.properties.errors instanceof Array) {
        const errorMessages = error.properties.errors.map(function (err) {
          return err.properties.explanation;
        }).join("\n");
        console.error("🔍 DETAILED ERRORS:", errorMessages);
        errorDetails = errorMessages;
      }
      
      return res.status(500).json({ 
        error: "Template Parsing Failed", 
        details: errorDetails 
      });
    }

    const buffer = doc.getZip().generate({ type: "nodebuffer" });
    const docxBase64 = buffer.toString("base64");

    console.log("✅ Document generated successfully.");
    return res.json({ docxBase64 });

  } catch (e) {
    console.error("💥 Fatal server error", e);
    return res.status(500).json({ error: "Server crash", details: e.message });
  }
});

app.get("/", (req, res) => {
  res.send("DOCX Engine is running (Debug Mode) ✔");
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log("DOCX engine running on port", PORT));
