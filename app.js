// ======================================================================
// DOCX ENGINE – FULLY FIXED VERSION WITH LOOP SUPPORT + IMAGE SUPPORT
// ======================================================================

const express = require("express");
const bodyParser = require("body-parser");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const ImageModule = require("docxtemplater-image-module-free");
const ExpressionModule = require("docxtemplater-expression-module");
const { Buffer } = require("buffer");

const app = express();

// Increase JSON size to allow screenshots
app.use(bodyParser.json({ limit: "200mb" }));

// -------------------------------
// IMAGE HANDLERS
// -------------------------------
function getImage(tagValue) {
    return Buffer.from(tagValue.data, "base64");
}

function getSize(img, tagValue) {
    return [tagValue.width || 550, null];
}

const imageOpts = {
    centered: false,
    getImage,
    getSize
};

// ======================================================================
// MAIN ENDPOINT — DOCX GENERATION
// ======================================================================

app.post("/fill", (req, res) => {
    try {
        console.log("📥 Request received");

        const { templateBase64, data, apiKey } = req.body;

        // Optional API key check
        if (process.env.API_KEY && apiKey !== process.env.API_KEY) {
            return res.status(403).json({ error: "Invalid API key" });
        }

        if (!templateBase64 || !data) {
            return res.status(400).json({ error: "Missing templateBase64 or data" });
        }

        // 1️⃣ Decode DOCX template
        const templateBuf = Buffer.from(templateBase64, "base64");
        const zip = new PizZip(templateBuf);

        // 2️⃣ Initialize Docxtemplater WITH loop support
        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
            delimiters: { start: "%%", end: "%%" },
            modules: [
                new ExpressionModule(),      // ⭐ REQUIRED FOR LOOPS
                new ImageModule(imageOpts)   // ⭐ REQUIRED FOR IMAGES
            ]
        });

        // 3️⃣ Set dynamic data
        doc.render(data);

        // 4️⃣ Generate DOCX output
        const outputBuffer = doc.getZip().generate({
            type: "nodebuffer",
            compression: "DEFLATE"
        });

        const docxBase64 = outputBuffer.toString("base64");

        console.log("✅ Document generated successfully");
        return res.json({ docxBase64 });

    } catch (error) {
        console.error("❌ DOCX Generation Error");

        let details = error.message;

        // MultiError handler
        if (error.properties && Array.isArray(error.properties.errors)) {
            details = error.properties.errors
                .map(e => e.properties.explanation)
                .join("\n");
            console.error("🔍 Template error details:\n" + details);
        }

        return res.status(500).json({
            error: "DOCX Engine Crash",
            details
        });
    }
});

// Health check
app.get("/", (req, res) => {
    res.send("DOCX Engine running with loop support ✔");
});

// Start server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log("🚀 DOCX Engine running on port", PORT));
