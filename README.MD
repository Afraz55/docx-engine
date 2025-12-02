# DOCX Engine for VAPT Report Automation

## Endpoints

### POST /fill
Fills a docx template with data + images and returns final docx file in base64.

Body format:
{
"templateBase64": "....",
"data": { ... },
"apiKey": "YOUR_SECRET"
}


## Environment Variables

- API_KEY (optional)

## Run locally:

npm install
node app.js