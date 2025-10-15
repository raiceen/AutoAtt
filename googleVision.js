import fetch from "node-fetch";
import fs from "fs";

const API_KEY = "AIzaSyACT6tmPEcYlXyV6UilO9POZaEwcIviqMQ"
const imagePath = "600489d5-bb97-4f33-af06-a914e32a782e.png";

// Convert image to Base64
const imageBase64 = fs.readFileSync(imagePath, { encoding: "base64" });

async function runOCR() {
  const body = {
    requests: [
      {
        image: { content: imageBase64 },
        features: [{ type: "TEXT_DETECTION" }] // or DOCUMENT_TEXT_DETECTION for better accuracy
      }
    ]
  };

  const url = `https://vision.googleapis.com/v1/images:annotate?key=${API_KEY}`;

  try {
    const res = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body)
    });

    const data = await res.json();

    if (data.responses?.[0]?.fullTextAnnotation?.text) {
      console.log("✅ OCR Result:\n", data.responses[0].fullTextAnnotation.text);
    } else {
      console.log("❌ No text found:", JSON.stringify(data, null, 2));
    }
  } catch (err) {
    console.error("Error calling Vision API:", err);
  }
}

runOCR();
