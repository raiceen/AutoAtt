# AutoAtt — Automated Attendance System (Prototype)

AutoAtt is a web-based prototype system that automates attendance recording using Optical Character Recognition (OCR) and Google integration.  
The goal of this project is to reduce the manual effort and time teachers spend inputting attendance into workbooks by automatically extracting text from uploaded attendance images and recording it into their Workbooks.

---

## This Project Is a Prototype

This system was developed for demonstration and academic purposes only.  
It showcases how automation can simplify classroom tasks using cloud APIs, but it is **not yet ready for full-scale deployment** because:

- It currently uses free-tier APIs (OCR.Space, Google+, and Google Drive) which have limited speed and request quotas.
- It does not yet include a centralized database for user management and record history.
- It requires manual setup of API credentials for security reasons.

Despite these limitations, the system successfully demonstrates how OCR and Google services can work together to automate real-world workflows.

---

## Environment Variables Are Necessary

Sensitive credentials like API keys and OAuth client secrets **must never be stored directly in the code**.  
Instead, they are placed inside a hidden file named `.env`, which is listed in `.gitignore` so that private keys are never uploaded to GitHub or shared publicly.

Each user must create their own `.env` file and fill in their own API credentials before running the system.

---

## Technologies Used

- **Frontend:** HTML, CSS, JavaScript  
- **Backend:** Node.js with Express.js  
- **APIs:** Google Drive, Google+, and OCR.Space API(free tier)
- **Authentication:** Google OAuth 2.0  

---

## Prerequisites

- Node.js (https://nodejs.org)
- Google Account
- An OCR API account (I used free OCR.Space free for prototype) (https://ocr.space/ocrapi/freekey)
- ngrok for using Google account on other devices(google dont allow localhost callbacks on other devices) (https://ngrok.com)

## System Setup (Step-by-Step)

### 1. Install Requirements

Clone this repository:

```bash/cmd
git clone https://github.com/raiceen/AutoAtt.git
cd AutoAtt
npm install
```
or
``` Download ZIP from GitHub
https://github.com/raiceen/AutoAtt.git
then unzip file
```

Create Google OAuth credentials
``` Go to https://console.cloud.google.com/apis/credentials
then Credentials
Create Credentials > OAuth client ID > Application type pick Web Application
add http://localhost:3000 to Authorized JavaScript origins
add http://localhost:3000/auth/google/callback
click Create
copy the Client ID paste the value to GOOGLE_CLIENT_ID in the .env
copy Client Secret passte the value to GOOGLE_CLIENT_SECRET in the .env
```

Get an OCR API key
```Go to https://ocr.space/ocrapi
buy a plan or register for free
check email then copy the key and paste the value to OCR_SPACE_API_KEY in the .env
```
# Your .env is now complete!

Go to project root
```
npm install
node server.js
then go to http://localhost:3000
```

# Allow other device to use the Google Feature (ngrok)

Create ngrok account
``` Go to https://ngrok.com
Domains > copy domain link (.....ngrok-free.dev)
Change the value of BASE_URL in the .env with https://<your-ngrok-domain>
```
``` Change Google Credentials
go to https://console.cloud.google.com/ > Credentials
change the Authorized JavaScript origins with https://<your-ngrok-domain>
change the Authorized redirect URIs with https://<your-ngrok-domain/auth/google/callback>
open 2 terminal
run
ngrok http 3000
also run
node server.js
```
