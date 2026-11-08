# 🌐 URL-First Excel ↔ Website Automation Assistant (Python + Streamlit + Playwright + Gemini)

A complete, human-in-the-loop **URL-First Two-Way Automation Assistant** in Python. Give the app any website URL (such as SSO Rajasthan, Raj eMitra, scholarship/pension portals, or internal web applications): it opens the site in a visible Chromium browser, **scans everything on the page**, and works in two directions:

1. **Step 1 & Step 2 (`Scan & Make Excel Format`)**:
   - Scans all inputs, dropdowns (including cascading `District → Block → Village`), radios, checkboxes, date pickers, file uploads, multi-step wizard forms, HTML tables, and pagination links.
   - Displays a full scan report (`Form fields found: N, Tables found: M, Uploads: K`) cleaned by Google Gemini (`llm.py`).
   - Auto-generates a ready-to-fill 3-sheet Excel workbook (`.xlsx`) via `template_maker.py`:
     - **Sheet 1 (`Form_Data_Template`)**: 1 column per form field in form order, required columns highlighted, header row styled and frozen, cell comments/hint row, **Excel DataValidation dropdown lists** populated with the website's actual options, plus `Photo_Document_Filename`, `Status`, `Application_Acknowledgement_No`, `Submission_Time`, and `Error_Message`.
     - **Sheet 2 (`Instructions`)**: Simple Hindi and English instructions.
     - **Sheet 3 (`Allowed Values`)**: Full list of valid dropdown/radio options scraped from the website.
   - Includes **"I already have my own Excel"**: upload your existing Excel/Word file, map columns via Gemini, and reshape it into the website's template format.
2. **Step 3 (`Fill Website from Excel`)**:
   - Pre-validates mandatory fields, 10-digit mobile, 12-digit Aadhaar, 6-digit PIN, date formats, allowed dropdown values, and duplicates with friendly messages like `"Row 12: mobile number 9 digit ka hai, 10 hona chahiye"`.
   - Fills the website row-by-row with fuzzy dropdown matching, cascading dropdown waits, photo auto-resize/compression (`<= 50 KB`), **Dry Run (Test Mode)**, **Semi-Automatic Mode** (stops before Submit for your review), crash recovery checkpoints, proof screenshots, and non-destructive Excel/PDF output.
3. **Step 4 (`Website to Excel`)**:
   - Multi-page HTML table & list extraction (`max_pages`), **Search-and-Fetch** from an Excel of Application IDs/Aadhaar numbers, Detail-Page scraping, and linked PDF/certificate downloads.
4. **Step 5 (`Fill + Verify`)**:
   - Fetches live portal status for submitted applications (`Submitted` → `Approved` / `Pending` / `Rejected`) and reconciles your Excel against website data into a color-coded **`Differences`** sheet.

---

## 📁 Project File Structure (Root Folder)

```text
.
├── app.py                  # Main Streamlit UI (4-Mode Switch: Scan & Make Excel Format / Fill Website / Website to Excel / Fill + Verify)
├── scanner.py              # Step 1: URL-First DOM & iframe scanner (fields, cascading dropdown prober, safe multi-step scan, tables)
├── template_maker.py       # Step 2: Auto-generates 3-sheet .xlsx template with Excel DataValidation dropdowns & reshapes user Excel
├── importer.py             # Loads Excel (.xlsx/.csv) multi-sheet & Word (.docx tables + "Label: Value" paragraphs)
├── validator.py            # Pre-flight validation (Mandatory, Mobile, Aadhaar, PIN, Dates, Allowed Dropdown Options, Duplicates)
├── browser.py              # Playwright Headed Chromium manager, persistent profile, Login/OTP/CAPTCHA pause, Ack No parser
├── filler.py               # Playwright form-filling engine (fuzzy dropdown matcher, cascading waits, multi-step forms, Unicode Hindi)
├── uploader.py             # Photo/document matcher & Pillow iterative KB/dimension compressor
├── runner.py               # Batch execution state, crash recovery checkpoints, proof screenshots, Semi-Auto vs Auto-Submit
├── extractor.py            # Step 4: HTML table detection, pagination, Search-and-Fetch by ID, Detail-page & PDF downloader
├── reconcile.py            # Step 5: Excel vs Website reconciliation & color-coded "Differences" workbook
├── report.py               # Non-destructive Excel result writer & fpdf2 PDF audit report generator
├── llm.py                  # Centralized Google Gemini API (google-genai) JSON scan cleaner & column mapper with 3x retry
├── utils.py                # PII masking (XXXXXXXX1234), value normalizers, Hindi/English/Hinglish voice command parser
├── requirements.txt        # Python dependencies
├── README.md               # Beginner setup guide, Privacy & Terms notice, and Hindi step-by-step walkthrough
├── templates/              # Saved portal JSON templates (reusable without re-scanning)
│   ├── SSO_Citizen_Registration.json
│   └── SSO_Status_Search_Extract.json
└── samples/                # Local demo website & sample test files
    ├── demo_portal.html    # Local SSO/eMitra test portal (Multi-Step Form + Cascading Dropdowns + Status Search + Paginated Table)
    ├── sample_applicants.xlsx
    ├── sample_applicants.docx
    └── photos/
        ├── 101_Ramesh_Kumar.jpg
        ├── 102_Sunita_Devi.jpg
        └── 103_Mahesh_Sharma.jpg
```

---

## 🚀 Beginner-Friendly Setup Instructions

### 1. Install Python 3.10+
Verify Python 3.10 or newer is installed:
```bash
python --version
```

### 2. Install Python Dependencies
From the root folder containing `requirements.txt`, run:
```bash
pip install -r requirements.txt
```

### 3. Install Playwright Chromium Browser
Download the visible Chromium browser binary:
```bash
playwright install chromium
```

### 4. Set Your Google Gemini API Key (`GEMINI_API_KEY`)
All AI calls live in `llm.py` using the official `google-genai` package. It checks `st.secrets["GEMINI_API_KEY"]` first, then the environment variable `GEMINI_API_KEY`.

**Windows PowerShell:**
```powershell
$env:GEMINI_API_KEY="your_gemini_api_key_here"
$env:GEMINI_MODEL="gemini-2.5-flash"
```

**Linux / macOS:**
```bash
export GEMINI_API_KEY="your_gemini_api_key_here"
export GEMINI_MODEL="gemini-2.5-flash"
```

### 5. Run the Streamlit App
```bash
streamlit run app.py
```

---

## 🇮🇳 हिंदी में स्टेप-बाय-स्टेप उपयोग गाइड (5 आसान चरण)

### (1) URL Scan (वेबसाइट स्कैन करना)
1. ऐप के ऊपर मोड स्विच में **"Scan & Make Excel Format"** चुनें।
2. वेबसाइट का URL डालें (या टेस्टिंग के लिए `samples/demo_portal.html` चुनें) और **"Open URL & Scan This Page"** दबाएं।
3. यदि वेबसाइट पर लॉगिन, OTP या CAPTCHA आता है, तो ऐप रुक जाएगा (*"Please login / enter OTP / captcha in the browser, then click Continue"*)। ब्राउज़र में स्वयं लॉगिन करें और ऐप में **Continue** दबाएं।
4. ऐप पेज के सभी फॉर्म फील्ड्स, ड्रॉपडाउन (`District → Block → Village`), फोटो अपलोड और टेबल्स को स्कैन करके रिपोर्ट दिखाएगा: `Form fields found: N, Tables found: M, Uploads: K`।

### (2) Excel Format Download (रेडी-टू-फिल एक्सेल टेम्पलेट डाउनलोड करना)
1. स्कैन पूरा होते ही **"Download Auto-Generated Website Excel Format (.xlsx)"** बटन दबाएं।
2. डाउनलोड हुई एक्सेल फाइल में 3 शीट होंगी:
   - **`Form_Data_Template`**: वेबसाइट के हर फील्ड के लिए एक कॉलम, अनिवार्य (Required) कॉलम हाइलाइट किए हुए, और ड्रॉपडाउन वाले कॉलम में वेबसाइट के असली विकल्पों की ड्रॉपडाउन लिस्ट (Data Validation)।
   - **`Instructions`**: हिंदी और इंग्लिश में फॉर्म भरने के नियम।
   - **`Allowed Values`**: सभी ड्रॉपडाउन के मान्य विकल्पों की सूची।
3. यदि आपके पास पहले से अपनी कोई एक्सेल फाइल है, तो **"I already have my own Excel"** में उसे अपलोड करें—Gemini आपके कॉलम को वेबसाइट के फॉर्मेट में अपने-आप बदल (Reshape) देगा।

### (3) Data Bharna (एक्सेल में डेटा भरना और जांचना)
1. डाउनलोड किए गए एक्सेल फॉर्मेट में अपने आवेदकों का डेटा भरें (मोबाइल 10 अंक, आधार 12 अंक, पिन कोड 6 अंक, जन्म तिथि `DD/MM/YYYY`) और फोटो फाइल का नाम लिखें।
2. अब ऐप में ऊपर **"Fill Website from Excel"** मोड चुनें और अपनी भरी हुई एक्सेल फाइल अपलोड करें।
3. ऐप फॉर्म भरने से पहले डेटा की जांच (Validation) करेगा। यदि कहीं गलती है (जैसे *"Row 12: mobile number 9 digit ka hai, 10 hona chahiye"*), तो वह लाल रंग में दिखेगी जिसे आप सीधे स्क्रीन पर ही सुधार सकते हैं।

### (4) Website Mein Fill (वेबसाइट पर फॉर्म भरना और सबमिट करना)
1. **"Fill Website from Excel"** के दूसरे टैब (*Row-by-Row & Batch Filler*) में जाएं।
2. यदि आप केवल पहली पंक्ति टेस्ट करना चाहते हैं, तो साइडबार में **Dry Run / Test Mode** चालू करें।
3. **"Fill Current Row (Stop Before Submit)"** दबाएं। ऐप अपने-आप नाम, पता, जिला → ब्लॉक → गाँव चुनकर और फोटो को 50KB से कम साइज में कंप्रेस करके अपलोड कर देगा, और **Submit** बटन से ठीक पहले रुक जाएगा।
4. ब्राउज़र में फॉर्म चेक करें और ऐप में **"Approve & Submit Row"** दबाएं। ऐप फॉर्म सबमिट करके **Application / Acknowledgement No** और **Screenshot Proof** सेव कर लेगा और आपकी एक्सेल की नई कॉपी में अपडेट कर देगा (आपकी मूल फाइल कभी ओवरराइट नहीं होती)।

### (5) Website Se Data Excel Mein (वेबसाइट से डेटा और स्टेटस एक्सेल में निकालना)
1. ऊपर **"Website to Excel"** मोड चुनें।
2. **टेबल निकालने के लिए**: *Multi-Page Table & List Extractor* टैब में अधिकतम पेज (`max_pages`) सेट करें और **"Extract Paginated Table"** दबाएं—सभी पेजों का डेटा साफ होकर एक्सेल में डाउनलोड हो जाएगा।
3. **एप्लीकेशन स्टेटस चेक करने के लिए (Search-and-Fetch)**: *Search-and-Fetch Status from Excel* टैब में अपनी एप्लीकेशन नंबर वाली एक्सेल अपलोड करें। ऐप एक-एक नंबर को वेबसाइट के स्टेटस पेज पर सर्च करेगा और उसका `Status (Approved/Pending/Rejected)`, `Approval Date` और `Sanctioned Amount` आपकी एक्सेल में लिख देगा।
4. **Fill + Verify (मिलान/Reconcile)**: **"Fill + Verify"** मोड में अपनी एक्सेल और वेबसाइट के ताज़ा डेटा की तुलना करें और रंगीन **`Differences`** एक्सेल शीट डाउनलोड करें।

---

## 🔒 Privacy & Terms of Use
- **Local Privacy**: All spreadsheets and documents stay on your computer. Gemini (`llm.py`) receives only column names, field labels, and 1–2 masked sample values (`XXXXXXXX1234`).
- **Authorised Use Only**: Use this assistant only on websites and data you are authorised to access, and always comply with the target portal's Terms of Use.
