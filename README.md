# Faculty Auto Marks Capture App

**Zero manual entry for faculty.**  
Upload cover-page photos of End-Term answer sheets (up to 70 students).  
The app automatically extracts:

- Student Name  
- Roll Number  
- Marks for every question / part  

Then calculates all question totals + Grand Total and lets you download a clean Excel file with live formulas.

---

## How faculty uses it

1. Take a clear photo of the **first page** (cover page with Name, Roll No and Marks Distribution table) of every answer booklet.
2. Open the Streamlit app.
3. Upload all photos at once (max 70).
4. Click **Process All Images**.
5. Review the extracted table (correct only the few OCR mistakes if any).
6. Download the Excel – all sums are already calculated with formulas.

No typing of marks is required.

---

## Excel layout (important)

The downloaded Excel has **two sheets**:

1. **Summary** – one row per student (Name, Roll, Grand Total)
2. **Detailed Marks (Q-wise)** – for every student you get:
   - Header with Name + Roll
   - Then **10 rows** (Q.No 1 to Q.No 10)
   - Columns: a | b | c | d | e | f | g | h | i | j | Total
   - Grand Total row at the bottom

This matches the original answer-sheet marks table.



## Deploy on Streamlit Community Cloud (free)

1. Create a new GitHub repository.
2. Upload **all files** from this folder:
   - `app.py`
   - `requirements.txt`
   - `README.md`
3. Go to [https://share.streamlit.io](https://share.streamlit.io) and sign in with GitHub.
4. Click **New app** → select your repository → main file `app.py` → Deploy.
5. Wait 1–2 minutes. Your app will be live at a public URL you can share with faculty.

### Important for Streamlit Cloud
Tesseract (OCR engine) is **not** pre-installed on Streamlit Cloud.  
You have two options:

**Option A (recommended for production)**  
Use a custom Docker image or a community space that includes Tesseract, **or**  
add a `packages.txt` file (see below) – Streamlit Cloud supports apt packages.

Create a file named `packages.txt` in the repo root with:

```
tesseract-ocr
tesseract-ocr-eng
poppler-utils
```

**Option B**  
Run the app locally (see below) where Tesseract is already installed.

---

## Run locally

```bash
# Install system dependency (Ubuntu/Debian)
sudo apt update && sudo apt install -y tesseract-ocr tesseract-ocr-eng
poppler-utils

# Install Python packages
pip install -r requirements.txt

# Launch
streamlit run app.py
```

Open http://localhost:8501 in your browser.

---

## Supported answer sheet format

Optimised for **Manav Rachna University** End-Term / Supplementary cover pages that contain:

- Name of Candidate  
- Roll No  
- Marks Distribution table (Q.No + parts a–j + Total)

Other similar university formats will also work reasonably well.

---

## Accuracy note

Handwritten marks are difficult for any OCR engine.  
The app does its best and shows a confidence status.  
Always perform a quick visual review before locking the final marks – this is still far faster than typing every mark manually.

---

## Files

| File              | Purpose                                      |
|-------------------|----------------------------------------------|
| `app.py`          | Main Streamlit application                   |
| `requirements.txt`| Python dependencies                          |
| `README.md`       | This file                                    |
| `packages.txt`    | System packages for Streamlit Cloud (OCR)    |

---

Made for faculty who want to stop adding marks by hand.


## Supported inputs
- Images (JPG/PNG/WEBP)
- Single-page PDF
- Multi-page PDF (one page = one student)
