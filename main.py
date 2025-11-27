# Extract text from the uploaded PDF and create CSV and XLSX files.
# The output files will be saved to /mnt/data/hr_list.csv and /mnt/data/hr_list.xlsx
# We'll parse each line looking for rows that start with a serial number and an email,
# and keep the remaining text as a combined "Title_and_Company" field when finer splitting isn't reliable.

from pathlib import Path
import re
import pandas as pd

pdf_path = Path("file_name.pdf")
assert pdf_path.exists(), "PDF not found at hr_list.pdf"

# Try using pdfplumber if available; otherwise fallback to PyPDF2 text extraction.
text = ""
try:
    import pdfplumber
    with pdfplumber.open(str(pdf_path)) as pdf:
        for p in pdf.pages:
            page_text = p.extract_text(x_tolerance=2, y_tolerance=2)
            if page_text:
                text += page_text + "\n"
except Exception as e:
    # fallback to PyPDF2
    try:
        import PyPDF2
        with open(pdf_path, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            for page in reader.pages:
                try:
                    text += page.extract_text() + "\n"
                except:
                    pass
    except Exception as e2:
        raise RuntimeError("Failed to extract text from PDF: " + str(e2))

# Normalize whitespace
lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

rows = []
# Regex: start with number(s), then name (greedy up to an email), then email, then rest
pattern = re.compile(r'^(\d+)\s+(.+?)\s+(\S+@\S+)\s+(.*)$')

for ln in lines:
    m = pattern.match(ln)
    if m:
        sno = m.group(1)
        name = m.group(2)
        email = m.group(3)
        rest = m.group(4).strip()
        # Heuristic: sometimes the rest contains both Title and Company. We'll keep it as-is in one field.
        title_and_company = rest
        rows.append({"SNo": sno, "Name": name, "Email": email, "Title_and_Company": title_and_company})
    else:
        # If a line doesn't match, it may be continuation of previous 'rest' (multiline row).
        # Attach it to the last row's Title_and_Company if plausible.
        if rows:
            rows[-1]["Title_and_Company"] += " " + ln

# Convert to DataFrame
df = pd.DataFrame(rows)

# Save CSV and XLSX
csv_path = "file_name.csv"
xlsx_path = "file_name.xlsx"
df.to_csv(csv_path, index=False)
df.to_excel(xlsx_path, index=False)

# Display a preview to the user (first 10 rows) using the UI helper
# import ace_tools as tools; tools.display_dataframe_to_user("HR list preview (first 10 rows)", df.head(10))

# Provide file paths for download
print(f"Saved CSV to: {csv_path}")
print(f"Saved Excel to: {xlsx_path}")
