import pypdf
import os

def dump_pdf(filepath):
    print(f"Dumping {filepath}...")
    reader = pypdf.PdfReader(filepath)
    text = ""
    for page in reader.pages:
        text += page.extract_text()
    
    txt_filename = filepath.replace(".pdf", ".txt")
    with open(txt_filename, "w", encoding="utf-8") as f:
        f.write(text)
    print(f"  Saved to {txt_filename}")

if __name__ == "__main__":
    for f in sorted(os.listdir(".")):
        if f.endswith(".pdf") and "Previred" in f:
            dump_pdf(f)
