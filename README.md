
# Word Document Comparator with Paragraph, Table, and Image Detection

This Python project compares two `.docx` Word files and generates a detailed change report that highlights:

✅ Paragraph-level differences  
✅ Table content and formatting differences  
✅ Image presence, dimension changes (width/height)  
✅ Word-by-word edits, extra spaces, font/style/spacing changes

---

## 🔍 What It Does

- Compares two versions of a Word document (pre/post)
- Detects:
  - Text changes
  - Font name, size, bold/italic, line spacing, alignment
  - Table data cell-by-cell
  - Added/removed/resized images
- Outputs a `.docx` summary report with color-coded differences

---

## 🛠 Technologies

- Python 3.x
- `python-docx`

---

## 🚀 How to Use

1. Clone the repo:
   ```bash
   git clone https://github.com/yourusername/docx-comparator.git
   cd docx-comparator
