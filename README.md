### 📄 **Script Description: `insert_images.py`**

This Python script automatically inserts images into an **active Microsoft Word document** by replacing specially formatted placeholders (e.g., `#0001`, `#0002`, etc.) with:

- The corresponding image file (from a specified folder),
- A **centered, auto-numbered caption** below the image (e.g., *Figure 1 — Main Window*),
- Proper paragraph formatting: **centered alignment**, **no first-line indent**, and **single line spacing**.

The caption text is extracted from the **filename** using the pattern:  
`#0001_Description.png` → caption becomes **"Figure 1 — Description"**.

The script uses **Word’s built-in `SEQ` (Sequence) fields** for numbering, enabling **cross-references** via Word’s *References → Cross-reference* feature.

> ✅ Ideal for technical documentation, reports, or any document requiring consistent, automated image insertion with professional captions.

---

### 📦 **Required Package**

- **`pywin32`** — provides Python access to Windows COM APIs, allowing automation of Microsoft Office applications.

---

### ⚙️ **Installation Instructions**

1. **Ensure Python is installed** (3.6 or newer; check with `python --version`).

2. **Install the required package** via pip:
   ```bash
   pip install pywin32
   ```

3. **(Optional but recommended)** After installing `pywin32`, run this command in an **elevated CMD/PowerShell** to register COM objects:
   ```cmd
   python Scripts/pywin32_postinstall.py -install
   ```
   *(This step is usually unnecessary on modern systems but may help if Word automation fails.)*

---

### ▶️ **How to Use**

1. Open a Word document containing placeholders like `#0001`, `#0002`, etc.
2. Place your images in the folder specified by `IMG_FOLDER` (e.g., `F:\`), named like:  
   `#0001_Main Window.png`, `#0002_Settings Dialog.jpg`, etc.
3. Run the script:
   ```bash
   python insert_images.py
   ```
4. The script will:
   - Replace each `#0001` with the matching image,
   - Add a formatted, auto-numbered caption below,
   - Preserve Word’s native caption system for cross-referencing.

> 💡 After running, press **Ctrl+A → F9** in Word to update all fields if numbers don’t appear immediately.

---

This tool bridges the gap between manual document editing and full-fledged document generation frameworks — perfect for quick, reliable, and professional image insertion! 🖼️📄
