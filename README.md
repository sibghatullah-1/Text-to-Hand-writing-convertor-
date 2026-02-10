
# Text to Handwritten PDF Converter

**Text to Handwritten PDF Converter** is a powerful desktop application designed to convert typed text or digital documents into realistic, handwritten-style PDF pages.

Built for students and professionals, it bridges the gap between digital convenience and analog charm. The application utilizes advanced rendering techniques to ensure natural variation, handling everything from spacing and ink color to complex mathematical superscripts and subscripts.

---

## ✨ Features

### 🖋️ realistic Handwriting Engine

* **Multi-Font Support:** Select one or more handwriting styles to introduce natural randomness.
* **Dynamic Variation:** Automatic font variation per character to avoid the "robotic" look of standard script fonts.
* **Micro-Jitter & Rotation:** Applies random rotation and position jitter to glyphs for a genuine handwritten feel.

### 📄 Comprehensive Input Support

* **Direct Paste:** Paste text directly from your clipboard.
* **File Import:** Seamlessly import content from Word documents (`.docx`) and PDFs.
* **Formatting Preservation:** Maintains indentation, tables, and special characters.

### 📐 & 🎨 Customization & Layout

* **Page Layout:** Adjust margins, line spacing, word spacing, and character-level letter spacing.
* **Ink Control:** Choose between Blue, Black, or Red ink.
* **"Messiness" Factor:** Adjustable sliders to control how neat or chaotic the handwriting appears.
* **Backgrounds:** Choose built-in templates (Lined Paper, Dark Mode, Plain White) or import custom background images.

### 🧮 Math-Friendly Rendering

Unlike many handwriting generators, Handwrite.ai supports scientific notation:

* **Superscripts:** Properly renders exponents (e.g., `x²`, `nⁿ`, `⁺`).
* **Subscripts:** Properly renders indices (e.g., `H₂O`, `x₀`, `₋`).

### 📑 Professional Output

* **High-Resolution PDF:** Generates multi-page documents suitable for printing or submission.
* **Smart Pagination:** Intelligent word wrapping and auto-pagination.
* **Responsive UI:** Background threading ensures the app remains responsive during PDF generation with a visual progress indicator.

---

## 🖥️ Tech Stack

| Component | Technology | Purpose |
| --- | --- | --- |
| **Language** | Python 3 | Core logic and backend |
| **GUI Framework** | CustomTkinter | Modern, dark-mode friendly UI |
| **Image Processing** | Pillow (PIL) | Text rendering, rotation, and image manipulation |
| **Doc Parsing** | python-docx | Extracting text from Word files |
| **PDF Parsing** | pypdf | Extracting text from existing PDFs |

---

## 🚀 How It Works

1. **Import Data:** Paste your text or upload a `.docx` / `.pdf` file.
2. **Configure Style:** Select your preferred handwriting fonts and page background.
3. **Fine-Tune:** Adjust the "Messiness," font size, margins, and ink color.
4. **Generate:** Click **Generate PDF**. The app processes the text character-by-character using background threads.
5. **Export:** Your realistic handwritten PDF is saved with a timestamped filename.

---

## 📂 Project Highlights

* **Character-by-Character Rendering:** Ensures no two letters look exactly aligned, mimicking human motor control.
* **Dynamic Resource Paths:** Designed to be PyInstaller-compatible for easy conversion to a standalone `.exe`.
* **Modern UI:** A clean, dark-mode interface built with CustomTkinter.

---

## 📌 Use Cases

* **University Assignments:** Submit digital work that requires a handwritten format.
* **Personal Notes:** Digitize typed notes into an aesthetic handwritten journal.
* **Practice Sheets:** Create tracing or copying sheets.
* **Presentation Stylization:** Create unique handouts or slides.

---

## ⚠️ Disclaimer

> **Ethical Use Warning:**
> This project is intended for **educational and personal use only**.
> Users are responsible for ensuring compliance with their institution’s academic integrity policies. Misusing this tool to bypass handwriting requirements for exams or graded assignments where authentic handwriting is mandated may constitute academic misconduct.

---

Would you like me to generate a `requirements.txt` file content to go along with this README?
