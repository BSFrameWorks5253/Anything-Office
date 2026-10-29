# Anything Office (Ao)

> **The Unified Desktop Workspace for Docs, Sheets, Slides & PDF**  
> *Engineered for the Anything Creative Ecosystem*

---

## 🌟 Overview

**Anything Office** unifies the capabilities of Microsoft Word, Excel, PowerPoint, and Adobe Acrobat into a single, high-performance desktop application. Built with sub-second boot times, zero-flash frameless window chrome, and an offline-first architecture.

- **Docs (Word & Writer)**: Flowable paragraph text buffer with A4 layout, inch measurement ruler, inline tables, and document stats.
- **Sheets (Excel & Calc)**: Reactive Directed Acyclic Graph (DAG) spreadsheet calculation engine with multi-sheet tabs and formulas (`SUM`, `AVERAGE`, `COUNT`, `MIN`, `MAX`, `IF`, `CONCAT`).
- **Slides (PowerPoint & Impress)**: Absolute coordinate presentation canvas with 8-point shape handles, slide deck thumbnail navigation, and presenter mode (`F5`).
- **PDF (Adobe Acrobat Suite)**: Native PDF reader with an annotation suite:
  - **Typewriter Text**: Dynamic inline text note placement.
  - **Highlight & Markup**: Yellow highlighter with Underline and Strikethrough options.
  - **Sticky Notes**: Interactive comment pins with threaded replies.
  - **Freehand Drawing (Ink)**: SVG ink sketching with customizable brush sizes and palettes.
  - **Digital Signatures**: 3-mode signature workflow (Draw, Type, Upload) with "Digitally Verified" audit stamp.
  - **Word PDF Reflow**: One-click conversion from PDF text and tables into editable Word paragraphs.
- **Universal Loader Receiver**: Automatically synchronizes startup splash artwork, typography, and duration from the master **Anything Hub**.

---

## 📥 Windows Setup (.exe) Releases

Official Windows setup releases are built as clean, standalone Windows Setup installers:

```
Anything-Office-Setup-1.0.0.exe
```

Download the latest installer directly from the [GitHub Releases](https://github.com/BSFrameworks5253/Anything-Office/releases) page.

---

## 🛠️ Development & Local Build

### Prerequisites
- [Node.js](https://nodejs.org/) v18+
- [Git](https://git-scm.com/)

### Clone & Install
```bash
git clone https://github.com/BSFrameworks5253/Anything-Office.git
cd Anything-Office
npm install
```

### Run Locally (Desktop App)
```bash
npm run build
npm start
```

### Build Windows Setup (.exe) Locally
```bash
npm run build:exe
```
The compiled installer will be output to:
```
release/Anything-Office-Setup-1.0.0.exe
```

---

## 📄 Supported Document Formats

| Mode | Supported Extensions |
| :--- | :--- |
| **Word Processor** | `.docx`, `.doc`, `.odt`, `.txt`, `.md`, `.html` |
| **Spreadsheets** | `.xlsx`, `.xls`, `.ods`, `.csv`, `.tsv` |
| **Presentations** | `.pptx`, `.ppt`, `.odp` |
| **PDF Documents** | `.pdf` |
| **Unified Workspace** | `.anything`, `.json` |

---

## 📜 License
MIT License • Anything Creative Ecosystem
