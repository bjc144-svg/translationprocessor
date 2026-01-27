# Translation Processor - Desktop Application

**Park Evaluation Services**

Automated translation processing tool for creating certified translation PDFs with headers, footers, and certification templates.

---

## Features

✅ **Drag-and-Drop File Selection** - Easy file upload
✅ **Automated Header/Footer** - Company branding and case information
✅ **Dual Certification** - Translator and Park employee certificates
✅ **Metadata Injection** - Case number, language pairs, dates, signatures
✅ **Translator Management** - Shared database of translators and signatures
✅ **PDF Output** - Single combined PDF ready for Plunet upload
✅ **Offline Capable** - Works without internet (metadata syncs when online)

---

## Installation

### Prerequisites

1. **Windows 10 or later**
2. **Python 3.8 or higher**
   - Download from: https://www.python.org/downloads/
   - **IMPORTANT**: Check "Add Python to PATH" during installation
3. **Microsoft Word** (for PDF conversion)

### Setup

1. Extract all files to a folder (e.g., `C:\TranslationProcessor`)

2. Double-click `start.bat` to launch the application

3. First launch will:
   - Create a virtual environment
   - Install required dependencies (takes 1-2 minutes)
   - Start the application

---

## Quick Start Guide

### First Time Setup

1. **Add Translators**:
   - Click the "Manage" button next to Translator dropdown
   - Click "New"
   - Enter translator name
   - Add language pairs (one per line, e.g., "Spanish > English")
   - Browse and select translator's signature image (PNG, JPG)
   - Click "Save"
   - Repeat for all translators

2. **Add Company Assets**:
   - Place your company logo in: `assets/park_logo.png`
   - Place Park employee signature in: `assets/park_signature.png`

### Processing a Translation

1. **Select Document**:
   - Drag-and-drop Word file onto the gray area, OR
   - Click "Browse for File..." and select the document

2. **Enter Case Information**:
   - **Case Number**: Enter the Plunet case number
   - **Source Language**: Select or type source language
   - **Target Language**: Usually "English"
   - **Translator**: Select from dropdown
   - **Date**: Auto-filled with today's date (can edit)

3. **Process**:
   - Click "Process Translation"
   - Choose where to save the output PDF
   - Wait for processing to complete (usually 10-30 seconds)
   - PDF is ready for upload to Plunet!

4. **Clear**:
   - Click "Clear" to reset form for next document

---

## File Structure

```
desktop-app/
├── start.bat                 # Application launcher
├── requirements.txt          # Python dependencies
├── README.md                 # This file
├── src/                      # Source code
│   ├── main.py               # Main application
│   ├── translator_manager.py # Translator database
│   ├── translator_manager_gui.py # Translator management UI
│   └── document_processor.py # Document processing logic
├── data/                     # User data
│   ├── translators.json      # Translator database (shared)
│   └── signatures/           # Translator signature images
├── templates/                # Certificate templates (optional)
└── assets/                   # Company assets (logos, signatures)
    ├── park_logo.png         # Your company logo
    └── park_signature.png    # Park employee signature
```

---

## Shared Translator Database

The translator database (`data/translators.json`) can be shared across users:

### Option 1: Network Shared Folder

1. Place the entire `data/` folder on a shared network drive
2. All users point to the same folder
3. Translators added by any user are immediately available to all

### Option 2: Cloud Storage (OneDrive, Dropbox, Google Drive)

1. Move the `data/` folder to your cloud storage folder
2. All users sync to the same cloud folder
3. Automatic synchronization when online

### Option 3: Manual Sync

1. User A exports `translators.json`
2. Send file to User B
3. User B imports the file
4. Repeat as needed

**Note**: Cloud sync (Option 2) is recommended for automatic updates.

---

## Output Format

The generated PDF contains:

1. **Translated Document**
   - Header: Company logo + contact info
   - Footer: "CERTIFIED TRANSLATION" + case # + page numbers + language pair
   - Original document content preserved (tables, images, formatting)

2. **Translator Certificate** (Page break)
   - Company logo
   - Certification statement
   - Translator name and signature
   - Language pair and case number

3. **Park Certificate** (Page break)
   - Company logo
   - Company certification statement
   - Park employee signature
   - Language pair and case number

---

## Customization

### Header & Footer

Edit in: `src/document_processor.py` → `add_header_footer()` method

- Change contact info
- Adjust fonts, sizes, alignment
- Modify footer text

### Certificate Templates

Edit in: `src/document_processor.py`:

- `create_translator_certificate()` - Translator cert
- `create_park_certificate()` - Park cert

Modify certification text, layout, or add additional fields.

### Company Assets

Replace files in `assets/` folder:

- `park_logo.png` - Company logo (recommended: 300x100 pixels)
- `park_signature.png` - Employee signature (recommended: 300x100 pixels)

---

## Troubleshooting

### "Python is not recognized"

- Python not installed or not in PATH
- Reinstall Python and check "Add Python to PATH"
- Restart computer after installation

### "Failed to convert to PDF"

- Microsoft Word must be installed
- If Word is not available, the app saves as `.docx` instead
- Manually convert to PDF using Word or online tools

### Drag-and-drop not working

- Make sure file is a `.docx` file (not `.doc`)
- Try using "Browse for File..." button instead
- Check file is not open in Word

### Translator list is empty

- Click "Manage" button
- Add translators with "New" button
- Make sure to click "Save" after adding

### Signature image not showing

- Check image format (PNG, JPG recommended)
- Image should be clear with transparent background
- Test by re-adding signature in Translator Management

### Changes not syncing between users

- Check that all users point to same `data/` folder
- Verify network/cloud folder permissions
- Click "Refresh" in Translator Management

---

## Version 2 Roadmap

Future enhancements planned:

- ✨ **Plunet API Integration** - Direct download/upload
- ✨ **Batch Processing** - Process multiple documents at once
- ✨ **Template Library** - Multiple header/footer templates
- ✨ **OCR Support** - Extract text from scanned documents
- ✨ **Quality Checks** - Validate document before processing
- ✨ **Statistics Dashboard** - Track processing metrics
- ✨ **Multi-language UI** - Interface in multiple languages

---

## Support

For issues or questions:

- Check this README first
- Review troubleshooting section
- Contact your IT department
- Email: eval@parkeval.com

---

## License

© 2025 Park Evaluation Services
Internal use only - Do not distribute

---

## Technical Details

**Built with:**
- Python 3.8+
- tkinter (GUI framework)
- python-docx (Word document processing)
- docx2pdf (PDF conversion)
- tkinterdnd2 (Drag-and-drop support)

**System Requirements:**
- Windows 10 or later
- 2GB RAM minimum
- 500MB disk space
- Microsoft Word 2013 or later (for PDF conversion)

**Data Storage:**
- Local JSON database
- No internet required for processing
- Optional cloud sync for translator database
