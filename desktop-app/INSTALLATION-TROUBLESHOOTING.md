# Installation Troubleshooting Guide

## Quick Fix - Use start-safe.bat

If `start.bat` fails, try **`start-safe.bat`** instead:
- Installs packages individually with better error messages
- Shows exactly which package failed
- More reliable on different Windows configurations

---

## Common Installation Issues

### Issue 1: "Failed to build pillow" or "Failed to build wheel"

**Cause:** Missing C++ build tools needed to compile some packages

**Solution A - Use Pre-built Packages (Easiest):**

1. Open Command Prompt as Administrator
2. Navigate to desktop-app folder:
   ```
   cd C:\Path\To\desktop-app
   ```
3. Run:
   ```
   python -m pip install --upgrade pip
   pip install --only-binary=:all: Pillow
   ```

**Solution B - Install Build Tools:**

1. Download Visual C++ Build Tools:
   https://visualstudio.microsoft.com/visual-cpp-build-tools/

2. Run installer and select:
   - ✓ "Desktop development with C++"
   - ✓ "MSVC v142 build tools"
   - ✓ "Windows 10 SDK"

3. Restart computer

4. Run `start-safe.bat` again

**Solution C - Skip Problematic Packages:**

The app can work without drag-and-drop support. Just use the Browse button.

---

### Issue 2: "No module named 'tkinterdnd2'"

**Cause:** Drag-and-drop package failed to install

**Solution:** This is now optional!

The app has been updated to work without drag-and-drop. You'll see:
- "Click Browse Button Below to Select Document"
- Instead of drag-and-drop area

Simply use the "Browse for File..." button instead.

If you want drag-and-drop:
```
pip install tkinterdnd2
```

---

### Issue 3: "No module named 'win32com'"

**Cause:** pywin32 not installed or not configured

**Solution:**

1. Install pywin32:
   ```
   pip install pywin32==306
   ```

2. Run post-install script:
   ```
   python venv\Scripts\pywin32_postinstall.py -install
   ```

3. If still fails, try:
   ```
   pip install --upgrade pywin32
   ```

---

### Issue 4: "Python is not recognized"

**Cause:** Python not installed or not in PATH

**Solution:**

1. Download Python from: https://www.python.org/downloads/

2. Run installer:
   - ✓ **CHECK "Add Python to PATH"** (very important!)
   - ✓ Install with default settings

3. Restart Command Prompt (or computer)

4. Verify installation:
   ```
   python --version
   ```

---

### Issue 5: PDF Conversion Fails

**Cause:** Microsoft Word not installed or not accessible

**Solution:**

1. **Install Microsoft Word** (required for PDF conversion)

2. If Word is installed but conversion still fails:
   - Run as Administrator
   - Check Word is not running with a document open
   - Try opening and closing Word once

3. **Alternative:** App will save as `.docx` with instructions
   - Manually open in Word
   - File → Save As → PDF

---

### Issue 6: "Access Denied" or Permission Errors

**Solution:**

1. Run `start-safe.bat` as Administrator:
   - Right-click `start-safe.bat`
   - Select "Run as administrator"

2. Or install in a folder where you have write permissions:
   - Move desktop-app to: `C:\Users\YourName\Documents\`
   - Avoid: `C:\Program Files\`

---

## Manual Installation

If automated scripts fail completely, install manually:

1. **Create virtual environment:**
   ```
   python -m venv venv
   ```

2. **Activate it:**
   ```
   venv\Scripts\activate
   ```

3. **Upgrade pip:**
   ```
   python -m pip install --upgrade pip setuptools wheel
   ```

4. **Install packages one at a time:**
   ```
   pip install python-docx==1.1.0
   pip install Pillow
   pip install pywin32==306
   pip install docx2pdf==0.1.8
   ```

5. **Run the app:**
   ```
   cd src
   python main.py
   ```

---

## Minimal Installation (If All Else Fails)

Install only the absolute essentials:

```
pip install python-docx
```

**Note:** Without other packages:
- No drag-and-drop (use Browse button)
- PDF conversion will fail (save as .docx instead)
- You'll need to manually convert to PDF using Word

---

## Verify Installation

Check if packages are installed correctly:

```python
python -c "import docx; print('python-docx: OK')"
python -c "import PIL; print('Pillow: OK')"
python -c "import win32com.client; print('pywin32: OK')"
python -c "import docx2pdf; print('docx2pdf: OK')"
```

Each line should print "OK". If you get errors, that package needs attention.

---

## Alternative: Use Online Python Environment

If installation is too problematic on your local machine:

1. Use a cloud-based Python environment:
   - Google Colab
   - Repl.it
   - PythonAnywhere

2. Upload the code

3. Run in browser

**Note:** Won't have Word for PDF conversion, but can generate .docx files.

---

## Getting Help

If still having issues:

1. **Check Python version:**
   ```
   python --version
   ```
   Should be 3.8 or higher

2. **Check Windows version:**
   - Windows 10 or 11 recommended
   - Windows 7/8 may have compatibility issues

3. **Try different Python version:**
   - Python 3.9 or 3.10 often more compatible
   - Avoid very latest Python (3.12+) as some packages may not support it yet

4. **Check antivirus:**
   - Some antivirus software blocks pip installations
   - Temporarily disable and try again

5. **Run Windows Update:**
   - Ensure Windows is up to date
   - Some packages need recent Windows updates

---

## Success Checklist

After successful installation, you should see:

✓ Virtual environment created (`venv/` folder exists)
✓ Application window opens with "Translation Processor" title
✓ File selection area (with Browse button)
✓ Metadata form (Case Number, Languages, etc.)
✓ Translator Management button works
✓ No error messages in console

---

## Contact Support

If none of these solutions work:

- Email: eval@parkeval.com
- Include:
  - Python version (`python --version`)
  - Windows version
  - Complete error message
  - Screenshot of error if possible
