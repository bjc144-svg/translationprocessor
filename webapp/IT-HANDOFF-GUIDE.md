# IT Handoff Guide
## Translation Processor Application - Production Deployment

**Prepared:** February 2, 2026
**Status:** Ready for Production Deployment
**Application Version:** 2.0 (Image-Based Processing)

---

## Table of Contents

1. [Application Review & Recommendations](#1-application-review--recommendations)
2. [Output Folder Management Strategy](#2-output-folder-management-strategy)
3. [Prerequisites & Dependencies](#3-prerequisites--dependencies)
4. [Deployment Instructions](#4-deployment-instructions)

---

## 1. Application Review & Recommendations

### Current Architecture

The Translation Processor is a Flask web application that:
- Accepts DOCX documents for translation processing
- Adds professional headers, footers, and certificates
- Converts pages to images to preserve exact layout (150 DPI → 120 DPI optimized)
- Supports division-based branding (Park vs EEI)
- Outputs professional PDF documents

### What We've Built

**Core Features:**
- ✅ Web-based interface (Flask)
- ✅ Document upload and processing
- ✅ Translator database management
- ✅ Division-specific branding (EEI vs Park)
- ✅ Per-page size preservation (handles mixed orientations)
- ✅ Image-based page conversion for layout accuracy
- ✅ Professional certificate generation
- ✅ Signature integration
- ✅ PDF compression optimization (30% size reduction via DPI)

**Recent Improvements:**
- ✅ Fixed pip upgrade error in start.bat
- ✅ Optimized image DPI for smaller file sizes (150 → 120 DPI)
- ✅ Reduced verbose logging for production
- ✅ Added path caching for LibreOffice/poppler
- ✅ Milestone-based logging (every 10th page)
- ✅ Removed DEBUG file output

### Recommendations for Production

#### Priority 1: CRITICAL (Implement Before Deployment)

**1.1 Change Flask Secret Key**

**Current Issue:** Secret key regenerates on every restart, invalidating sessions
**File:** `webapp/app.py` line 20

```python
# Current (BAD for production):
app.secret_key = secrets.token_hex(16)

# Change to (GOOD for production):
app.secret_key = 'your-permanent-secret-key-min-32-chars-here-abc123xyz789'
```

**How to generate:**
```python
import secrets
print(secrets.token_hex(32))
```

**Impact if not fixed:** Users will be logged out whenever server restarts

---

**1.2 Disable Debug Mode**

**Current Issue:** Debug mode is likely enabled (not confirmed in current view)
**File:** `webapp/app.py` (bottom of file)

```python
# Bad for production:
app.run(host='0.0.0.0', port=5000, debug=True)

# Good for production:
app.run(host='0.0.0.0', port=5000, debug=False)
```

**Impact if not fixed:** Security vulnerabilities, detailed error messages exposed to users

---

**1.3 Implement Output Folder Cleanup**

**Current Issue:** Output files accumulate indefinitely
**See Section 2 for detailed implementation**

---

#### Priority 2: RECOMMENDED (Implement Within First Month)

**2.1 Add Application Logging**

**Why:** Currently prints to console only, hard to troubleshoot production issues

**Implementation:**
Add to `app.py` after imports:

```python
import logging
from logging.handlers import RotatingFileHandler

# Configure logging
if not app.debug:
    # Create logs directory
    log_dir = Path(__file__).parent / 'logs'
    log_dir.mkdir(exist_ok=True)

    # File handler (10 MB max, keep 5 backups)
    file_handler = RotatingFileHandler(
        log_dir / 'translation_processor.log',
        maxBytes=10 * 1024 * 1024,  # 10 MB
        backupCount=5
    )
    file_handler.setFormatter(logging.Formatter(
        '[%(asctime)s] %(levelname)s in %(module)s: %(message)s'
    ))
    file_handler.setLevel(logging.INFO)
    app.logger.addHandler(file_handler)
    app.logger.setLevel(logging.INFO)
    app.logger.info('Translation Processor startup')
```

**Benefit:** Troubleshoot issues without accessing console

---

**2.2 Add Error Tracking**

**Why:** Know when errors occur without users reporting them

**Implementation:**
Add error handler to `app.py`:

```python
@app.errorhandler(Exception)
def handle_error(e):
    app.logger.error(f'Unhandled exception: {str(e)}', exc_info=True)
    flash('An unexpected error occurred. Please contact IT support.', 'error')
    return redirect(url_for('index'))
```

**Benefit:** Centralized error logging

---

**2.3 Add Health Check Endpoint**

**Why:** Monitor application health from monitoring systems

**Implementation:**
Add to `app.py`:

```python
@app.route('/health')
def health_check():
    """Health check endpoint for monitoring"""
    try:
        # Check critical folders exist
        checks = {
            'uploads_folder': UPLOAD_FOLDER.exists(),
            'output_folder': OUTPUT_FOLDER.exists(),
            'data_folder': DATA_FOLDER.exists(),
            'assets_folder': ASSETS_FOLDER.exists(),
        }

        if all(checks.values()):
            return jsonify({'status': 'healthy', 'checks': checks}), 200
        else:
            return jsonify({'status': 'unhealthy', 'checks': checks}), 500
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500
```

**Benefit:** Automated monitoring integration

---

#### Priority 3: OPTIONAL (Future Enhancements)

**3.1 Database Integration**

**Current:** Uses JSON files for translator data
**Consider:** SQLite or PostgreSQL for better concurrency and querying

**When to implement:** If you have >50 translators or need reporting

---

**3.2 Job Queue System**

**Current:** Processes documents synchronously (blocks during processing)
**Consider:** Celery + Redis for background processing

**When to implement:** If processing takes >30 seconds or you have >10 concurrent users

---

**3.3 User Authentication**

**Current:** No authentication (anyone with URL can access)
**Consider:**
- Windows Authentication (via IIS)
- Flask-Login with LDAP
- Azure AD integration

**When to implement:** If security requirements mandate authentication

---

**3.4 Advanced Compression (Phase 2)**

**Current:** 120 DPI (30% size reduction)
**Available:** Hybrid approach (JPEG + pikepdf) for 75-80% reduction

**When to implement:** If file sizes are still too large after testing

**See:** `/root/.claude/plans/rippling-splashing-thacker.md` for full analysis

---

### Code Quality Assessment

**Strengths:**
- ✅ Clean separation of concerns (utils modules)
- ✅ Proper error handling in upload/processing
- ✅ Secure filename handling
- ✅ File size validation
- ✅ Session management for file tracking
- ✅ Responsive web interface
- ✅ Milestone-based logging (production-friendly)

**Areas for Improvement:**
- ⚠️ No automated tests (consider adding unit tests)
- ⚠️ Limited logging (addressed in recommendations)
- ⚠️ No file cleanup (addressed in Section 2)
- ⚠️ No database (acceptable for current scale)

---

## 2. Output Folder Management Strategy

### The Problem

**Current Behavior:**
- Uploaded files: Deleted after processing ✅
- Output files: **Never deleted** ❌

**Impact:**
- Output folder grows indefinitely
- 10-page doc ≈ 17 MB (after DPI optimization)
- 100 docs/day = 1.7 GB/day = 51 GB/month
- Will fill disk within months

### Recommended Strategy: Multi-Tiered Approach

#### Option 1: Automated Cleanup Script (RECOMMENDED)

**Best for:** Most organizations
**Retention:** Keep files for 30 days, then auto-delete

**Implementation:**

Create `webapp/cleanup_output.py`:

```python
#!/usr/bin/env python3
"""
Cleanup script for Translation Processor output folder
Deletes files older than specified retention period
"""
from pathlib import Path
from datetime import datetime, timedelta
import logging

# Configuration
OUTPUT_FOLDER = Path(__file__).parent / 'output'
UPLOADS_FOLDER = Path(__file__).parent / 'uploads'
RETENTION_DAYS = 30  # Keep files for 30 days
DRY_RUN = False  # Set to True to test without deleting

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(Path(__file__).parent / 'logs' / 'cleanup.log'),
        logging.StreamHandler()
    ]
)

def cleanup_folder(folder, retention_days, file_pattern='*'):
    """Delete files older than retention period"""
    if not folder.exists():
        logging.warning(f"Folder does not exist: {folder}")
        return

    cutoff_date = datetime.now() - timedelta(days=retention_days)
    deleted_count = 0
    freed_space = 0

    logging.info(f"Cleaning up {folder} - deleting files older than {cutoff_date.strftime('%Y-%m-%d')}")

    for file in folder.glob(file_pattern):
        if file.is_file():
            file_modified = datetime.fromtimestamp(file.stat().st_mtime)

            if file_modified < cutoff_date:
                file_size = file.stat().st_size

                if DRY_RUN:
                    logging.info(f"[DRY RUN] Would delete: {file.name} ({file_size / 1024 / 1024:.2f} MB)")
                else:
                    try:
                        file.unlink()
                        logging.info(f"Deleted: {file.name} ({file_size / 1024 / 1024:.2f} MB)")
                        deleted_count += 1
                        freed_space += file_size
                    except Exception as e:
                        logging.error(f"Failed to delete {file.name}: {e}")

    if not DRY_RUN:
        logging.info(f"Cleanup complete: {deleted_count} files deleted, {freed_space / 1024 / 1024:.2f} MB freed")
    else:
        logging.info(f"[DRY RUN] Would delete {deleted_count} files, freeing {freed_space / 1024 / 1024:.2f} MB")

if __name__ == '__main__':
    logging.info("=== Translation Processor Cleanup Started ===")

    # Cleanup output folder (PDF files)
    cleanup_folder(OUTPUT_FOLDER, RETENTION_DAYS, '*.pdf')

    # Cleanup uploads folder (orphaned files)
    cleanup_folder(UPLOADS_FOLDER, 1, '*.docx')  # Delete uploads after 1 day

    logging.info("=== Cleanup Complete ===")
```

**Windows Task Scheduler Setup:**

Create `webapp/run_cleanup.bat`:
```batch
@echo off
REM Run cleanup script using virtual environment Python
cd /d "%~dp0"
call venv\Scripts\activate
python cleanup_output.py
deactivate
```

**Schedule in Windows:**
1. Open Task Scheduler
2. Create Basic Task
3. Name: "Translation Processor Cleanup"
4. Trigger: Daily at 2:00 AM
5. Action: Start a program
6. Program: `C:\TranslationProcessor\webapp\run_cleanup.bat`
7. Settings:
   - Run whether user is logged on or not
   - Run with highest privileges
   - Configure for: Windows Server 2016

**Linux Cron Setup:**
```bash
# Edit crontab
crontab -e

# Add line (runs daily at 2 AM):
0 2 * * * cd /opt/translationprocessor/webapp && ./venv/bin/python cleanup_output.py
```

---

#### Option 2: User-Initiated Cleanup (SIMPLE)

**Best for:** Small teams, low volume
**Behavior:** User downloads file, then it's deleted

**Implementation:**

Modify `webapp/app.py` download route:

```python
@app.route('/download')
def download():
    """Download processed PDF and optionally delete it"""
    output_file = session.get('output_file')

    if not output_file:
        flash('No file available for download', 'error')
        return redirect(url_for('index'))

    output_path = app.config['OUTPUT_FOLDER'] / output_file

    if not output_path.exists():
        flash('Output file not found', 'error')
        return redirect(url_for('index'))

    # Option A: Delete immediately after download
    @after_this_request
    def remove_file(response):
        try:
            output_path.unlink()
            app.logger.info(f"Deleted output file after download: {output_file}")
        except Exception as e:
            app.logger.error(f"Failed to delete output file: {e}")
        return response

    return send_file(
        output_path,
        as_attachment=True,
        download_name=output_file
    )

# Add this import at top of file:
from flask import after_this_request
```

**Pros:** Simple, no disk space issues
**Cons:** Files lost if user doesn't download, can't retrieve later

---

#### Option 3: Archive to Network Share (ENTERPRISE)

**Best for:** Large organizations, compliance requirements
**Behavior:** Move files to archive after 7 days, delete from archive after 1 year

**Implementation:**

Create `webapp/archive_output.py`:

```python
#!/usr/bin/env python3
"""
Archive old output files to network share
"""
from pathlib import Path
from datetime import datetime, timedelta
import shutil
import logging

# Configuration
OUTPUT_FOLDER = Path(__file__).parent / 'output'
ARCHIVE_PATH = Path('//file-server/translations/archive')  # Network share
ARCHIVE_AFTER_DAYS = 7
DELETE_ARCHIVE_AFTER_DAYS = 365

logging.basicConfig(level=logging.INFO)

def archive_old_files():
    """Move files older than threshold to archive"""
    cutoff_date = datetime.now() - timedelta(days=ARCHIVE_AFTER_DAYS)

    # Create year/month folder structure in archive
    archive_folder = ARCHIVE_PATH / datetime.now().strftime('%Y') / datetime.now().strftime('%m')
    archive_folder.mkdir(parents=True, exist_ok=True)

    for file in OUTPUT_FOLDER.glob('*.pdf'):
        if file.is_file():
            file_modified = datetime.fromtimestamp(file.stat().st_mtime)

            if file_modified < cutoff_date:
                try:
                    # Move to archive
                    archive_file = archive_folder / file.name
                    shutil.move(str(file), str(archive_file))
                    logging.info(f"Archived: {file.name}")
                except Exception as e:
                    logging.error(f"Failed to archive {file.name}: {e}")

def cleanup_old_archives():
    """Delete archived files older than threshold"""
    cutoff_date = datetime.now() - timedelta(days=DELETE_ARCHIVE_AFTER_DAYS)
    deleted_count = 0

    for file in ARCHIVE_PATH.rglob('*.pdf'):
        if file.is_file():
            file_modified = datetime.fromtimestamp(file.stat().st_mtime)

            if file_modified < cutoff_date:
                try:
                    file.unlink()
                    deleted_count += 1
                except Exception as e:
                    logging.error(f"Failed to delete archived file {file.name}: {e}")

    logging.info(f"Deleted {deleted_count} old archived files")

if __name__ == '__main__':
    archive_old_files()
    cleanup_old_archives()
```

**Schedule:** Run daily via Task Scheduler

---

### Recommended Configuration

**For most organizations:**
- ✅ Option 1 (Automated Cleanup) with 30-day retention
- ✅ Add application-level logging to track what was deleted
- ✅ Schedule cleanup to run daily at 2 AM
- ✅ Keep cleanup logs for troubleshooting

**Storage math:**
- 100 docs/day × 17 MB = 1.7 GB/day
- With 30-day retention: ~51 GB disk space needed
- With 7-day retention: ~12 GB disk space needed

---

## 3. Prerequisites & Dependencies

### Hardware Requirements

**Minimum:**
- CPU: 2 cores
- RAM: 4 GB
- Disk: 100 GB (with 30-day retention)
- Network: 100 Mbps LAN

**Recommended:**
- CPU: 4 cores
- RAM: 8 GB
- Disk: 250 GB SSD (faster processing)
- Network: 1 Gbps LAN

**Scaling guidance:**
- Add 2 GB RAM per 10 concurrent users
- Add 50 GB disk per additional month of retention

---

### Software Prerequisites

#### Required on Server

**1. Operating System**
- Windows Server 2016 or later
- Windows 10/11 Pro (for small deployments)
- OR Ubuntu Server 20.04+ (requires LibreOffice instead of Word)

**2. Python**
- Version: 3.8 or later (3.10+ recommended)
- Download: https://www.python.org/downloads/
- **CRITICAL:** Check "Add Python to PATH" during installation

**3. Microsoft Office**
- Product: Microsoft Word (any version from 2013+)
- License: Must be installed and activated on server
- Purpose: DOCX to PDF conversion
- Alternative: LibreOffice (free, but requires code changes)

**4. LibreOffice**
- Version: 7.0 or later
- Download: https://www.libreoffice.org/download/
- Purpose: DOCX to PDF conversion (primary method)
- Install: Standard installation, note install path

**5. Poppler Utilities**
- **Windows:**
  - Download: https://github.com/oschwartz10612/poppler-windows/releases/
  - Extract to: `C:\Program Files\poppler\`
  - Files needed: `pdftoppm.exe`, `pdfinfo.exe` in `bin` folder
- **Linux:**
  ```bash
  sudo apt-get install poppler-utils
  ```
- Purpose: PDF to image conversion

---

#### Python Packages (Auto-Installed)

These install automatically via `pip install -r requirements.txt`:

- Flask (web framework)
- python-docx (DOCX manipulation)
- Pillow (image processing)
- pdf2image (PDF conversion)
- pywin32 (Windows COM automation)
- docx2pdf (DOCX to PDF conversion)

**Note:** requirements.txt already exists in webapp folder

---

### Network Requirements

**Ports:**
- 5000/TCP (default Flask port)
- OR 80/443 if using IIS reverse proxy

**Firewall Rules:**
- Allow inbound on port 5000 from internal network
- No internet access required for application

**Firewall Command (Windows):**
```cmd
netsh advfirewall firewall add rule name="Translation Processor" dir=in action=allow protocol=TCP localport=5000
```

---

### File System Permissions

**Application folder needs:**
- Read: All files
- Write: `uploads/`, `output/`, `data/`, `logs/` folders
- Execute: Python scripts

**If running as Windows Service:**
- Service account needs read/write access to application folder
- Recommended: Create dedicated service account

---

### Browser Requirements (Client-Side)

**Supported:**
- Chrome 90+
- Edge 90+
- Firefox 90+
- Safari 14+

**Not Supported:**
- Internet Explorer

---

## 4. Deployment Instructions

### Quick Start (Development/Testing)

**Time: 30 minutes**

1. **Install Python 3.10+**
   ```
   Download from python.org
   Check "Add Python to PATH"
   ```

2. **Install LibreOffice**
   ```
   Download from libreoffice.org
   Use default installation path
   ```

3. **Install Poppler (Windows)**
   ```
   Download poppler-windows from GitHub
   Extract to C:\Program Files\poppler\
   ```

4. **Copy application files**
   ```
   Extract to: C:\TranslationProcessor\webapp\
   ```

5. **Run start.bat**
   ```
   Double-click webapp\start.bat
   First run installs all dependencies (5-10 minutes)
   ```

6. **Configure firewall**
   ```cmd
   netsh advfirewall firewall add rule name="Translation Processor" ^
     dir=in action=allow protocol=TCP localport=5000
   ```

7. **Test locally**
   ```
   Open browser: http://localhost:5000
   Upload test document
   Process and verify output
   ```

8. **Get server IP**
   ```cmd
   ipconfig
   Note IPv4 address (e.g., 192.168.1.100)
   ```

9. **Test from client**
   ```
   From another computer: http://192.168.1.100:5000
   ```

10. **Provide URL to users**
    ```
    http://[SERVER-IP]:5000
    ```

**Done!** Application is running and accessible on local network.

---

### Production Deployment (Windows Service)

**Time: 2 hours**
**For: Always-on production use**

#### Step 1: Complete Quick Start

Follow Quick Start steps 1-7 above.

---

#### Step 2: Apply Production Settings

**2.1: Change Secret Key**

Edit `webapp/app.py` line 20:

```python
# Generate a secure key first:
# python -c "import secrets; print(secrets.token_hex(32))"

# Then replace line 20:
app.secret_key = 'paste-your-generated-key-here-min-64-chars'
```

**2.2: Disable Debug Mode**

Edit `webapp/app.py` (near bottom):

```python
# Change debug=True to debug=False:
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=False)
```

**2.3: Add Logging (Optional but Recommended)**

See Priority 2 recommendations in Section 1.

---

#### Step 3: Install NSSM (Windows Service Manager)

**NSSM makes Python scripts run as Windows Services**

1. Download NSSM: https://nssm.cc/download
2. Extract `nssm.exe` (64-bit) to `C:\nssm\`
3. Add to PATH or use full path

---

#### Step 4: Create Windows Service

Open Command Prompt **as Administrator**:

```cmd
cd C:\TranslationProcessor\webapp

# Create service using NSSM
C:\nssm\nssm.exe install TranslationProcessor ^
  "C:\TranslationProcessor\webapp\venv\Scripts\python.exe" ^
  "C:\TranslationProcessor\webapp\app.py"

# Configure service
C:\nssm\nssm.exe set TranslationProcessor AppDirectory "C:\TranslationProcessor\webapp"
C:\nssm\nssm.exe set TranslationProcessor DisplayName "Translation Processor"
C:\nssm\nssm.exe set TranslationProcessor Description "Translation document processor web application"
C:\nssm\nssm.exe set TranslationProcessor Start SERVICE_AUTO_START

# Configure service recovery (auto-restart on failure)
sc failure TranslationProcessor reset= 86400 actions= restart/60000/restart/60000/restart/60000
```

---

#### Step 5: Start Service

```cmd
# Start service
net start TranslationProcessor

# Verify status
sc query TranslationProcessor

# Check if accessible
curl http://localhost:5000
```

---

#### Step 6: Configure Service Recovery

1. Open Services (`services.msc`)
2. Find "Translation Processor"
3. Right-click → Properties
4. Recovery tab:
   - First failure: Restart the Service (1 minute delay)
   - Second failure: Restart the Service (1 minute delay)
   - Subsequent failures: Restart the Service (1 minute delay)
5. Click OK

---

#### Step 7: Test Automatic Startup

```cmd
# Reboot server
shutdown /r /t 0

# After reboot, verify service started automatically
sc query TranslationProcessor

# Test application
curl http://localhost:5000
```

---

#### Step 8: Configure Output Cleanup

**8.1: Copy cleanup script**

Copy the cleanup script from Section 2 to `webapp/cleanup_output.py`

**8.2: Test cleanup script**

```cmd
cd C:\TranslationProcessor\webapp
venv\Scripts\activate
python cleanup_output.py
# Should complete without errors
```

**8.3: Create batch file**

Create `webapp/run_cleanup.bat`:
```batch
@echo off
cd /d "%~dp0"
call venv\Scripts\activate
python cleanup_output.py
deactivate
```

**8.4: Schedule cleanup**

1. Open Task Scheduler
2. Create Basic Task
3. Name: "Translation Processor Cleanup"
4. Description: "Clean up old translation output files"
5. Trigger: Daily at 2:00 AM
6. Action: Start a program
7. Program: `C:\TranslationProcessor\webapp\run_cleanup.bat`
8. Settings:
   - ✅ Run whether user is logged on or not
   - ✅ Run with highest privileges
   - ✅ Configure for: Windows Server 2016
9. Finish

**8.5: Test scheduled task**

```cmd
# Right-click task in Task Scheduler → Run
# Check logs/cleanup.log for results
```

---

### Optional: HTTPS via IIS Reverse Proxy

**Prerequisites:**
- IIS installed with URL Rewrite and ARR modules
- SSL certificate (can use self-signed for internal)

**Steps:**

1. **Install IIS components** (if not already installed)
   - URL Rewrite: https://www.iis.net/downloads/microsoft/url-rewrite
   - ARR: https://www.iis.net/downloads/microsoft/application-request-routing

2. **Enable ARR proxy**
   - IIS Manager → Server → Application Request Routing Cache
   - Server Proxy Settings → Enable proxy

3. **Create IIS site**
   - IIS Manager → Add Website
   - Site name: Translation Processor
   - Physical path: C:\inetpub\translations (create empty folder)
   - Binding: Port 80 (or 443 with SSL)

4. **Configure URL Rewrite**
   - Select site → URL Rewrite → Add Rules → Reverse Proxy
   - Server name: localhost:5000
   - Enable: "Rewrite the domain names"

5. **Add SSL certificate** (optional)
   - Site → Bindings → Add → HTTPS (port 443)
   - Select your SSL certificate
   - Add redirect rule HTTP → HTTPS

6. **Test**
   ```
   https://your-server-name
   ```

---

### Monitoring & Verification

**Check service health:**
```cmd
sc query TranslationProcessor
```

**Check application response:**
```cmd
curl http://localhost:5000
```

**View application logs:**
```
C:\TranslationProcessor\webapp\logs\translation_processor.log
```

**View cleanup logs:**
```
C:\TranslationProcessor\webapp\logs\cleanup.log
```

**Monitor disk space:**
```cmd
dir C:\TranslationProcessor\webapp\output
```

---

### Troubleshooting

**Service won't start:**
1. Check Python installation: `python --version`
2. Check virtual environment: `C:\TranslationProcessor\webapp\venv\Scripts\python.exe --version`
3. Test manually: `cd C:\TranslationProcessor\webapp && python app.py`
4. Check Event Viewer: Application logs
5. Check file permissions on application folder

**PDF conversion fails:**
1. Verify LibreOffice installed: `"C:\Program Files\LibreOffice\program\soffice.exe" --version`
2. Verify Poppler installed: `"C:\Program Files\poppler\bin\pdftoppm.exe" -v`
3. Check output folder permissions
4. Try processing manually to see detailed error

**Users can't connect:**
1. Verify service running: `sc query TranslationProcessor`
2. Test locally: `curl http://localhost:5000`
3. Check firewall: `netsh advfirewall firewall show rule name="Translation Processor"`
4. Verify server IP: `ipconfig`
5. Test from client: `telnet [server-ip] 5000`

**Cleanup not running:**
1. Check Task Scheduler → Task History
2. Run cleanup.bat manually
3. Check logs/cleanup.log for errors
4. Verify task credentials have folder permissions

---

### Rollback Procedure

If deployment fails or issues arise:

1. **Stop service:**
   ```cmd
   net stop TranslationProcessor
   ```

2. **Restore backup** (if you made one):
   ```cmd
   xcopy /E /I "C:\Backup\webapp" "C:\TranslationProcessor\webapp"
   ```

3. **Remove service** (if needed):
   ```cmd
   sc delete TranslationProcessor
   ```

4. **Reinstall from scratch:**
   - Follow Quick Start again with backup files

---

## Summary Checklist for IT

### Pre-Deployment
- [ ] Python 3.8+ installed on server
- [ ] LibreOffice installed
- [ ] Poppler utilities installed (Windows)
- [ ] Application files copied to server
- [ ] start.bat runs successfully
- [ ] Test document processes successfully locally
- [ ] Secret key changed in app.py
- [ ] Debug mode disabled in app.py

### Production Configuration
- [ ] Windows Service created and starts automatically
- [ ] Service recovery configured (auto-restart)
- [ ] Firewall rule added (port 5000)
- [ ] Output cleanup script installed
- [ ] Cleanup scheduled in Task Scheduler
- [ ] Application logging enabled (recommended)
- [ ] Health check endpoint added (recommended)

### Testing
- [ ] Service starts on server reboot
- [ ] Application accessible from server (localhost:5000)
- [ ] Application accessible from client computers
- [ ] Document upload works
- [ ] Document processing works (creates PDF)
- [ ] PDF downloads correctly
- [ ] Translator database accessible
- [ ] Cleanup script runs without errors

### Post-Deployment
- [ ] Monitor disk space (output folder)
- [ ] Monitor service uptime
- [ ] Check logs regularly for errors
- [ ] Backup data folder (translators.json, signatures)
- [ ] Document server details (IP, credentials)
- [ ] Train users on application
- [ ] Provide support contact to users

### Optional (If Using IIS/HTTPS)
- [ ] IIS installed with URL Rewrite and ARR
- [ ] Reverse proxy configured
- [ ] SSL certificate installed and working
- [ ] HTTP → HTTPS redirect working
- [ ] Custom domain configured (if applicable)

---

## Support Contacts

**Application Support:**
- Email: eval@parkeval.com (Park division)
- Email: eval@educei.com (EEI division)
- Phone: 212-581-8877

**Technical Issues:**
- Check logs first: `logs/translation_processor.log`
- Check cleanup logs: `logs/cleanup.log`
- Test manually: Run `start.bat` to see console output

---

## Appendix: File Structure

```
C:\TranslationProcessor\webapp\
├── app.py                      # Main Flask application
├── start.bat                   # Quick start script
├── requirements.txt            # Python dependencies
├── cleanup_output.py           # Output folder cleanup script (create this)
├── run_cleanup.bat             # Cleanup batch wrapper (create this)
├── venv\                       # Python virtual environment
├── assets\                     # Company logos and images
│   ├── park_logo.png
│   ├── eei_logo.png
│   ├── park_contact.png
│   └── park_signature.png
├── data\                       # Application data
│   ├── translators.json        # Translator database (BACKUP THIS)
│   ├── languages.json          # Language list
│   └── signatures\             # Translator signatures (BACKUP THIS)
├── uploads\                    # Temporary upload folder (auto-cleaned)
├── output\                     # Generated PDFs (configure cleanup)
├── logs\                       # Application logs (create this folder)
│   ├── translation_processor.log
│   └── cleanup.log
├── static\                     # Web assets (CSS, JS)
├── templates\                  # HTML templates
└── utils\                      # Python utility modules
    ├── document_processor.py
    ├── translator_manager.py
    └── language_manager.py
```

---

**Document prepared by:** Claude (AI Assistant)
**For:** Park Evaluation Services / EEI
**Version:** 1.0
**Date:** February 2, 2026
