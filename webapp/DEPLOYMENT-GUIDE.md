## Deployment Guide for IT Teams

**Translation Processor Web Application**

This guide helps IT teams deploy the Translation Processor web application for production use.

---

## Overview

- **Type:** Flask web application (Python)
- **Port:** 5000 (default, configurable)
- **Requirements:** Python 3.8+, Microsoft Word (for PDF conversion)
- **Users:** Multiple concurrent users supported
- **Data:** Shared translator database (JSON file)

---

## Deployment Options

### Option 1: Local Network Server (Recommended for Testing)

**Best for:** Quick deployment, small teams (< 10 users)

**Steps:**

1. **Install Python 3.8+ on Windows Server/PC**
   ```
   Download from: https://www.python.org/downloads/
   Check "Add Python to PATH" during installation
   ```

2. **Extract application files to server**
   ```
   Example: C:\Translation Processor\webapp\
   ```

3. **Run start.bat**
   ```
   Double-click start.bat
   First run installs dependencies automatically
   ```

4. **Configure Windows Firewall**
   ```cmd
   netsh advfirewall firewall add rule name="Translation Processor" dir=in action=allow protocol=TCP localport=5000
   ```

5. **Get server IP address**
   ```cmd
   ipconfig
   ```
   Note the IPv4 address (e.g., 192.168.1.100)

6. **Provide URL to users**
   ```
   http://192.168.1.100:5000
   ```

**Pros:**
- Quick setup (30 minutes)
- No complex configuration
- Works on existing Windows computers

**Cons:**
- Server must stay running
- No automatic restart if crashes
- Limited scalability

---

### Option 2: Windows Service (Recommended for Production)

**Best for:** Production use, always-on availability

**Steps:**

1. **Complete Option 1 steps 1-2**

2. **Install additional tools**
   ```cmd
   cd C:\TranslationProcessor\webapp
   venv\Scripts\activate
   pip install pywin32
   pip install pyinstaller
   ```

3. **Create Windows Service**
   ```cmd
   python -m win32serviceutil install
   sc create TranslationProcessor binPath="C:\TranslationProcessor\webapp\venv\Scripts\python.exe C:\TranslationProcessor\webapp\app.py" start=auto
   ```

4. **Start service**
   ```cmd
   sc start TranslationProcessor
   ```

5. **Verify service**
   ```cmd
   sc query TranslationProcessor
   ```

6. **Configure to start on boot**
   - Open Services (services.msc)
   - Find "Translation Processor"
   - Set Startup Type: Automatic
   - Set Recovery: Restart the Service

**Pros:**
- Runs automatically on server startup
- Restarts on failure
- Runs in background
- Professional deployment

**Cons:**
- Slightly more complex setup
- Requires administrator privileges

---

### Option 3: IIS Reverse Proxy (Recommended for HTTPS)

**Best for:** HTTPS access, professional URLs, existing IIS infrastructure

**Prerequisites:**
- Windows Server with IIS
- URL Rewrite module
- Application Request Routing (ARR)

**Steps:**

1. **Complete Option 2 (Windows Service)**

2. **Install IIS components**
   - URL Rewrite: https://www.iis.net/downloads/microsoft/url-rewrite
   - Application Request Routing: https://www.iis.net/downloads/microsoft/application-request-routing

3. **Create IIS site**
   - Open IIS Manager
   - Add Website
   - Name: Translation Processor
   - Physical Path: (any folder, not used)
   - Binding: Port 80 (or 443 for HTTPS)

4. **Configure URL Rewrite**
   - Select site in IIS
   - Open "URL Rewrite"
   - Add Rule → Reverse Proxy
   - Server: localhost:5000
   - Enable: "Rewrite the domain names of links in HTTP responses"

5. **Enable Proxy in ARR**
   - IIS → Server → Application Request Routing
   - Server Proxy Settings
   - Enable Proxy
   - Save

6. **Add SSL Certificate (optional but recommended)**
   - Request/import SSL certificate
   - Bind to HTTPS (port 443)
   - Redirect HTTP to HTTPS

7. **Configure DNS**
   - Create A record: `translations.yourcompany.com`
   - Points to server IP

**Access:**
```
https://translations.yourcompany.com
```

**Pros:**
- HTTPS encryption
- Professional URL
- Integrated with existing infrastructure
- Can add authentication (Windows Auth, etc.)

**Cons:**
- More complex setup
- Requires IIS knowledge
- SSL certificate needed

---

### Option 4: Azure/Cloud Deployment

**Best for:** Remote teams, cloud-first organizations, scalability

**Azure App Service Steps:**

1. **Create Azure App Service**
   - Login to Azure Portal
   - Create Resource → Web App
   - Runtime: Python 3.10
   - Operating System: Windows (for Word support)
   - Region: Choose nearest

2. **Configure Application Settings**
   ```
   SCM_DO_BUILD_DURING_DEPLOYMENT=true
   WEBSITE_HTTPLOGGING_RETENTION_DAYS=7
   ```

3. **Deploy Code**

   **Option A - Git Deployment:**
   ```cmd
   git init
   git add .
   git commit -m "Initial commit"
   git remote add azure https://[your-app].scm.azurewebsites.net/[your-app].git
   git push azure master
   ```

   **Option B - ZIP Deployment:**
   ```cmd
   Compress-Archive -Path * -DestinationPath app.zip
   az webapp deployment source config-zip --resource-group [RG] --name [APP] --src app.zip
   ```

4. **Configure Custom Domain (optional)**
   - Add custom domain in Azure Portal
   - Configure DNS CNAME record
   - Add SSL certificate (free with App Service)

5. **Install Word on App Service (for PDF conversion)**
   - This is challenging on Azure
   - Alternative: Use cloud-based conversion service
   - Or use Linux + LibreOffice

**Pros:**
- Always available
- Auto-scaling
- Managed infrastructure
- Global access
- HTTPS by default

**Cons:**
- Monthly cost ($50-200/month)
- PDF conversion may need alternative
- Requires Azure knowledge

---

## Security Configuration

### 1. Change Secret Key

**File:** `app.py`

```python
# Replace this line:
app.secret_key = secrets.token_hex(16)

# With a fixed key for production:
app.secret_key = 'your-long-random-secure-key-here-minimum-32-characters'
```

Generate secure key:
```python
import secrets
print(secrets.token_hex(32))
```

### 2. Disable Debug Mode

**File:** `app.py`

```python
# Change this:
app.run(host='0.0.0.0', port=5000, debug=True)

# To this:
app.run(host='0.0.0.0', port=5000, debug=False)
```

### 3. Add Authentication (optional)

If you need user authentication:

```python
# Install Flask-Login
pip install Flask-Login

# Add to app.py:
from flask_login import LoginManager, login_required

# Add @login_required decorator to routes
@app.route('/')
@login_required
def index():
    # ...
```

Or use Windows Authentication via IIS.

### 4. HTTPS Only

**IIS:** Configure URL Rewrite to redirect HTTP to HTTPS

**Azure:** Enable "HTTPS Only" in configuration

### 5. File Upload Restrictions

Already configured in `app.py`:
- Maximum file size: 50 MB
- Allowed extensions: .docx only

### 6. Firewall Rules

**Allow inbound on port 5000:**
```cmd
netsh advfirewall firewall add rule name="Translation Processor" dir=in action=allow protocol=TCP localport=5000
```

**Restrict to specific IPs (if needed):**
```cmd
netsh advfirewall firewall set rule name="Translation Processor" new remoteip=192.168.1.0/24
```

---

## Monitoring & Maintenance

### Health Checks

**Check if service is running:**
```cmd
curl http://localhost:5000
```

**Windows Service status:**
```cmd
sc query TranslationProcessor
```

**IIS status:**
```cmd
Get-Website -Name "Translation Processor"
```

### Log Files

**Application logs:**
- Check console output (if running via start.bat)
- Check Windows Event Viewer (if running as service)

**IIS logs:**
- `C:\inetpub\logs\LogFiles\`

**Create custom logging:**

Add to `app.py`:
```python
import logging

logging.basicConfig(
    filename='app.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
```

### Disk Space Monitoring

**Monitor these folders:**
- `uploads/` - Temporary uploaded files
- `output/` - Generated PDFs

**Set up automated cleanup:**

**cleanup.bat:**
```batch
@echo off
REM Delete files older than 7 days
forfiles /p "C:\TranslationProcessor\webapp\uploads" /m * /d -7 /c "cmd /c del @path"
forfiles /p "C:\TranslationProcessor\webapp\output" /m * /d -7 /c "cmd /c del @path"
```

**Schedule as Windows Task:**
- Open Task Scheduler
- Create Basic Task
- Trigger: Daily
- Action: Start a program → cleanup.bat

### Backup

**Critical files to backup:**
- `data/translators.json` - Translator database
- `data/signatures/` - Signature images
- `assets/` - Company logos/signatures

**Backup script (backup.bat):**
```batch
@echo off
set BACKUP_DIR=\\backup-server\translations\%DATE%
mkdir "%BACKUP_DIR%"
xcopy /E /I "C:\TranslationProcessor\webapp\data" "%BACKUP_DIR%\data"
xcopy /E /I "C:\TranslationProcessor\webapp\assets" "%BACKUP_DIR%\assets"
```

---

## Performance Optimization

### For Multiple Users

**Option 1: Gunicorn (Linux) or Waitress (Windows)**

```cmd
pip install waitress

# Replace app.run() in app.py with:
from waitress import serve
serve(app, host='0.0.0.0', port=5000, threads=4)
```

**Option 2: Multiple Workers**

Run multiple instances behind a load balancer:
- Instance 1: Port 5000
- Instance 2: Port 5001
- Instance 3: Port 5002
- Load balancer distributes traffic

### Caching

Add caching for translator list:

```python
from flask_caching import Cache

cache = Cache(app, config={'CACHE_TYPE': 'simple'})

@app.route('/api/translators')
@cache.cached(timeout=300)  # Cache for 5 minutes
def api_translators():
    # ...
```

---

## Troubleshooting

### Service Won't Start

1. Check Python installation:
   ```cmd
   python --version
   ```

2. Check dependencies:
   ```cmd
   cd C:\TranslationProcessor\webapp
   venv\Scripts\activate
   pip install -r requirements.txt
   ```

3. Test manually:
   ```cmd
   python app.py
   ```

4. Check Event Viewer for errors

### Users Can't Connect

1. Verify server is running:
   ```cmd
   curl http://localhost:5000
   ```

2. Check firewall:
   ```cmd
   netsh advfirewall firewall show rule name="Translation Processor"
   ```

3. Verify IP address:
   ```cmd
   ipconfig
   ```

4. Test from client:
   ```cmd
   telnet [SERVER-IP] 5000
   ```

### PDF Conversion Fails

1. Verify Word is installed on server
2. Try running as Administrator
3. Check output folder permissions
4. Test Word COM automation:
   ```python
   import win32com.client
   word = win32com.client.Dispatch("Word.Application")
   print("Word version:", word.Version)
   word.Quit()
   ```

---

## Rollback Plan

If deployment fails:

1. **Stop service:**
   ```cmd
   sc stop TranslationProcessor
   ```

2. **Restore backup:**
   ```cmd
   xcopy /E /I "\\backup-server\translations\[DATE]\data" "C:\TranslationProcessor\webapp\data"
   ```

3. **Restart service:**
   ```cmd
   sc start TranslationProcessor
   ```

---

## Support

### Internal IT Support

- Check logs first
- Verify service is running
- Test with simple document
- Check disk space

### Vendor Support

Email: eval@parkeval.com
Phone: 212-581-8877

---

## Checklist

**Pre-Deployment:**
- [ ] Python 3.8+ installed
- [ ] Microsoft Word installed (on server)
- [ ] Firewall configured
- [ ] Company logo/signature added to `assets/`
- [ ] Test translators added
- [ ] Secret key changed in app.py
- [ ] Debug mode disabled

**Post-Deployment:**
- [ ] Application accessible from server (localhost:5000)
- [ ] Application accessible from client computers
- [ ] File upload works
- [ ] Document processing works
- [ ] PDF download works
- [ ] Translator management works
- [ ] Monitoring configured
- [ ] Backup scheduled
- [ ] Cleanup scheduled
- [ ] Users trained

**Production Ready:**
- [ ] Windows Service configured
- [ ] Auto-restart enabled
- [ ] HTTPS configured (if needed)
- [ ] Custom domain configured (if needed)
- [ ] Authentication added (if needed)
- [ ] Documentation provided to users
- [ ] Support contact established

---

© 2025 Park Evaluation Services
