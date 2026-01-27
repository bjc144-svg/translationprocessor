# Translation Processor - Web Application

**Park Evaluation Services**

Web-based translation processing tool for creating certified translation PDFs with headers, footers, and certification templates.

---

## ✨ Features

✅ **Web-Based Interface** - Access from any browser on your network
✅ **Multi-User Support** - Multiple users can access simultaneously
✅ **Drag-and-Drop Upload** - Easy file upload
✅ **Automated Processing** - Headers, footers, and certificates
✅ **Shared Translator Database** - All users see the same translators
✅ **PDF Output** - Single combined PDF ready for Plunet
✅ **No Installation Needed** - Users just open a web browser

---

## 📋 Prerequisites

- **Windows Server or Windows 10/11**
- **Python 3.8+** ([Download](https://www.python.org/downloads/))
- **Microsoft Word** (for PDF conversion)
- **Network access** (if users need remote access)

---

## 🚀 Quick Start (Localhost Testing)

### 1. Install Python

1. Download Python from https://www.python.org/downloads/
2. **IMPORTANT**: Check "Add Python to PATH" during installation
3. Complete installation

### 2. Run the Application

1. Extract/navigate to the `webapp` folder
2. Double-click **`start.bat`**
3. Wait for dependencies to install (first time only)
4. Application starts on http://localhost:5000

### 3. Open in Browser

- On the server computer: http://localhost:5000
- You should see the Translation Processor interface

### 4. Test Functionality

1. Click "Manage Translators" → Add a test translator
2. Go back to "Process Translation"
3. Upload a test Word document
4. Fill in case information
5. Click "Process Translation"
6. Download the generated PDF

---

## 🌐 Network Access (For Your IT Team)

Once working on localhost, your IT team can deploy it properly:

### Option 1: Same Computer (Current Setup)

The app already listens on `0.0.0.0:5000`, making it accessible from the network.

**Users access via:**
```
http://[SERVER-IP]:5000
```

**Example:**
```
http://192.168.1.100:5000
```

### Option 2: Dedicated Server

Your IT team can:
1. Install on a Windows Server
2. Run as a Windows Service
3. Configure firewall rules
4. Set up proper DNS name (e.g., http://translations.yourcompany.com)

### Option 3: Cloud Hosting

Deploy to:
- Azure App Service
- AWS EC2
- Google Cloud
- Your company's existing web server

---

## 📁 File Structure

```
webapp/
├── app.py                     # Main Flask application
├── start.bat                  # Startup script
├── requirements.txt           # Python dependencies
│
├── templates/                 # HTML templates
│   ├── layout.html           # Base template
│   ├── index.html            # Main processing page
│   ├── translators.html      # Translator management
│   └── download.html         # Download page
│
├── static/                    # Static assets
│   ├── css/
│   │   └── style.css         # Stylesheet
│   └── js/
│       └── app.js            # JavaScript
│
├── utils/                     # Backend utilities
│   ├── document_processor.py # Document processing
│   └── translator_manager.py # Translator database
│
├── data/                      # Application data
│   ├── translators.json      # Translator database
│   └── signatures/           # Translator signatures
│
├── uploads/                   # Temporary uploads
├── output/                    # Generated PDFs
└── assets/                    # Company assets
    ├── park_logo.png         # Your logo
    └── park_signature.png    # Employee signature
```

---

## 👥 User Workflow

### For End Users:

1. **Open Browser** → Navigate to `http://[SERVER-IP]:5000`
2. **Upload Document** → Drag & drop or click to browse
3. **Enter Case Info** → Fill in case number, languages, translator
4. **Process** → Click "Process Translation"
5. **Download PDF** → Save and upload to Plunet

### For Administrators:

1. **Manage Translators** → Add/remove translators
2. **Upload Signatures** → Add translator signature images
3. **Monitor** → Check server logs for issues

---

## ⚙️ Configuration

### Add Company Assets

Place these files in the `assets/` folder:

- **`park_logo.png`** - Your company logo (300x100px recommended)
- **`park_signature.png`** - Park employee signature (300x100px)

### Customize Templates

Edit files in `utils/document_processor.py`:
- Header/footer layout
- Certificate text
- Styling and formatting

---

## 🔧 Troubleshooting

### Server won't start

**Check Python installation:**
```
python --version
```
Should show Python 3.8 or higher

**Try manual start:**
```
cd webapp
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

### Can't access from other computers

1. **Check Windows Firewall:**
   - Open "Windows Defender Firewall with Advanced Security"
   - Create inbound rule for port 5000

2. **Get server IP address:**
   ```
   ipconfig
   ```
   Look for IPv4 Address (e.g., 192.168.1.100)

3. **Test from another computer:**
   ```
   http://192.168.1.100:5000
   ```

### PDF conversion fails

- Ensure Microsoft Word is installed on the server
- Try running server as Administrator
- Check `output/` folder permissions

### File upload fails

- Check `uploads/` folder exists and is writable
- Verify file is .docx format (not .doc)
- Check file size (maximum 50 MB)

---

## 🔒 Security Considerations

### For Production Deployment:

1. **Change Secret Key** (in `app.py`):
   ```python
   app.secret_key = 'your-secure-random-key-here'
   ```

2. **Disable Debug Mode** (in `app.py`):
   ```python
   app.run(host='0.0.0.0', port=5000, debug=False)
   ```

3. **Use HTTPS** (configure with reverse proxy like IIS/Nginx)

4. **Add Authentication** (if needed for your network)

5. **Set File Size Limits** (already configured at 50 MB)

6. **Regular Cleanup** of uploads/ and output/ folders

---

## 📊 IT Team Deployment Guide

### Windows Server Deployment

1. **Install Python 3.8+** on server

2. **Create Windows Service:**
   ```
   pip install pywin32
   python -m win32serviceutil install
   ```

3. **Configure Firewall:**
   ```
   netsh advfirewall firewall add rule name="Translation Processor" dir=in action=allow protocol=TCP localport=5000
   ```

4. **Set up IIS Reverse Proxy** (optional, for HTTPS):
   - Install IIS with Application Request Routing
   - Configure reverse proxy to http://localhost:5000
   - Add SSL certificate

5. **Configure DNS:**
   - Create A record: `translations.yourcompany.com`
   - Points to server IP

6. **Set up Monitoring:**
   - Check if service is running
   - Monitor disk space (uploads/output folders)
   - Check application logs

### Azure/Cloud Deployment

1. **Create App Service** (Windows)
2. **Configure Python runtime**
3. **Deploy code** via Git or ZIP
4. **Set environment variables**
5. **Configure custom domain and SSL**

---

## 🔄 Maintenance

### Regular Tasks:

1. **Clean up old files:**
   - Delete old files in `uploads/`
   - Archive or delete old PDFs in `output/`

2. **Backup translator database:**
   - Copy `data/translators.json`
   - Copy `data/signatures/` folder

3. **Update Python packages:**
   ```
   pip install --upgrade -r requirements.txt
   ```

### Monitoring:

- Check application logs for errors
- Monitor disk space
- Verify PDF generation is working
- Test from different browsers/devices

---

## 🆘 Support

### Common Issues:

| Issue | Solution |
|-------|----------|
| "Python not found" | Install Python, check PATH |
| "Port 5000 already in use" | Change port in app.py or stop other app |
| "Can't connect remotely" | Check firewall, verify IP address |
| "PDF conversion fails" | Install Microsoft Word on server |
| "Upload fails" | Check file is .docx and under 50MB |

### Getting Help:

- Check server console for error messages
- Review `uploads/` and `output/` folder permissions
- Test with simple Word document first
- Contact: eval@parkeval.com

---

## 🚀 Version 2 Features (Future)

Planned enhancements:
- ✨ Plunet API integration (direct download/upload)
- ✨ User authentication and permissions
- ✨ Batch processing
- ✨ Processing history and analytics
- ✨ Email notifications
- ✨ Custom template library
- ✨ Mobile-responsive design improvements

---

## 📄 License

© 2025 Park Evaluation Services
Internal use only - Do not distribute

---

## 📞 Contact

**Park Evaluation Services**
Phone: 212-581-8877
Email: eval@parkeval.com
Web: www.parkeval.com
