# Translation Processor

A web-based application for processing translation documents with certified headers, footers, and certificates.

## Quick Start

1. **Install Python** (if not already installed)
   - Download from: https://www.python.org/downloads/
   - During installation, check "Add Python to PATH"
   - Verify: Open Command Prompt and type `python --version`

2. **Navigate to the webapp folder**
   ```
   cd webapp
   ```

3. **Run the application**
   - Double-click `start.bat` (or `start-fixed.bat` if you have network issues)
   - First run will install dependencies (takes 1-2 minutes)
   - Server will start on port 5000

4. **Access the application**
   - On the same computer: http://localhost:5000
   - From other computers: http://[YOUR-IP]:5000

## Features

- Upload Word documents (.doc or .docx)
- Add professional headers and footers with Park branding
- Generate translator and Park certificates
- Export as combined PDF
- Manage translators and their signatures
- Support for multiple languages

## Project Structure

```
translationprocessor/
├── README.md           (this file)
└── webapp/            (the application)
    ├── start.bat      (run this to start the app)
    ├── app.py         (main application)
    ├── requirements.txt
    ├── templates/     (HTML templates)
    ├── static/        (CSS, JS, images)
    ├── utils/         (document processing)
    ├── data/          (languages, translators)
    ├── assets/        (logos, signatures)
    ├── uploads/       (temporary uploads)
    └── output/        (generated PDFs)
```

## Documentation

For detailed documentation, see:
- `webapp/README.md` - Complete application guide
- `webapp/DEPLOYMENT-GUIDE.md` - Network and deployment instructions

## Troubleshooting

**Can't access from other computers?**
- Windows Firewall may be blocking port 5000
- See `webapp/DEPLOYMENT-GUIDE.md` for network troubleshooting

**Python not found?**
- Make sure Python is installed and added to PATH
- Restart Command Prompt after installing Python

**Dependencies won't install?**
- Make sure you have internet connection
- Try running `start.bat` as Administrator

## Support

For issues or questions, refer to the detailed documentation in the `webapp` folder.
