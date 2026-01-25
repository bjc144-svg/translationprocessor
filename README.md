# Translation Processor - Smoke Test

This is a minimal test to verify that a Python web server can run on your network and be accessed by other machines.

## Purpose
Before building the full translation processing application, we need to verify:
- ✓ Python web server runs on your machine
- ✓ Other machines on your network can access it
- ✓ Security software doesn't block it

## Installation Instructions

### Step 1: Install Python
1. Download Python from: https://www.python.org/downloads/
2. **IMPORTANT**: During installation, check the box "Add Python to PATH"
3. Install with default settings
4. Verify installation by opening Command Prompt and typing: `python --version`

### Step 2: Download Project Files
You should have these files in a folder:
- `app.py` (the web server)
- `requirements.txt` (dependencies)
- `start.bat` (startup script)
- `README.md` (this file)

### Step 3: Run the Server
1. Double-click `start.bat`
2. First time will take a minute to install dependencies
3. You should see output showing the server address

### Step 4: Test on the Same Machine
1. Open a web browser (Chrome, Edge, Firefox)
2. Go to: `http://localhost:5000`
3. You should see a success page with server information

### Step 5: Test from Another Machine
1. Note the IP address shown on the success page (something like `192.168.1.100`)
2. On another computer on the same network, open a web browser
3. Go to: `http://[IP-ADDRESS]:5000` (replace with actual IP)
4. **SUCCESS**: If you see the same page, we're good to proceed! ✓
5. **BLOCKED**: If connection fails, security software may be blocking it

## Troubleshooting

### "Python is not recognized"
- Python is not installed or not in PATH
- Reinstall Python and check "Add Python to PATH"

### Can't access from other machines
1. Check Windows Firewall:
   - Open "Windows Defender Firewall"
   - Click "Allow an app through firewall"
   - May need to allow Python or port 5000
2. Try adding a firewall rule:
   - Open Command Prompt as Administrator
   - Run: `netsh advfirewall firewall add rule name="Python Web Test" dir=in action=allow protocol=TCP localport=5000`
3. Check antivirus/security software settings
4. Verify both machines are on the same network

### Server won't start
- Make sure port 5000 isn't already in use
- Try restarting your computer
- Check if antivirus is blocking Python

## Next Steps

**If smoke test succeeds:**
- Report back and we'll proceed with building the actual translation processor!

**If smoke test fails:**
- We'll pivot to a desktop application approach instead
- Desktop app installs on each machine, no network/server needed

## Stopping the Server

Press `CTRL+C` in the Command Prompt window running the server, or simply close the window.
