# Firewall Fix Guide

## Quick Summary
**Problem:** You can access the server locally, but other users on the network cannot.
**Cause:** Windows Firewall is blocking incoming connections on port 5000.
**Solution:** Add a firewall rule to allow port 5000.

---

## Automated Fix (Recommended)

### Option 1: Use the Automated Script

1. Locate the file: **`fix-firewall.bat`**
2. **Right-click** on it
3. Select **"Run as administrator"**
4. Follow the prompts
5. Done! ✓

### Option 2: Check Current Firewall Status

1. Run **`check-firewall.bat`** (no admin required)
2. This will show if the firewall rule exists
3. Follow the recommendations shown

---

## Manual Fix

If the automated script doesn't work, follow these manual steps:

### Method 1: Command Line

1. Press `Windows + X`
2. Select **"Command Prompt (Admin)"** or **"PowerShell (Admin)"**
3. Copy and paste this command:
   ```
   netsh advfirewall firewall add rule name="Translation Processor" dir=in action=allow protocol=TCP localport=5000
   ```
4. Press Enter
5. You should see: "Ok."

### Method 2: Windows Firewall GUI

1. Press `Windows + R`
2. Type: `wf.msc` and press Enter
3. Click **"Inbound Rules"** in the left panel
4. Click **"New Rule..."** in the right panel
5. Select **"Port"** → Click Next
6. Select **"TCP"** and enter port: `5000` → Click Next
7. Select **"Allow the connection"** → Click Next
8. Check all profiles (Domain, Private, Public) → Click Next
9. Name: `Translation Processor` → Click Finish

---

## Verification Steps

After adding the firewall rule:

1. **Start the server** (run `start.bat`)
2. **Get your correct IP address**:
   - Run `show-all-ips.bat` to see all available IPs
   - Look for one like 192.168.x.x or 10.x.x.x
   - NOT 127.0.0.1 or 169.254.x.x
3. **Run diagnostic tool**:
   - Run `diagnose-connection.bat` on the SERVER
   - This will verify everything is configured correctly
   - It checks server status, firewall, and shows correct IP
4. **From another computer on the network**:
   - Run `test-connection-from-client.bat` on the CLIENT machine
   - Enter the server's IP address when prompted
   - This will test if connection is possible
5. **If test passes**, open a web browser and go to:
   - `http://[YOUR-IP]:5000`
   - Example: `http://192.168.1.100:5000`
6. You should see the success page!

---

## Still Not Working?

If the firewall rule is added but connections still fail:

### Check 1: Third-Party Security Software
- Norton, McAfee, Kaspersky, etc. may have their own firewalls
- Check their settings and allow Python or port 5000

### Check 2: Network Configuration
- Ensure both computers are on the **same network**
- Not connected through VPN
- Not on separate VLANs (corporate networks)

### Check 3: Server is Actually Running
- Run `check-firewall.bat` to verify port 5000 is listening
- Make sure you see output in the Command Prompt window
- Don't close the Command Prompt window while the server is running

### Check 4: Corporate/Managed Network
- If on a corporate network, IT may block custom services
- Contact your network administrator
- May need to use a different approach (desktop app instead)

### Check 5: Test with Firewall Temporarily Disabled
**WARNING: Only for testing!**

1. Open Windows Defender Firewall
2. Click "Turn Windows Defender Firewall on or off"
3. Temporarily turn off for Private networks
4. Test if other computers can connect
5. **Turn firewall back on immediately**
6. If this worked, the issue is definitely firewall-related

---

## Removing the Firewall Rule (If Needed)

To remove the firewall rule later:

```batch
netsh advfirewall firewall delete rule name="Translation Processor" protocol=TCP localport=5000
```

---

## Alternative: Desktop Application Approach

If firewall issues cannot be resolved (e.g., corporate restrictions), we can pivot to a **desktop application** that:
- Installs on each user's machine individually
- No network/server required
- No firewall issues
- Each instance runs independently

Let us know if you need to go this route!
