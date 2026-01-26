# Router Isolation / AP Isolation Fix Guide

## Problem: Wired and Wireless Devices Cannot Communicate

**Symptom:** Server on wired Ethernet cannot be accessed by devices on WiFi (or vice versa), even though both are on the same network (192.168.1.x).

**Cause:** Router has "AP Isolation", "Client Isolation", or "Wireless Isolation" enabled. This is a security feature that prevents devices from communicating with each other, even on the same network.

---

## Quick Test to Confirm This is the Issue

**From the WiFi client (192.168.1.155), open Command Prompt and run:**

```
ping 192.168.1.244
```

**Result:**
- If ping **FAILS** (timeout/unreachable): This is router isolation ✓
- If ping **WORKS**: Problem is something else (firewall, port block)

---

## Solution 1: Disable Router Isolation (Recommended)

### Step 1: Access Your Router

1. Open web browser
2. Go to your router's IP (usually one of these):
   - http://192.168.1.1
   - http://192.168.0.1
   - http://10.0.0.1
3. Login with admin credentials
   - Check router label for default username/password
   - Common: admin/admin, admin/password, admin/(blank)

### Step 2: Find Isolation Settings

Look for these settings (location varies by router brand):

**Common Names:**
- AP Isolation
- Client Isolation
- Wireless Isolation
- Station Isolation
- Wireless Privacy
- Multicast to Unicast

**Common Locations:**
- Wireless → Advanced Settings
- Wireless → Security
- Advanced → Wireless Settings
- LAN → DHCP Settings
- Guest Network Settings (if using guest network)

### Step 3: Disable Isolation

1. Find the isolation setting
2. **Disable** or **Uncheck** it
3. Save settings
4. Reboot router (usually required)
5. Wait 2-3 minutes for router to fully restart

### Step 4: Test Again

1. From WiFi client, ping the server: `ping 192.168.1.244`
2. If ping works, try browser: `http://192.168.1.244:5000`

---

## Solution 2: Move Server to WiFi

If you cannot access router settings:

1. **Disconnect wired Ethernet** on server
2. **Connect to same WiFi network** as other users
3. Server will get new IP address
4. Run `show-all-ips.bat` to find new IP
5. Give new IP to users
6. Test connection

**Pros:**
- No router configuration needed
- Works immediately

**Cons:**
- WiFi may be slower/less stable than wired
- Server IP may change if it reconnects

---

## Solution 3: Move Clients to Wired

If users can connect via Ethernet:

1. Have users disconnect from WiFi
2. Connect Ethernet cables
3. They'll get new IPs in 192.168.1.x range
4. Test connection to `http://192.168.1.244:5000`

**Pros:**
- Faster, more stable connection
- No router configuration needed

**Cons:**
- Requires physical Ethernet cables
- Less convenient

---

## Solution 4: Desktop Application Approach

If router settings cannot be changed and network isolation is required:

**Abandon server/client model and build desktop applications instead:**

- Each user installs application on their own machine
- No network communication needed
- No firewall/router issues
- Applications run independently
- Data can be shared via files, USB drives, or cloud storage

This is the best approach for:
- Corporate/managed networks
- Public WiFi
- Environments with strict security policies
- When router access is not available

---

## Router-Specific Guides

### Netgear Routers
1. Go to http://192.168.1.1 or http://routerlogin.net
2. Login (default: admin/password)
3. Navigate to: **Advanced → Wireless Settings**
4. Uncheck **"Enable Wireless Isolation"**
5. Click **Apply**

### TP-Link Routers
1. Go to http://192.168.0.1 or http://tplinkwifi.net
2. Login (default: admin/admin)
3. Navigate to: **Wireless → Wireless Settings → Advanced**
4. Uncheck **"Enable AP Isolation"**
5. Click **Save**

### Linksys Routers
1. Go to http://192.168.1.1 or http://myrouter.local
2. Login (default: admin/admin)
3. Navigate to: **Wireless → Basic Wireless Settings**
4. Uncheck **"AP Isolation"**
5. Click **Save Settings**

### ASUS Routers
1. Go to http://192.168.1.1 or http://router.asus.com
2. Login
3. Navigate to: **Wireless → Professional**
4. Set **"AP Isolated"** to **"No"**
5. Click **Apply**

### Google WiFi / Nest WiFi
1. Open Google Home app on phone
2. Tap on WiFi network
3. Tap **Settings → Advanced Networking**
4. Toggle **"Client Isolation"** to **OFF**

---

## Verification

After making changes:

1. **Ping test from client:**
   ```
   ping 192.168.1.244
   ```
   Should show replies, not timeouts

2. **Port test from client:**
   - Run `test-connection-from-client.bat`
   - Enter server IP: 192.168.1.244
   - Should pass both ping and port 5000 tests

3. **Browser test:**
   - Open browser on client
   - Navigate to: `http://192.168.1.244:5000`
   - Should see green success page

---

## Still Not Working?

If disabling isolation doesn't work:

1. **Router may have multiple isolation settings:**
   - Check both 2.4GHz and 5GHz bands
   - Check Guest Network settings separately
   - Look for VLAN settings

2. **Managed/Enterprise Router:**
   - May require IT admin access
   - May have security policies that cannot be changed
   - Consider desktop app approach

3. **Firewall at router level:**
   - Some routers have built-in firewalls that block ports
   - Check firewall settings in router
   - May need to create port forwarding rule (though this shouldn't be needed for local network)

4. **Contact Router Manufacturer Support:**
   - Provide router model number
   - Ask how to disable client isolation
   - Ask how to allow wired/wireless communication

---

## Alternative: Test with Both Devices on Same Connection Type

**Quick test to confirm isolation is the issue:**

1. Connect server to WiFi (instead of wired)
2. Run `show-all-ips.bat` to get new IP
3. Have client try new IP
4. If it works → Confirmed router isolation issue
5. If still doesn't work → Different problem

---

## Need Help?

If you can't access router settings or isolation setting doesn't exist:
- Check router manual for isolation/privacy settings
- Look up your router model + "disable AP isolation"
- Consider the desktop application approach instead
