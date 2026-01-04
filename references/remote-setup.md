# Remote Connection Setup

Connect Office applications on one machine to a bridge server running on another machine. Useful for scenarios like controlling your work laptop's Outlook from your personal laptop.

## Architecture

```
Personal Laptop (192.168.1.100)          Work Laptop
┌─────────────────────────────┐          ┌─────────────────────┐
│ Bridge Server    :3847      │◄────────►│ Office Application  │
│ Add-in Server    :3003      │          │  └─ Add-in (loaded  │
│                             │          │     from personal)  │
└─────────────────────────────┘          └─────────────────────┘
        Same WiFi Network
```

**Host machine** (personal laptop): Runs bridge server and serves add-in files
**Client machine** (work laptop): Runs Office with the add-in connected to host

## Prerequisites

- Both machines on the same network
- Host machine IP address (e.g., `192.168.1.100`)
- Ability to sideload Office add-ins on the client machine

## Setup Steps

### 1. Find Host Machine IP

```bash
# macOS
ipconfig getifaddr en0

# Windows
ipconfig | findstr IPv4

# Linux
hostname -I | awk '{print $1}'
```

Note your IP (e.g., `192.168.1.100`).

### 2. Generate SSL Certificates for IP

The default certificates are for `localhost`. Generate new ones for your IP:

```bash
cd office-bridge

# Create certs directory if needed
mkdir -p ~/.office-addin-dev-certs

# Generate self-signed cert for your IP
openssl req -x509 -newkey rsa:2048 -nodes \
  -keyout ~/.office-addin-dev-certs/localhost.key \
  -out ~/.office-addin-dev-certs/localhost.crt \
  -days 365 \
  -subj "/CN=192.168.1.100" \
  -addext "subjectAltName=IP:192.168.1.100,DNS:localhost"
```

Replace `192.168.1.100` with your actual IP.

### 3. Create Remote Manifest

Copy and modify the manifest for your app. Example for Outlook:

```bash
cd addins/outlook

# Copy manifest
cp manifest.xml manifest-remote.xml
```

Edit `manifest-remote.xml` and replace all `localhost` with your IP:

```xml
<!-- Before -->
<IconUrl DefaultValue="https://localhost:3003/assets/icon-32.png"/>
<SourceLocation DefaultValue="https://localhost:3003/taskpane.html"/>

<!-- After -->
<IconUrl DefaultValue="https://192.168.1.100:3003/assets/icon-32.png"/>
<SourceLocation DefaultValue="https://192.168.1.100:3003/taskpane.html"/>
```

Also update the `<AppDomain>` for the bridge:

```xml
<!-- Before -->
<AppDomain>https://localhost:3847</AppDomain>

<!-- After -->
<AppDomain>https://192.168.1.100:3847</AppDomain>
```

### 4. Start Servers on Host Machine

```bash
# Terminal 1: Bridge server
cd office-bridge
./server.sh

# Terminal 2: Add-in dev server (e.g., Outlook)
cd office-bridge/addins/outlook
npm run dev-server
```

The servers now accept connections from any IP on the network.

### 5. Trust Certificates on Client Machine

On the work laptop, open a browser and visit:

1. `https://192.168.1.100:3847` - Accept the certificate warning
2. `https://192.168.1.100:3003` - Accept the certificate warning

This tells the browser (and Office) to trust the self-signed certificates.

### 6. Sideload the Remote Manifest

**Outlook:**
1. In Outlook, go to **...** → **Get Add-ins**
2. Select **My add-ins** → **Custom Add-ins**
3. Click **Add from File**
4. Select the `manifest-remote.xml` from the host machine (copy it over or access via network share)

**Word/Excel/PowerPoint:**
```bash
# On client machine, create the manifest directory
mkdir -p ~/Library/Containers/com.microsoft.Outlook/Data/Documents/wef

# Copy the remote manifest (via network share, USB, etc.)
cp /path/to/manifest-remote.xml ~/Library/Containers/com.microsoft.Outlook/Data/Documents/wef/
```

### 7. Verify Connection

1. Open the Office application on the client machine
2. Click the Office Bridge add-in
3. The add-in should connect to the bridge on the host machine
4. Verify on host machine:

```bash
cd office-bridge && npx tsx <<'EOF'
import { connect } from "./src/client.js";
const bridge = await connect();
const sessions = await bridge.sessions();
console.log("Connected:", sessions.map(s => `[${s.app}] ${s.filename}`));
await bridge.close();
EOF
```

## Troubleshooting

### Add-in won't load

- Ensure both servers are running on host
- Verify the IP in the manifest matches the host IP
- Check that port 3003 (add-in) and 3847 (bridge) aren't blocked by firewall

### Certificate errors

- Visit both URLs in the browser first and accept warnings
- Regenerate certificates if IP changed
- On macOS, you may need to add the cert to Keychain and trust it

### Connection timeout

- Verify both machines are on the same network
- Check host firewall allows incoming connections on ports 3003 and 3847
- Try pinging the host IP from the client

### "WebSocket connection failed"

- The add-in auto-detects the bridge host from where it was loaded
- Verify the add-in is loading from the host IP (check network tab in dev tools)
- Ensure bridge server is running on host

## Security Considerations

- This setup uses self-signed certificates (not for production)
- Traffic is encrypted but certificate is not validated
- Only use on trusted networks
- The bridge accepts connections from any IP when running in this mode

## Firewall Configuration

If connections fail, ensure the host machine allows incoming connections:

**macOS:**
```bash
# Temporarily allow (resets on reboot)
sudo pfctl -d  # Disable packet filter

# Or add specific rules
sudo /usr/libexec/ApplicationFirewall/socketfilterfw --add /path/to/node
sudo /usr/libexec/ApplicationFirewall/socketfilterfw --unblockapp /path/to/node
```

**Windows:**
```powershell
# Allow Node.js through firewall
netsh advfirewall firewall add rule name="Office Bridge" dir=in action=allow program="C:\Program Files\nodejs\node.exe" enable=yes
```

## Port Reference

| Component | Port |
|-----------|------|
| Bridge Server | 3847 |
| Word Add-in | 3000 |
| Excel Add-in | 3001 |
| PowerPoint Add-in | 3002 |
| Outlook Add-in | 3003 |
