"""
Smart Schematic — PDF Link Generator

Generate a clickable PDF where reference designators, part numbers, and net names
are hyperlinks that trigger Allegro Free Viewer navigation or OnePDM lookup
via custom brdnav:// protocol.

Links are handled silently by brdnav_handler.pyw (no browser opens).
The handler auto-starts the server on first click.
"""

import fitz  # PyMuPDF
import re
import os
import sys
import json
import argparse
import platform
try:
    import winreg
except ImportError:
    winreg = None  # Not available on Linux — protocol registration skipped

# === Configuration ===
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_PDF = os.path.join(BASE_DIR, "input.pdf")
OUTPUT_PDF = os.path.join(BASE_DIR, "input_linked.pdf")
COMPONENT_DATA = os.path.join(BASE_DIR, "component_data.json")
# brdnav://nav/REFDES?layer=... — custom protocol, no browser opens.
# All links share host "nav" so Acrobat only asks permission once.
PROTOCOL = "brdnav"
PROTOCOL_HOST = "nav"

REFDES_PATTERN = re.compile(r'^(?:SW|[RUCLJDQ][A-Z]?)\d{1,5}$')
PART_NUMBER_PATTERN = re.compile(r'^[MX]\d{5,8}-\d{2,4}$')

# OnePDM datasheet lookup — links open local server endpoint
DATASHEET_PROTOCOL = "brdnav"
DATASHEET_HOST = "datasheet"

# Net name highlight in Allegro
NET_PROTOCOL = "brdnav"
NET_HOST = "net"
NET_DATA = os.path.join(BASE_DIR, "net_data.json")

# HTTP mode — uses http://127.0.0.1:PORT/ instead of brdnav://
# This avoids ASR "Block Adobe Reader from creating child processes" rule.
HTTP_PORT = 5588
_use_http = True

# Project ID — derived from BRD filename, embedded in links for multi-project support
_project_id = ""

# Component data (x, y, layer) — loaded in main() to support CLI args
_comp_data = {}
_net_whitelist = set()


def _register_project(proj_id, brd_abs, comp_data_path):
    """Register project in .brdnav_config.json so server can find it on cold start."""
    config_path = os.path.join(BASE_DIR, ".brdnav_config.json")
    try:
        with open(config_path) as f:
            registry = json.load(f)
    except Exception:
        registry = {}
    if "projects" not in registry:
        registry = {"projects": {}}
    entry = {"brd": brd_abs}
    if comp_data_path and os.path.exists(comp_data_path):
        entry["component_data"] = os.path.abspath(comp_data_path)
    registry["projects"][proj_id] = entry
    registry["active"] = proj_id
    try:
        with open(config_path, "w") as f:
            json.dump(registry, f, indent=2)
        print(f"  Registered project {proj_id} in .brdnav_config.json")
    except Exception as e:
        print(f"  Warning: Could not register project: {e}")


def extract_embedded_pdf(input_path):
    """Extract the actual PDF from the IRM-protected wrapper.
    
    Handles two cases:
    1. Embedded PDF is clean (no encryption) — use directly
    2. Embedded PDF has IRM /CFM /None encryption — strip the /Encrypt reference
    
    Raises RuntimeError if the PDF is truly IRM-encrypted and text cannot be extracted.
    """
    doc = fitz.open(input_path)
    if doc.embfile_count() == 0:
        print("No embedded files found - using input PDF directly")
        doc.close()
        return input_path

    embedded_data = doc.embfile_get(0)
    doc.close()

    temp_path = os.path.join(BASE_DIR, "_temp_embedded.pdf")

    # Detect IRM before stripping /Encrypt
    is_irm = b'/Encrypt' in embedded_data or b'/MicrosoftIRMServices' in embedded_data

    # Check if embedded PDF has IRM encryption with /CFM /None (not actually encrypted)
    if b'/Encrypt' in embedded_data and b'/CFM /None' in embedded_data:
        print("  Embedded PDF has IRM wrapper (/CFM /None), stripping...")
        embedded_data = re.sub(
            rb'/Encrypt \d+ 0 R',
            lambda m: b' ' * len(m.group(0)),
            embedded_data)

    with open(temp_path, 'wb') as f:
        f.write(embedded_data)

    # Only run clean-save and IRM check for encrypted PDFs
    if is_irm:
        # Clean-save through MuPDF to recompress any corrupted zlib streams
        try:
            clean_doc = fitz.open(temp_path)
            clean_bytes = clean_doc.tobytes(garbage=4, deflate=True, clean=True)
            clean_doc.close()
            with open(temp_path, 'wb') as f:
                f.write(clean_bytes)
            print("  Clean-saved PDF (recompressed streams)")
        except Exception as e:
            print(f"  Warning: clean-save failed ({e}), using raw extracted PDF")

    test_doc = fitz.open(temp_path)
    page_count = test_doc.page_count
    print(f"Extracted embedded PDF: {page_count} pages")

    if is_irm:
        # Verify text is actually extractable (detect IRM-encrypted content)
        sample_words = 0
        for i in range(min(5, page_count)):
            sample_words += len(test_doc[i].get_text("words"))
        test_doc.close()

        if sample_words == 0 and page_count > 0:
            raise RuntimeError(
                "IRM_ENCRYPTED: This PDF is protected by Microsoft IRM (Information Rights Management).\n"
                "The content streams are encrypted and cannot be read by this tool.\n"
                "\n"
                "To remove protection in Adobe Acrobat:\n"
                "  1. Open the PDF in Adobe Acrobat\n"
                "  2. Menu > Protection > Security Properties\n"
                "  3. Change Security Method to 'No Security'\n"
                "  4. Save the PDF and use it as input"
            )
    else:
        test_doc.close()

    return temp_path

    return temp_path


def rebuild_pdf(source_path):
    """Rebuild PDF by importing all pages into a new clean document.
    
    The original Allegro-generated PDF has structural issues that prevent
    link annotations from being clickable in Adobe Acrobat.
    """
    src = fitz.open(source_path)
    dst = fitz.open()
    dst.insert_pdf(src)
    src.close()
    print(f"Rebuilt {dst.page_count} pages into clean PDF")
    return dst


def find_refdes_on_page(page):
    """Find all refdes words on a page with their bounding boxes."""
    results = []
    words = page.get_text("words")

    for word_info in words:
        x0, y0, x1, y1, word = word_info[0], word_info[1], word_info[2], word_info[3], word_info[4]
        clean_word = word.strip().rstrip('.,;:)')

        if REFDES_PATTERN.match(clean_word):
            rect = fitz.Rect(x0, y0, x1, y1)
            results.append((clean_word, rect))

    return results


def find_part_numbers_on_page(page):
    """Find all Microsoft part numbers (Mxxxxxxx-xxx, Xxxxxxx-xxx) on a page."""
    results = []
    words = page.get_text("words")

    for word_info in words:
        x0, y0, x1, y1, word = word_info[0], word_info[1], word_info[2], word_info[3], word_info[4]
        clean_word = word.strip().rstrip('.,;:)')

        if PART_NUMBER_PATTERN.match(clean_word):
            rect = fitz.Rect(x0, y0, x1, y1)
            results.append((clean_word, rect))

    return results


def find_nets_on_page(page):
    """Find all net name words on a page that match the BRD whitelist."""
    results = []
    if not _net_whitelist:
        return results
    words = page.get_text("words")

    for word_info in words:
        x0, y0, x1, y1, word = word_info[0], word_info[1], word_info[2], word_info[3], word_info[4]
        clean_word = word.strip().rstrip('.,;:)')

        # Skip words already matched as refdes or part number
        if REFDES_PATTERN.match(clean_word):
            continue
        if PART_NUMBER_PATTERN.match(clean_word):
            continue

        if clean_word in _net_whitelist:
            rect = fitz.Rect(x0, y0, x1, y1)
            results.append((clean_word, rect))

    return results


def _build_uri(protocol, host, path, **params):
    """Build a link URI using either brdnav:// or http://localhost."""
    from urllib.parse import quote
    encoded_path = quote(path, safe='')
    if _project_id:
        params["proj"] = _project_id
    if _use_http:
        qs = "&".join(f"{k}={v}" for k, v in params.items() if v)
        base = f"http://127.0.0.1:{HTTP_PORT}/{host}/{encoded_path}"
        return f"{base}?{qs}" if qs else base
    else:
        qs = "&".join(f"{k}={v}" for k, v in params.items() if v)
        base = f"{protocol}://{host}/{encoded_path}"
        return f"{base}?{qs}" if qs else base


def add_links_to_pdf(doc):
    """Add clickable URI links for each refdes, part number, and net name in the PDF."""
    total_refdes_links = 0
    total_pn_links = 0
    total_net_links = 0
    unique_refdes = set()
    unique_pn = set()
    unique_nets = set()

    for page_num in range(doc.page_count):
        page = doc[page_num]
        page_refdes = 0
        page_pn = 0
        page_net = 0
        print(f"  Processing page {page_num + 1}/{doc.page_count}…", flush=True)

        # Refdes links
        refdes_items = find_refdes_on_page(page)
        for refdes, rect in refdes_items:
            unique_refdes.add(refdes)
            link_rect = rect + (-1, -1, 1, 1)
            comp = _comp_data.get(refdes)
            layer = comp["layer"] if comp else ""
            uri = _build_uri(PROTOCOL, PROTOCOL_HOST, refdes, layer=layer)
            link = {
                'kind': fitz.LINK_URI,
                'from': link_rect,
                'uri': uri,
            }
            page.insert_link(link)
            page_refdes += 1

        # Part number links
        pn_items = find_part_numbers_on_page(page)
        for pn, rect in pn_items:
            unique_pn.add(pn)
            link_rect = rect + (-1, -1, 1, 1)
            uri = _build_uri(DATASHEET_PROTOCOL, DATASHEET_HOST, pn)
            link = {
                'kind': fitz.LINK_URI,
                'from': link_rect,
                'uri': uri,
            }
            page.insert_link(link)
            page_pn += 1

        # Net name links
        net_items = find_nets_on_page(page)
        for net_name, rect in net_items:
            unique_nets.add(net_name)
            link_rect = rect + (-1, -1, 1, 1)
            uri = _build_uri(NET_PROTOCOL, NET_HOST, net_name)
            link = {
                'kind': fitz.LINK_URI,
                'from': link_rect,
                'uri': uri,
            }
            page.insert_link(link)
            page_net += 1

        total_refdes_links += page_refdes
        total_pn_links += page_pn
        total_net_links += page_net
        if page_refdes or page_pn or page_net:
            print(f"  Page {page_num + 1}: {page_refdes} refdes, {page_pn} part#, {page_net} nets")

    return (total_refdes_links, len(unique_refdes),
            total_pn_links, len(unique_pn),
            total_net_links, len(unique_nets))


def generate_viewer_data(source_pdf, output_dir):
    """Generate page images + link position JSON for the HTML viewer."""
    pages_dir = os.path.join(output_dir, "pages")
    os.makedirs(pages_dir, exist_ok=True)

    doc = fitz.open(source_pdf)
    zoom = 2  # 144 DPI — crisp on high-DPI screens
    mat = fitz.Matrix(zoom, zoom)

    viewer_data = {"page_count": doc.page_count, "zoom": zoom, "pages": {}}

    for page_num in range(doc.page_count):
        page = doc[page_num]

        # Render page image
        pix = page.get_pixmap(matrix=mat)
        img_path = os.path.join(pages_dir, f"{page_num + 1}.png")
        pix.save(img_path)

        # Collect link positions
        links = []
        for refdes, rect in find_refdes_on_page(page):
            comp = _comp_data.get(refdes)
            layer = comp["layer"] if comp else ""
            uri = f"/nav/{refdes}?layer={layer}" if layer else f"/nav/{refdes}"
            links.append({
                "text": refdes, "type": "refdes",
                "x": round(rect.x0 * zoom), "y": round(rect.y0 * zoom),
                "w": round(rect.width * zoom), "h": round(rect.height * zoom),
                "uri": uri,
            })

        for pn, rect in find_part_numbers_on_page(page):
            links.append({
                "text": pn, "type": "part",
                "x": round(rect.x0 * zoom), "y": round(rect.y0 * zoom),
                "w": round(rect.width * zoom), "h": round(rect.height * zoom),
                "uri": f"/datasheet/{pn}",
            })

        for net_name, rect in find_nets_on_page(page):
            links.append({
                "text": net_name, "type": "net",
                "x": round(rect.x0 * zoom), "y": round(rect.y0 * zoom),
                "w": round(rect.width * zoom), "h": round(rect.height * zoom),
                "uri": f"/net/{net_name}",
            })

        viewer_data["pages"][str(page_num + 1)] = {
            "width": pix.width, "height": pix.height, "links": links,
        }
        print(f"  Page {page_num + 1}: {pix.width}x{pix.height}px, {len(links)} links")

    doc.close()

    links_path = os.path.join(output_dir, "links.json")
    with open(links_path, "w") as f:
        json.dump(viewer_data, f)
    print(f"  Saved viewer data to {output_dir}")


def register_brdnav_protocol():
    if winreg is None:
        print("  Skipped protocol registration (not on Windows)")
        return
    handler = os.path.join(BASE_DIR, "brdnav_handler.pyw")
    pythonw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
    if not os.path.exists(pythonw):
        pythonw = sys.executable
    command = f'"{pythonw}" "{handler}" "%1"'

    try:
        key_path = rf"Software\Classes\{PROTOCOL}"
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
            winreg.SetValue(key, "", winreg.REG_SZ,
                            "URL:Smart Schematic Protocol")
            winreg.SetValueEx(key, "URL Protocol", 0, winreg.REG_SZ, "")
        with winreg.CreateKey(
                winreg.HKEY_CURRENT_USER,
                rf"{key_path}\shell\open\command") as key:
            winreg.SetValue(key, "", winreg.REG_SZ, command)
        print(f"  Registered {PROTOCOL}:// protocol handler")
    except Exception as e:
        print(f"  Warning: Could not register protocol: {e}")


def main():
    parser = argparse.ArgumentParser(
        description="Smart Schematic — generate linked PDF with clickable navigation links")
    parser.add_argument("--input", default=None,
                        help="Input PDF file path (default: input.pdf)")
    parser.add_argument("--output", default=None,
                        help="Acrobat PDF output path (default: <input>_linked_Acrobat.pdf)")
    parser.add_argument("--brd", default=None,
                        help="BRD file path — used to set project ID in links")
    parser.add_argument("--component-data", default=None,
                        help="Component data JSON path (default: component_data.json)")
    parser.add_argument("--net-data", default=None,
                        help="Net data JSON path (default: net_data.json next to BRD or in BASE_DIR)")
    args = parser.parse_args()

    global _comp_data
    global _net_whitelist
    global _use_http
    global _project_id
    input_pdf = args.input or INPUT_PDF
    comp_data_path = args.component_data or COMPONENT_DATA

    # Derive project ID from BRD filename (used in links for multi-project support)
    if args.brd:
        import hashlib
        brd_abs = os.path.abspath(args.brd)
        _project_id = hashlib.md5(brd_abs.encode()).hexdigest()[:8]
        print(f"Project ID: {_project_id} (from {os.path.basename(args.brd)})")
        _register_project(_project_id, brd_abs, comp_data_path)

    # Build output paths with new naming convention
    if args.output:
        output_pdf = args.output
    else:
        pdf_dir = os.path.dirname(os.path.abspath(input_pdf))
        base = os.path.splitext(os.path.basename(input_pdf))[0]
        output_pdf = os.path.join(pdf_dir, f"{base}_linked_Acrobat.pdf")

    if os.path.exists(comp_data_path):
        with open(comp_data_path) as f:
            _comp_data = json.load(f)
        print(f"Loaded {len(_comp_data)} components from {comp_data_path}")

    # Resolve net data path: CLI arg > next to component_data > next to BRD > BASE_DIR
    net_data_path = args.net_data
    if not net_data_path:
        # Try next to component_data.json
        if comp_data_path and os.path.exists(comp_data_path):
            candidate = os.path.join(os.path.dirname(os.path.abspath(comp_data_path)), "net_data.json")
            if os.path.exists(candidate):
                net_data_path = candidate
        # Try next to BRD file
        if not net_data_path and args.brd and os.path.exists(args.brd):
            candidate = os.path.join(os.path.dirname(os.path.abspath(args.brd)), "net_data.json")
            if os.path.exists(candidate):
                net_data_path = candidate
        # Fall back to BASE_DIR
        if not net_data_path:
            net_data_path = NET_DATA

    if net_data_path and os.path.exists(net_data_path):
        with open(net_data_path) as f:
            _net_whitelist = set(json.load(f))
        print(f"Loaded {len(_net_whitelist)} net names from {net_data_path}")

    print("=" * 60)
    print("Smart Schematic — PDF Link Generator")
    print("=" * 60)

    print("\n[Step 1] Registering brdnav:// protocol...")
    register_brdnav_protocol()

    print("\n[Step 2] Extracting embedded PDF...")
    source_pdf = extract_embedded_pdf(input_pdf)

    print("\n[Step 3] Rebuilding PDF (clean structure)...")
    doc = rebuild_pdf(source_pdf)

    print("\n[Step 4] Adding refdes + part number + net name links...")
    result = add_links_to_pdf(doc)
    total_refdes, unique_refdes, total_pn, unique_pn, total_net, unique_net = result

    print("\n[Step 5a] Saving Acrobat PDF (brdnav:// links)...")
    try:
        doc.save(output_pdf)
        print(f"  Saved: {output_pdf}")
    except Exception:
        alt = output_pdf.replace(".pdf", "_new.pdf")
        doc.save(alt)
        output_pdf = alt
        print(f"  Output locked, saved as: {alt}")
    doc.close()

    # Cleanup temp file
    temp_file = os.path.join(BASE_DIR, "_temp_embedded.pdf")
    if os.path.exists(temp_file):
        try:
            os.remove(temp_file)
        except Exception:
            pass  # file may be locked by OneDrive sync

    print("\n" + "=" * 60)
    print("DONE!")
    print(f"  Output PDF:    {output_pdf}")
    print(f"  Refdes links:  {total_refdes} ({unique_refdes} unique)")
    print(f"  Part # links:  {total_pn} ({unique_pn} unique)")
    print(f"  Net links:     {total_net} ({unique_net} unique)")
    print("=" * 60)


if __name__ == "__main__":
    main()
