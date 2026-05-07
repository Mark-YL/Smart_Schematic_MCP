"""Extract component placement data directly from Allegro BRD binary.

Supports two extraction methods:
1. SKILL-format placement records (older Allegro versions):
   list("REFDES" "FOOTPRINT" list(X Y) LAYER "NONE")
   Layer is INVERTED: 't' = Bottom, 'nil' = Top.
   Coordinates are in mm (board-internal coordinate system).

2. Embedded ZIP XML (newer Allegro 17.x+):
   Extracts refdes names from objects/Part.xml inside embedded ZIP.
   Coordinates/layer not available in XML — set to defaults.

No Allegro Viewer needed — pure Python binary parsing.
"""
import re
import json
import os
import sys
import struct
import mmap
import zipfile
import io
import argparse


def _extract_skill_components(data):
    """Extract components from SKILL-format placement records."""
    # SKILL placement pattern:
    # list("REFDES" "FOOTPRINT" list(X Y) LAYER "NONE")
    # Layer is INVERTED: t = Bottom, nil = Top (confirmed by validation)
    pattern = re.compile(
        rb'list\("([A-Z][A-Z0-9_]+)"\s+'     # refdes
        rb'"([^"]+)"\s+'                       # footprint
        rb'list\(([\d.]+)\s+([\d.]+)\)\s+'     # x, y (already in mm)
        rb'(t|nil)\s+'                         # layer: t=Bottom, nil=Top
        rb'"([^"]*)"'                          # extra field (usually "NONE")
    )

    components = {}
    for m in pattern.finditer(data):
        refdes = m.group(1).decode('ascii', errors='ignore')
        x_mm = float(m.group(3).decode('ascii'))
        y_mm = float(m.group(4).decode('ascii'))
        layer_flag = m.group(5).decode('ascii')
        # INVERTED: 't' in BRD = Bottom, 'nil' = Top
        layer = "Bottom" if layer_flag == "t" else "Top"

        components[refdes] = {
            "layer": layer,
            "x": f"{x_mm:.4f}",
            "y": f"{y_mm:.4f}",
        }

    return components


def _extract_binary_layers(filepath):
    """Extract component layers from binary record flags.
    
    In Allegro BRD binary, component records have the pattern:
      ... [flag_byte] ... 00 00 06 00 00 00 00 00 [REFDES\0] ...
    
    The flag byte at offset -17 from the marker encodes the side:
      - Odd values (LSB=1): Top  (e.g. 0xFB, 0xFD, 0x01)
      - Even values (LSB=0): Bottom (e.g. 0xFA, 0xFC, 0x00)
    
    Validated 15/15 correct against SKILL-extracted layer data.
    """
    layers = {}
    fsize = os.path.getsize(filepath)
    
    with open(filepath, 'rb') as f:
        mm = mmap.mmap(f.fileno(), 0, access=mmap.ACCESS_READ)
        
        # Only search the binary section (before embedded ZIP)
        zip_pos = mm.find(b'PK\x03\x04')
        search_end = zip_pos if zip_pos > 0 else len(mm)
        
        marker = b'\x00\x00\x06\x00\x00\x00\x00\x00'
        refdes_pattern = re.compile(rb'^[A-Z][A-Z0-9_]{1,20}$')
        
        pos = 0
        while True:
            pos = mm.find(marker, pos, search_end)
            if pos < 0:
                break
            
            # Read refdes after marker
            refdes_start = pos + 8
            refdes_end = mm.find(b'\x00', refdes_start, refdes_start + 30)
            if refdes_end <= refdes_start:
                pos += 1
                continue
            
            refdes_bytes = mm[refdes_start:refdes_end]
            if not refdes_pattern.match(refdes_bytes):
                pos += 1
                continue
            
            refdes = refdes_bytes.decode('ascii')
            
            # Read flag byte at offset -17 from marker
            flag_off = pos - 17
            if flag_off < 0:
                pos += 1
                continue
            
            flag_byte = mm[flag_off]
            layer = "Top" if (flag_byte & 1) else "Bottom"
            # Keep first occurrence only — later duplicates may have stale flags
            if refdes not in layers:
                layers[refdes] = layer
            
            pos = refdes_end + 1
        
        mm.close()
    
    return layers


def _extract_xml_components(filepath):
    """Extract component refdes from embedded ZIP XML (Part.xml)."""
    fsize = os.path.getsize(filepath)
    components = {}

    with open(filepath, 'rb') as f:
        mm = mmap.mmap(f.fileno(), 0, access=mmap.ACCESS_READ)

        # Find all EOCD (PK\x05\x06) markers — BRDs may have multiple ZIPs
        pos = 0
        eocds = []
        while True:
            pos = mm.find(b'PK\x05\x06', pos)
            if pos < 0:
                break
            eocds.append(pos)
            pos += 1

        # Find the first PK\x03\x04 (local file header) in the later portion
        # The main XML ZIP is typically in the last ~10-20% of the file
        search_start = max(0, int(fsize * 0.7))
        first_pk = mm.find(b'PK\x03\x04', search_start)
        if first_pk < 0:
            first_pk = mm.find(b'PK\x03\x04')

        if first_pk < 0:
            mm.close()
            return components

        # Try each EOCD to find one that contains Part.xml
        for eocd_pos in eocds:
            if eocd_pos < first_pk:
                continue
            try:
                zip_data = mm[first_pk:eocd_pos + 22]
                zf = zipfile.ZipFile(io.BytesIO(bytes(zip_data)))
                if 'objects/Part.xml' in zf.namelist():
                    part_xml = zf.read('objects/Part.xml').decode('utf-8', errors='replace')
                    # Parse <o N="REFDES"> patterns
                    for m in re.finditer(r'<o\s+N="([A-Z][A-Z0-9_]+)"', part_xml):
                        refdes = m.group(1)
                        if refdes not in components:
                            components[refdes] = {
                                "layer": "",  # unknown from XML
                                "x": "0",
                                "y": "0",
                            }
                    break
            except (zipfile.BadZipFile, Exception):
                continue

        mm.close()

    return components


def extract_components_from_brd(brd_path):
    """Extract component refdes, layer, x, y from BRD binary."""
    print(f"Reading {os.path.basename(brd_path)}...")
    file_size = os.path.getsize(brd_path)
    print(f"  File size: {file_size/1024/1024:.1f} MB")

    with open(brd_path, 'rb') as f:
        data = f.read()

    # Method 1: SKILL-format placement records
    components = _extract_skill_components(data)
    if components:
        print(f"  SKILL extraction: {len(components)} components")
    else:
        print("  No SKILL placement records found")

    # Method 2: XML supplement/fallback — fill in missing components
    print("  Checking embedded ZIP XML for additional components...")
    xml_components = _extract_xml_components(brd_path)
    if xml_components:
        added = 0
        for refdes, cdata in xml_components.items():
            if refdes not in components:
                components[refdes] = cdata
                added += 1
        if added:
            print(f"  XML added {added} extra refdes (no layer/coords)")
        else:
            print(f"  XML: no new refdes beyond SKILL extraction")

    # Method 3: Binary flag extraction — fill in missing layers AND add new components
    print(f"  Extracting layers from binary record flags...")
    binary_layers = _extract_binary_layers(brd_path)
    if binary_layers:
        filled = 0
        for refdes, layer in binary_layers.items():
            if refdes in components and not components[refdes]['layer']:
                components[refdes]['layer'] = layer
                filled += 1
            elif refdes not in components:
                components[refdes] = {"layer": layer, "x": "0", "y": "0"}
                filled += 1
        if filled:
            print(f"  Binary flags: added/filled {filled} components")
        else:
            print(f"  Binary flags: no additional layers found")
    else:
        print(f"  Binary flags: no records found")

    # Layer stats
    top = sum(1 for v in components.values() if v['layer'] == 'Top')
    bot = sum(1 for v in components.values() if v['layer'] == 'Bottom')
    unk = sum(1 for v in components.values() if v['layer'] == '')
    print(f"  Top: {top}, Bottom: {bot}" + (f", Unknown: {unk}" if unk else ""))

    return components


def extract_nets_from_brd(brd_path):
    """Extract net names from BRD binary (reuse existing logic)."""
    print(f"Extracting net names from {os.path.basename(brd_path)}...")
    with open(brd_path, 'rb') as f:
        data = f.read()

    xml_nets = set()
    for m in re.finditer(rb'<o\s+N="([^"]+)"', data):
        name = m.group(1).decode("ascii", errors="ignore").strip()
        if name:
            xml_nets.add(name)

    m_nets = set()
    for m_match in re.finditer(rb'<m>([^<]+)</m>', data):
        name = m_match.group(1).decode("ascii", errors="ignore").strip()
        if name:
            m_nets.add(name)

    raw = xml_nets | m_nets

    net_pattern = re.compile(r'^[A-Za-z0-9_+\-/\.\[\]()#@$%^&*~|<>]+$')
    noise_patterns = [
        re.compile(r'[/\\]'),
        re.compile(r'\.(?:dra|pad|psm|brd|dat|txt|log|xml|html|jpg|png)', re.I),
        re.compile(r'^\d+$'),
        re.compile(r'^[a-z]{1}$', re.I),
    ]

    nets = set()
    for name in raw:
        if len(name) < 2 or len(name) > 80:
            continue
        if not net_pattern.match(name):
            continue
        if any(p.search(name) for p in noise_patterns):
            continue
        name = re.sub(r'~(\d+)~', r'<\1>', name)
        nets.add(name)

    result = sorted(nets)
    print(f"  Extracted {len(result)} net names")
    return result


def main():
    parser = argparse.ArgumentParser(
        description="Extract component data from Allegro BRD file (no Allegro needed)")
    parser.add_argument("--brd", required=True, help="BRD file path")
    parser.add_argument("--output", default=None,
                        help="Output JSON path (default: component_data.json next to BRD)")
    parser.add_argument("--nets", action="store_true",
                        help="Also extract net names")
    parser.add_argument("--validate", default=None,
                        help="Validate against existing component_data.json")
    args = parser.parse_args()

    brd_path = os.path.abspath(args.brd)
    brd_dir = os.path.dirname(brd_path)

    # Extract components
    components = extract_components_from_brd(brd_path)

    if not components:
        print("ERROR: No components found in BRD file")
        sys.exit(1)

    # Save
    output = args.output or os.path.join(brd_dir, "component_data.json")
    with open(output, "w") as f:
        json.dump(components, f, indent=2)
    print(f"Saved {len(components)} components to {output}")

    # Extract nets
    if args.nets:
        nets = extract_nets_from_brd(brd_path)
        net_output = os.path.join(brd_dir, "net_data.json")
        with open(net_output, "w") as f:
            json.dump(nets, f)
        print(f"Saved {len(nets)} net names to {net_output}")

    # Validate
    if args.validate:
        print(f"\nValidating against {args.validate}...")
        with open(args.validate) as f:
            truth = json.load(f)

        matched = 0
        mismatched = 0
        missing = 0
        extra = 0

        for refdes, tdata in truth.items():
            if refdes not in components:
                missing += 1
                if missing <= 5:
                    print(f"  MISSING: {refdes}")
                continue
            cdata = components[refdes]
            if cdata['layer'] != tdata['layer']:
                mismatched += 1
                if mismatched <= 5:
                    print(f"  LAYER MISMATCH: {refdes} expected={tdata['layer']} got={cdata['layer']}")
            else:
                # Check coordinates (allow small tolerance for rounding)
                try:
                    tx = float(tdata['x'] or 0)
                    ty = float(tdata['y'] or 0)
                    cx = float(cdata['x'] or 0)
                    cy = float(cdata['y'] or 0)
                except (ValueError, TypeError):
                    matched += 1
                    continue
                if abs(tx - cx) > 0.1 or abs(ty - cy) > 0.1:
                    mismatched += 1
                    if mismatched <= 5:
                        print(f"  COORD MISMATCH: {refdes} expected=({tx},{ty}) got=({cx},{cy})")
                else:
                    matched += 1

        for refdes in components:
            if refdes not in truth:
                extra += 1

        total = len(truth)
        print(f"\nValidation results:")
        print(f"  Ground truth: {total} components")
        print(f"  Extracted:    {len(components)} components")
        print(f"  Matched:      {matched} ({matched*100//total}%)")
        print(f"  Mismatched:   {mismatched}")
        print(f"  Missing:      {missing}")
        print(f"  Extra:        {extra}")


if __name__ == "__main__":
    main()
