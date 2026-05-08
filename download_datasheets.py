"""
Download Datasheets — Extract part numbers from schematic PDF, associate with
nearest refdes, then download datasheets from OnePDM in the background.

Creates a datasheets/ folder under the project folder with subfolders:
    U1000_M1386584-002/
    C1001_M1335148-001/
    ...

Usage:
    python download_datasheets.py --pdf <schematic.pdf> [--output <folder>]
"""

import os
import re
import sys
import json
import threading
import time
import argparse
from pathlib import Path

import fitz  # PyMuPDF

# Part number pattern: M1234567-001 or X1234567-01
PART_NUMBER_PATTERN = re.compile(r'^[MX]\d{5,8}-\d{2,4}$')
# Refdes pattern: letter(s) + digits (e.g., U1000, C1001, R1234)
REFDES_PATTERN = re.compile(r'^[A-Z]{1,3}\d{2,6}$')


def extract_refdes_part_pairs(pdf_path, comp_data_path=None):
    """Extract refdes-to-part-number associations from PDF.
    
    On each page, finds refdes and part numbers, then associates them
    by vertical proximity (closest refdes to each part number).
    
    If comp_data_path is provided, only considers text that exists as a
    known refdes in component_data.json (filters out footprint names like BGA167).
    
    Returns dict: {refdes: part_number}
    """
    # Load valid refdes set from component_data.json
    valid_refdes = None
    if comp_data_path and os.path.exists(comp_data_path):
        with open(comp_data_path) as f:
            valid_refdes = set(json.load(f).keys())

    doc = fitz.open(pdf_path)
    associations = {}  # refdes -> part_number

    # IC prefixes — these are primary components that "own" the part number
    IC_PREFIXES = ('U', 'IC', 'Q')

    # First pass: count page appearances for each part number
    # Part numbers on many pages are board-level (title block), not component-level
    from collections import Counter
    pn_page_count = Counter()
    for page_num in range(doc.page_count):
        page = doc[page_num]
        seen = set()
        for w in page.get_text("words"):
            clean = w[4].strip().rstrip('.,;:)')
            if PART_NUMBER_PATTERN.match(clean):
                seen.add(clean)
        for pn in seen:
            pn_page_count[pn] += 1

    # Exclude part numbers appearing on >20% of pages (board/title block PNs)
    max_pages = max(3, doc.page_count * 0.2)
    excluded_pns = {pn for pn, cnt in pn_page_count.items() if cnt > max_pages}

    # Second pass: build associations

    for page_num in range(doc.page_count):
        page = doc[page_num]
        words = page.get_text("words")

        refdes_items = []  # (text, x_center, y_center)
        pn_items = []      # (text, x_center, y_center)

        for w in words:
            x0, y0, x1, y1, text = w[0], w[1], w[2], w[3], w[4]
            clean = text.strip().rstrip('.,;:)')
            cx = (x0 + x1) / 2
            cy = (y0 + y1) / 2

            if REFDES_PATTERN.match(clean):
                # Only accept if it's a known component refdes
                if valid_refdes is None or clean in valid_refdes:
                    refdes_items.append((clean, cx, cy))
            elif PART_NUMBER_PATTERN.match(clean):
                if clean not in excluded_pns:
                    pn_items.append((clean, cx, cy))

        if not pn_items:
            continue

        # Separate IC-type refdes from passives
        ic_refs = [(r, x, y) for r, x, y in refdes_items
                   if any(r.startswith(p) for p in IC_PREFIXES)]

        # Associate each part number with best refdes
        for pn, px, py in pn_items:
            best_refdes = None

            # Strategy 1: If only one IC on this page, check X-alignment
            # In schematics, IC refdes and its part number share similar X coordinate
            if len(ic_refs) == 1:
                ic_ref, ic_x, ic_y = ic_refs[0]
                x_diff = abs(px - ic_x)
                if x_diff < 80:  # horizontally aligned (same column)
                    best_refdes = ic_ref
            elif ic_refs:
                # Multiple ICs — prefer X-aligned IC, then nearest
                best_score = float('inf')
                for ref, rx, ry in ic_refs:
                    x_diff = abs(px - rx)
                    dist = ((px - rx) ** 2 + (py - ry) ** 2) ** 0.5
                    # Prefer X-alignment (weight X difference heavily)
                    score = x_diff * 3 + dist
                    if score < best_score and dist < 800:
                        best_score = score
                        best_refdes = ref

            # Strategy 2: Fall back to nearest any refdes (tight threshold)
            if not best_refdes:
                best_dist = 100
                for ref, rx, ry in refdes_items:
                    dist = ((px - rx) ** 2 + (py - ry) ** 2) ** 0.5
                    if dist < best_dist:
                        best_dist = dist
                        best_refdes = ref

            if best_refdes and best_refdes not in associations:
                associations[best_refdes] = pn

    doc.close()
    return associations


def download_worker(tasks, output_dir, progress_file, progress_callback=None, cancel_event=None):
    """Background worker that downloads datasheets from OnePDM.
    
    tasks: list of (refdes, part_number) tuples
    output_dir: base datasheets/ directory
    progress_file: JSON file tracking download status
    progress_callback: optional callable(current, total, message) for progress updates
    cancel_event: optional threading.Event — set to cancel download
    """
    # Load or init progress
    progress = {}
    if os.path.exists(progress_file):
        with open(progress_file) as f:
            progress = json.load(f)

    # Initialize OnePDM client (headless — no browser window)
    onepdm = None
    try:
        from onepdm_client import get_client
        onepdm = get_client(headless=True)
    except Exception as e:
        print(f"[datasheets] OnePDM not available: {e}", flush=True)
        # Mark all as failed
        for refdes, pn in tasks:
            progress[f"{refdes}_{pn}"] = {"status": "error", "error": str(e)}
        with open(progress_file, 'w') as f:
            json.dump(progress, f, indent=2)
        return

    total = len(tasks)
    downloaded = 0
    failed = 0

    for i, (refdes, pn) in enumerate(tasks):
        # Check for cancellation
        if cancel_event and cancel_event.is_set():
            print(f"[datasheets] Download cancelled by user at {i}/{total}", flush=True)
            break

        folder_name = f"{refdes}_{pn}"
        key = folder_name

        # Skip if already done
        if key in progress and progress[key].get("status") == "done":
            downloaded += 1
            continue

        dest_dir = os.path.join(output_dir, folder_name)
        os.makedirs(dest_dir, exist_ok=True)

        print(f"[datasheets] ({i+1}/{total}) {folder_name}...", flush=True)
        if progress_callback:
            progress_callback(i+1, total, folder_name)
        progress[key] = {"status": "downloading", "started": time.time()}

        try:
            result = onepdm.download_datasheet(pn, dest_dir)
            if result:
                progress[key] = {
                    "status": "done",
                    "file": result,
                    "timestamp": time.time()
                }
                downloaded += 1
                print(f"[datasheets]   ✅ {os.path.basename(result)}", flush=True)
            else:
                progress[key] = {"status": "not_found", "timestamp": time.time()}
                failed += 1
                print(f"[datasheets]   ❌ No datasheet found", flush=True)
        except Exception as e:
            progress[key] = {"status": "error", "error": str(e), "timestamp": time.time()}
            failed += 1
            print(f"[datasheets]   ❌ Error: {e}", flush=True)
            # If browser crashed, try to recover
            if "browser" in str(e).lower() or "session" in str(e).lower():
                print("[datasheets] Browser crashed, reinitializing...", flush=True)
                try:
                    onepdm = None
                    from onepdm_client import get_client
                    import onepdm_client
                    onepdm_client._client = None
                    onepdm = get_client()
                except Exception:
                    print("[datasheets] Failed to reinitialize, stopping.", flush=True)
                    break

        # Save progress after each download
        with open(progress_file, 'w') as f:
            json.dump(progress, f, indent=2)

    # Cleanup
    if onepdm:
        try:
            onepdm.close()
        except Exception:
            pass

    print(f"\n[datasheets] Complete: {downloaded} downloaded, {failed} failed, {total} total",
          flush=True)


def main():
    parser = argparse.ArgumentParser(
        description="Download datasheets from OnePDM for all parts in schematic PDF")
    parser.add_argument("--pdf", required=True, help="Schematic PDF path")
    parser.add_argument("--component-data", default=None,
                        help="component_data.json path (filters footprint text from refdes)")
    parser.add_argument("--output", default=None,
                        help="Output folder (default: datasheets/ next to PDF)")
    parser.add_argument("--foreground", action="store_true",
                        help="Run in foreground (default: background thread)")
    parser.add_argument("--list-only", action="store_true",
                        help="Just list refdes/part pairs, don't download")
    args = parser.parse_args()

    pdf_path = os.path.abspath(args.pdf)
    if not os.path.exists(pdf_path):
        print(f"ERROR: PDF not found: {pdf_path}")
        sys.exit(1)

    pdf_dir = os.path.dirname(pdf_path)
    output_dir = args.output or os.path.join(pdf_dir, "datasheets")
    os.makedirs(output_dir, exist_ok=True)

    # Find component_data.json
    comp_data_path = args.component_data
    if not comp_data_path:
        auto_path = os.path.join(pdf_dir, "component_data.json")
        if os.path.exists(auto_path):
            comp_data_path = auto_path

    print(f"Scanning PDF for part numbers: {os.path.basename(pdf_path)}")
    if comp_data_path:
        print(f"  Using component_data.json to filter valid refdes")
    associations = extract_refdes_part_pairs(pdf_path, comp_data_path)

    if not associations:
        print("No refdes + part number pairs found in PDF")
        sys.exit(0)

    # Deduplicate by part number (same part may appear multiple times)
    # Keep first refdes for each unique part number
    seen_pn = {}
    tasks = []
    for refdes, pn in sorted(associations.items()):
        if pn not in seen_pn:
            seen_pn[pn] = refdes
            tasks.append((refdes, pn))

    print(f"Found {len(associations)} associations, {len(tasks)} unique parts")
    print(f"Output: {output_dir}")
    print()

    if args.list_only:
        for refdes, pn in tasks:
            print(f"  {refdes}_{pn}")
        return

    # Save association map
    map_file = os.path.join(output_dir, "_part_map.json")
    with open(map_file, 'w') as f:
        json.dump(associations, f, indent=2)

    progress_file = os.path.join(output_dir, "_progress.json")

    if args.foreground:
        download_worker(tasks, output_dir, progress_file)
    else:
        print("Starting background download...")
        print(f"  Progress: {progress_file}")
        print(f"  Folders:  {output_dir}/<REFDES>_<PART>/")
        print()
        t = threading.Thread(
            target=download_worker,
            args=(tasks, output_dir, progress_file),
            daemon=True
        )
        t.start()
        # Keep alive until done
        try:
            while t.is_alive():
                time.sleep(1)
        except KeyboardInterrupt:
            print("\n[datasheets] Interrupted — progress saved.")


if __name__ == "__main__":
    main()
