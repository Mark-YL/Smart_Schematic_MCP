"""
Smart Schematic — Auto Linked PDF Watcher

Monitors the inbox/ folder for new PDF + BRD file pairs and automatically
generates linked PDFs. Uses pure-Python BRD parsing (no Allegro needed)
via extract_brd_components.py, then generate_linked_pdf.py for PDF linking.

Usage:
    python auto_linked_pdf.py [--inbox <folder>]

Drop files into inbox/ (or subfolders like inbox/ProjectA/):
    MyProject.pdf + MyProject.brd  →  MyProject_linked_Acrobat.pdf
"""

import os
import sys
import time
import json
import queue
import logging
import argparse
import subprocess
import threading
import textwrap
from pathlib import Path
from watchdog.observers import Observer

try:
    import pystray
    from PIL import Image, ImageDraw
    HAS_TRAY = True
except ImportError:
    HAS_TRAY = False
from watchdog.events import FileSystemEventHandler

# Global tray icon reference for notifications
_tray_icon = None

# Activity log — recent events for the manager UI
_activity_log = []  # list of (timestamp, message) tuples
_activity_lock = threading.Lock()
_ACTIVITY_FILE = os.path.join(os.environ.get('TEMP', '.'), 'ss_activity.json')
_DS_CANCEL_FILE = os.path.join(os.environ.get('TEMP', '.'), 'ss_ds_cancel')
_DS_RESUME_FILE = os.path.join(os.environ.get('TEMP', '.'), 'ss_ds_resume')
_SCAN_TRIGGER_FILE = os.path.join(os.environ.get('TEMP', '.'), 'ss_scan_trigger')
_ds_cancel_event = threading.Event()
_SETTINGS_FILE = os.path.join(os.environ.get('TEMP', '.'), 'ss_settings.json')
_last_ds_context = {}  # stores last download context for resume

def _load_settings():
    """Load user settings (returns dict with defaults)."""
    defaults = {'download_datasheets': True}
    try:
        if os.path.exists(_SETTINGS_FILE):
            with open(_SETTINGS_FILE, encoding='utf-8') as f:
                saved = json.load(f)
            defaults.update(saved)
    except Exception:
        pass
    return defaults

def _save_settings(settings):
    """Persist user settings."""
    try:
        with open(_SETTINGS_FILE, 'w', encoding='utf-8') as f:
            json.dump(settings, f)
    except Exception:
        pass

def _log_activity(msg):
    """Record an activity event (kept in memory + written to temp file for UI)."""
    import json
    entry = (time.strftime("%H:%M:%S"), msg)
    with _activity_lock:
        _activity_log.append(entry)
        if len(_activity_log) > 50:
            _activity_log[:] = _activity_log[-50:]
        try:
            with open(_ACTIVITY_FILE, 'w', encoding='utf-8') as f:
                json.dump(_activity_log, f)
        except OSError:
            pass

def notify(title, message):
    """Show a Windows notification balloon via the tray icon."""
    _log_activity(f"[{title}] {message}")
    if _tray_icon:
        try:
            _tray_icon.notify(message, title)
        except Exception:
            pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PYTHON = sys.executable

# Files to ignore
IGNORE_PATTERNS = ("_linked_", "_temp_", "~$", ".tmp", ".crdownload")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S"
)
log = logging.getLogger("auto_linked_pdf")


def is_ignored(filename):
    return any(p in filename for p in IGNORE_PATTERNS)


def wait_for_stable_file(path, checks=3, interval=2):
    """Wait until file size stops changing (fully copied)."""
    prev_size = -1
    stable_count = 0
    for _ in range(checks * 5):
        try:
            size = os.path.getsize(path)
        except OSError:
            return False
        if size == prev_size and size > 0:
            stable_count += 1
            if stable_count >= checks:
                return True
        else:
            stable_count = 0
        prev_size = size
        time.sleep(interval)
    return False


def find_matching_pair(folder):
    """Find PDF + BRD pairs in a folder.
    
    Matching rules:
    1. Exact stem match (MyProject.pdf + MyProject.brd)
    2. One PDF + one BRD in the same folder → auto-pair
    3. Multiple of either → ambiguous, skip
    
    Returns list of (pdf, brd) tuples.
    """
    pdfs = []
    brds = []
    for f in os.listdir(folder):
        lower = f.lower()
        if is_ignored(f):
            continue
        full = os.path.join(folder, f)
        if not os.path.isfile(full):
            continue
        if lower.endswith(".pdf"):
            pdfs.append(f)
        elif lower.endswith(".brd"):
            brds.append(f)

    pairs = []
    matched_pdfs = set()
    matched_brds = set()

    # Pass 1: exact stem match
    for pdf in pdfs:
        pdf_stem = Path(pdf).stem.lower()
        for brd in brds:
            if Path(brd).stem.lower() == pdf_stem:
                output_name = f"{Path(pdf).stem}_linked_Acrobat.pdf"
                output_path = os.path.join(folder, output_name)
                if os.path.exists(output_path):
                    src_time = max(os.path.getmtime(os.path.join(folder, pdf)),
                                   os.path.getmtime(os.path.join(folder, brd)))
                    if os.path.getmtime(output_path) > src_time:
                        continue
                pairs.append((pdf, brd))
                matched_pdfs.add(pdf)
                matched_brds.add(brd)

    # Pass 2: unmatched files — auto-pair if exactly 1 PDF + 1 BRD remain
    remaining_pdfs = [p for p in pdfs if p not in matched_pdfs]
    remaining_brds = [b for b in brds if b not in matched_brds]

    if len(remaining_pdfs) == 1 and len(remaining_brds) == 1:
        pdf = remaining_pdfs[0]
        brd = remaining_brds[0]
        output_name = f"{Path(pdf).stem}_linked_Acrobat.pdf"
        output_path = os.path.join(folder, output_name)
        if os.path.exists(output_path):
            src_time = max(os.path.getmtime(os.path.join(folder, pdf)),
                           os.path.getmtime(os.path.join(folder, brd)))
            if os.path.getmtime(output_path) > src_time:
                return pairs
        pairs.append((pdf, brd))
        log.info(f"  Auto-paired: {pdf} ↔ {brd} (only pair in folder)")
    elif len(remaining_pdfs) > 0 and len(remaining_brds) > 1:
        log.warning(f"Ambiguous: {len(remaining_pdfs)} PDF(s) + {len(remaining_brds)} BRDs in {folder} — use subfolders")
    elif len(remaining_pdfs) > 1 and len(remaining_brds) > 0:
        log.warning(f"Ambiguous: {len(remaining_pdfs)} PDFs + {len(remaining_brds)} BRD(s) in {folder} — use subfolders")

    return pairs


def ensure_server_running():
    """Start refdes_server.py if not already running on port 5588."""
    import socket
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.settimeout(1)
        s.connect(("127.0.0.1", 5588))
        s.close()
        return  # already running
    except (ConnectionRefusedError, OSError):
        pass

    log.info("Starting navigation server (refdes_server.py)...")
    server_script = os.path.join(BASE_DIR, "refdes_server.py")
    subprocess.Popen(
        [PYTHON, server_script],
        creationflags=subprocess.CREATE_NEW_CONSOLE,
        cwd=BASE_DIR
    )
    time.sleep(2)
    log.info("  Navigation server started on http://localhost:5588/")


def process_pair(folder, pdf_name, brd_name):
    """Extract components from BRD + generate linked PDF for a pair."""
    pdf_path = os.path.join(folder, pdf_name)
    brd_path = os.path.join(folder, brd_name)
    comp_data = os.path.join(folder, "component_data.json")
    net_data = os.path.join(folder, "net_data.json")

    _log_activity(f"[Processing] {folder}\\{pdf_name} + {brd_name}")
    notify("Processing", f"{pdf_name} + {brd_name}")
    log.info(f"{'='*50}")
    log.info(f"Processing: {pdf_name} + {brd_name}")
    log.info(f"Folder: {folder}")

    # Wait for files to finish copying
    for path, name in [(pdf_path, pdf_name), (brd_path, brd_name)]:
        log.info(f"  Waiting for {name} to stabilize...")
        if not wait_for_stable_file(path):
            log.error(f"  {name} never stabilized — skipping")
            notify("Error", f"{name} never stabilized")
            return False

    # Step 1: Extract component data from BRD binary (no Allegro needed)
    log.info(f"  [Step 1] Extracting components from {brd_name} (pure Python)...")
    try:
        from extract_brd_components import extract_components_from_brd, extract_nets_from_brd
        components = extract_components_from_brd(brd_path)
        if not components:
            log.error(f"  No components found in {brd_name}")
            notify("Error", f"No components found in {brd_name}")
            return False
        with open(comp_data, "w") as f:
            json.dump(components, f, indent=2)
        log.info(f"  Extracted {len(components)} components → {os.path.basename(comp_data)}")

        # Also extract net names
        nets = extract_nets_from_brd(brd_path)
        if nets:
            with open(net_data, "w") as f:
                json.dump(nets, f)
            log.info(f"  Extracted {len(nets)} net names → {os.path.basename(net_data)}")
    except Exception as e:
        log.error(f"  Component extraction failed: {e}")
        notify("Error", f"Component extraction failed: {e}")
        return False

    # Step 2: Generate linked PDF
    log.info(f"  [Step 2] Generating linked PDF...")
    gen_cmd = [
        PYTHON, os.path.join(BASE_DIR, "generate_linked_pdf.py"),
        "--input", pdf_path,
        "--brd", brd_path,
        "--component-data", comp_data,
    ]
    if os.path.exists(net_data):
        gen_cmd.extend(["--net-data", net_data])

    try:
        result = subprocess.run(gen_cmd, capture_output=True, text=True, timeout=300)
        if result.returncode != 0:
            log.error(f"  generate_linked_pdf failed:\n{result.stderr}")
            notify("Error", f"Linked PDF generation failed")
            return False
        log.info(f"  generate_linked_pdf output:\n{result.stdout}")
    except subprocess.TimeoutExpired:
        log.error(f"  generate_linked_pdf timed out (5 min)")
        notify("Error", "Linked PDF generation timed out (5 min)")
        return False

    log.info(f"  ✅ Done! Check folder for *_linked_Acrobat.pdf")

    # Auto-start navigation server so PDF links work immediately
    ensure_server_running()

    # Step 3: Download datasheets in background (if enabled)
    settings = _load_settings()
    if not settings.get('download_datasheets', True):
        log.info(f"  [Step 3] Datasheet download disabled — skipping")
    else:
        _launch_datasheet_download(pdf_path, comp_data, folder)

def _launch_datasheet_download(pdf_path, comp_data, folder):
    """Launch or resume datasheet download for a project folder."""
    log.info(f"  [Step 3] Starting background datasheet download...")
    try:
        from download_datasheets import extract_refdes_part_pairs, download_worker
        associations = extract_refdes_part_pairs(pdf_path, comp_data)
        if associations:
            ds_dir = os.path.join(folder, "datasheets")
            os.makedirs(ds_dir, exist_ok=True)
            seen_pn = {}
            tasks = []
            for refdes, pn in sorted(associations.items()):
                if pn not in seen_pn:
                    seen_pn[pn] = refdes
                    tasks.append((refdes, pn))
            log.info(f"  Found {len(tasks)} unique parts — downloading in background")
            _log_activity(f"[Datasheets] Starting download of {len(tasks)} parts")
            progress_file = os.path.join(ds_dir, "_progress.json")
            with open(os.path.join(ds_dir, "_part_map.json"), 'w') as f:
                json.dump(associations, f, indent=2)

            # Save context for resume
            _last_ds_context.update({
                'pdf_path': pdf_path, 'comp_data': comp_data, 'folder': folder
            })

            def _ds_progress(current, total, name):
                _log_activity(f"[Datasheets] {current}/{total} - {name}")
                if os.path.exists(_DS_CANCEL_FILE):
                    _ds_cancel_event.set()
                    _log_activity("[Datasheets] Download stopped by user")
                    try:
                        os.unlink(_DS_CANCEL_FILE)
                    except OSError:
                        pass

            _ds_cancel_event.clear()
            if os.path.exists(_DS_CANCEL_FILE):
                os.unlink(_DS_CANCEL_FILE)

            t = threading.Thread(
                target=download_worker,
                args=(tasks, ds_dir, progress_file),
                kwargs={"progress_callback": _ds_progress,
                        "cancel_event": _ds_cancel_event},
                daemon=True
            )
            t.start()
        else:
            log.info(f"  No part numbers found in PDF — skipping datasheet download")
    except Exception as e:
        log.warning(f"  Datasheet download setup failed: {e}")

    return True


class InboxHandler(FileSystemEventHandler):
    """Watchdog handler — queues new PDF/BRD files for processing."""

    def __init__(self, job_queue):
        self.job_queue = job_queue
        self._recent = {}  # debounce: path → timestamp

    def _debounce(self, path):
        now = time.time()
        last = self._recent.get(path, 0)
        if now - last < 5:
            return True  # duplicate event
        self._recent[path] = now
        return False

    def on_created(self, event):
        if event.is_directory:
            return
        self._handle(event.src_path)

    def on_moved(self, event):
        if event.is_directory:
            return
        self._handle(event.dest_path)

    def _handle(self, path):
        filename = os.path.basename(path)
        if is_ignored(filename):
            return
        lower = filename.lower()
        if not (lower.endswith(".pdf") or lower.endswith(".brd")):
            return
        if self._debounce(path):
            return

        folder = os.path.dirname(path)
        log.info(f"Detected: {filename}")
        # Delay to allow companion file to arrive
        self.job_queue.put(folder)


def worker(job_queue):
    """Single-threaded worker — processes one job at a time."""
    processed_recently = {}  # folder → timestamp, avoid rapid re-processing
    while True:
        folder = job_queue.get()
        # Debounce: wait a bit for companion files, skip if recently processed
        time.sleep(5)

        # Drain duplicate folder entries
        seen = {folder}
        while not job_queue.empty():
            try:
                f = job_queue.get_nowait()
                seen.add(f)
            except queue.Empty:
                break

        for folder in seen:
            now = time.time()
            if now - processed_recently.get(folder, 0) < 30:
                continue

            pairs = find_matching_pair(folder)
            if not pairs:
                log.info(f"No complete PDF+BRD pair in {folder} yet — waiting...")
                continue

            for pdf, brd in pairs:
                notify("Generating Linked PDF", f"⚙️ {pdf} + {brd}")
                success = process_pair(folder, pdf, brd)
                if success is not False:
                    notify("Linked PDF Ready", f"✅ {pdf.replace('.pdf', '_linked_Acrobat.pdf')}")
                else:
                    notify("Generation Failed", f"❌ {pdf} + {brd}")
                processed_recently[folder] = time.time()


def startup_sweep(inbox):
    """Process any existing pairs on startup."""
    log.info("Startup sweep — checking for existing file pairs...")
    count = 0
    for root, dirs, files in os.walk(inbox):
        pairs = find_matching_pair(root)
        for pdf, brd in pairs:
            process_pair(root, pdf, brd)
            count += 1
    if count == 0:
        log.info("  No pending pairs found")
    return count


def main():
    parser = argparse.ArgumentParser(description="Auto-generate linked PDFs from watched folders")
    parser.add_argument("--inbox", default=os.path.join(BASE_DIR, "inbox"),
                        help="Primary folder to watch (default: inbox/)")
    parser.add_argument("--watch", action="append", default=[],
                        help="Additional folders to watch (can be repeated)")
    parser.add_argument("--no-sweep", action="store_true",
                        help="Skip startup sweep of existing files")
    args = parser.parse_args()

    watch_dirs = [os.path.abspath(args.inbox)]
    for w in args.watch:
        watch_dirs.append(os.path.abspath(w))

    # Restore previously saved folders from UI
    _saved_folders = os.path.join(os.environ.get('TEMP', '.'), 'ss_watch_folders.json')
    if os.path.exists(_saved_folders):
        try:
            saved = json.load(open(_saved_folders, encoding='utf-8'))
            for f in saved:
                if f not in watch_dirs and os.path.isdir(f):
                    watch_dirs.append(f)
        except Exception:
            pass

    # Ensure all directories exist
    for d in watch_dirs:
        os.makedirs(d, exist_ok=True)

    print("=" * 55)
    print("  Smart Schematic — Auto Linked PDF Watcher")
    print("=" * 55)
    for d in watch_dirs:
        print(f"  Watching: {d}")
    print(f"  Drop PDF + BRD files into any watched folder")
    print(f"  Subfolders supported (e.g., ProjectA\\)")
    print(f"  Press Ctrl+C to stop")
    print("=" * 55)

    # Start navigation server so existing linked PDFs work immediately
    ensure_server_running()

    # Process existing files in background (after tray icon is ready)
    if not args.no_sweep:
        def _deferred_sweep():
            time.sleep(3)  # Wait for tray icon to initialize
            _log_activity("[Startup] Scanning existing files...")
            for d in watch_dirs:
                startup_sweep(d)
            _log_activity("[Startup] Scan complete")
        threading.Thread(target=_deferred_sweep, daemon=True).start()
    else:
        _log_activity("[Startup] Watcher ready (no initial scan)")

    # Start worker thread
    job_queue = queue.Queue()
    t = threading.Thread(target=worker, args=(job_queue,), daemon=True)
    t.start()

    # Start watchdog observer for all directories
    handler = InboxHandler(job_queue)
    observer = Observer()
    for d in watch_dirs:
        observer.schedule(handler, d, recursive=True)
        log.info(f"Scheduled watch: {d}")
    observer.start()

    log.info("Watcher running — waiting for files...")

    # System tray icon — must run on main thread (Windows requirement)
    if HAS_TRAY:
        def _create_tray_icon():
            img = Image.new('RGBA', (32, 32), (0, 0, 0, 0))
            draw = ImageDraw.Draw(img)
            bolt = [(16, 2), (8, 16), (14, 16), (12, 30), (24, 14), (17, 14), (20, 2)]
            draw.polygon(bolt, fill=(255, 200, 0, 255))
            return img

        # Use python.exe (not pythonw) for subprocess dialogs
        _python_exe = sys.executable.replace("pythonw.exe", "python.exe")

        # For subprocess dialogs: hide console window via STARTUPINFO
        _startupinfo = subprocess.STARTUPINFO()
        _startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        _startupinfo.wShowWindow = 0  # SW_HIDE

        _ui_proc = None  # Track the manager UI subprocess

        def _open_manager_ui(icon, item=None):
            """Open the folder manager UI in a separate process."""
            nonlocal _ui_proc
            # Prevent duplicate windows
            if _ui_proc is not None and _ui_proc.poll() is None:
                return
            import subprocess as sp, tempfile, json
            # Write current folders to a temp JSON for the UI to read
            state_file = os.path.join(os.environ.get('TEMP', '.'), 'ss_watch_state.json')
            result_file = os.path.join(os.environ.get('TEMP', '.'), 'ss_watch_result.json')
            settings_file = _SETTINGS_FILE
            activity_file = _ACTIVITY_FILE
            with open(state_file, 'w', encoding='utf-8') as f:
                json.dump(watch_dirs, f)
            # Remove stale result
            if os.path.exists(result_file):
                os.unlink(result_file)

            fd, tmp = tempfile.mkstemp(suffix=".pyw")
            with os.fdopen(fd, 'w', encoding='utf-8') as f:
                f.write("# -*- coding: utf-8 -*-\n")
                f.write("import tkinter as tk\n")
                f.write("from tkinter import filedialog, messagebox\n")
                f.write("import json, os, sys\n\n")
                f.write(f"STATE_FILE = {repr(state_file)}\n")
                f.write(f"RESULT_FILE = {repr(result_file)}\n")
                f.write(f"SETTINGS_FILE = {repr(settings_file)}\n")
                f.write(f"ACTIVITY_FILE = {repr(activity_file)}\n")
                f.write(f"DS_CANCEL_FILE = {repr(_DS_CANCEL_FILE)}\n")
                f.write(f"DS_RESUME_FILE = {repr(_DS_RESUME_FILE)}\n")
                f.write(f"SCAN_TRIGGER_FILE = {repr(_SCAN_TRIGGER_FILE)}\n")
                f.write(f"SCRIPT_FILE = {repr(tmp)}\n\n")
                f.write(textwrap.dedent("""\
                    folders = json.load(open(STATE_FILE, encoding='utf-8'))
                    settings = {}
                    if os.path.exists(SETTINGS_FILE):
                        try:
                            settings = json.load(open(SETTINGS_FILE, encoding='utf-8'))
                        except Exception:
                            pass

                    root = tk.Tk()
                    root.title('Smart Schematic v2.1 - Folder Monitor')
                    root.attributes('-topmost', False)
                    root.resizable(True, True)
                    root.geometry('560x480')
                    root.configure(bg='#f0f0f0')

                    # Header
                    hdr = tk.Frame(root, bg='#2c3e50', height=40)
                    hdr.pack(fill='x')
                    hdr.pack_propagate(False)
                    tk.Label(hdr, text='\\u26a1 Smart Schematic - Folder Monitor',
                             font=('Segoe UI', 11, 'bold'), fg='white', bg='#2c3e50').pack(
                             side='left', padx=10, pady=8)

                    # --- Watched Folders Section ---
                    tk.Label(root, text='Watched Folders:', font=('Segoe UI', 9, 'bold'),
                             bg='#f0f0f0', anchor='w').pack(fill='x', padx=10, pady=(8,2))

                    frm = tk.Frame(root, bg='#f0f0f0')
                    frm.pack(fill='both', expand=True, padx=10, pady=(0,5))

                    scrollbar = tk.Scrollbar(frm)
                    scrollbar.pack(side='right', fill='y')

                    listbox = tk.Listbox(frm, font=('Segoe UI', 9), selectmode='single',
                                         yscrollcommand=scrollbar.set, bg='white', relief='flat',
                                         highlightthickness=1, highlightcolor='#3498db', height=5)
                    listbox.pack(fill='both', expand=True)
                    scrollbar.config(command=listbox.yview)

                    for fld in folders:
                        listbox.insert('end', fld)

                    # Folder buttons
                    btn_frm = tk.Frame(root, bg='#f0f0f0')
                    btn_frm.pack(fill='x', padx=10, pady=(0, 5))
                    btn_style = dict(font=('Segoe UI', 9), relief='flat', cursor='hand2', padx=12, pady=3)
                    state = {'changed': False}
                    status_var = tk.StringVar(value=f'Monitoring {len(folders)} folder(s)')

                    def add_folder():
                        d = filedialog.askdirectory(title='Select folder to monitor',
                                                   initialdir=os.path.expanduser('~'))
                        if d:
                            d = os.path.normpath(d)
                            if d not in folders:
                                folders.append(d)
                                listbox.insert('end', d)
                                state['changed'] = True
                                status_var.set(f'Monitoring {len(folders)} folder(s)')

                    def remove_folder():
                        sel = listbox.curselection()
                        if not sel:
                            messagebox.showwarning('Remove Folder', 'Select a folder to remove.')
                            return
                        idx = sel[0]
                        folders.pop(idx)
                        listbox.delete(idx)
                        state['changed'] = True
                        status_var.set(f'Monitoring {len(folders)} folder(s)')

                    tk.Button(btn_frm, text='+ Add Folder', command=add_folder,
                              bg='#27ae60', fg='white', **btn_style).pack(side='left', padx=(0,5))
                    tk.Button(btn_frm, text='- Remove', command=remove_folder,
                              bg='#e74c3c', fg='white', **btn_style).pack(side='left', padx=(0,5))
                    def scan_folders():
                        with open(SCAN_TRIGGER_FILE, 'w') as sf:
                            sf.write('scan')
                        messagebox.showinfo('Scan', 'Scanning watched folders for existing PDF + BRD pairs...')
                    tk.Button(btn_frm, text='Scan Now', command=scan_folders,
                              bg='#3498db', fg='white', **btn_style).pack(side='right')

                    # --- Settings Section ---
                    settings_frm = tk.Frame(root, bg='#f0f0f0')
                    settings_frm.pack(fill='x', padx=10, pady=(5, 0))
                    ds_var = tk.BooleanVar(value=settings.get('download_datasheets', True))
                    def on_ds_toggle():
                        settings['download_datasheets'] = ds_var.get()
                        state['changed'] = True
                    tk.Checkbutton(settings_frm, text='Download datasheets from OnePDM',
                                   variable=ds_var, command=on_ds_toggle,
                                   font=('Segoe UI', 9), bg='#f0f0f0',
                                   activebackground='#f0f0f0').pack(side='left')
                    def stop_download():
                        with open(DS_CANCEL_FILE, 'w') as cf:
                            cf.write('cancel')
                        messagebox.showinfo('Stop Download', 'Datasheet download will stop after the current item.')
                    def resume_download():
                        with open(DS_RESUME_FILE, 'w') as rf:
                            rf.write('resume')
                        messagebox.showinfo('Resume Download', 'Datasheet download will resume shortly.')
                    tk.Button(settings_frm, text='Resume', command=resume_download,
                              bg='#2980b9', fg='white', font=('Segoe UI', 8), relief='flat',
                              cursor='hand2', padx=8, pady=1).pack(side='right', padx=(0,5))
                    tk.Button(settings_frm, text='Stop', command=stop_download,
                              bg='#e67e22', fg='white', font=('Segoe UI', 8), relief='flat',
                              cursor='hand2', padx=8, pady=1).pack(side='right')

                    # --- Activity Log Section ---
                    log_hdr = tk.Frame(root, bg='#f0f0f0')
                    log_hdr.pack(fill='x', padx=10, pady=(5,2))
                    tk.Label(log_hdr, text='Activity Log:', font=('Segoe UI', 9, 'bold'),
                             bg='#f0f0f0', anchor='w').pack(side='left')
                    def clear_log():
                        with open(ACTIVITY_FILE, 'w', encoding='utf-8') as af:
                            json.dump([], af)
                    tk.Button(log_hdr, text='Clear', command=clear_log,
                              font=('Segoe UI', 7), relief='flat', bg='#bdc3c7',
                              cursor='hand2', padx=6).pack(side='right')
                    def open_folder():
                        sel = listbox.curselection()
                        folder = folders[sel[0]] if sel else (folders[0] if folders else None)
                        if folder and os.path.isdir(folder):
                            os.startfile(folder)
                    tk.Button(log_hdr, text='Open Folder', command=open_folder,
                              font=('Segoe UI', 7), relief='flat', bg='#8e44ad', fg='white',
                              cursor='hand2', padx=6).pack(side='right', padx=(0,5))
                    def reprocess():
                        sel = listbox.curselection()
                        folder = folders[sel[0]] if sel else None
                        if not folder:
                            messagebox.showwarning('Re-process', 'Select a folder first.')
                            return
                        with open(SCAN_TRIGGER_FILE, 'w') as sf:
                            sf.write('scan')
                        messagebox.showinfo('Re-process', f'Re-scanning {folder}...')
                    tk.Button(log_hdr, text='Re-process', command=reprocess,
                              font=('Segoe UI', 7), relief='flat', bg='#16a085', fg='white',
                              cursor='hand2', padx=6).pack(side='right', padx=(0,5))

                    log_frm = tk.Frame(root, bg='#f0f0f0')
                    log_frm.pack(fill='both', expand=True, padx=10, pady=(0,5))

                    log_scroll = tk.Scrollbar(log_frm)
                    log_scroll.pack(side='right', fill='y')

                    log_text = tk.Text(log_frm, font=('Consolas', 9), height=8, wrap='word',
                                       yscrollcommand=log_scroll.set, bg='#1e1e1e', fg='#00ff88',
                                       relief='flat', state='disabled', padx=6, pady=4)
                    log_text.pack(fill='both', expand=True)
                    log_scroll.config(command=log_text.yview)

                    # --- Download Progress Bar ---
                    import re as _re
                    prog_frm = tk.Frame(root, bg='#f0f0f0')
                    prog_frm.pack(fill='x', padx=10, pady=(0,3))
                    prog_lbl = tk.Label(prog_frm, text='', font=('Segoe UI', 8),
                                        bg='#f0f0f0', fg='#2d3436', anchor='w')
                    prog_lbl.pack(side='left')
                    prog_canvas = tk.Canvas(prog_frm, height=12, bg='#ecf0f1',
                                            highlightthickness=0)
                    prog_canvas.pack(fill='x', expand=True, padx=(5,0))
                    prog_bar = prog_canvas.create_rectangle(0, 0, 0, 12, fill='#27ae60', width=0)

                    def refresh_activity():
                        try:
                            if os.path.exists(ACTIVITY_FILE):
                                with open(ACTIVITY_FILE, encoding='utf-8') as af:
                                    entries = json.load(af)
                            else:
                                entries = []
                        except (json.JSONDecodeError, OSError):
                            entries = []
                        log_text.config(state='normal')
                        log_text.delete('1.0', 'end')
                        if not entries:
                            log_text.insert('end', '  (no activity yet - waiting for files)')
                        else:
                            for ts, msg in entries[-20:]:
                                log_text.insert('end', f'[{ts}] {msg}\\n')
                        log_text.see('end')
                        log_text.config(state='disabled')
                        # Update progress bar from last datasheet entry
                        ds_progress = None
                        for ts, msg in reversed(entries):
                            m = _re.search(r'\\[Datasheets\\] (\\d+)/(\\d+)', msg)
                            if m:
                                ds_progress = (int(m.group(1)), int(m.group(2)))
                                break
                        if ds_progress:
                            cur, total = ds_progress
                            pct = cur / total if total > 0 else 0
                            w = prog_canvas.winfo_width()
                            prog_canvas.coords(prog_bar, 0, 0, int(w * pct), 12)
                            prog_lbl.config(text=f'Datasheets: {cur}/{total}')
                            if cur >= total:
                                prog_canvas.itemconfig(prog_bar, fill='#27ae60')
                            else:
                                prog_canvas.itemconfig(prog_bar, fill='#3498db')
                        else:
                            prog_lbl.config(text='')
                            prog_canvas.coords(prog_bar, 0, 0, 0, 12)
                        log_text.see('end')
                        log_text.config(state='disabled')
                        root.after(2000, refresh_activity)

                    refresh_activity()

                    # --- Status bar ---
                    status_bar = tk.Label(root, textvariable=status_var, font=('Segoe UI', 8),
                                          bg='#dfe6e9', fg='#2d3436', anchor='w', padx=8, pady=3)
                    status_bar.pack(fill='x', side='bottom')

                    # Close button
                    close_frm = tk.Frame(root, bg='#f0f0f0')
                    close_frm.pack(fill='x', padx=10, pady=(0, 8))

                    def on_close():
                        if state['changed']:
                            with open(RESULT_FILE, 'w', encoding='utf-8') as rf:
                                json.dump(folders, rf)
                            with open(SETTINGS_FILE, 'w', encoding='utf-8') as sf:
                                json.dump(settings, sf)
                        try:
                            os.unlink(SCRIPT_FILE)
                        except OSError:
                            pass
                        try:
                            os.unlink(STATE_FILE)
                        except OSError:
                            pass
                        root.destroy()

                    tk.Button(close_frm, text='Close', command=on_close,
                              bg='#95a5a6', fg='white', **btn_style).pack(side='right')

                    root.protocol('WM_DELETE_WINDOW', on_close)
                    root.mainloop()
                """))

            def _wait_for_ui():
                nonlocal _ui_proc
                _ui_proc = sp.Popen([_python_exe, tmp], startupinfo=_startupinfo)
                _ui_proc.wait(timeout=300)
                if os.path.exists(result_file):
                    try:
                        new_folders = json.load(open(result_file, encoding='utf-8'))
                        os.unlink(result_file)
                    except Exception:
                        return
                    # Determine added/removed
                    added = [f for f in new_folders if f not in watch_dirs]
                    removed = [f for f in watch_dirs if f not in new_folders]
                    # Remove watches
                    for f in removed:
                        watch_dirs.remove(f)
                        log.info(f"Removed watch: {f}")
                    # Add watches
                    for f in added:
                        if os.path.isdir(f):
                            observer.schedule(handler, f, recursive=True)
                            watch_dirs.append(f)
                            log.info(f"Added watch: {f}")
                    if added or removed:
                        icon.title = f"Smart Schematic v2.1 — Watching {len(watch_dirs)} folders"
                        notify("Folders Updated",
                               f"Now watching {len(watch_dirs)} folder(s)")
                        # Persist folder list for next restart
                        try:
                            with open(_saved_folders, 'w', encoding='utf-8') as sf:
                                json.dump(watch_dirs, sf)
                        except Exception:
                            pass

            threading.Thread(target=_wait_for_ui, daemon=True).start()

        def _resume_watcher():
            """Background thread that watches for resume/scan signals from UI."""
            while True:
                time.sleep(2)
                if os.path.exists(_DS_RESUME_FILE):
                    try:
                        os.unlink(_DS_RESUME_FILE)
                    except OSError:
                        pass
                    if _last_ds_context:
                        _log_activity("[Datasheets] Resuming download...")
                        _launch_datasheet_download(
                            _last_ds_context['pdf_path'],
                            _last_ds_context['comp_data'],
                            _last_ds_context['folder']
                        )
                    else:
                        _log_activity("[Datasheets] No previous download to resume")
                if os.path.exists(_SCAN_TRIGGER_FILE):
                    try:
                        os.unlink(_SCAN_TRIGGER_FILE)
                    except OSError:
                        pass
                    _log_activity("[Scan] Scanning watched folders...")
                    for d in watch_dirs:
                        startup_sweep(d)

        threading.Thread(target=_resume_watcher, daemon=True).start()

        def _on_quit(icon, item):
            icon.stop()
            observer.stop()
            os._exit(0)

        menu = pystray.Menu(
            pystray.MenuItem("Smart Schematic v2.1 — Watcher", None, enabled=False),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Manage Folders...", _open_manager_ui, default=True),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Quit", _on_quit)
        )
        tray_icon = pystray.Icon("SmartSchematic", _create_tray_icon(),
                                  f"Smart Schematic v2.1 — Watching {len(watch_dirs)} folders", menu)
        global _tray_icon
        _tray_icon = tray_icon
        log.info("System tray icon active")
        tray_icon.run()  # Blocks on main thread until Quit
        observer.stop()
        observer.join()
        return

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        log.info("Stopping watcher...")
        observer.stop()
    observer.join()


if __name__ == "__main__":
    main()
