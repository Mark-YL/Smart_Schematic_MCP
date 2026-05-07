"""
Smart Schematic — Navigation Server

Local HTTP server that bridges clickable PDF links to Allegro Free Viewer and OnePDM.
When a refdes link is clicked in the PDF, this server receives the request
and automates Allegro to zoom/highlight the component.

PDF links: http://localhost:5588/REFDES?x=...&y=...&layer=...
Spawns navigate_worker.py subprocess for each navigation (fresh UIA context).

Usage:  python refdes_server.py   (keep running while using the linked PDF)
"""

import http.server
import subprocess
import os
import sys
import re
import time
import ctypes
import ctypes.wintypes
import json
import argparse
import threading
from urllib.parse import urlparse, parse_qs

PORT = 5588
BRD_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "input.brd")
from allegro_utils import get_allegro_viewer
ALLEGRO_VIEWER = get_allegro_viewer()
COMPONENT_DATA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "component_data.json")
WORKER_SCRIPT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "navigate_worker.py")

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    pass

import win32gui

user32 = ctypes.windll.user32

# Component data cache — loaded in main() to support CLI args
_comp_cache = {}
args_comp_data = None  # set in main() for per-project comp_data comparison

_CLOSE_HTML = (
    '<html><head><title>x</title>'
    '<script>window.close();'
    'setTimeout(function(){window.close()},200);'
    '</script></head>'
    '<body style="background:#000;margin:0"></body></html>'
)


# Track browser HWND to minimize it after responding
_browser_hwnd = None


def _close_browser_tab():
    """Close the browser tab by sending Ctrl+W to the foreground browser window."""
    global _browser_hwnd
    try:
        fg = user32.GetForegroundWindow()
        if fg:
            cls = win32gui.GetClassName(fg)
            title = win32gui.GetWindowText(fg)
            # Common browser window classes / titles
            is_browser = any(k in cls for k in (
                "Chrome", "MozillaWindow", "Edge", "Opera", "ApplicationFrameWindow"
            )) or any(k in title.lower() for k in (
                "chrome", "firefox", "edge", "opera", "brave", "localhost", "127.0.0.1"
            ))
            if is_browser:
                _browser_hwnd = fg
                WM_KEYDOWN = 0x0100
                WM_KEYUP = 0x0101
                VK_CONTROL = 0x11
                VK_W = 0x57
                user32.PostMessageW(fg, WM_KEYDOWN, VK_CONTROL, 0)
                user32.PostMessageW(fg, WM_KEYDOWN, VK_W, 0)
                time.sleep(0.05)
                user32.PostMessageW(fg, WM_KEYUP, VK_W, 0)
                user32.PostMessageW(fg, WM_KEYUP, VK_CONTROL, 0)
                return
        # Also try the cached HWND
        if _browser_hwnd and win32gui.IsWindow(_browser_hwnd):
            win32gui.ShowWindow(_browser_hwnd, 6)
    except Exception:
        pass

# Allegro detection now handled per-project in _find_allegro_hwnd_for / _launch_allegro_for


def _navigate(refdes, layer, proj_id=None):
    """Send navigation command to the correct project's persistent worker."""
    _load_projects_config()  # reload in case new projects were added
    pid = proj_id if proj_id else "_default"
    for attempt in range(2):
        t_get = time.perf_counter()
        worker = _get_worker(proj_id)
        print(f"  TIMING: get_worker={time.perf_counter()-t_get:.1f}s", flush=True)
        if not worker:
            print("(worker_start_err) ", end="", flush=True)
            return False
        with worker["lock"]:
            try:
                worker["proc"].stdin.write(f"{refdes} {layer or ''}\n")
                worker["proc"].stdin.flush()
            except Exception as e:
                print(f"(write_err:{e}) ", end="", flush=True)
                _kill_worker_proc(worker)
                with _workers_lock:
                    if _workers.get(pid) is worker:
                        _workers.pop(pid, None)
                if attempt == 0:
                    continue
                return False
            t_resp = time.perf_counter()
            ok = _read_worker_response(worker)
            print(f"  TIMING: worker_navigate={time.perf_counter()-t_resp:.1f}s", flush=True)
            if ok:
                # Cache detected layer if it was unknown and worker detected it
                detected = worker.get("_last_detected_layer")
                if detected and not layer:
                    _cache_detected_layer(refdes, detected, proj_id)
                return True
            if attempt == 0:
                with _workers_lock:
                    if _workers.get(pid) is worker:
                        _workers.pop(pid, None)
                print("(retrying...) ", end="", flush=True)
                continue
            return False
    return False


def _cache_detected_layer(refdes, layer, proj_id):
    """Cache a detected layer in component_data.json for future use.
    
    Uses atomic write (temp file + os.replace) to prevent corruption.
    Also updates the in-memory cache immediately.
    """
    try:
        _load_projects_config()
        pid = proj_id if proj_id else "_default"
        entry = _projects_config.get(pid)
        if not entry:
            return
        comp_path = entry.get("component_data")
        if not comp_path or not os.path.exists(comp_path):
            return
        
        import tempfile
        
        with open(comp_path) as f:
            comp_data = json.load(f)
        if refdes in comp_data and not comp_data[refdes].get("layer"):
            comp_data[refdes]["layer"] = layer
            # Atomic write: temp file → flush → fsync → os.replace
            dir_name = os.path.dirname(comp_path)
            fd, tmp_path = tempfile.mkstemp(dir=dir_name, suffix=".tmp")
            try:
                with os.fdopen(fd, 'w') as f:
                    json.dump(comp_data, f, indent=2)
                    f.flush()
                    os.fsync(f.fileno())
                os.replace(tmp_path, comp_path)
            except Exception:
                try:
                    os.remove(tmp_path)
                except Exception:
                    pass
                raise
            
            # Update in-memory cache immediately
            if _comp_cache and refdes in _comp_cache:
                _comp_cache[refdes]["layer"] = layer
            
            print(f"  [layer-cache] {refdes}={layer} saved", flush=True)
    except Exception as e:
        print(f"  [layer-cache] err: {e}", flush=True)


def _navigate_net(net_name, proj_id=None):
    """Send net highlight command to the correct project's persistent worker."""
    _load_projects_config()  # reload in case new projects were added
    pid = proj_id if proj_id else "_default"
    for attempt in range(2):
        worker = _get_worker(proj_id)
        if not worker:
            print("(worker_start_err) ", end="", flush=True)
            return False
        with worker["lock"]:
            try:
                worker["proc"].stdin.write(f"NET:{net_name}\n")
                worker["proc"].stdin.flush()
            except Exception as e:
                print(f"(write_err:{e}) ", end="", flush=True)
                _kill_worker_proc(worker)
                with _workers_lock:
                    if _workers.get(pid) is worker:
                        _workers.pop(pid, None)
                if attempt == 0:
                    continue
                return False
            ok = _read_worker_response(worker)
            if ok:
                return True
            if attempt == 0:
                with _workers_lock:
                    if _workers.get(pid) is worker:
                        _workers.pop(pid, None)
                print("(retrying...) ", end="", flush=True)
                continue
            return False
    return False


# ---- Multi-project worker management ----
# Each project gets its own persistent worker targeting its Allegro window.
# Workers are keyed by proj_id (from the &proj= query param).
_workers = {}       # proj_id -> {"proc": Popen, "lock": Lock, "brd_name": str, "brd_file": str}
_workers_lock = threading.Lock()  # protects _workers dict mutations
_nav_lock = threading.Lock()  # serializes all Allegro navigation requests

_projects_config = {}  # proj_id -> {"brd": path, "component_data": path}
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def _load_projects_config():
    """Load project registry from .brdnav_config.json."""
    global _projects_config
    config_path = os.path.join(_SCRIPT_DIR, ".brdnav_config.json")
    try:
        with open(config_path) as f:
            registry = json.load(f)
        _projects_config = registry.get("projects", {})
    except Exception:
        _projects_config = {}


def _get_project_info(proj_id):
    """Get BRD info for a project ID. Returns (brd_file, brd_name, comp_cache).
    Only falls back to active project when no explicit proj_id was given."""
    # Try explicit project ID first
    entry = _projects_config.get(proj_id) if proj_id else None

    if proj_id and not entry:
        # proj_id was provided (from PDF link) but not in registry — don't silently
        # fall back to a different BRD. Log warning and return None so caller can handle.
        print(f"  WARNING: Project {proj_id} not found in registry!", flush=True)
        print(f"  Available: {list(_projects_config.keys())}", flush=True)
        return None, None, None

    # Fallback: use active project only for legacy links (no proj_id)
    if not entry:
        config_path = os.path.join(_SCRIPT_DIR, ".brdnav_config.json")
        try:
            with open(config_path) as f:
                registry = json.load(f)
            active = registry.get("active", "")
            if active and active in _projects_config:
                entry = _projects_config[active]
        except Exception:
            pass

    if entry:
        brd = entry.get("brd", BRD_FILE)
        comp_path = entry.get("component_data")
        comp = _comp_cache
        if comp_path and os.path.exists(comp_path) and comp_path != (args_comp_data or COMPONENT_DATA):
            try:
                with open(comp_path) as f:
                    comp = json.load(f)
            except Exception:
                pass
        return brd, os.path.basename(brd), comp
    return BRD_FILE, os.path.basename(BRD_FILE), _comp_cache


def _get_worker(proj_id):
    """Get or create a worker for the given project. Returns worker dict or None."""
    if not proj_id:
        proj_id = "_default"

    with _workers_lock:
        w = _workers.get(proj_id)
        if w and w["proc"].poll() is None:
            # Worker process alive — check if Allegro is still running
            brd_name = w.get("brd_name", "")
            if brd_name:
                current_hwnd = _find_allegro_hwnd_for(brd_name)
                if not current_hwnd:
                    print(f"  (Allegro closed for {brd_name}, killing stale worker)", flush=True)
                    _kill_worker_proc(w)
                    _workers.pop(proj_id, None)
                elif current_hwnd != w.get("allegro_hwnd"):
                    # Allegro was restarted — HWND changed, worker UIA handles are stale
                    print(f"  (Allegro restarted for {brd_name}, recreating worker)", flush=True)
                    _kill_worker_proc(w)
                    _workers.pop(proj_id, None)
                else:
                    return w
            else:
                return w

        # Need to create a new worker for this project
        brd_file, brd_name, _ = _get_project_info(proj_id if proj_id != "_default" else None)

        if brd_file is None:
            print(f"  ERROR: Cannot create worker — project {proj_id} not registered", flush=True)
            return None

        # Ensure Allegro is running for this BRD
        t_allegro = time.perf_counter()
        allegro_hwnd = _find_allegro_hwnd_for(brd_name)
        if not allegro_hwnd:
            _launch_allegro_for(brd_file, brd_name)
            allegro_hwnd = _find_allegro_hwnd_for(brd_name)
        print(f"  TIMING: allegro_ensure={time.perf_counter()-t_allegro:.1f}s", flush=True)

        python_exe = sys.executable.replace("pythonw.exe", "python.exe")
        try:
            t_worker = time.perf_counter()
            proc = subprocess.Popen(
                [python_exe, WORKER_SCRIPT, "--persistent",
                 "--brd-name", brd_name],
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                creationflags=0x08000000,  # CREATE_NO_WINDOW
                env={**os.environ, "PYTHONUNBUFFERED": "1"},
            )
            # Wait for READY signal
            deadline = time.time() + 30
            while time.time() < deadline:
                line = proc.stdout.readline().strip()
                if line == "READY":
                    print(f"  TIMING: worker_ready={time.perf_counter()-t_worker:.1f}s", flush=True)
                    print(f"  Worker ready (proj={proj_id}, brd={brd_name})", flush=True)
                    worker = {"proc": proc, "lock": threading.Lock(),
                              "brd_name": brd_name, "brd_file": brd_file,
                              "allegro_hwnd": allegro_hwnd}
                    _workers[proj_id] = worker
                    return worker
                if line:
                    print(f"  Worker init: {line}", flush=True)
                if proc.poll() is not None:
                    break
            print(f"  TIMING: worker_timeout={time.perf_counter()-t_worker:.1f}s", flush=True)
            print("  Worker startup timeout", flush=True)
            try:
                proc.kill()
            except Exception:
                pass
            return None
        except Exception as e:
            print(f"  Worker start error: {e}", flush=True)
            return None


def _find_allegro_hwnd_for(brd_name):
    """Find Allegro window HWND matching a specific BRD filename."""
    brd_lower = brd_name.lower()
    best = None
    best_area = 0
    def cb(h, _):
        nonlocal best, best_area
        if win32gui.IsWindowVisible(h):
            title = win32gui.GetWindowText(h).lower()
            if "allegro" in title and "viewer" in title and brd_lower in title:
                r = win32gui.GetWindowRect(h)
                a = (r[2] - r[0]) * (r[3] - r[1])
                if a > best_area:
                    best = h
                    best_area = a
        return True
    win32gui.EnumWindows(cb, None)
    return best


def _launch_allegro_for(brd_file, brd_name):
    """Launch Allegro with a specific BRD file. Waits for design to fully load."""
    t0 = time.perf_counter()
    subprocess.Popen(
        [ALLEGRO_VIEWER, brd_file],
        creationflags=subprocess.DETACHED_PROCESS,
    )
    print(f"  (launching Allegro for {brd_name}...) ", flush=True)
    t_window = None
    # Wait for window with BRD name in title (means design is loaded)
    for i in range(360):
        time.sleep(0.5)
        hwnd = _find_allegro_hwnd_for(brd_name)
        if hwnd:
            if t_window is None:
                t_window = time.perf_counter()
                print(f"  TIMING: allegro_window_visible={t_window-t0:.1f}s", flush=True)
            title = win32gui.GetWindowText(hwnd).lower()
            if brd_name.lower() in title:
                print(f"  TIMING: allegro_design_loaded={time.perf_counter()-t0:.1f}s", flush=True)
                print(f"  (Allegro ready, design loaded) ", flush=True)
                return True
            if i % 20 == 0:
                print(f"  (Allegro loading design... {time.perf_counter()-t0:.0f}s) ", flush=True)
    print(f"  TIMING: allegro_timeout={time.perf_counter()-t0:.1f}s", flush=True)
    print("  (Allegro launch timeout) ", flush=True)
    return False


def _read_worker_response(worker, timeout=120):
    """Read worker stdout until DONE/ERR line, with thread-based timeout.
    
    Returns True on success, False on failure/timeout.
    Also checks for 'layer_detected=<Top|Bottom>' in output and caches it.
    """
    proc = worker["proc"]
    result = [False]
    detected_layer = [None]

    def _reader():
        try:
            while True:
                line = proc.stdout.readline()
                if not line:
                    break
                line = line.strip()
                if line.startswith("DONE"):
                    result[0] = True
                    break
                if line.startswith("ERR:"):
                    print(f"({line}) ", end="", flush=True)
                    result[0] = "ERR"
                    break
                if line:
                    # Check for layer detection output
                    if "layer_detected=" in line:
                        parts = line.split("layer_detected=")
                        if len(parts) > 1:
                            layer_str = parts[1].split()[0]
                            if layer_str in ("Top", "Bottom"):
                                detected_layer[0] = layer_str
                    print(f"({line}) ", end="", flush=True)
        except Exception:
            pass

    t = threading.Thread(target=_reader, daemon=True)
    t.start()
    t.join(timeout=timeout)
    if t.is_alive():
        print("(TIMEOUT) ", end="", flush=True)
        _kill_worker_proc(worker)
        return False
    if result[0] == "ERR":
        _kill_worker_proc(worker)
        return False
    
    # Store detected layer for caller to retrieve
    worker["_last_detected_layer"] = detected_layer[0]
    return result[0]


def _kill_worker_proc(worker):
    """Kill a specific worker process."""
    if worker and worker.get("proc"):
        try:
            worker["proc"].kill()
        except Exception:
            pass


# ---- OnePDM datasheet lookup ----
_onepdm_client = None


def _get_onepdm_client():
    """Lazy-init the OnePDM client."""
    global _onepdm_client
    if _onepdm_client is None:
        try:
            from onepdm_client import get_client
            _onepdm_client = get_client()
        except ImportError:
            print("  [onepdm] onepdm_client.py not found")
    return _onepdm_client


def _open_onepdm_datasheet(part_number):
    """Open the part page in OnePDM and navigate to Get all files."""
    client = _get_onepdm_client()
    if client:
        if client.open_part_page(part_number):
            return
        # If user closed the browser, don't fall back to opening another tab
        if getattr(client, '_user_closed', False):
            print(f"(OnePDM browser closed, skipped) ", end="", flush=True)
            return
    # Fallback: open OnePDM and copy part number to clipboard
    import webbrowser
    webbrowser.open("https://onepdm.plm.microsoft.com/onepdm/")
    try:
        subprocess.Popen(
            ["powershell", "-Command", f"Set-Clipboard -Value '{part_number}'"],
            creationflags=0x08000000,
        ).wait(timeout=5)
    except Exception:
        pass
    print(f"(fallback: opened OnePDM, '{part_number}' copied to clipboard) ", end="", flush=True)


class _Handler(http.server.BaseHTTPRequestHandler):
    def do_GET(self):
        parsed = urlparse(self.path)
        path_parts = parsed.path.strip("/").split("/", 1)

        # Handle /nav/REFDES endpoint (Edge PDF http:// links)
        if len(path_parts) == 2 and path_parts[0].lower() == "nav":
            refdes = path_parts[1].split("?")[0].upper()
            params = parse_qs(parsed.query)
            layer = params.get("layer", [None])[0]
            proj_id = params.get("proj", [None])[0]
            if not layer:
                _, _, comp = _get_project_info(proj_id)
                if comp:
                    c = comp.get(refdes)
                    if c:
                        layer = c.get("layer")
            self.send_response(200)
            self.send_header("Content-type", "text/html")
            self.send_header("Connection", "close")
            self.end_headers()
            self.wfile.write(_CLOSE_HTML.encode())
            self.wfile.flush()
            t0 = time.time()
            print(f"  -> {refdes} ({layer or '?'}) [proj={proj_id or 'default'}] ", end="", flush=True)
            with _nav_lock:
                _navigate(refdes, layer, proj_id)
            dt = time.time() - t0
            print(f"[{dt:.1f}s]")
            return

        # Handle /datasheet/PARTNUMBER endpoint
        if len(path_parts) == 2 and path_parts[0].lower() == "datasheet":
            part_number = path_parts[1].upper()
            print(f"  -> datasheet: {part_number} ", end="", flush=True)
            _open_onepdm_datasheet(part_number)
            self.send_response(200)
            self.send_header("Content-type", "text/html")
            self.send_header("Connection", "close")
            self.end_headers()
            self.wfile.write(_CLOSE_HTML.encode())
            self.wfile.flush()
            print("done")
            return

        # Handle /net/NETNAME endpoint — highlight net in Allegro
        if len(path_parts) >= 2 and path_parts[0].lower() == "net":
            from urllib.parse import unquote
            net_name = unquote(path_parts[1])
            net_params = parse_qs(parsed.query)
            proj_id = net_params.get("proj", [None])[0]
            print(f"  -> net: {net_name} [proj={proj_id or 'default'}] ", end="", flush=True)

            # Respond immediately
            self.send_response(200)
            self.send_header("Content-type", "text/html")
            self.send_header("Connection", "close")
            self.end_headers()
            self.wfile.write(_CLOSE_HTML.encode())
            self.wfile.flush()

            t0 = time.time()
            with _nav_lock:
                _navigate_net(net_name, proj_id)
            dt = time.time() - t0
            print(f"[{dt:.1f}s]")
            return

        refdes = parsed.path.strip("/").upper()
        params = parse_qs(parsed.query)

        if not refdes or not re.match(r"^[A-Z]+\d+[A-Z]?$", refdes):
            self.send_response(204)
            self.end_headers()
            return

        layer = params.get("layer", [None])[0]
        proj_id = params.get("proj", [None])[0]
        if not layer:
            _, _, comp = _get_project_info(proj_id)
            if comp:
                c = comp.get(refdes)
                if c:
                    layer = c.get("layer")

        # Respond immediately (browser auto-closes)
        self.send_response(200)
        self.send_header("Content-type", "text/html")
        self.send_header("Connection", "close")
        self.end_headers()
        self.wfile.write(_CLOSE_HTML.encode())
        self.wfile.flush()

        print(f"  -> {refdes} ({layer or '?'}) [proj={proj_id or 'default'}] ", end="", flush=True)

        t0 = time.time()
        with _nav_lock:
            _navigate(refdes, layer, proj_id)
        dt = time.time() - t0
        print(f"[{dt:.1f}s]")

    def log_message(self, fmt, *args):
        pass


def main():
    parser = argparse.ArgumentParser(description="Smart Schematic — Navigation Server")
    parser.add_argument("--brd", default=None,
                        help="BRD file path (default: input.brd in tool dir)")
    parser.add_argument("--component-data", default=None,
                        help="Component data JSON path")
    args = parser.parse_args()

    global BRD_FILE, _comp_cache, args_comp_data
    args_comp_data = args.component_data
    if args.brd:
        BRD_FILE = args.brd

    # Load component data
    comp_data_path = args.component_data or COMPONENT_DATA
    if os.path.exists(comp_data_path):
        with open(comp_data_path) as f:
            _comp_cache = json.load(f)

    # Load multi-project config
    _load_projects_config()

    print("=" * 55)
    print("  Smart Schematic v2.1 — Navigation Server (multi-project)")
    print(f"  Listening on http://localhost:{PORT}/")
    print("=" * 55)
    print(f"  Default BRD : {BRD_FILE}")
    print(f"  Projects    : {len(_projects_config)} registered")
    for pid, pinfo in _projects_config.items():
        print(f"    {pid} -> {os.path.basename(pinfo.get('brd', '?'))}")
    print(f"  Viewer      : {ALLEGRO_VIEWER}")
    print(f"  Cache       : {len(_comp_cache)} components")
    print(f"  Worker      : persistent per-project (fast mode)")
    print()

    # Check for any running Allegro instances
    brd_name = os.path.basename(BRD_FILE).lower()
    if _find_allegro_hwnd_for(brd_name):
        print(f"  Allegro  : found for {os.path.basename(BRD_FILE)}")
    else:
        print("  Allegro  : will launch on first click")
    print(f"  OnePDM   : datasheet lookup on demand (SSO on first click)")

    print()
    print("  Click a refdes in PDF -> Allegro zooms to it")
    print("  Click a part number   -> opens datasheet from OnePDM")
    print("  Click a net name      -> Allegro highlights the net")
    print("  Press Ctrl+C to stop")
    print("=" * 55)

    srv = http.server.ThreadingHTTPServer(("127.0.0.1", PORT), _Handler)
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\n  Stopping workers...")
        for w in _workers.values():
            _kill_worker_proc(w)
        print("  Server stopped.")
        srv.server_close()


if __name__ == "__main__":
    main()
