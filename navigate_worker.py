"""
Smart Schematic — Allegro Navigation Worker

Standalone worker: navigate Allegro Free Viewer to a component.
Uses PostMessage to send clicks directly to Allegro's window message queue,
bypassing foreground/session restrictions. Works from any terminal.

Called by refdes_server.py as a subprocess (fresh UIA context per request).

Usage:
    python navigate_worker.py <refdes> <layer> [<state_file>]

State file is a JSON dict that persists across calls to avoid repeated
display/grid setup.
"""

import sys
import os
import json
import time
import ctypes
import ctypes.wintypes
import threading

ctypes.windll.shcore.SetProcessDpiAwareness(2)

import uiautomation as auto
import win32gui
import win32con

# Reduce UIA global search timeout from 10s to 2s for faster failure detection
# Must use SetGlobalSearchTimeout — setting auto.TIME_OUT_SECOND only changes
# the __init__ alias, not the core module's variable used by actual searches.
auto.SetGlobalSearchTimeout(2)

user32 = ctypes.windll.user32
user32.WindowFromPoint.argtypes = [ctypes.wintypes.POINT]
user32.WindowFromPoint.restype = ctypes.wintypes.HWND

# BRD name filter for multi-instance Allegro support
_brd_name_filter = ""  # set via --brd-name arg


class _BusyPopup:
    """Topmost progress popup shown during first-time UIA setup.
    Runs on a daemon thread so it doesn't block the main automation."""

    _STEPS = [
        "Discovering UI panels\u2026",
        "Scanning layer table\u2026",
        "Opening command bar\u2026",
        "Highlighting component\u2026",
        "Switching layers\u2026",
        "Setting view\u2026",
    ]

    def __init__(self, hwnd=None):
        self._root = None
        self._hwnd = hwnd
        self._cur_val = 0.0
        self._closing = False
        self._ready = threading.Event()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()
        self._ready.wait(3)

    def _run(self):
        import tkinter as tk
        from tkinter import ttk

        # Get Allegro window center to place popup on same monitor
        mon_cx, mon_cy, mon_w, mon_h = None, None, None, None
        if self._hwnd:
            try:
                rect = win32gui.GetWindowRect(self._hwnd)
                mon_cx = (rect[0] + rect[2]) // 2
                mon_cy = (rect[1] + rect[3]) // 2
                # Get monitor work area via Win32
                import ctypes
                MONITOR_DEFAULTTONEAREST = 2
                hmon = ctypes.windll.user32.MonitorFromPoint(
                    ctypes.wintypes.POINT(mon_cx, mon_cy),
                    MONITOR_DEFAULTTONEAREST)
                class MONITORINFO(ctypes.Structure):
                    _fields_ = [("cbSize", ctypes.c_ulong),
                                ("rcMonitor", ctypes.wintypes.RECT),
                                ("rcWork", ctypes.wintypes.RECT),
                                ("dwFlags", ctypes.c_ulong)]
                mi = MONITORINFO()
                mi.cbSize = ctypes.sizeof(MONITORINFO)
                ctypes.windll.user32.GetMonitorInfoW(hmon, ctypes.byref(mi))
                rc = mi.rcWork
                mon_cx = (rc.left + rc.right) // 2
                mon_cy = (rc.top + rc.bottom) // 2
                mon_w = rc.right - rc.left
                mon_h = rc.bottom - rc.top
            except Exception:
                pass

        r = tk.Tk()
        self._root = r
        r.overrideredirect(True)
        r.attributes("-topmost", True)
        r.attributes("-alpha", 0.95)
        r.configure(bg="#1a3c5a")

        frm = tk.Frame(r, bg="#1a3c5a", padx=60, pady=44)
        frm.pack(fill="both", expand=True)

        tk.Label(frm, text="\u26a1 Smart Schematic v2.1", font=("Segoe UI", 28, "bold"),
                 fg="#3cb4e8", bg="#1a3c5a").pack(anchor="w")
        tk.Label(frm, text="Initial setup \u2014 please don't touch Allegro",
                 font=("Segoe UI", 16), fg="#ffffff", bg="#1a3c5a"
                 ).pack(anchor="w", pady=(10, 24))

        style = ttk.Style(r)
        style.theme_use("default")
        style.configure("blue.Horizontal.TProgressbar",
                        troughcolor="#0d2840", background="#3cb4e8",
                        thickness=30)
        self._pbar = ttk.Progressbar(frm, length=640, mode="determinate",
                                     maximum=len(self._STEPS),
                                     style="blue.Horizontal.TProgressbar")
        self._pbar.pack(fill="x", pady=(0, 4))

        self._pct_label = tk.Label(frm, text="0%",
                                   font=("Segoe UI", 12, "bold"), fg="#ffffff",
                                   bg="#1a3c5a", anchor="e")
        self._pct_label.pack(fill="x", pady=(0, 10))

        self._step_label = tk.Label(frm, text=self._STEPS[0],
                                    font=("Segoe UI", 14), fg="#88bbdd",
                                    bg="#1a3c5a", anchor="w")
        self._step_label.pack(fill="x")

        r.update_idletasks()
        w = max(r.winfo_reqwidth(), 740)
        h = r.winfo_reqheight()
        if mon_cx is not None:
            x = mon_cx - w // 2
            y = mon_cy - h // 2
        else:
            sx, sy = r.winfo_screenwidth(), r.winfo_screenheight()
            x = (sx - w) // 2
            y = (sy - h) // 2
        r.geometry(f"{w}x{h}+{x}+{y}")

        # Auto-creep: advance bar with deceleration so it never looks stuck
        max_val = float(len(self._STEPS))
        def _tick():
            if self._closing:
                return
            # Decelerate: always move 3% of remaining distance
            remaining = max_val - self._cur_val
            increment = max(remaining * 0.03, 0.005)
            target = self._cur_val + increment
            if target < max_val and target > self._cur_val:
                self._cur_val = target
                pct = int(self._cur_val / max_val * 100)
                try:
                    self._pbar.configure(value=self._cur_val)
                    self._pct_label.configure(text=f"{pct}%")
                except Exception:
                    pass
            r.after(300, _tick)
        r.after(300, _tick)

        self._ready.set()
        r.mainloop()

    def step(self, idx):
        """Advance progress bar to step idx (0-based). Never goes backward."""
        if self._root:
            try:
                new_val = float(idx + 1)
                if new_val <= self._cur_val:
                    return  # don't go backward
                self._cur_val = new_val
                label = self._STEPS[idx] if idx < len(self._STEPS) else "Finishing\u2026"
                pct = int(self._cur_val / len(self._STEPS) * 100)
                self._root.after(0, lambda: (
                    self._pbar.configure(value=self._cur_val),
                    self._pct_label.configure(text=f"{pct}%"),
                    self._step_label.configure(text=label),
                ))
            except Exception:
                pass

    def close(self):
        if self._root:
            try:
                self._closing = True
                self._root.after(0, self._root.quit)
            except Exception:
                pass

# ---- Win32 constants ----
WM_LBUTTONDOWN = 0x0201
WM_LBUTTONUP = 0x0202
WM_LBUTTONDBLCLK = 0x0203
MK_LBUTTON = 0x0001

# ---- AutomationId constants ----

_COMP_FILTER_AUTOID = (
    "qt_scrollarea_viewport.ADTableContainerWidget.ADTableFilterLineEdit"
)
_COMP_TABLE_AUTOID = (
    "qt_scrollarea_viewport.ADTableContainerWidget."
    "ADTableTab.qt_tabwidget_stackedwidget.ComponentsTable"
)
_ALL_OFF_AUTOID = (
    "VisibilityTabWidget.qt_tabwidget_stackedwidget."
    "VisibilityLayersWidget.VisibilityLayersViewsWidget.All_Off"
)
_LAYER_TABLE_AUTOID = (
    "VisLayersObjectsAccordion.VisibilityLayersObjArWidget.VisLayerTable"
)
_OBJECTS_ACCORDION_AUTOID = "VisLayersObjectsAccordion"
_LAYERS_VSCROLL_AUTOID = (
    "VisibilityTabWidget.qt_tabwidget_stackedwidget."
    "VisibilityLayersWidget.VisLayersScrollArea."
    "qt_scrollarea_vcontainer.VisVerticalSlider"
)
_LAYERS_VIEWPORT_AUTOID = (
    "VisibilityTabWidget.qt_tabwidget_stackedwidget."
    "VisibilityLayersWidget.VisLayersScrollArea.qt_scrollarea_viewport"
)
_VIEW_FROM_AUTOID = (
    "VisibilityTabWidget.qt_tabwidget_stackedwidget."
    "VisDispScrollArea.qt_scrollarea_viewport."
    "VisibilityDisplayWidget.VisDispGeneralAccordion."
    "DisplayGeneral.DisplayGeneralViewFrom"
)
_BRIGHT_SLIDER_AUTOID = (
    "VisibilityTabWidget.qt_tabwidget_stackedwidget."
    "VisDispScrollArea.qt_scrollarea_viewport."
    "VisibilityDisplayWidget.VisDispShadowAccordion."
    "DisplayShadow.OriginBrightnessSlider"
)
_GLOBAL_OPACITY_SLIDER_AUTOID = (
    "VisibilityTabWidget.qt_tabwidget_stackedwidget."
    "VisDispScrollArea.qt_scrollarea_viewport."
    "VisibilityDisplayWidget.VisDispOpacityAccordion."
    "DisplayOpacity.OriginGlobalOpacitySlider"
)

_LAYER_ROW = {"Top": 4, "Bottom": 17}
_PIN_COL = 2
_ETCH_COL = 4  # Pins=2, Vias=3, Traces=4, Shapes=5, Text=6, DRC=7


# ---- Panel detection helper ----

def _find_left_panel(main_hwnd):
    """Find the left visibility panel via Win32 child window enumeration.
    Returns (left, top, right, bottom) or None. Fast (~1ms), no UIA."""
    wr = win32gui.GetWindowRect(main_hwnd)
    wl = wr[0]
    candidates = []
    def cb(h, _):
        if win32gui.IsWindowVisible(h):
            r = win32gui.GetWindowRect(h)
            w = r[2] - r[0]
            ht = r[3] - r[1]
            if w < 600 and ht > 400 and 0 <= (r[0] - wl) < 50:
                candidates.append(r)
        return True
    win32gui.EnumChildWindows(main_hwnd, cb, None)
    if candidates:
        candidates.sort(key=lambda r: r[3] - r[1], reverse=True)
        return candidates[0]
    return None


# ---- PostMessage helpers ----

def _find_allegro():
    """Find the Allegro main window with BRD loaded, return UIA control.
    If _brd_name_filter is set, only match windows with that BRD filename."""
    best_hwnd = None
    best_area = 0
    brd_filter = _brd_name_filter.lower()
    def cb(h, _):
        nonlocal best_hwnd, best_area
        if win32gui.IsWindowVisible(h):
            title = win32gui.GetWindowText(h).lower()
            if "allegro" in title and "viewer" in title and ".brd" in title:
                # If filtering by BRD name, skip windows that don't match
                if brd_filter and brd_filter not in title:
                    return True
                r = win32gui.GetWindowRect(h)
                a = (r[2] - r[0]) * (r[3] - r[1])
                if a > best_area:
                    best_hwnd = h
                    best_area = a
        return True
    win32gui.EnumWindows(cb, None)
    if not best_hwnd:
        return None
    return auto.ControlFromHandle(best_hwnd)


def _find_floating_panel(main_hwnd):
    """Find the offscreen floating Components panel HWND."""
    panel = None
    def cb(h, _):
        nonlocal panel
        if win32gui.IsWindowVisible(h):
            r = win32gui.GetWindowRect(h)
            if r[1] < 0:
                panel = h
        return True
    win32gui.EnumChildWindows(main_hwnd, cb, None)
    return panel


def _find_child_hwnd_at(parent_hwnd, screen_x, screen_y):
    """Find the smallest visible child HWND containing screen point."""
    best_hwnd = parent_hwnd
    best_area = float("inf")
    def cb(h, _):
        nonlocal best_hwnd, best_area
        if win32gui.IsWindowVisible(h):
            r = win32gui.GetWindowRect(h)
            if r[0] <= screen_x <= r[2] and r[1] <= screen_y <= r[3]:
                a = (r[2] - r[0]) * (r[3] - r[1])
                if a < best_area:
                    best_area = a
                    best_hwnd = h
        return True
    win32gui.EnumChildWindows(parent_hwnd, cb, None)
    return best_hwnd


def _post_click(parent_hwnd, screen_x, screen_y):
    """Send a mouse click via PostMessage at the given screen coordinates."""
    target = _find_child_hwnd_at(parent_hwnd, screen_x, screen_y)
    cpt = ctypes.wintypes.POINT(0, 0)
    user32.ClientToScreen(target, ctypes.byref(cpt))
    lx = screen_x - cpt.x
    ly = screen_y - cpt.y
    lp = (ly << 16) | (lx & 0xFFFF)
    win32gui.PostMessage(target, WM_LBUTTONDOWN, MK_LBUTTON, lp)
    time.sleep(0.05)
    win32gui.PostMessage(target, WM_LBUTTONUP, 0, lp)


def _post_dblclick(target_hwnd, screen_x, screen_y):
    """Send a double-click via PostMessage to a specific HWND."""
    cpt = ctypes.wintypes.POINT(0, 0)
    user32.ClientToScreen(target_hwnd, ctypes.byref(cpt))
    lx = screen_x - cpt.x
    ly = screen_y - cpt.y
    lp = (ly << 16) | (lx & 0xFFFF)
    win32gui.PostMessage(target_hwnd, WM_LBUTTONDOWN, MK_LBUTTON, lp)
    time.sleep(0.05)
    win32gui.PostMessage(target_hwnd, WM_LBUTTONUP, 0, lp)
    time.sleep(0.05)
    win32gui.PostMessage(target_hwnd, WM_LBUTTONDBLCLK, MK_LBUTTON, lp)
    time.sleep(0.05)
    win32gui.PostMessage(target_hwnd, WM_LBUTTONUP, 0, lp)


def _post_key(main_hwnd, vk_code):
    """Send a key press to the Allegro canvas via real mouse click + keybd_event."""
    KEYEVENTF_KEYUP = 0x0002
    MOUSEEVENTF_LEFTDOWN = 0x0002
    MOUSEEVENTF_LEFTUP = 0x0004
    # Get canvas center
    wr = win32gui.GetWindowRect(main_hwnd)
    cx = (wr[0] + wr[2]) // 2
    cy = (wr[1] + wr[3]) // 2
    # Bring window to foreground
    win32gui.SetForegroundWindow(main_hwnd)
    time.sleep(0.2)
    # Real mouse click on canvas to shift focus from command bar
    user32.SetCursorPos(cx, cy)
    time.sleep(0.1)
    user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(0.05)
    user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    time.sleep(0.2)
    # Send key via keybd_event (reliable for accelerator keys)
    scan = user32.MapVirtualKeyW(vk_code, 0)
    user32.keybd_event(vk_code, scan, 0, 0)
    time.sleep(0.05)
    user32.keybd_event(vk_code, scan, KEYEVENTF_KEYUP, 0)


# ---- navigation steps ----

def _switch_tab(allegro, main_hwnd, tab_index):
    """Switch left-panel tab. 0=Visibility, 2=Display."""
    try:
        vis_widget = allegro.GroupControl(
            AutomationId="VisibilityTabWidget", timeout=1
        )
        tab_ctrl = vis_widget.TabControl(timeout=1)
        tabs = tab_ctrl.GetChildren()
        if tab_index < len(tabs):
            br = tabs[tab_index].BoundingRectangle
            _post_click(main_hwnd,
                        (br.left + br.right) // 2,
                        (br.top + br.bottom) // 2)
            time.sleep(0.15)
    except Exception as e:
        print(f"tab_err:{e}")


def _expand_accordion(allegro, main_hwnd):
    """Expand the layers/objects accordion if collapsed."""
    try:
        acc = allegro.GroupControl(
            AutomationId=_OBJECTS_ACCORDION_AUTOID, timeout=1
        )
        br = acc.BoundingRectangle
        if br.bottom - br.top <= 30:
            try:
                label = acc.TextControl(
                    AutomationId="AccordionHeaderLabel", timeout=1
                )
                lbr = label.BoundingRectangle
                cx = (lbr.left + lbr.right) // 2
                cy = (lbr.top + lbr.bottom) // 2
            except Exception:
                cx = br.left + 14
                cy = br.top + 10
            _post_click(main_hwnd, cx, cy)
            time.sleep(0.4)
    except Exception:
        pass


def set_display_defaults(allegro, main_hwnd):
    """Set brightness and opacity sliders to 100%."""
    _switch_tab(allegro, main_hwnd, 2)
    for aid in (_BRIGHT_SLIDER_AUTOID, _GLOBAL_OPACITY_SLIDER_AUTOID):
        try:
            slider = allegro.SliderControl(AutomationId=aid, timeout=1)
            br = slider.BoundingRectangle
            if br.right - br.left > 0:
                _post_click(main_hwnd, br.right - 3,
                            (br.top + br.bottom) // 2)
                time.sleep(0.15)
        except Exception:
            pass
    _switch_tab(allegro, main_hwnd, 0)
    print("disp_ok")


def set_view_from(allegro, main_hwnd, target, skip_tab_back=False):
    """Set View From to Top or Bottom."""
    _switch_tab(allegro, main_hwnd, 2)
    try:
        vf = allegro.GroupControl(AutomationId=_VIEW_FROM_AUTOID, timeout=1)
        br = vf.BoundingRectangle
        w = br.right - br.left
        h = br.bottom - br.top
        if target == "Top":
            x = br.left + w // 4
        else:
            x = br.left + 3 * w // 4
        _post_click(main_hwnd, x, br.top + h // 2)
        time.sleep(0.2)
        print(f"view_{target}")
    except Exception as e:
        print(f"view_err:{e}")
    finally:
        if not skip_tab_back:
            _switch_tab(allegro, main_hwnd, 0)


def _open_command_bar_menu(allegro):
    """Ensure command bar is open via View > Command menu (idempotent).
    Checks if menu item is already checked — won't toggle OFF.
    Returns True if command bar is/was open."""
    import time as _t
    _t0 = _t.perf_counter()
    try:
        menubar = allegro.MenuBarControl(ClassName="SPBMenuBar", timeout=2)
        print(f"M:menubar={_t.perf_counter()-_t0:.2f}s", flush=True)
        for item in menubar.GetChildren():
            if item.Name == "View":
                item.GetExpandCollapsePattern().Expand()
                time.sleep(0.3)
                break
        print(f"M:expand={_t.perf_counter()-_t0:.2f}s", flush=True)

        popup_hwnd = None
        def find_popup(h, _):
            nonlocal popup_hwnd
            if win32gui.IsWindowVisible(h):
                if "DropShadow" in win32gui.GetClassName(h):
                    popup_hwnd = h
            return True
        win32gui.EnumWindows(find_popup, None)
        print(f"M:popup={_t.perf_counter()-_t0:.2f}s", flush=True)

        if popup_hwnd:
            menu_ctrl = auto.ControlFromHandle(popup_hwnd)
            for mi in menu_ctrl.GetChildren():
                if mi.Name == "Command":
                    # Check if already checked (command bar already visible)
                    already_on = False
                    try:
                        tp = mi.GetTogglePattern()
                        already_on = (tp.ToggleState == 1)  # On
                    except Exception:
                        pass
                    if not already_on:
                        try:
                            lap = mi.GetLegacyIAccessiblePattern()
                            already_on = bool(lap.State & 0x10)  # CHECKED
                        except Exception:
                            pass
                    if already_on:
                        # Already on — close menu without toggling
                        win32gui.PostMessage(
                            allegro.NativeWindowHandle, 0x0100, 0x1B, 0)
                        time.sleep(0.2)
                        print(f"M:already_on={_t.perf_counter()-_t0:.2f}s",
                              flush=True)
                        return True
                    mi.GetInvokePattern().Invoke()
                    time.sleep(0.3)
                    print(f"M:invoke={_t.perf_counter()-_t0:.2f}s", flush=True)
                    return True
    except Exception as e:
        print(f"M:err:{e} {_t.perf_counter()-_t0:.2f}s", flush=True)
    return False


def _force_toggle_view_command(allegro):
    """Force command bar ON by toggling View > Command off then on.
    Handles case where search bar has replaced command bar."""
    try:
        menubar = allegro.MenuBarControl(ClassName="SPBMenuBar", timeout=2)
        view_item = None
        for item in menubar.GetChildren():
            if item.Name == "View":
                view_item = item
                break
        if not view_item:
            return False
        # Click Command twice: first click toggles off, second toggles on
        for i in range(2):
            view_item.GetExpandCollapsePattern().Expand()
            time.sleep(0.3)
            popup_hwnd = None
            def find_popup(h, _):
                nonlocal popup_hwnd
                if win32gui.IsWindowVisible(h):
                    if "DropShadow" in win32gui.GetClassName(h):
                        popup_hwnd = h
                return True
            win32gui.EnumWindows(find_popup, None)
            if popup_hwnd:
                menu_ctrl = auto.ControlFromHandle(popup_hwnd)
                for mi in menu_ctrl.GetChildren():
                    if mi.Name == "Command":
                        mi.GetInvokePattern().Invoke()
                        time.sleep(0.3)
                        break
            print(f"force_toggle_{i+1}", flush=True)
        return True
    except Exception as e:
        print(f"force_toggle_err:{e}", flush=True)
        return False


def _get_command_bar(allegro, main_hwnd, skip_initial_check=False):
    """Find or open the Allegro command bar.
    Returns (combo_center_pos, combo_hwnd) or (None, None).
    Gets position + hwnd immediately while Allegro is responsive.
    Handles search bar replacing command bar by force-toggling via menu.
    """
    import time as _t
    _COMBO_AID = "qt_scrollarea_viewport.CommmandPane.SPBCommandLine"

    combo = None
    if not skip_initial_check:
        try:
            combo = allegro.ComboBoxControl(AutomationId=_COMBO_AID, timeout=1)
            if combo.BoundingRectangle.right <= 0:
                combo = None
        except Exception:
            combo = None

    if not combo:
        # Open via View > Command menu
        _open_command_bar_menu(allegro)
        try:
            combo = allegro.ComboBoxControl(AutomationId=_COMBO_AID, timeout=2)
            if combo.BoundingRectangle.right <= 0:
                combo = None
        except Exception:
            combo = None

    if not combo:
        # Force toggle: turn Command off then on (handles search bar replacing it)
        print("cmd_force_toggle", flush=True)
        _force_toggle_view_command(allegro)
        try:
            combo = allegro.ComboBoxControl(AutomationId=_COMBO_AID, timeout=2)
        except Exception:
            return None, None

    # Get position via BoundingRectangle (may block on cold start)
    try:
        cbr = combo.BoundingRectangle
        cx = (cbr.left + cbr.right) // 2
        cy = (cbr.top + cbr.bottom) // 2
        pt = ctypes.wintypes.POINT(cx, cy)
        hwnd = user32.WindowFromPoint(pt)
        return (cx, cy), hwnd
    except Exception:
        return None, None


def _run_command(allegro, main_hwnd, cmd_text):
    """Run a command in Allegro command bar via keyboard typing + PostMessage Enter."""
    pos, chw = _get_command_bar(allegro, main_hwnd)
    if not pos:
        print("cmd_err:no_combo")
        return False
    try:
        return _type_command(main_hwnd, chw, pos, cmd_text)
    except Exception as e:
        print(f"cmd_err:{e}")
        return False


def highlight_refdes_cmd(allegro, main_hwnd, refdes):
    """Highlight + zoom via 'Symbol REFDES' command — fast."""
    if _run_command(allegro, main_hwnd, f"symbol {refdes}"):
        print("hl_ok")
    else:
        print("hl_err:cmd_failed")


def set_pin_layer(allegro, main_hwnd, target, accordion_expanded=False):
    """All Off + enable pin on target layer."""
    try:
        # Click All Off
        off_btn = allegro.ButtonControl(AutomationId=_ALL_OFF_AUTOID, timeout=1)
        obr = off_btn.BoundingRectangle
        _post_click(main_hwnd,
                    (obr.left + obr.right) // 2,
                    (obr.top + obr.bottom) // 2)
        time.sleep(0.2)

        if not accordion_expanded:
            _expand_accordion(allegro, main_hwnd)

        # Click pin cell
        lt = allegro.TableControl(AutomationId=_LAYER_TABLE_AUTOID, timeout=1)
        gp = lt.GetGridPattern()
        row_idx = _LAYER_ROW.get(target, 4)
        cell = gp.GetItem(row_idx, _PIN_COL)
        r = cell.BoundingRectangle
        if r.right - r.left > 0:
            _post_click(main_hwnd,
                        (r.left + r.right) // 2,
                        (r.top + r.bottom) // 2)
            time.sleep(0.15)
            print(f"pin_{target}")
    except Exception as e:
        print(f"pin_err:{e}")


def turn_off_grid(allegro, main_hwnd):
    """Turn off grid dots via command bar."""
    if _run_command(allegro, main_hwnd, "setwindow pcb -nogrid"):
        print("grid_off")
    else:
        print("grid_err:cmd_failed")


# ---- Cached fast-path (persistent worker only) ----

_cache = {}


def _build_cache(allegro, main_hwnd, popup=None):
    """Discover UI element positions (one-time cost).
    If Allegro is unresponsive (slow UIA), builds minimal cache for command only.
    """
    import time as _t
    _t0 = _t.perf_counter()
    c = {}
    c["win_rect"] = win32gui.GetWindowRect(main_hwnd)

    # --- Visibility tab: tabs, All Off ---
    try:
        vis_widget = allegro.GroupControl(
            AutomationId="VisibilityTabWidget", timeout=2)
        tab_ctrl = vis_widget.TabControl(timeout=2)
        tabs = tab_ctrl.GetChildren()
        for i, tab in enumerate(tabs):
            br = tab.BoundingRectangle
            c[f"tab_{i}"] = ((br.left + br.right) // 2,
                             (br.top + br.bottom) // 2)
    except Exception as e:
        print(f"cache_tab_err:{e}", flush=True)
    tabs_elapsed = _t.perf_counter() - _t0
    print(f"B:tabs={tabs_elapsed:.2f}s", flush=True)

    # Build minimal cache on first click for fast response (>3s for tabs)
    # Full cache builds on the next click
    if tabs_elapsed > 8.0:
        print("B:slow_allegro, minimal cache", flush=True)
        c["slow_start"] = True
        # Open command bar via menu (will be slow but necessary)
        if _open_command_bar_menu(allegro):
            c["cmd_bar_opened"] = True
            print(f"cmd_bar_opened B:cmdbar_minimal={_t.perf_counter()-_t0:.2f}s",
                  flush=True)
        print(f"cache_built_minimal({len(c)}) total={_t.perf_counter()-_t0:.2f}s",
              flush=True)
        return c

    # Ensure Visibility tab is active
    if "tab_0" in c:
        _post_click(main_hwnd, *c["tab_0"])
        time.sleep(0.2)

    # All Off button — find and CLICK immediately to speed up rendering
    _t1 = _t.perf_counter()
    try:
        off_btn = allegro.ButtonControl(
            AutomationId=_ALL_OFF_AUTOID, timeout=2)
        br = off_btn.BoundingRectangle
        c["all_off"] = ((br.left + br.right) // 2,
                        (br.top + br.bottom) // 2)
        # Click All Off now to reduce layer rendering load
        _post_click(main_hwnd, *c["all_off"])
        time.sleep(0.2)
    except Exception:
        pass
    print(f"B:alloff={_t.perf_counter()-_t1:.2f}s", flush=True)

    # --- Display tab: View From, sliders ---
    _t1 = _t.perf_counter()
    if "tab_2" in c:
        _post_click(main_hwnd, *c["tab_2"])
        time.sleep(0.2)

    # View From control
    try:
        vf = allegro.GroupControl(
            AutomationId=_VIEW_FROM_AUTOID, timeout=2)
        br = vf.BoundingRectangle
        w = br.right - br.left
        h = br.bottom - br.top
        c["vf_Top"] = (br.left + w // 4, br.top + h // 2)
        c["vf_Bottom"] = (br.left + 3 * w // 4, br.top + h // 2)
    except Exception:
        pass
    print(f"B:viewfrom={_t.perf_counter()-_t1:.2f}s", flush=True)

    # Set sliders to 100%
    _t1 = _t.perf_counter()
    for aid in (_BRIGHT_SLIDER_AUTOID, _GLOBAL_OPACITY_SLIDER_AUTOID):
        try:
            slider = allegro.SliderControl(AutomationId=aid, timeout=1)
            br = slider.BoundingRectangle
            if br.right - br.left > 0:
                _post_click(main_hwnd, br.right - 3,
                            (br.top + br.bottom) // 2)
                time.sleep(0.1)
        except Exception:
            pass
    print(f"B:sliders={_t.perf_counter()-_t1:.2f}s", flush=True)

    # --- Back to Visibility tab ---
    if "tab_0" in c:
        _post_click(main_hwnd, *c["tab_0"])
        time.sleep(0.15)

    # --- Cache accordion + layer table + PIN CELL POSITIONS ---
    _t1 = _t.perf_counter()
    try:
        acc = allegro.GroupControl(
            AutomationId=_OBJECTS_ACCORDION_AUTOID, timeout=2)
        br = acc.BoundingRectangle
        if br.bottom - br.top <= 30:
            try:
                label = acc.TextControl(
                    AutomationId="AccordionHeaderLabel", timeout=1)
                lbr = label.BoundingRectangle
                cx = (lbr.left + lbr.right) // 2
                cy = (lbr.top + lbr.bottom) // 2
            except Exception:
                cx = br.left + 14
                cy = br.top + 10
            c["accordion_hdr"] = (cx, cy)
            _post_click(main_hwnd, cx, cy)
            time.sleep(0.3)
        else:
            c["accordion_hdr"] = (br.left + 14, br.top + 10)
        c["accordion_uia"] = acc

        lt = allegro.TableControl(
            AutomationId=_LAYER_TABLE_AUTOID, timeout=2)
        c["layer_table_uia"] = lt
        gp = lt.GetGridPattern()
        row_count = gp.RowCount
        col_count = gp.ColumnCount
        print(f"layer_table: {row_count}r x {col_count}c", flush=True)

        # Fast layer row detection: verify hardcoded rows, scan only if wrong
        layer_rows = dict(_LAYER_ROW)  # {"Top": 4, "Bottom": 17}
        try:
            # Verify Top at expected row
            top_cell = gp.GetItem(_LAYER_ROW["Top"], 0)
            if (top_cell.Name or "").strip().upper() != "TOP":
                for r in range(min(row_count, 8)):
                    n = (gp.GetItem(r, 0).Name or "").strip().upper()
                    if n == "TOP":
                        layer_rows["Top"] = r
                        break
            # Verify Bottom at expected row
            bot_ok = False
            if _LAYER_ROW["Bottom"] < row_count:
                bot_cell = gp.GetItem(_LAYER_ROW["Bottom"], 0)
                bot_ok = (bot_cell.Name or "").strip().upper() == "BOTTOM"
            if not bot_ok:
                for r in range(row_count - 1, max(row_count - 8, 0), -1):
                    n = (gp.GetItem(r, 0).Name or "").strip().upper()
                    if n == "BOTTOM":
                        layer_rows["Bottom"] = r
                        break
            print(f"layer_rows_detected={layer_rows}", flush=True)
        except Exception:
            pass
        c["_layer_rows"] = layer_rows

        # Cache viewport for scroll-safe clicking later
        try:
            vp = allegro.GroupControl(
                AutomationId=_LAYERS_VIEWPORT_AUTOID, timeout=0.5)
            c["_layers_viewport_uia"] = vp
        except Exception:
            pass

        # Cache pin AND etch cell positions directly (skip slow scrollbar search)
        _need_scroll = False
        _vp_build = c.get("_layers_viewport_uia")
        for view_name, row_idx in layer_rows.items():
            try:
                cell = gp.GetItem(row_idx, _PIN_COL)
                r = cell.BoundingRectangle
                if r.right - r.left > 0:
                    c[f"pin_{view_name}"] = ((r.left + r.right) // 2,
                                             (r.top + r.bottom) // 2)
                    # Check if cell is outside viewport (needs scroll later)
                    if _vp_build:
                        vr = _vp_build.BoundingRectangle
                        cy = (r.top + r.bottom) // 2
                        if cy < vr.top or cy > vr.bottom:
                            _need_scroll = True
                else:
                    _need_scroll = True
                cell_e = gp.GetItem(row_idx, _ETCH_COL)
                r_e = cell_e.BoundingRectangle
                if r_e.right - r_e.left > 0:
                    c[f"etch_{view_name}"] = ((r_e.left + r_e.right) // 2,
                                               (r_e.top + r_e.bottom) // 2)
            except Exception:
                pass

        # Only search for scrollbar if some cells were off-screen
        if _need_scroll:
            try:
                vsb = acc.ScrollBarControl(
                    AutomationId=_LAYERS_VSCROLL_AUTOID, timeout=0.3)
                _rvp_build = vsb.GetPattern(auto.PatternId.RangeValuePattern)
                if _rvp_build:
                    c["_layers_vscroll_rvp"] = _rvp_build
                    # Re-cache cells that were off-screen by scrolling
                    for view_name, row_idx in layer_rows.items():
                        key_p = f"pin_{view_name}"
                        if key_p not in c:
                            if view_name == "Top":
                                _rvp_build.SetValue(_rvp_build.Minimum)
                            else:
                                _rvp_build.SetValue(_rvp_build.Maximum)
                            time.sleep(0.05)
                            cell = gp.GetItem(row_idx, _PIN_COL)
                            r = cell.BoundingRectangle
                            if r.right - r.left > 0:
                                c[key_p] = ((r.left + r.right) // 2,
                                            (r.top + r.bottom) // 2)
                            cell_e = gp.GetItem(row_idx, _ETCH_COL)
                            r_e = cell_e.BoundingRectangle
                            if r_e.right - r_e.left > 0:
                                c[f"etch_{view_name}"] = (
                                    (r_e.left + r_e.right) // 2,
                                    (r_e.top + r_e.bottom) // 2)
                    _rvp_build.SetValue(_rvp_build.Minimum)
            except Exception:
                pass
        print(f"pin_cached={[k for k in c if k.startswith('pin_')]}", flush=True)
        print(f"etch_cached={[k for k in c if k.startswith('etch_')]}", flush=True)
    except Exception as e:
        print(f"pin_cache_err:{e}", flush=True)
    except Exception as e:
        print(f"pin_cache_err:{e}", flush=True)
    print(f"B:accordion+table={_t.perf_counter()-_t1:.2f}s", flush=True)

    # Command bar: detect by checking window HEIGHT at expected position
    # Command bar combo is narrow (~25-35px), canvas/panels are tall (>100px)
    _t1 = _t.perf_counter()
    wr = c["win_rect"]
    approx_x = (wr[0] + wr[2]) // 2
    y_cmd = wr[3] - 72
    pt_cmd = ctypes.wintypes.POINT(approx_x, y_cmd)
    child_cmd = user32.WindowFromPoint(pt_cmd)
    cmd_bar_found = False
    if child_cmd and child_cmd != main_hwnd:
        try:
            rect = win32gui.GetWindowRect(child_cmd)
            child_h = rect[3] - rect[1]
            child_w = rect[2] - rect[0]
            print(f"cmdbar_probe: hwnd={child_cmd} h={child_h} w={child_w} rect={rect}",
                  flush=True)
            # Command bar combo is narrow (height < 60px) and wide (> 200px)
            if child_h < 60 and child_w > 200:
                c["cmd_combo_pos"] = (approx_x, y_cmd)
                c["cmd_combo_hwnd"] = child_cmd
                cmd_bar_found = True
                print(f"cmd_bar_detected=({approx_x},{y_cmd}) B:cmdbar={_t.perf_counter()-_t1:.2f}s",
                      flush=True)
        except Exception:
            pass
    if not cmd_bar_found:
        # Command bar not open — open via menu
        if _open_command_bar_menu(allegro):
            c["cmd_bar_opened"] = True
            print(f"cmd_bar_opened B:cmdbar={_t.perf_counter()-_t1:.2f}s", flush=True)
        else:
            print(f"cmd_bar_failed B:cmdbar={_t.perf_counter()-_t1:.2f}s", flush=True)

    # Save panel rect, DPI, and original positions for cross-monitor correction
    panel = _find_left_panel(main_hwnd)
    if panel:
        c["_panel_rect"] = panel
        print(f"panel_rect={panel} ({panel[2]-panel[0]}x{panel[3]-panel[1]})", flush=True)
    try:
        c["_dpi"] = ctypes.windll.user32.GetDpiForWindow(main_hwnd)
    except Exception:
        c["_dpi"] = 0
    # Snapshot panel-relative positions for DPI correction later
    if panel:
        orig = {}
        skip = {"win_rect", "cmd_combo_pos", "_panel_rect", "_dpi", "_orig_positions",
                "accordion_uia", "layer_table_uia", "_layers_vscroll_rvp",
                "_layers_viewport_uia", "_layer_rows", "_force_layer_reapply",
                "cmd_bar_opened", "cmd_combo_hwnd", "cmd_target_hwnd", "slow_start"}
        for k, v in c.items():
            if k in skip or k.startswith("_"):
                continue
            if isinstance(v, tuple) and len(v) == 2 and \
               all(isinstance(x, (int, float)) for x in v):
                orig[k] = (v[0] - panel[0], v[1] - panel[1])
        c["_orig_positions"] = orig

    print(f"cache_built({len(c)}) keys={[k for k in c if k not in ('accordion_uia','layer_table_uia')]}",
          flush=True)
    return c


def _is_cache_valid(main_hwnd):
    """Check if Allegro window hasn't moved/resized (cache still usable).
    On move: offset all positions. On resize: detect DPI change and apply corrections.
    """
    if not _cache or "win_rect" not in _cache:
        return False
    try:
        new_rect = win32gui.GetWindowRect(main_hwnd)
        old_rect = _cache["win_rect"]
        if new_rect == old_rect:
            return True

        old_size = (old_rect[2] - old_rect[0], old_rect[3] - old_rect[1])
        new_size = (new_rect[2] - new_rect[0], new_rect[3] - new_rect[1])
        is_resize = old_size != new_size

        # Check for DPI change (cross-monitor move)
        orig_dpi = _cache.get("_dpi", 0)
        try:
            new_dpi = ctypes.windll.user32.GetDpiForWindow(main_hwnd)
        except Exception:
            new_dpi = 0
        dpi_changed = orig_dpi and new_dpi and orig_dpi != new_dpi

        if dpi_changed and is_resize and _cache.get("_orig_positions"):
            # Cross-monitor: recompute from panel-relative originals + DPI corrections
            time.sleep(0.1)  # let Qt settle
            new_panel = _find_left_panel(main_hwnd)
            if new_panel:
                orig_pos = _cache["_orig_positions"]
                dpi_diff = new_dpi - orig_dpi
                table_y_corr = round(dpi_diff * 0.604)
                x_corr_left = round(dpi_diff * 0.32)
                x_corr_right = round(dpi_diff * 0.14)
                _X_LEFT = {"vf_Top", "accordion_hdr"}
                _X_RIGHT = {"all_off", "vf_Bottom"}

                accordion_rel_y = orig_pos.get("accordion_hdr", (0, 0))[1]

                for k, rel in orig_pos.items():
                    x_corr = 0
                    if k in _X_LEFT:
                        x_corr = x_corr_left
                    elif k in _X_RIGHT:
                        x_corr = x_corr_right
                    new_x = new_panel[0] + rel[0] + x_corr
                    new_y = new_panel[1] + rel[1]
                    if table_y_corr and accordion_rel_y and rel[1] > accordion_rel_y:
                        new_y += table_y_corr
                    _cache[k] = (new_x, new_y)

                _cache["_panel_rect"] = new_panel
                _cache["_dpi"] = new_dpi
                print(f"cache_dpi_rescale: dpi={orig_dpi}->{new_dpi} "
                      f"panel={new_panel} tbl_y={table_y_corr} "
                      f"xL={x_corr_left} xR={x_corr_right} "
                      f"{old_size}->{new_size}", flush=True)
            else:
                # Fallback: simple offset
                dx = new_rect[0] - old_rect[0]
                dy = new_rect[1] - old_rect[1]
                for k, v in list(_cache.items()):
                    if k == "cmd_combo_pos" or k.startswith("_"):
                        continue
                    if isinstance(v, tuple) and len(v) == 2 and \
                       all(isinstance(x, (int, float)) for x in v):
                        _cache[k] = (v[0] + dx, v[1] + dy)
                print(f"cache_offset_fallback: dx={dx} dy={dy}", flush=True)
        else:
            # Same-monitor move/resize: simple offset
            dx = new_rect[0] - old_rect[0]
            dy = new_rect[1] - old_rect[1]
            for k, v in list(_cache.items()):
                if k == "cmd_combo_pos" or k.startswith("_"):
                    continue
                if isinstance(v, tuple) and len(v) == 2 and \
                   all(isinstance(x, (int, float)) for x in v):
                    _cache[k] = (v[0] + dx, v[1] + dy)
            if is_resize:
                print(f"cache_offset: dx={dx} dy={dy} RESIZE {old_size}->{new_size}",
                      flush=True)
            else:
                print(f"cache_offset: dx={dx} dy={dy} MOVE", flush=True)

        # Command bar: recalculate from new rect (center X, bottom - 72)
        new_cmd_x = (new_rect[0] + new_rect[2]) // 2
        new_cmd_y = new_rect[3] - 72
        _cache["cmd_combo_pos"] = (new_cmd_x, new_cmd_y)
        pt = ctypes.wintypes.POINT(new_cmd_x, new_cmd_y)
        new_hwnd = user32.WindowFromPoint(pt)
        if new_hwnd and new_hwnd != main_hwnd:
            _cache["cmd_combo_hwnd"] = new_hwnd
        _cache.pop("cmd_target_hwnd", None)

        # On RESIZE: force layer re-apply on next click
        if is_resize:
            _cache["_force_layer_reapply"] = True

        _cache["win_rect"] = new_rect
        return True
    except Exception:
        return False


def _scroll_layer_to(allegro_ctrl, layer_name, cell=None):
    """Scroll the layers panel to make the given layer visible.
    If cell is provided, scrolls by exact pixel offset.
    Otherwise falls back to min (Top) / max (Bottom).
    """
    rvp = _cache.get("_layers_vscroll_rvp")
    if not rvp:
        try:
            vsb = allegro_ctrl.ScrollBarControl(
                AutomationId=_LAYERS_VSCROLL_AUTOID, timeout=1)
            rvp = vsb.GetPattern(auto.PatternId.RangeValuePattern)
            if rvp:
                _cache["_layers_vscroll_rvp"] = rvp
        except Exception:
            pass
    if rvp:
        try:
            vp = _cache.get("_layers_viewport_uia")
            if cell and vp:
                vr = vp.BoundingRectangle
                cr = cell.BoundingRectangle
                cy = (cr.top + cr.bottom) // 2
                vp_center = (vr.top + vr.bottom) // 2
                new_val = rvp.Value + (cy - vp_center)
                new_val = max(rvp.Minimum, min(rvp.Maximum, new_val))
                rvp.SetValue(new_val)
            else:
                target_val = rvp.Minimum if layer_name == "Top" else rvp.Maximum
                rvp.SetValue(target_val)
            time.sleep(0.15)
        except Exception:
            pass


def _fast_click(main_hwnd, cache_key):
    """PostMessage click at a cached screen position."""
    pos = _cache.get(cache_key)
    if pos:
        _post_click(main_hwnd, pos[0], pos[1])
        return True
    return False


def _click_layer_cell(main_hwnd, cell_type, layer_name):
    """Click a pin/etch layer cell with scroll-safe operation.
    Checks if cell is within viewport; if not, scrolls by the exact
    pixel offset needed to center the cell in the viewport.
    """
    col = _PIN_COL if cell_type == "pin" else _ETCH_COL
    cache_key = f"{cell_type}_{layer_name}"
    row_idx = _cache.get("_layer_rows", _LAYER_ROW).get(layer_name)
    if row_idx is None:
        return _fast_click(main_hwnd, cache_key)
    lt = _cache.get("layer_table_uia")
    if lt is None:
        return _fast_click(main_hwnd, cache_key)
    try:
        import time as _t
        _t0 = _t.perf_counter()
        gp = lt.GetGridPattern()
        cell = gp.GetItem(row_idx, col)

        # Check if cell is within the visible viewport
        need_scroll = False
        vp = _cache.get("_layers_viewport_uia")
        if vp:
            try:
                vr = vp.BoundingRectangle
                cr = cell.BoundingRectangle
                cy = (cr.top + cr.bottom) // 2
                if cy < vr.top or cy > vr.bottom:
                    need_scroll = True
                    # Find scrollbar (cached or fresh)
                    rvp = _cache.get("_layers_vscroll_rvp")
                    if not rvp:
                        try:
                            allegro_ctrl = auto.ControlFromHandle(main_hwnd)
                            vsb = allegro_ctrl.ScrollBarControl(
                                AutomationId=_LAYERS_VSCROLL_AUTOID, timeout=1)
                            rvp = vsb.GetPattern(auto.PatternId.RangeValuePattern)
                            if rvp:
                                _cache["_layers_vscroll_rvp"] = rvp
                                print(f"vscroll_found_fresh: val={rvp.Value} max={rvp.Maximum}",
                                      flush=True)
                        except Exception:
                            pass
                    if rvp:
                        try:
                            # Calculate exact scroll offset to center cell in viewport
                            vp_center = (vr.top + vr.bottom) // 2
                            offset = cy - vp_center
                            new_val = rvp.Value + offset
                            new_val = max(rvp.Minimum, min(rvp.Maximum, new_val))
                            rvp.SetValue(new_val)
                            time.sleep(0.15)
                            print(f"vscroll: {rvp.Value:.0f} offset={offset} "
                                  f"cell_y={cy} vp=({vr.top},{vr.bottom})",
                                  flush=True)
                        except Exception as e:
                            print(f"vscroll_err:{e}", flush=True)
            except Exception:
                pass

        r = cell.BoundingRectangle
        if r.right - r.left > 0:
            cx, cy = (r.left + r.right) // 2, (r.top + r.bottom) // 2
            _post_click(main_hwnd, cx, cy)
            _cache[cache_key] = (cx, cy)
            print(f"layer_click:{cache_key} pos=({cx},{cy}) "
                  f"scrolled={need_scroll} ({_t.perf_counter()-_t0:.2f}s)", flush=True)
            return True
    except Exception as e:
        print(f"layer_cell_err({cache_key}):{e}", flush=True)
    return _fast_click(main_hwnd, cache_key)


def _fast_run_command(main_hwnd, cmd_text):
    """Run Allegro command. Tries cached position → direct typing → fresh search."""
    import time as _t
    _t0 = _t.perf_counter()

    # Path 1: cached combo position + keyboard typing (click to focus, then type)
    combo_pos = _cache.get("cmd_combo_pos")
    if combo_pos:
        need_refresh = False
        # After a recent recovery, re-verify the command bar is still valid
        if _cache.pop("cmd_needs_reverify", False):
            need_refresh = True
            print("cmd_reverify_check", flush=True)
        else:
            # Check if the command bar widget has been replaced (e.g., search bar)
            cached_hwnd = _cache.get("cmd_combo_hwnd")
            if cached_hwnd:
                pt = ctypes.wintypes.POINT(combo_pos[0], combo_pos[1])
                current_hwnd = user32.WindowFromPoint(pt)
                if current_hwnd != cached_hwnd:
                    # Check if current_hwnd is a child of our Allegro window
                    # If not, it's another app covering Allegro — re-focus and retry
                    parent = current_hwnd
                    is_allegro_child = False
                    for _ in range(10):
                        parent = user32.GetParent(parent)
                        if parent == 0:
                            break
                        if parent == main_hwnd:
                            is_allegro_child = True
                            break
                    if not is_allegro_child:
                        # Another window is in front — bring Allegro back
                        try:
                            win32gui.SetForegroundWindow(main_hwnd)
                            time.sleep(0.5)
                        except Exception:
                            pass
                        # Recheck after refocus
                        current_hwnd = user32.WindowFromPoint(pt)
                        if current_hwnd == cached_hwnd:
                            need_refresh = False
                        else:
                            need_refresh = True
                            print(f"cmd_bar_widget_changed: was={cached_hwnd} now={current_hwnd}",
                                  flush=True)
                    else:
                        # Hwnd changed but still an Allegro child (e.g., resize)
                        _cache["cmd_combo_hwnd"] = current_hwnd
                        _cache.pop("cmd_target_hwnd", None)
                        need_refresh = False
                        print(f"cmd_hwnd_updated: {cached_hwnd} -> {current_hwnd}",
                              flush=True)

        if need_refresh:
            # Quick verify: use ControlFromPoint to check what's at the position
            refreshed = False
            try:
                elem = auto.ControlFromPoint(combo_pos[0], combo_pos[1])
                is_cmd = False
                check = elem
                for _ in range(3):
                    aid = getattr(check, "AutomationId", "") or ""
                    if "CommmandPane" in aid:
                        is_cmd = True
                        break
                    check = check.GetParentControl()
                    if not check:
                        break
                if is_cmd:
                    new_hwnd = user32.WindowFromPoint(
                        ctypes.wintypes.POINT(combo_pos[0], combo_pos[1]))
                    _cache["cmd_combo_hwnd"] = new_hwnd
                    refreshed = True
                    print(f"cmd_quick_verify=ok hwnd={new_hwnd} "
                          f"{_t.perf_counter()-_t0:.2f}s", flush=True)
                else:
                    print(f"cmd_quick_verify=not_cmdbar "
                          f"aid={getattr(elem, 'AutomationId', '?')} "
                          f"{_t.perf_counter()-_t0:.2f}s", flush=True)
            except Exception as e:
                print(f"cmd_quick_verify_err:{e} {_t.perf_counter()-_t0:.2f}s",
                      flush=True)
            if not refreshed:
                _cache.pop("cmd_combo_pos", None)
                _cache.pop("cmd_combo_hwnd", None)
                _cache.pop("cmd_target_hwnd", None)
                combo_pos = None

        if combo_pos:
            try:
                r = _type_command(main_hwnd, None, combo_pos, cmd_text)
                print(f"cmd_path1={_t.perf_counter()-_t0:.2f}s", flush=True)
                return r
            except Exception as e:
                print(f"cmd_path1_err:{e} {_t.perf_counter()-_t0:.2f}s", flush=True)

    # Path 2: command bar already opened (during cache build) — type via PostMessage
    if _cache.get("cmd_bar_opened"):
        try:
            _cache.pop("cmd_bar_opened", None)  # only try once
            WM_CHAR = 0x0102
            WM_KEYDOWN = 0x0100
            WM_KEYUP = 0x0101
            wr = _cache.get("win_rect")
            if wr:
                approx_x = (wr[0] + wr[2]) // 2
                approx_y = wr[3] - 72
                target = _find_child_hwnd_at(main_hwnd, approx_x, approx_y)
                cpt = ctypes.wintypes.POINT(0, 0)
                user32.ClientToScreen(target, ctypes.byref(cpt))
                lx = approx_x - cpt.x
                ly = approx_y - cpt.y
                lp = (ly << 16) | (lx & 0xFFFF)
                # Triple-click to select all text in edit field
                for _ in range(3):
                    win32gui.PostMessage(target, WM_LBUTTONDOWN, 0x0001, lp)
                    win32gui.PostMessage(target, WM_LBUTTONUP, 0, lp)
                    time.sleep(0.05)
                time.sleep(0.15)
                # Type the command — prefix with space (absorbed by dropdown)
                win32gui.PostMessage(target, WM_CHAR, ord(' '), 0)
                time.sleep(0.1)
                # Now type real command
                for ch in cmd_text:
                    win32gui.PostMessage(target, WM_CHAR, ord(ch), 0)
                time.sleep(0.05)
                # Enter
                win32gui.PostMessage(target, WM_KEYDOWN, 0x0D, 0)
                time.sleep(0.02)
                win32gui.PostMessage(target, WM_KEYUP, 0x0D, 0)
                time.sleep(0.15)
                print(f"cmd_path2_direct={_t.perf_counter()-_t0:.2f}s", flush=True)

                # Cache estimated position + hwnd for future Path 1 clicks (no UIA)
                _cache["cmd_combo_pos"] = (approx_x, approx_y)
                _cache["cmd_combo_hwnd"] = target
                print(f"cmd_pos_cached=({approx_x},{approx_y})", flush=True)
                return True
        except Exception as e:
            print(f"cmd_path2_err:{e} {_t.perf_counter()-_t0:.2f}s", flush=True)

    # Path 3: find/open command bar from scratch (slow fallback)
    print(f"cmd_path3_start={_t.perf_counter()-_t0:.2f}s", flush=True)
    try:
        allegro = auto.ControlFromHandle(main_hwnd)
        pos, chw = _get_command_bar(allegro, main_hwnd)
        if not pos:
            print("cmd_err:no_combo", flush=True)
            return False
        _cache["cmd_combo_pos"] = pos
        _cache["cmd_combo_hwnd"] = chw
        r = _type_command(main_hwnd, chw, pos, cmd_text)
        print(f"cmd_path3={_t.perf_counter()-_t0:.2f}s", flush=True)
        return r
    except Exception as e:
        print(f"cmd_err:{e}", flush=True)
        return False


def _type_command(main_hwnd, combo_hwnd, combo_pos, cmd_text):
    """Type a command into the ComboBox via triple-click + space prefix + type."""
    WM_CHAR = 0x0102
    WM_KEYDOWN = 0x0100
    WM_KEYUP = 0x0101
    WM_LBUTTONDOWN = 0x0201
    WM_LBUTTONUP = 0x0202
    cx, cy = combo_pos
    # Use cached target hwnd if available and still valid, else re-discover
    target = None
    cached_target = _cache.get("cmd_target_hwnd")
    if cached_target and win32gui.IsWindow(cached_target):
        try:
            if win32gui.IsWindowVisible(cached_target):
                target = cached_target
        except Exception:
            pass
    if not target:
        target = _find_child_hwnd_at(main_hwnd, cx, cy)
    _cache["cmd_target_hwnd"] = target
    try:
        trect = win32gui.GetWindowRect(target)
        th = trect[3] - trect[1]
        tw = trect[2] - trect[0]
        print(f"type_target: hwnd={target} h={th} w={tw} at=({cx},{cy})", flush=True)
    except Exception:
        pass
    cpt = ctypes.wintypes.POINT(0, 0)
    user32.ClientToScreen(target, ctypes.byref(cpt))
    lx = cx - cpt.x
    ly = cy - cpt.y
    lp = (ly << 16) | (lx & 0xFFFF)
    # Triple-click to select all text in edit field
    for _ in range(3):
        win32gui.PostMessage(target, WM_LBUTTONDOWN, 0x0001, lp)
        win32gui.PostMessage(target, WM_LBUTTONUP, 0, lp)
        time.sleep(0.05)
    time.sleep(0.15)
    # Space prefix — absorbed by dropdown, prevents first real char from being eaten
    win32gui.PostMessage(target, WM_CHAR, ord(' '), 0)
    time.sleep(0.1)
    # Type the real command
    for ch in cmd_text:
        win32gui.PostMessage(target, WM_CHAR, ord(ch), 0)
    time.sleep(0.05)
    # Press Enter
    win32gui.PostMessage(target, WM_KEYDOWN, 0x0D, 0)
    time.sleep(0.02)
    win32gui.PostMessage(target, WM_KEYUP, 0x0D, 0)
    time.sleep(0.15)
    return True


def _detect_layer_from_props(main_hwnd, expected_refdes=None):
    """Detect component side (Top/Bottom) from Allegro's Properties panel.
    
    After a component is highlighted via 'symbol <refdes>', we click the canvas
    center to SELECT the component (symbol only zooms, doesn't select).
    Then we read the Properties panel for COMP_MIRROR/SIDE attributes.
    
    If expected_refdes is given, verifies the Properties panel actually shows
    that refdes before trusting the detected layer.
    
    Returns "Top", "Bottom", or "" if unable to detect.
    """
    try:
        # Click canvas center to select the component that symbol zoomed to
        MOUSEEVENTF_LEFTDOWN = 0x0002
        MOUSEEVENTF_LEFTUP = 0x0004
        wr = win32gui.GetWindowRect(main_hwnd)
        # Canvas center — offset slightly from window center to avoid panels
        panel_rect = _cache.get("_panel_rect")
        if panel_rect:
            # Canvas starts after the left panel
            canvas_left = panel_rect[2]  # right edge of panel
            cx = (canvas_left + wr[2]) // 2
        else:
            cx = (wr[0] + wr[2]) // 2
        cy = (wr[1] + wr[3]) // 2
        
        try:
            win32gui.SetForegroundWindow(main_hwnd)
        except Exception:
            pass
        time.sleep(0.1)
        user32.SetCursorPos(cx, cy)
        time.sleep(0.05)
        user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        time.sleep(0.05)
        user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        time.sleep(0.4)  # wait for Properties panel to update
        
        win = auto.ControlFromHandle(main_hwnd)
        props = None
        for child in win.GetChildren():
            try:
                if child.AutomationId == "SPBMainWindow.PropertiesBar":
                    props = child
                    break
            except Exception:
                continue
        
        if not props:
            return ""
        
        # Collect all text from Properties panel
        def _collect_text(ctrl, depth=0):
            texts = []
            if depth > 5:
                return texts
            try:
                name = ctrl.Name
                if name:
                    texts.append(name)
            except Exception:
                pass
            try:
                vp = ctrl.GetValuePattern()
                if vp and vp.Value:
                    texts.append(vp.Value)
            except Exception:
                pass
            if depth < 4:
                try:
                    for ch in ctrl.GetChildren():
                        texts.extend(_collect_text(ch, depth + 1))
                except Exception:
                    pass
            return texts
        
        all_text = _collect_text(props)
        combined = " ".join(all_text).upper()
        
        # Debug: log what the Properties panel contains
        print(f"  props_text({len(all_text)} items): {combined[:300]}", flush=True)
        
        # Verify refdes if requested — use word-boundary matching
        if expected_refdes:
            ref_upper = expected_refdes.upper()
            import re as _re
            if not _re.search(r'(?<![A-Z0-9])' + _re.escape(ref_upper) + r'(?![A-Z0-9])', combined):
                print(f"  props_refdes_mismatch: expected {expected_refdes}", flush=True)
                return ""
        
        # Check for explicit side indicators
        if "COMP_MIRROR" in combined:
            after = combined.split("COMP_MIRROR")[1][:20]
            if "YES" in after:
                return "Bottom"
            elif "NO" in after:
                return "Top"
        
        if "PLACED ON" in combined:
            after = combined.split("PLACED ON")[1][:20]
            if "TOP" in after:
                return "Top"
            elif "BOTTOM" in after or "BOT" in after:
                return "Bottom"
        
        if "SIDE" in combined:
            after = combined.split("SIDE")[1][:20]
            if "TOP" in after:
                return "Top"
            elif "BOTTOM" in after or "BOT" in after:
                return "Bottom"
        
        # Fallback: assembly/silkscreen layer names
        if "ASSEMBLY_BOTTOM" in combined or "SILKSCREEN_BOTTOM" in combined:
            return "Bottom"
        if "ASSEMBLY_TOP" in combined or "SILKSCREEN_TOP" in combined:
            return "Top"
        
    except Exception as e:
        print(f"  props_detect_err: {e}", flush=True)
    
    return ""




def _do_navigate_cached(main_hwnd, refdes, layer, state):
    """Fast navigation using cached UI positions (no UIA tree search)."""
    global _cache
    import time as _t
    _t0 = _t.perf_counter()

    # Show progress popup during first-time setup
    popup = None
    is_first_setup = not _is_cache_valid(main_hwnd)
    if is_first_setup:
        try:
            popup = _BusyPopup(main_hwnd)
        except Exception:
            pass

    try:
        _do_navigate_cached_inner(main_hwnd, refdes, layer, state, popup, _t, _t0)
    except Exception:
        import traceback
        traceback.print_exc()
    finally:
        if popup:
            popup.close()


def _do_navigate_cached_inner(main_hwnd, refdes, layer, state, popup, _t, _t0):
    """Inner implementation — wrapped by _do_navigate_cached for popup cleanup."""
    global _cache
    is_first_setup = popup is not None or not _is_cache_valid(main_hwnd)

    # Build cache on first call or if window moved
    if is_first_setup:
        if popup:
            popup.step(0)  # "Discovering UI panels…"
        for attempt in range(3):
            try:
                if win32gui.IsIconic(main_hwnd):
                    win32gui.ShowWindow(main_hwnd, 9)
                win32gui.SetForegroundWindow(main_hwnd)
                time.sleep(0.2)
            except Exception:
                pass

            allegro = auto.ControlFromHandle(main_hwnd)
            if popup:
                popup.step(1)  # "Scanning layer table…"
            _cache = _build_cache(allegro, main_hwnd, popup=popup)
            if ("all_off" in _cache and "tab_0" in _cache) or \
               _cache.get("slow_start"):
                break
            print(f"cache_incomplete(attempt={attempt+1}/3), retrying...",
                  flush=True)
            time.sleep(2)
        if not _cache.get("slow_start"):
            state["display_defaults_set"] = True
            state["accordion_expanded"] = True
            state.pop("last_pin_layer", None)
    print(f"T:cache={_t.perf_counter()-_t0:.2f}s", flush=True)

    # If Allegro was unresponsive, just run the command and skip everything else
    if _cache.get("slow_start"):
        if popup:
            popup.step(3)  # "Highlighting component…"
        _t5 = _t.perf_counter()
        if _fast_run_command(main_hwnd, f"symbol {refdes}"):
            print(f"hl_ok T:cmd={_t.perf_counter()-_t5:.2f}s total={_t.perf_counter()-_t0:.2f}s",
                  flush=True)
        else:
            print(f"hl_err T:cmd={_t.perf_counter()-_t5:.2f}s total={_t.perf_counter()-_t0:.2f}s",
                  flush=True)
        state["first_highlight_done"] = True
        _cache.pop("win_rect", None)
        return

    # After resize: force layer re-apply by clearing tracked state
    if _cache.pop("_force_layer_reapply", False):
        state.pop("last_pin_layer", None)
        state.pop("last_view_from", None)

    # When layer is unknown, highlight the component, then try to detect
    # its side from the Properties panel. If detected, apply proper view switch.
    layer_known = layer and layer.lower() in ("top", "bottom")
    is_cold_start = not state.get("first_highlight_done")

    if not layer_known:
        if is_cold_start:
            if popup:
                popup.step(4)  # "Switching layers…"
            # Cold start with unknown layer: enable both pin layers
            _fast_click(main_hwnd, "tab_0")
            time.sleep(0.1)
            _fast_click(main_hwnd, "all_off")
            time.sleep(0.15)
            for _lname in ("Top", "Bottom"):
                _click_layer_cell(main_hwnd, "pin", _lname)
                time.sleep(0.1)
            print(f"T:pin_both={_t.perf_counter()-_t0:.2f}s", flush=True)

        # Highlight command
        if popup:
            popup.step(3)  # "Highlighting component…"
        _t5 = _t.perf_counter()
        if _fast_run_command(main_hwnd, f"symbol {refdes}"):
            print(f"hl_ok T:cmd={_t.perf_counter()-_t5:.2f}s total={_t.perf_counter()-_t0:.2f}s",
                  flush=True)
        else:
            print(f"hl_err T:cmd={_t.perf_counter()-_t5:.2f}s total={_t.perf_counter()-_t0:.2f}s",
                  flush=True)
        state["first_highlight_done"] = True

        # Try to detect layer from Properties panel after highlight
        _t6 = _t.perf_counter()
        detected = _detect_layer_from_props(main_hwnd, expected_refdes=refdes)
        if detected:
            print(f"layer_detected={detected} T:{_t.perf_counter()-_t6:.2f}s", flush=True)
            # Apply proper layer switching now that we know the side
            layer = detected
            target_view = "Top" if layer.lower() != "bottom" else "Bottom"
            need_view_change = state.get("last_view_from") != target_view
            
            # Switch pin layer to detected side only
            _fast_click(main_hwnd, "tab_0")
            time.sleep(0.1)
            _fast_click(main_hwnd, "all_off")
            time.sleep(0.15)
            _click_layer_cell(main_hwnd, "pin", target_view)
            time.sleep(0.1)
            state["last_pin_layer"] = layer
            
            # Switch View From
            if need_view_change:
                _fast_click(main_hwnd, "tab_2")
                time.sleep(0.3)
                _fast_click(main_hwnd, f"vf_{target_view}")
                time.sleep(0.25)
                state["last_view_from"] = target_view
                state["accordion_expanded"] = False
                _fast_click(main_hwnd, "tab_0")
                time.sleep(0.1)
            print(f"T:layer_apply={_t.perf_counter()-_t6:.2f}s", flush=True)
        else:
            print(f"layer_unknown T:{_t.perf_counter()-_t6:.2f}s", flush=True)
        return

    target_view = "Top" if layer.lower() != "bottom" else "Bottom"
    need_view_change = state.get("last_view_from") != target_view
    need_layer_change = state.get("last_pin_layer") != layer
    is_cold_start = not state.get("first_highlight_done")

    if is_cold_start:
        # Cold start: highlight FIRST (user sees result quickly),
        # then fix layers/view reliably with UIA afterward
        # 1. Highlight + zoom (fast keyboard typing)
        if popup:
            popup.step(3)  # "Highlighting component…"
        _t5 = _t.perf_counter()
        if _fast_run_command(main_hwnd, f"symbol {refdes}"):
            print(f"hl_ok T:cmd={_t.perf_counter()-_t5:.2f}s total={_t.perf_counter()-_t0:.2f}s", flush=True)
        else:
            print(f"hl_err T:cmd={_t.perf_counter()-_t5:.2f}s total={_t.perf_counter()-_t0:.2f}s", flush=True)
        state["first_highlight_done"] = True

        # 2. Pin layer — cached positions (accordion already expanded by cache build)
        if need_layer_change:
            if popup:
                popup.step(4)  # "Switching layers…"
            _t1 = _t.perf_counter()
            _fast_click(main_hwnd, "tab_0")
            time.sleep(0.2)
            _fast_click(main_hwnd, "all_off")
            time.sleep(0.2)
            pin_key = f"pin_{target_view}"
            pin_clicked = False
            if pin_key in _cache:
                _click_layer_cell(main_hwnd, "pin", target_view)
                time.sleep(0.15)
                pin_clicked = True
                print(f"T:pin_cached={_t.perf_counter()-_t1:.2f}s", flush=True)
            else:
                # Fall back to fresh UIA
                try:
                    allegro_ctrl = auto.ControlFromHandle(main_hwnd)
                    _expand_accordion(allegro_ctrl, main_hwnd)
                    lt = allegro_ctrl.TableControl(
                        AutomationId=_LAYER_TABLE_AUTOID, timeout=5)
                    gp = lt.GetGridPattern()
                    row_idx = _cache.get("_layer_rows", _LAYER_ROW).get(target_view, 4)
                    cell = gp.GetItem(row_idx, _PIN_COL)
                    # Use scrollbar to ensure cell is visible
                    _scroll_layer_to(allegro_ctrl, target_view, cell)
                    r = cell.BoundingRectangle
                    if r.right - r.left > 0:
                        _post_click(main_hwnd,
                                    (r.left + r.right) // 2,
                                    (r.top + r.bottom) // 2)
                        pin_clicked = True
                except Exception as e:
                    print(f"pin_cold_err:{e}", flush=True)
                print(f"T:pin_uia={_t.perf_counter()-_t1:.2f}s clicked={pin_clicked}", flush=True)
            state["last_pin_layer"] = layer

        # 3. View From (after pin, with longer delays for slow Allegro)
        if popup:
            popup.step(5)  # "Setting view…"
        _t4 = _t.perf_counter()
        if need_view_change:
            _fast_click(main_hwnd, "tab_2")
            time.sleep(0.5)
            _fast_click(main_hwnd, f"vf_{target_view}")
            time.sleep(0.3)
            state["last_view_from"] = target_view
            state["accordion_expanded"] = False
            # Switch back to Layers tab
            _fast_click(main_hwnd, "tab_0")
            time.sleep(0.1)
        print(f"T:viewfrom={_t.perf_counter()-_t4:.2f}s", flush=True)
        return

    # Normal path (subsequent calls — all fast)
    # Pin layer FIRST (before view-from tab switch which collapses accordion)
    if need_layer_change:
        _t1 = _t.perf_counter()
        _fast_click(main_hwnd, "tab_0")
        time.sleep(0.2)
        _fast_click(main_hwnd, "all_off")
        time.sleep(0.2)
        print(f"T:alloff={_t.perf_counter()-_t1:.2f}s", flush=True)
        _t2 = _t.perf_counter()
        pin_key = f"pin_{target_view}"
        pin_clicked = False
        # Use cached position (fast) — positions are reliable after dynamic row detection
        if pin_key in _cache:
            _click_layer_cell(main_hwnd, "pin", target_view)
            time.sleep(0.15)
            pin_clicked = True
            print(f"T:pin_fast={_t.perf_counter()-_t2:.2f}s", flush=True)
        else:
            # Fallback: fresh UIA (slow but reliable)
            try:
                allegro = auto.ControlFromHandle(main_hwnd)
                _expand_accordion(allegro, main_hwnd)
                state["accordion_expanded"] = True
                lt = allegro.TableControl(
                    AutomationId=_LAYER_TABLE_AUTOID, timeout=3)
                gp = lt.GetGridPattern()
                row_idx = _cache.get("_layer_rows", _LAYER_ROW).get(target_view, 4)
                cell = gp.GetItem(row_idx, _PIN_COL)
                # Use scrollbar to ensure cell is visible
                _scroll_layer_to(allegro, target_view, cell)
                r = cell.BoundingRectangle
                if r.right - r.left > 0:
                    _post_click(main_hwnd,
                                (r.left + r.right) // 2,
                                (r.top + r.bottom) // 2)
                    _cache[pin_key] = ((r.left + r.right) // 2,
                                        (r.top + r.bottom) // 2)
                    pin_clicked = True
                print(f"T:pin_uia={_t.perf_counter()-_t2:.2f}s", flush=True)
            except Exception as e:
                print(f"pin_err:{e}", flush=True)
        time.sleep(0.1)
        state["last_pin_layer"] = layer
        print(f"T:pin_total={_t.perf_counter()-_t1:.2f}s clicked={pin_clicked}", flush=True)

    # View From (switch to Display tab — this collapses accordion)
    _t4 = _t.perf_counter()
    if need_view_change:
        _fast_click(main_hwnd, "tab_2")
        time.sleep(0.3)
        _fast_click(main_hwnd, f"vf_{target_view}")
        time.sleep(0.25)
        state["last_view_from"] = target_view
        state["accordion_expanded"] = False  # tab switch collapses it
        # Switch back to Layers tab
        _fast_click(main_hwnd, "tab_0")
        time.sleep(0.1)
    print(f"T:viewfrom={_t.perf_counter()-_t4:.2f}s", flush=True)

    # Highlight + zoom
    _t5 = _t.perf_counter()
    if _fast_run_command(main_hwnd, f"symbol {refdes}"):
        print(f"hl_ok T:cmd={_t.perf_counter()-_t5:.2f}s total={_t.perf_counter()-_t0:.2f}s", flush=True)
    else:
        print(f"hl_err T:cmd={_t.perf_counter()-_t5:.2f}s total={_t.perf_counter()-_t0:.2f}s", flush=True)


def _persistent_main():
    """Persistent mode: stay alive, read commands from stdin.
    Eliminates Python startup + import overhead on every click.
    """
    global _brd_name_filter
    # Parse --brd-name argument
    for i, arg in enumerate(sys.argv):
        if arg == "--brd-name" and i + 1 < len(sys.argv):
            _brd_name_filter = sys.argv[i + 1]
            break

    import time as _t
    _t0 = _t.perf_counter()
    # Retry finding Allegro — it may still be loading the design
    allegro = None
    for attempt in range(15):
        allegro = _find_allegro()
        if allegro:
            break
        time.sleep(2)
    if not allegro:
        print("ERR:no_allegro", flush=True)
        sys.exit(1)
    print(f"W:find_allegro={_t.perf_counter()-_t0:.2f}s", flush=True)

    main_hwnd = allegro.NativeWindowHandle
    _t1 = _t.perf_counter()
    floating_hwnd = None
    for attempt in range(10):
        floating_hwnd = _find_floating_panel(main_hwnd)
        if floating_hwnd:
            break
        time.sleep(2)
    print(f"W:find_panel={_t.perf_counter()-_t1:.2f}s", flush=True)
    if not floating_hwnd:
        print("ERR:no_floating_panel", flush=True)
        sys.exit(1)

    # Watchdog: detect when Allegro is closing (unresponsive) and force-terminate.
    # BRD is read-only, so no data is lost. Prevents the 40s slow shutdown.
    # Pauses during active navigation to avoid false triggers from UIA operations.
    _wd_active = threading.Event()  # set = worker is busy, skip watchdog check
    def _watchdog():
        import ctypes, ctypes.wintypes
        _unresponsive_start = None
        while True:
            time.sleep(0.3)
            if not win32gui.IsWindow(main_hwnd):
                sys.exit(0)
            if _wd_active.is_set():
                _unresponsive_start = None
                continue
            # Check if window is responding (300ms timeout)
            result = ctypes.wintypes.DWORD()
            ok = ctypes.windll.user32.SendMessageTimeoutW(
                main_hwnd, 0, 0, 0, 0x0002, 300,  # SMTO_ABORTIFHUNG
                ctypes.byref(result))
            if ok == 0:
                if _unresponsive_start is None:
                    _unresponsive_start = time.time()
                elif time.time() - _unresponsive_start > 10:
                    # Unresponsive >10s while idle — force-terminate Allegro
                    try:
                        pid = ctypes.wintypes.DWORD()
                        ctypes.windll.user32.GetWindowThreadProcessId(
                            main_hwnd, ctypes.byref(pid))
                        if pid.value:
                            handle = ctypes.windll.kernel32.OpenProcess(
                                1, False, pid.value)  # PROCESS_TERMINATE
                            if handle:
                                ctypes.windll.kernel32.TerminateProcess(handle, 0)
                                ctypes.windll.kernel32.CloseHandle(handle)
                    except Exception:
                        pass
                    sys.exit(0)
            else:
                _unresponsive_start = None
    wd = threading.Thread(target=_watchdog, daemon=True)
    wd.start()

    state = {}
    print("READY", flush=True)

    global _cache
    while True:
        try:
            line = sys.stdin.readline()
            if not line:  # EOF — server closed pipe
                break
            line = line.strip()
            if not line:
                continue

            _wd_active.set()  # pause watchdog during navigation

            # Bring this Allegro window to foreground before any action.
            # Critical for multi-instance: prevents WindowFromPoint/ControlFromPoint
            # from hitting the wrong Allegro window at overlapping screen positions.
            try:
                if win32gui.IsIconic(main_hwnd):
                    win32gui.ShowWindow(main_hwnd, 9)
                win32gui.SetForegroundWindow(main_hwnd)
                time.sleep(0.15)
            except Exception:
                pass

            # NET:netname — highlight a net with etch layers visible
            if line.startswith("NET:"):
                net_name = line[4:]
                _t1 = _t.perf_counter()

                # Ensure cache is built for layer clicks
                if not _is_cache_valid(main_hwnd):
                    net_popup = None
                    try:
                        net_popup = _BusyPopup(main_hwnd)
                    except Exception:
                        pass
                    try:
                        allegro = auto.ControlFromHandle(main_hwnd)
                        _cache = _build_cache(allegro, main_hwnd)
                    finally:
                        if net_popup:
                            net_popup.close()

                # All Off + turn on etch for Top & Bottom
                if not _cache.get("slow_start") and "all_off" in _cache:
                    _fast_click(main_hwnd, "tab_0")
                    time.sleep(0.1)
                    _fast_click(main_hwnd, "all_off")
                    time.sleep(0.15)
                    for layer in ("Top", "Bottom"):
                        _click_layer_cell(main_hwnd, "etch", layer)
                        time.sleep(0.1)
                    print(f"T:etch_on={_t.perf_counter()-_t1:.2f}s", flush=True)

                # Run net highlight command
                _t2 = _t.perf_counter()
                if _fast_run_command(main_hwnd, f"net {net_name}"):
                    print(f"net_ok T:cmd={_t.perf_counter()-_t2:.2f}s total={_t.perf_counter()-_t1:.2f}s", flush=True)
                else:
                    print(f"net_err T:cmd={_t.perf_counter()-_t2:.2f}s", flush=True)
                # Invalidate pin layer state so next refdes click re-applies pin visibility
                state.pop("last_pin_layer", None)
                print("DONE", flush=True)
                _wd_active.clear()  # resume watchdog
                continue

            parts = line.split(maxsplit=1)
            refdes = parts[0]
            layer = parts[1].strip() if len(parts) > 1 else ""

            # Fast cached navigation (no UIA tree search after first call)
            _do_navigate_cached(main_hwnd, refdes, layer, state)
            print("DONE", flush=True)
            _wd_active.clear()  # resume watchdog
        except Exception as e:
            print(f"ERR:{e}", flush=True)
            _wd_active.clear()  # resume watchdog


def main():
    if "--persistent" in sys.argv:
        _persistent_main()
        return

    # One-shot mode (backward compatible)
    if len(sys.argv) < 3:
        print("Usage: navigate_worker.py <refdes> <layer> [<state_file>]")
        print("       navigate_worker.py --persistent")
        sys.exit(1)

    refdes = sys.argv[1]
    layer = sys.argv[2] if sys.argv[2] else "Top"
    state_file = sys.argv[3] if len(sys.argv) > 3 else None

    # Load persistent state
    state = {}
    if state_file and os.path.exists(state_file):
        try:
            with open(state_file) as f:
                state = json.load(f)
        except Exception:
            state = {}

    allegro = _find_allegro()
    if not allegro:
        print("ERR:no_allegro")
        sys.exit(1)

    main_hwnd = allegro.NativeWindowHandle
    floating_hwnd = None
    for attempt in range(10):
        floating_hwnd = _find_floating_panel(main_hwnd)
        if floating_hwnd:
            break
        time.sleep(2)
    if not floating_hwnd:
        print("ERR:no_floating_panel")
        sys.exit(1)

    # One-shot uses slow UIA path (no cache benefit for single call)
    target_view = "Top" if layer.lower() != "bottom" else "Bottom"
    need_view_change = state.get("last_view_from") != target_view
    need_layer_change = state.get("last_pin_layer") != layer

    if not state.get("display_defaults_set"):
        set_display_defaults(allegro, main_hwnd)
        state["display_defaults_set"] = True

    if need_view_change:
        set_view_from(allegro, main_hwnd, target_view,
                      skip_tab_back=need_layer_change)
        state["last_view_from"] = target_view

    if need_layer_change:
        if need_view_change:
            _switch_tab(allegro, main_hwnd, 0)
        accordion_ok = state.get("accordion_expanded", False)
        set_pin_layer(allegro, main_hwnd, target_view,
                      accordion_expanded=accordion_ok)
        state["last_pin_layer"] = layer
        state["accordion_expanded"] = True

    highlight_refdes_cmd(allegro, main_hwnd, refdes)

    # Save state
    if state_file:
        try:
            with open(state_file, "w") as f:
                json.dump(state, f)
        except Exception:
            pass

    print("DONE")


if __name__ == "__main__":
    main()
