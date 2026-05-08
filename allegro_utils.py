"""
Smart Schematic — Allegro Viewer Utilities
Auto-detects Allegro Free Viewer installation path.
"""

import os
import glob


def find_allegro_viewer():
    """Auto-detect Allegro Free Viewer executable.
    
    Search order:
    1. Common Cadence install paths (PCBViewers_20xx, newest first)
    2. SPB paths (SPB_xx.x)
    3. PATH environment variable
    
    Returns the full path to allegro_free_viewer.exe, or None if not found.
    """
    exe_name = "allegro_free_viewer.exe"

    # 1. Scan C:\Cadence for PCBViewers_* folders (newest first)
    cadence_root = r"C:\Cadence"
    if os.path.isdir(cadence_root):
        # PCBViewers_2025, PCBViewers_2024, PCBViewers_2023, ...
        viewer_dirs = sorted(
            glob.glob(os.path.join(cadence_root, "PCBViewers_*")),
            reverse=True
        )
        for d in viewer_dirs:
            exe = os.path.join(d, "tools", "bin", exe_name)
            if os.path.isfile(exe):
                return exe

        # SPB_xx.x (e.g., SPB_17.4, SPB_23.1)
        spb_dirs = sorted(
            glob.glob(os.path.join(cadence_root, "SPB_*")),
            reverse=True
        )
        for d in spb_dirs:
            exe = os.path.join(d, "tools", "bin", exe_name)
            if os.path.isfile(exe):
                return exe

    # 2. Check PATH
    path_dirs = os.environ.get("PATH", "").split(os.pathsep)
    for d in path_dirs:
        exe = os.path.join(d, exe_name)
        if os.path.isfile(exe):
            return exe

    return None


# Module-level cached result
_cached_path = None


def get_allegro_viewer():
    """Get Allegro viewer path (cached). Raises FileNotFoundError if not found."""
    global _cached_path
    if _cached_path is None:
        _cached_path = find_allegro_viewer()
    if _cached_path is None:
        raise FileNotFoundError(
            "Allegro Free Viewer not found. "
            "Please install Cadence Allegro Free Viewer or add it to PATH.\n"
            "Expected location: C:\\Cadence\\PCBViewers_20xx\\tools\\bin\\allegro_free_viewer.exe"
        )
    return _cached_path
