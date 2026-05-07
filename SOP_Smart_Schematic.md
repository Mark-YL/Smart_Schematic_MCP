# Smart Schematic v2.1 — Standard Operating Procedure (SOP)

## Purpose

This SOP describes how to set up and use the **Smart Schematic** tool, which generates interactive linked PDFs from Allegro BRD and schematic PDF files. The tool can be used standalone (auto-watcher) or via GitHub Copilot MCP integration.

---

## Table of Contents

1. [Prerequisites](#1-prerequisites)
2. [Install the Tool](#2-install-the-tool)
3. [Register the MCP Server](#3-register-the-mcp-server)
4. [Verify MCP Server Connection](#4-verify-mcp-server-connection)
5. [Using Smart Schematic via Copilot (MCP)](#5-using-smart-schematic-via-copilot-mcp)
6. [Using Smart Schematic Standalone (Auto-Watcher)](#6-using-smart-schematic-standalone-auto-watcher)
7. [Using the Linked PDF](#7-using-the-linked-pdf)
8. [Troubleshooting](#8-troubleshooting)

---

## 1. Prerequisites

| Item | Details |
|------|---------|
| **Python** | 3.10 or later. Download from https://www.python.org or the Microsoft Store. |
| **Allegro Free Viewer** | Required for component navigation. Install from Cadence. Expected at `C:\Cadence\PCBViewers_20xx\tools\bin\allegro_free_viewer.exe`. |
| **Edge Browser** | Required for OnePDM datasheet downloads. |
| **Git** | For cloning the repository. |
| **VS Code + GitHub Copilot** | For MCP integration (optional — standalone watcher works without it). |

---

## 2. Install the Tool

### Step 1: Clone the Repository

```powershell
cd C:\Dev
git clone https://dev.azure.com/MSFTDEVICES/DVSE%20Subsystems/_git/Motherboard.MCP
```

### Step 2: Clone the BRD Location Tool

```powershell
cd C:\Dev
git clone <BRD_Location_tool_repo_url> BRD_Location_tool
```

### Step 3: Install Python Packages

```powershell
pip install PyMuPDF watchdog pystray Pillow selenium uiautomation pywin32
```

### Step 4: Verify Installation

```powershell
cd C:\Dev\BRD_Location_tool
python -c "import fitz, watchdog, pystray, PIL; print('All packages OK')"
```

---

## 3. Register the MCP Server

The MCP server allows GitHub Copilot to use Smart Schematic tools directly from chat.

### Step 1: Locate Your MCP Config File

The config file is at:
```
%USERPROFILE%\.copilot\mcp-config.json
```

If it doesn't exist, create it.

### Step 2: Add the Smart Schematic Entry

Open the file and add the `smart-schematic` entry under `mcpServers`:

```json
{
  "mcpServers": {
    "smart-schematic": {
      "type": "stdio",
      "command": "python",
      "args": ["C:\\Dev\\Motherboard.MCP\\servers\\smart_schematic\\smart_schematic_mcp_server.py"]
    }
  }
}
```

> **Note:** Adjust the path if your Motherboard.MCP repo is cloned to a different location.

### Step 3: Set the Tool Directory (Optional)

If your `BRD_Location_tool` is not at `C:\Dev\BRD_Location_tool`, set the environment variable:

```json
{
  "mcpServers": {
    "smart-schematic": {
      "type": "stdio",
      "command": "python",
      "args": ["C:\\Dev\\Motherboard.MCP\\servers\\smart_schematic\\smart_schematic_mcp_server.py"],
      "env": {
        "SMART_SCHEMATIC_DIR": "D:\\your\\path\\to\\BRD_Location_tool"
      }
    }
  }
}
```

### Step 4: Restart VS Code

After saving the config file, restart VS Code for the MCP server to be registered.

---

## 4. Verify MCP Server Connection

### In VS Code Copilot Chat

1. Open Copilot Chat (Ctrl+Shift+I)
2. Type: `List the tools available from smart-schematic`
3. You should see 6 tools:

| Tool | Description |
|------|-------------|
| `schematic_generate_linked_pdf` | Generate linked PDF from BRD + schematic pair |
| `schematic_extract_components` | Extract component data from Allegro BRD binary |
| `schematic_download_datasheets` | Download datasheets from OnePDM |
| `schematic_list_parts` | List refdes → part number associations |
| `schematic_start_watcher` | Start the auto-watcher |
| `schematic_datasheet_status` | Check datasheet download progress |

### In Copilot CLI

```powershell
# Test the MCP server manually
python C:\Dev\Motherboard.MCP\servers\smart_schematic\smart_schematic_mcp_server.py
```

It should output JSON-RPC initialization messages. Press Ctrl+C to stop.

---

## 5. Using Smart Schematic via Copilot (MCP)

### Generate a Linked PDF

Ask Copilot:
```
Generate a linked PDF for the schematic at C:\Projects\MyBoard\schematic.pdf 
with the BRD file C:\Projects\MyBoard\MyBoard.brd
```

Copilot will:
1. Extract component placements from the BRD (pure Python, no Allegro needed)
2. Generate the linked PDF with clickable reference designators
3. Output `schematic_linked_Acrobat.pdf` in the same folder

### Extract Components Only

```
Extract components from C:\Projects\MyBoard\MyBoard.brd
```

This creates `component_data.json` with refdes, layer (Top/Bottom), and coordinates.

### Download Datasheets

```
Download datasheets for all components in C:\Projects\MyBoard\schematic.pdf
```

Downloads datasheets from OnePDM into a `datasheets/` subfolder, organized by refdes and part number.

### List Part Numbers

```
List all part numbers in C:\Projects\MyBoard\schematic.pdf
```

Shows the refdes → part number mapping found in the schematic PDF.

### Start the Watcher

```
Start watching C:\Projects for new BRD+PDF pairs
```

---

## 6. Using Smart Schematic Standalone (Auto-Watcher)

The auto-watcher monitors folders and automatically generates linked PDFs when BRD+PDF pairs are detected.

### Start Manually

```powershell
pythonw C:\Dev\BRD_Location_tool\auto_linked_pdf.py --inbox "C:\Users\<you>\OneDrive - Microsoft\2026_Projects"
```

### Set Up Auto-Start at Login

```powershell
schtasks /create /tn "SmartSchematic_Watcher" ^
    /tr "pythonw C:\Dev\BRD_Location_tool\auto_linked_pdf.py --inbox \"C:\Users\<you>\OneDrive - Microsoft\2026_Projects\"" ^
    /sc onlogon /rl highest
```

### Using the System Tray Icon

Once running, a ⚡ lightning bolt icon appears in the system tray.

| Action | What It Does |
|--------|-------------|
| **Click** the icon | Opens the Folder Monitor UI |
| **+ Add Folder** | Browse and add a folder to monitor |
| **- Remove** | Stop monitoring the selected folder |
| **Activity Log** | Shows real-time processing status |
| **Right-click → Quit** | Stops the watcher |

### How File Matching Works

1. Place your `.pdf` and `.brd` files in the **same folder**
2. Files are matched by name: `MyBoard.pdf` + `MyBoard.brd` (case-insensitive)
3. If only one PDF and one BRD exist in a folder, they auto-pair
4. The watcher generates `MyBoard_linked_Acrobat.pdf` automatically

### Example Folder Structure

```
2026_Projects/
├── ProjectA/
│   ├── schematic.pdf          ← input
│   ├── schematic.brd          ← input
│   ├── schematic_linked_Acrobat.pdf  ← output (auto-generated)
│   ├── component_data.json    ← output
│   └── datasheets/            ← output (auto-downloaded)
│       ├── U1001_TPS51363/
│       │   └── TPS51363_datasheet.pdf
│       └── ...
└── ProjectB/
    ├── board.pdf
    └── board.brd
```

---

## 7. Using the Linked PDF

### Opening the Linked PDF

1. Open the `*_linked_Acrobat.pdf` file in **Adobe Acrobat** or **Microsoft Edge**
2. Make sure the navigation server is running:
   ```powershell
   pythonw C:\Dev\BRD_Location_tool\refdes_server.py
   ```
   (The auto-watcher starts this automatically)

### Clicking Reference Designators

- Click any **U1001**, **R2345**, **C6789** etc. in the PDF
- Allegro Free Viewer opens (first time) or navigates (subsequent clicks) to that component
- The viewer zooms to the component location and highlights it
- Layers are set automatically (Top or Bottom based on component placement)

### Clicking Part Numbers

- Click any **part number** link → opens the datasheet PDF (if downloaded)

### Clicking Net Names

- Click any **net name** → highlights the net in Allegro Free Viewer

### Navigation Performance

| Action | Time |
|--------|------|
| First click (launches Allegro) | ~15-20 seconds |
| Subsequent clicks (same project) | ~1-2 seconds |
| Switching between projects | ~15-20 seconds (new Allegro instance) |

---

## 8. Troubleshooting

| Problem | Cause | Solution |
|---------|-------|----------|
| No tray icon visible | Watcher not running | Run `pythonw auto_linked_pdf.py --inbox <folder>` |
| "No components found" | BRD format variation | File may use a different internal structure — contact Mark Lai |
| Linked PDF not generated | Missing pair | Ensure both `.pdf` and `.brd` are in the same folder with matching names |
| Allegro re-launches every click | Watchdog killed it | Update to v2.1 (watchdog timeout increased to 10s) |
| All layers off in Allegro | Normal behavior | Layers are turned off for faster loading; the tool enables the correct layer automatically |
| Datasheets not downloading | Not on corpnet | OnePDM requires Microsoft corpnet access + Edge browser |
| MCP server not showing in Copilot | Config error | Check `~/.copilot/mcp-config.json` syntax and restart VS Code |
| Navigation links don't work | Server not running | Start `refdes_server.py` — it listens on `http://localhost:5588` |
| "Project not found in registry" | New project | The linked PDF auto-registers its project on generation; regenerate if needed |

### Getting Help

- **Tool repo:** `MSFTDEVICES/DVSE Subsystems/_git/Motherboard.MCP` → `servers/smart_schematic/`
- **Owner:** Mark Lai (kil@microsoft.com)

---

## Appendix: Available MCP Tools Reference

### schematic_generate_linked_pdf

```
Parameters:
  pdf_path   (required) — Path to the schematic PDF file
  brd_path   (required) — Path to the Allegro .brd file
  output_dir (optional) — Output directory (default: same as PDF)
```

### schematic_extract_components

```
Parameters:
  brd_path (required) — Path to the Allegro .brd file
```

### schematic_download_datasheets

```
Parameters:
  pdf_path       (required) — Path to the schematic PDF file
  component_data (required) — Path to component_data.json
  output_dir     (required) — Output directory for datasheets
  max_count      (optional) — Max datasheets to download (0 = all)
```

### schematic_list_parts

```
Parameters:
  pdf_path       (required) — Path to the schematic PDF file
  component_data (required) — Path to component_data.json
```

### schematic_start_watcher

```
Parameters:
  inbox_dir  (optional) — Primary folder to watch
  extra_dirs (optional) — Additional folders to watch
```

### schematic_datasheet_status

```
Parameters:
  progress_file (required) — Path to _progress.json in datasheets folder
```
