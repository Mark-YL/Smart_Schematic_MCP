# Smart Schematic v2.1 — MCP Server & Auto-Watcher

**Owner:** Mark Lai (kil@microsoft.com)

## What It Does

Smart Schematic generates **interactive linked PDFs** from Allegro BRD files and schematic PDFs.

- Click any **reference designator** (e.g. U1001) in the PDF → Allegro Free Viewer zooms to that component
- Click any **part number** → opens the datasheet
- Click any **net name** → highlights the net in the viewer
- Automatically detects Top/Bottom layer from BRD binary data

---

## Connect to This MCP Server (Copilot CLI)

### 1. Clone this repo

```powershell
git clone https://github.com/Mark-YL/Smart_Schematic_MCP.git C:\Dev\Smart_Schematic_MCP
```

### 2. Install Python dependencies

```powershell
pip install PyMuPDF watchdog pystray Pillow selenium uiautomation pywin32
```

### 3. Register the MCP server

Edit (or create) `%USERPROFILE%\.copilot\mcp-config.json` and add:

```json
{
  "mcpServers": {
    "smart-schematic": {
      "type": "stdio",
      "command": "python",
      "args": ["C:\\Dev\\Smart_Schematic_MCP\\smart_schematic_mcp_server.py"]
    }
  }
}
```

> **Note:** Adjust the path if you cloned to a different location.

### 4. Restart Copilot CLI

The 6 Smart Schematic tools are now available in your Copilot conversations:

| Tool | What It Does |
|------|-------------|
| `schematic_generate_linked_pdf` | Generate linked PDF from BRD + schematic pair |
| `schematic_extract_components` | Extract component data from Allegro BRD binary |
| `schematic_download_datasheets` | Download datasheets from OnePDM |
| `schematic_list_parts` | List refdes → part number associations |
| `schematic_start_watcher` | Start the auto-watcher |
| `schematic_datasheet_status` | Check datasheet download progress |

### 5. Try it

Ask Copilot:
```
Generate a linked PDF for C:\Projects\MyBoard\schematic.pdf with C:\Projects\MyBoard\MyBoard.brd
```

---

## Quick Start (Standalone)

### Option A: Automatic (Recommended)

Just drop your `.brd` + `.pdf` files into a watched folder. The tool handles everything.

1. **Look for the ⚡ tray icon** in your Windows system tray (it starts at login)
2. **Click the tray icon** → opens the Folder Monitor UI
3. **Add your project folder** using the "+ Add Folder" button
4. **Drop files**: place `MyProject.pdf` + `MyProject.brd` in the same folder
5. **Wait** — the linked PDF (`MyProject_linked_Acrobat.pdf`) generates automatically

The Activity Log in the UI shows real-time progress.

### Option B: Via Copilot Chat (MCP Tools)

Ask Copilot in VS Code or CLI to use the Smart Schematic tools:

```
"Generate a linked PDF for C:\Projects\Minos\m1331388-008.pdf and M1331388-008.brd"
```

```
"Extract components from the BRD file at C:\Projects\board.brd"
```

```
"Download datasheets for all components in m1331388-008.pdf"
```

---

## Setup

### Prerequisites

| Requirement | Purpose |
|---|---|
| Python 3.10+ | Core runtime |
| `pip install PyMuPDF watchdog pystray Pillow` | PDF linking, file watching, tray icon |
| `pip install selenium` | Datasheet downloads (OnePDM) |
| Allegro Free Viewer | Component navigation (not needed for PDF generation) |
| Edge browser | OnePDM datasheet automation |

### Installation

1. **Clone the tool repo:**
   ```
   git clone <BRD_Location_tool repo> C:\Dev\BRD_Location_tool
   ```

2. **Install Python packages:**
   ```
   pip install PyMuPDF watchdog pystray Pillow selenium
   ```

3. **Register the MCP server** (for Copilot integration):
   
   Add to `~/.copilot/mcp-config.json`:
   ```json
   {
     "servers": {
       "smart-schematic": {
         "type": "stdio",
         "command": "python",
         "args": ["C:\\Dev\\Motherboard.MCP\\servers\\smart_schematic\\smart_schematic_mcp_server.py"]
       }
     }
   }
   ```

4. **Start the auto-watcher** (runs at login via Scheduled Task):
   ```
   pythonw auto_linked_pdf.py --inbox "C:\Users\<you>\OneDrive - Microsoft\2026_Projects"
   ```
   Or set up the Windows Scheduled Task:
   ```
   schtasks /create /tn "SmartSchematic_Watcher" /tr "pythonw C:\Dev\BRD_Location_tool\auto_linked_pdf.py --inbox \"<your folder>\"" /sc onlogon
   ```

---

## How It Works

```
┌─────────────┐     ┌──────────────────┐     ┌────────────────────┐
│  Drop files │────▶│  Auto-Watcher    │────▶│  Linked PDF Output │
│  .pdf + .brd│     │  (file monitor)  │     │  _linked_Acrobat   │
└─────────────┘     └──────────────────┘     └────────────────────┘
                           │
                           ▼
                    ┌──────────────────┐
                    │  BRD Parser      │  ← Pure Python, no Allegro needed
                    │  (binary extract)│
                    └──────────────────┘
                           │
                           ▼
                    ┌──────────────────┐
                    │  Datasheet DL    │  ← Background, from OnePDM
                    │  (per component) │
                    └──────────────────┘
```

### File Matching Rules

The watcher pairs files automatically:
1. **Exact name match**: `MyBoard.pdf` + `MyBoard.brd` (case-insensitive)
2. **Only pair in folder**: If exactly 1 PDF + 1 BRD exist → auto-paired
3. Files can arrive in any order — the watcher waits for the companion

### Output Files

| File | Description |
|---|---|
| `*_linked_Acrobat.pdf` | Interactive PDF with clickable links |
| `component_data.json` | Extracted component placements (refdes, layer, x, y) |
| `net_data.json` | Extracted net names |
| `datasheets/` | Downloaded component datasheets (organized by refdes_partnum) |

---

## System Tray Icon

The ⚡ tray icon provides:

| Action | What it does |
|---|---|
| **Click** (or double-click) | Opens Folder Monitor UI |
| **Right-click → Manage Folders** | Same as click |
| **Right-click → Quit** | Stops the watcher |

### Folder Monitor UI

- **Watched Folders list** — shows all monitored directories
- **+ Add Folder** — browse to add a new folder to watch
- **- Remove** — stop watching the selected folder
- **Activity Log** — live feed showing:
  - PDF+BRD pair detection
  - Linked PDF generation progress
  - Datasheet download progress (e.g. `5/85 - C1234_GRM155R60J106ME15`)

---

## MCP Server Tools

Available when registered as an MCP server in Copilot:

| Tool | Parameters | Description |
|---|---|---|
| `schematic_generate_linked_pdf` | `pdf_path`, `brd_path`, `output_dir` | Generate linked PDF from a BRD + schematic pair |
| `schematic_extract_components` | `brd_path` | Extract component data from Allegro BRD binary |
| `schematic_download_datasheets` | `pdf_path`, `component_data`, `output_dir`, `max_count` | Download datasheets from OnePDM |
| `schematic_list_parts` | `pdf_path`, `component_data` | List refdes→part number associations |
| `schematic_start_watcher` | `inbox_dir`, `extra_dirs` | Start the auto-watcher |
| `schematic_datasheet_status` | `progress_file` | Check datasheet download progress |

---

## Troubleshooting

| Issue | Solution |
|---|---|
| No tray icon visible | Run `pythonw auto_linked_pdf.py --inbox <folder>` manually |
| Linked PDF not generated | Check Activity Log — ensure both .pdf and .brd are in the same folder |
| "No components found" | BRD format may not be supported — contact Mark |
| Datasheets not downloading | Requires Microsoft corpnet + Edge browser |
| Navigation not working | Install Allegro Free Viewer and start `refdes_server.py` |

---

## Configuration

Set environment variable `SMART_SCHEMATIC_DIR` to override the tool location:

```json
{
  "env": {
    "SMART_SCHEMATIC_DIR": "C:\\Dev\\BRD_Location_tool"
  }
}
```

Default: `~/Dev/BRD_Location_tool`
