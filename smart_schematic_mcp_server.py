"""
smart_schematic_mcp_server.py -- MCP server for Smart Schematic tool.
Stdio transport (JSON-RPC over stdin/stdout). All logging to stderr.

Generates linked PDFs from Allegro BRD + schematic PDF pairs,
navigates to components in Allegro Free Viewer, and downloads datasheets from OnePDM.
"""
import sys, os, json, subprocess, time, traceback, threading

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stdin, 'reconfigure'):
    sys.stdin.reconfigure(encoding='utf-8')

# --- Path to BRD_Location_tool ---
BRD_TOOL_DIR = os.environ.get("SMART_SCHEMATIC_DIR",
    os.path.join(os.path.expanduser('~'), 'Dev', 'BRD_Location_tool'))
sys.path.insert(0, BRD_TOOL_DIR)

def log(msg):
    print(f"[smart-schematic] {msg}", file=sys.stderr, flush=True)

# ===================================================================
# Tool Implementations
# ===================================================================

def tool_generate_linked_pdf(args):
    """Generate a linked PDF from a BRD + schematic PDF pair."""
    pdf_path = args.get("pdf_path", "")
    brd_path = args.get("brd_path", "")
    output_dir = args.get("output_dir", "")

    if not pdf_path or not brd_path:
        return "Error: both pdf_path and brd_path are required"
    if not os.path.exists(pdf_path):
        return f"Error: PDF not found: {pdf_path}"
    if not os.path.exists(brd_path):
        return f"Error: BRD not found: {brd_path}"

    try:
        from extract_brd_components import extract_components_from_brd
        from generate_linked_pdf import generate_linked_pdf

        # Step 1: Extract components from BRD
        log(f"Extracting components from {os.path.basename(brd_path)}...")
        components = extract_components_from_brd(brd_path)

        if not components:
            return f"Error: No components extracted from {brd_path}"

        # Save component data
        if not output_dir:
            output_dir = os.path.dirname(pdf_path)
        comp_data_path = os.path.join(output_dir, "component_data.json")
        with open(comp_data_path, 'w') as f:
            json.dump(components, f, indent=2)

        # Step 2: Generate linked PDF
        log(f"Generating linked PDF...")
        out_pdf = os.path.join(output_dir,
            os.path.splitext(os.path.basename(pdf_path))[0] + "_linked.pdf")
        generate_linked_pdf(pdf_path, comp_data_path, out_pdf)

        return (f"✅ Linked PDF generated successfully!\n"
                f"  Components: {len(components)}\n"
                f"  Output: {out_pdf}\n"
                f"  Component data: {comp_data_path}")
    except Exception as e:
        return f"Error: {e}\n{traceback.format_exc()}"


def tool_extract_components(args):
    """Extract component data from an Allegro BRD file."""
    brd_path = args.get("brd_path", "")
    if not brd_path:
        return "Error: brd_path is required"
    if not os.path.exists(brd_path):
        return f"Error: BRD not found: {brd_path}"

    try:
        from extract_brd_components import extract_components_from_brd
        t0 = time.time()
        components = extract_components_from_brd(brd_path)
        elapsed = time.time() - t0

        # Summary
        layers = {}
        for ref, data in components.items():
            layer = data.get("layer", "unknown")
            layers[layer] = layers.get(layer, 0) + 1

        result = f"Extracted {len(components)} components in {elapsed:.1f}s\n"
        result += f"  File: {os.path.basename(brd_path)}\n"
        for layer, count in sorted(layers.items()):
            result += f"  {layer}: {count} components\n"
        result += f"\nSample (first 10):\n"
        for ref in list(sorted(components.keys()))[:10]:
            d = components[ref]
            result += f"  {ref}: layer={d.get('layer','?')}, x={d.get('x',0):.1f}, y={d.get('y',0):.1f}\n"
        return result
    except Exception as e:
        return f"Error: {e}\n{traceback.format_exc()}"


def tool_download_datasheets(args):
    """Download datasheets from OnePDM for components in a schematic PDF."""
    pdf_path = args.get("pdf_path", "")
    comp_data_path = args.get("component_data", "")
    output_dir = args.get("output_dir", "")
    max_count = args.get("max_count", 0)

    if not pdf_path:
        return "Error: pdf_path is required"
    if not os.path.exists(pdf_path):
        return f"Error: PDF not found: {pdf_path}"

    try:
        from download_datasheets import extract_refdes_part_pairs, download_worker

        # Extract associations
        pairs = extract_refdes_part_pairs(pdf_path, comp_data_path or None)
        if not pairs:
            return "No refdes-part associations found in PDF"

        # Deduplicate
        seen_pn = {}
        tasks = []
        for refdes, pn in sorted(pairs.items()):
            if pn not in seen_pn:
                seen_pn[pn] = refdes
                tasks.append((refdes, pn))

        if max_count and max_count > 0:
            tasks = tasks[:max_count]

        if not output_dir:
            output_dir = os.path.join(os.path.dirname(pdf_path), "datasheets")
        os.makedirs(output_dir, exist_ok=True)
        progress_file = os.path.join(output_dir, "_progress.json")

        # Run download in background thread
        t = threading.Thread(
            target=download_worker,
            args=(tasks, output_dir, progress_file),
            daemon=True
        )
        t.start()

        return (f"✅ Datasheet download started in background (headless)\n"
                f"  Total unique parts: {len(seen_pn)}\n"
                f"  Downloading: {len(tasks)} parts\n"
                f"  Output: {output_dir}\n"
                f"  Progress: {progress_file}")
    except Exception as e:
        return f"Error: {e}\n{traceback.format_exc()}"


def tool_list_parts(args):
    """List refdes-to-part associations found in a schematic PDF."""
    pdf_path = args.get("pdf_path", "")
    comp_data_path = args.get("component_data", "")

    if not pdf_path:
        return "Error: pdf_path is required"
    if not os.path.exists(pdf_path):
        return f"Error: PDF not found: {pdf_path}"

    try:
        from download_datasheets import extract_refdes_part_pairs
        pairs = extract_refdes_part_pairs(pdf_path, comp_data_path or None)

        # Deduplicate by part number
        seen_pn = {}
        for refdes, pn in sorted(pairs.items()):
            if pn not in seen_pn:
                seen_pn[pn] = refdes

        result = f"Found {len(pairs)} associations, {len(seen_pn)} unique parts\n\n"
        for pn, ref in sorted(seen_pn.items(), key=lambda x: x[1]):
            result += f"  {ref}_{pn}\n"
        return result
    except Exception as e:
        return f"Error: {e}\n{traceback.format_exc()}"


def tool_start_watcher(args):
    """Start the auto-watcher to monitor directories for new BRD+PDF pairs."""
    watch_dirs = args.get("watch_dirs", [])
    if not watch_dirs:
        watch_dirs = [os.path.join(BRD_TOOL_DIR, "inbox")]

    # Validate dirs
    for d in watch_dirs:
        if not os.path.isdir(d):
            return f"Error: directory not found: {d}"

    try:
        cmd = [sys.executable, os.path.join(BRD_TOOL_DIR, "auto_linked_pdf.py"),
               "--watch"] + watch_dirs
        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        return (f"✅ Watcher started (PID {proc.pid})\n"
                f"  Monitoring: {', '.join(watch_dirs)}\n"
                f"  Will auto-generate linked PDFs when BRD+PDF pairs are detected")
    except Exception as e:
        return f"Error starting watcher: {e}"


def tool_datasheet_status(args):
    """Check datasheet download progress."""
    progress_file = args.get("progress_file", "")
    if not progress_file:
        return "Error: progress_file path is required"
    if not os.path.exists(progress_file):
        return "No progress file found — download may not have started yet"

    try:
        with open(progress_file) as f:
            progress = json.load(f)
        done = sum(1 for v in progress.values() if v.get("status") == "done")
        failed = sum(1 for v in progress.values() if v.get("status") in ("not_found", "error"))
        downloading = sum(1 for v in progress.values() if v.get("status") == "downloading")
        total = len(progress)
        return (f"Datasheet download progress:\n"
                f"  ✅ Downloaded: {done}\n"
                f"  ❌ Failed: {failed}\n"
                f"  ⏳ In progress: {downloading}\n"
                f"  Total tracked: {total}")
    except Exception as e:
        return f"Error reading progress: {e}"


# ===================================================================
# MCP Tool Definitions & Handler Map
# ===================================================================
TOOLS = [
    {
        "name": "schematic_generate_linked_pdf",
        "description": "Generate a linked PDF from an Allegro BRD file and schematic PDF. Clickable reference designators navigate to components in Allegro Free Viewer.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "pdf_path": {"type": "string", "description": "Path to the schematic PDF file"},
                "brd_path": {"type": "string", "description": "Path to the Allegro .brd file"},
                "output_dir": {"type": "string", "description": "Output directory (default: same as PDF)"}
            },
            "required": ["pdf_path", "brd_path"]
        }
    },
    {
        "name": "schematic_extract_components",
        "description": "Extract component placement data (refdes, layer, coordinates) from an Allegro BRD binary file. Pure Python — no Allegro installation needed.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "brd_path": {"type": "string", "description": "Path to the Allegro .brd file"}
            },
            "required": ["brd_path"]
        }
    },
    {
        "name": "schematic_download_datasheets",
        "description": "Download datasheets from OnePDM for all components found in a schematic PDF. Runs headless (no browser window) in background.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "pdf_path": {"type": "string", "description": "Path to the schematic PDF file"},
                "component_data": {"type": "string", "description": "Path to component_data.json (for refdes validation)"},
                "output_dir": {"type": "string", "description": "Output directory for datasheets"},
                "max_count": {"type": "integer", "description": "Max number of datasheets to download (0 = all)"}
            },
            "required": ["pdf_path"]
        }
    },
    {
        "name": "schematic_list_parts",
        "description": "List all refdes-to-part-number associations found in a schematic PDF. Useful to preview what datasheets will be downloaded.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "pdf_path": {"type": "string", "description": "Path to the schematic PDF file"},
                "component_data": {"type": "string", "description": "Path to component_data.json (for refdes validation)"}
            },
            "required": ["pdf_path"]
        }
    },
    {
        "name": "schematic_start_watcher",
        "description": "Start the auto-watcher that monitors directories for new BRD+PDF pairs and automatically generates linked PDFs and downloads datasheets.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "watch_dirs": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "List of directories to monitor"
                }
            },
            "required": []
        }
    },
    {
        "name": "schematic_datasheet_status",
        "description": "Check the progress of a background datasheet download.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "progress_file": {"type": "string", "description": "Path to _progress.json file in datasheets folder"}
            },
            "required": ["progress_file"]
        }
    }
]

TOOL_HANDLERS = {
    "schematic_generate_linked_pdf": tool_generate_linked_pdf,
    "schematic_extract_components": tool_extract_components,
    "schematic_download_datasheets": tool_download_datasheets,
    "schematic_list_parts": tool_list_parts,
    "schematic_start_watcher": tool_start_watcher,
    "schematic_datasheet_status": tool_datasheet_status,
}

# ===================================================================
# MCP JSON-RPC Protocol
# ===================================================================
def handle_request(req):
    method = req.get("method", "")
    params = req.get("params", {})
    req_id = req.get("id")

    if method == "initialize":
        return {
            "jsonrpc": "2.0",
            "id": req_id,
            "result": {
                "protocolVersion": "2024-11-05",
                "capabilities": {"tools": {}},
                "serverInfo": {
                    "name": "smart-schematic",
                    "version": "1.0.0"
                }
            }
        }
    elif method == "notifications/initialized":
        return None
    elif method == "tools/list":
        return {
            "jsonrpc": "2.0",
            "id": req_id,
            "result": {"tools": TOOLS}
        }
    elif method == "tools/call":
        tool_name = params.get("name", "")
        tool_args = params.get("arguments", {})
        handler = TOOL_HANDLERS.get(tool_name)
        if not handler:
            return {
                "jsonrpc": "2.0",
                "id": req_id,
                "result": {
                    "content": [{"type": "text", "text": f"Unknown tool: {tool_name}"}],
                    "isError": True
                }
            }
        try:
            log(f"Calling tool: {tool_name}")
            result_text = handler(tool_args)
            log(f"Tool {tool_name} completed")
            return {
                "jsonrpc": "2.0",
                "id": req_id,
                "result": {
                    "content": [{"type": "text", "text": result_text}],
                    "isError": False
                }
            }
        except Exception as e:
            error_msg = f"Error in {tool_name}: {e}\n{traceback.format_exc()}"
            log(error_msg)
            return {
                "jsonrpc": "2.0",
                "id": req_id,
                "result": {
                    "content": [{"type": "text", "text": error_msg}],
                    "isError": True
                }
            }
    elif method == "ping":
        return {"jsonrpc": "2.0", "id": req_id, "result": {}}
    else:
        if req_id is not None:
            return {
                "jsonrpc": "2.0",
                "id": req_id,
                "error": {"code": -32601, "message": f"Method not found: {method}"}
            }
        return None


def main():
    log(f"Smart Schematic MCP server starting (tool_dir: {BRD_TOOL_DIR})")
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            req = json.loads(line)
        except json.JSONDecodeError:
            continue
        response = handle_request(req)
        if response is not None:
            out = json.dumps(response)
            sys.stdout.write(out + '\n')
            sys.stdout.flush()


if __name__ == '__main__':
    main()
