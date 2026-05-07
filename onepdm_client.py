"""
OnePDM Client — Opens part pages in OnePDM and navigates to "Get all files".

Uses Aras client-side JavaScript API for instant Part lookup + open.
No visible search typing or menu navigation — just direct item open.

Flow:
  1. First click: launch Edge → OnePDM SSO → wait for SPA
  2. JavaScript: aras IOM lookup Part by item_number → get ID → uiShowItem
  3. Navigate to More → Files in the Part page
"""

import time
import tempfile
import threading

_ONEPDM_URL = "https://onepdm.plm.microsoft.com/onepdm/"


class OnePdmClient:
    """Manages a persistent OnePDM browser session."""

    def __init__(self, headless=False):
        self._driver = None
        self._lock = threading.Lock()
        self._temp_dir = None
        self._ready = False
        self._part_id_cache = {}  # part_number -> aras item id
        self._download_dir = None  # custom download directory
        self._headless = headless

    def _ensure_browser(self):
        """Open Edge and navigate to OnePDM if not already open."""
        if self._driver:
            try:
                _ = self._driver.title
                return True
            except Exception:
                # Browser was closed — clean up, fall through to relaunch
                print("  [onepdm] browser was closed, relaunching...", flush=True)
                self._driver = None
                self._ready = False

        try:
            from selenium import webdriver
            from selenium.webdriver.edge.options import Options
            from selenium.webdriver.common.by import By
            from selenium.webdriver.support.ui import WebDriverWait
            from selenium.webdriver.support import expected_conditions as EC
        except ImportError:
            print("  [onepdm] selenium not installed")
            return False

        print("  [onepdm] Opening OnePDM browser...", end="", flush=True)
        t0 = time.time()
        self._temp_dir = tempfile.mkdtemp(prefix="edge_onepdm_")
        options = Options()
        options.add_argument(f"--user-data-dir={self._temp_dir}")
        options.add_argument("--no-first-run")
        options.add_argument("--no-default-browser-check")
        # Headless mode — no visible browser window
        if getattr(self, '_headless', False):
            options.add_argument("--headless=new")
            options.add_argument("--disable-gpu")
            options.add_argument("--window-size=1920,1080")
        # Set download directory if specified
        if self._download_dir:
            import os
            os.makedirs(self._download_dir, exist_ok=True)
            prefs = {
                "download.default_directory": self._download_dir,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
            }
            options.add_experimental_option("prefs", prefs)

        try:
            self._driver = webdriver.Edge(options=options)
            self._driver.maximize_window()
            self._driver.get(_ONEPDM_URL)

            # Wait for Aras SPA fully loaded: search box + IOM ready
            wait = WebDriverWait(self._driver, 60)
            # First wait for search input (stable SPA indicator)
            from selenium.webdriver.common.by import By
            from selenium.webdriver.support import expected_conditions as EC
            wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, 'input[placeholder="Search by Item ID/Keyword"]')
            ))
            # Then wait for IOM to be fully initialized
            wait.until(lambda d: d.execute_script(
                "try { aras.newIOMInnovator(); return true; } "
                "catch(e) { return false; }"
            ))
            print(f" ready ({time.time()-t0:.1f}s)")
            self._ready = True
            return True
        except Exception as e:
            err = str(e).lower()
            if "no such window" in err or "target window already closed" in err:
                self._user_closed = True
            print(f" error: {e}")
            self._driver = None
            return False

    def _bring_to_front(self):
        """Bring the OnePDM browser window to the foreground."""
        if not self._driver:
            return
        try:
            self._driver.switch_to.window(self._driver.window_handles[0])
            self._driver.execute_script("window.focus();")
        except Exception:
            pass

    def _open_part_direct(self, part_number):
        """Look up Part by item_number via IOM and open it using Aras UI API."""
        driver = self._driver
        driver.switch_to.default_content()

        # Use cached ID if available
        cached_id = self._part_id_cache.get(part_number)
        if cached_id:
            driver.execute_script(
                "aras.uiShowItem('Part', arguments[0]);", cached_id
            )
            return True

        for attempt in range(5):
            result = driver.execute_script("""
                try {
                    var inn = aras.newIOMInnovator();
                    var part = inn.newItem('Part', 'get');
                    part.setProperty('item_number', arguments[0]);
                    part.setAttribute('select', 'id,item_number,name');
                    part = part.apply();
                    if (part.isError()) return 'error:' + part.getErrorString();
                    var id = part.getID();
                    aras.uiShowItem('Part', id);
                    return 'ok:' + id;
                } catch(e) {
                    return 'exception:' + e.message;
                }
            """, part_number)

            if result.startswith('ok:'):
                self._part_id_cache[part_number] = result[3:]
                return True
            if 'authorization' in result.lower() and attempt < 4:
                time.sleep(2)
                continue
            break

        print(f"({result}) ", end="", flush=True)
        return False

    def _navigate_and_expand_files(self):
        """After Part page opens, click More → Files, then expandAll in one flow.
        Returns the Part iframe element for reuse, or None."""
        from selenium.webdriver.common.by import By
        driver = self._driver

        # Phase 1: Poll for More button in Part iframe, click More → Files
        part_iframe = None
        for _ in range(80):
            time.sleep(0.2)
            driver.switch_to.default_content()
            for f in driver.find_elements(By.TAG_NAME, "iframe"):
                try:
                    if not f.is_displayed():
                        continue
                    driver.switch_to.default_content()
                    driver.switch_to.frame(f)
                    result = driver.execute_script("""
                        var btn = document.querySelector('button[title="More"]');
                        if (!btn) return 'no_more';
                        btn.click();
                        var lis = document.querySelectorAll('li.aras-list-item');
                        for (var i = 0; i < lis.length; i++) {
                            var label = lis[i].querySelector('.aras-list-item__label');
                            if (label && label.textContent.trim() === 'Files') {
                                lis[i].click();
                                return 'ok';
                            }
                        }
                        return 'no_files';
                    """)
                    if result == 'ok':
                        part_iframe = f
                        break
                except Exception:
                    continue
            if part_iframe:
                break

        if not part_iframe:
            driver.switch_to.default_content()
            return

        # Phase 2: Already inside Part iframe — poll for dialog directly (no re-scan)
        dialog_iframe = None
        for _ in range(40):
            time.sleep(0.2)
            try:
                # Stay in Part iframe context
                dlg = driver.find_elements(By.CSS_SELECTOR, 'dialog.aras-dialog[open]')
                if dlg:
                    inner = dlg[0].find_elements(By.TAG_NAME, "iframe")
                    if inner:
                        dialog_iframe = inner[0]
                        break
            except Exception:
                # If stale, re-enter Part iframe
                try:
                    driver.switch_to.default_content()
                    driver.switch_to.frame(part_iframe)
                except Exception:
                    break

        if not dialog_iframe:
            driver.switch_to.default_content()
            return

        # Phase 3: Switch into dialog iframe, wait for data, expandAll
        try:
            driver.switch_to.frame(dialog_iframe)
        except Exception:
            driver.switch_to.default_content()
            return

        for _ in range(60):
            time.sleep(0.3)
            ready = driver.execute_script(
                "var rows = document.querySelectorAll('tr.aras-grid-row'); "
                "if (rows.length === 0) return false; "
                "var grid = document.querySelector('.aras-grid'); "
                "if (grid && grid.textContent.includes('Processing')) return false; "
                "return true;"
            )
            if ready:
                break

        driver.execute_script(
            "if (typeof mainPage !== 'undefined') mainPage.expandAll();"
        )

        for _ in range(15):
            time.sleep(0.3)
            if driver.execute_script(
                "return document.querySelectorAll('.aras-treegrid-expand-button.aras-icon-plus').length;"
            ) == 0:
                break

        driver.switch_to.default_content()

    def _find_file_urls(self, filename_pattern=None):
        """After Files are expanded, extract file download URLs from the grid.
        Returns list of dicts: [{"name": "...", "url": "..."}]."""
        from selenium.webdriver.common.by import By
        driver = self._driver
        driver.switch_to.default_content()

        files = []
        for f in driver.find_elements(By.TAG_NAME, "iframe"):
            try:
                if not f.is_displayed():
                    continue
                driver.switch_to.default_content()
                driver.switch_to.frame(f)
                # Look for dialogs with file grids
                dlgs = driver.find_elements(
                    By.CSS_SELECTOR, 'dialog.aras-dialog[open]')
                for dlg in dlgs:
                    inner = dlg.find_elements(By.TAG_NAME, "iframe")
                    if not inner:
                        continue
                    driver.switch_to.frame(inner[0])
                    result = driver.execute_script("""
                        var files = [];
                        var rows = document.querySelectorAll('tr.aras-grid-row');
                        for (var i = 0; i < rows.length; i++) {
                            var cells = rows[i].querySelectorAll('td');
                            for (var j = 0; j < cells.length; j++) {
                                var text = cells[j].textContent.trim();
                                if (text && (text.toLowerCase().endsWith('.pdf')
                                    || text.toLowerCase().endsWith('.doc')
                                    || text.toLowerCase().endsWith('.docx'))) {
                                    files.push(text);
                                }
                            }
                        }
                        return files;
                    """)
                    if result:
                        for name in result:
                            files.append({"name": name})
                    driver.switch_to.parent_frame()
            except Exception:
                pass
        driver.switch_to.default_content()

        if filename_pattern:
            import re
            pat = re.compile(filename_pattern, re.IGNORECASE)
            files = [f for f in files if pat.search(f["name"])]

        return files

    def _download_file_via_iom(self, part_number, dest_dir):
        """Download the datasheet file for a part using IOM API.
        Looks for PDF files in the Part's file relationships.
        Returns the local file path on success, None otherwise."""
        import os
        driver = self._driver
        driver.switch_to.default_content()

        pn = part_number.upper().strip()
        cached_id = self._part_id_cache.get(pn)

        # Get file list via IOM
        result = driver.execute_script("""
            try {
                var inn = aras.newIOMInnovator();
                var part;
                if (arguments[1]) {
                    part = inn.getItemById('Part', arguments[1]);
                } else {
                    part = inn.newItem('Part', 'get');
                    part.setProperty('item_number', arguments[0]);
                    part = part.apply();
                }
                if (part.isError()) return 'error:' + part.getErrorString();

                // Get Part CAD relationships
                var files = [];
                var relTypes = ['Part Document', 'Part CAD'];
                for (var r = 0; r < relTypes.length; r++) {
                    var rels = part.getRelationships(relTypes[r]);
                    if (!rels || rels.isError()) continue;
                    var count = rels.getItemCount();
                    for (var i = 0; i < count; i++) {
                        var rel = rels.getItemByIndex(i);
                        var relItem = rel.getRelatedItem();
                        if (!relItem || relItem.isError()) continue;
                        var fname = relItem.getProperty('filename', '');
                        var fid = relItem.getID();
                        if (fname) files.push(fname + '|' + fid);
                    }
                }

                // Also try File relationships directly
                var fileRel = inn.newItem('Part', 'get');
                fileRel.setProperty('item_number', arguments[0]);
                fileRel.setAttribute('select', 'id');
                var managedFile = inn.newItem('File', 'get');
                managedFile.setAttribute('select', 'filename,id');
                fileRel.addRelationship(managedFile);
                var res2 = fileRel.apply();
                if (!res2.isError()) {
                    var fItems = res2.getRelationships();
                    if (fItems) {
                        var fc = fItems.getItemCount();
                        for (var j = 0; j < fc; j++) {
                            var fi = fItems.getItemByIndex(j);
                            var fn = fi.getProperty('filename', '');
                            if (fn) files.push(fn + '|' + fi.getID());
                        }
                    }
                }

                return 'ok:' + files.join(';;');
            } catch(e) {
                return 'exception:' + e.message;
            }
        """, pn, cached_id)

        if not result or not result.startswith('ok:'):
            print(f"  [onepdm] IOM file query failed: {result}")
            return None

        file_list_str = result[3:]
        if not file_list_str:
            print(f"  [onepdm] No files found for {pn}")
            return None

        # Find PDF files (prefer 'datasheet' or 'spec' in name)
        file_entries = []
        for entry in file_list_str.split(';;'):
            if '|' not in entry:
                continue
            fname, fid = entry.rsplit('|', 1)
            file_entries.append({"name": fname, "id": fid})

        pdf_files = [f for f in file_entries
                     if f["name"].lower().endswith('.pdf')]
        if not pdf_files:
            print(f"  [onepdm] No PDF files found for {pn}")
            return None

        # Prefer files with 'datasheet' or 'spec' in name
        target = None
        for f in pdf_files:
            name_lower = f["name"].lower()
            if 'datasheet' in name_lower or 'spec' in name_lower:
                target = f
                break
        if not target:
            target = pdf_files[0]  # Use first PDF

        # Download via Aras vault
        os.makedirs(dest_dir, exist_ok=True)
        dest_path = os.path.join(dest_dir, target["name"])
        if os.path.exists(dest_path):
            print(f"  [onepdm] Using cached: {target['name']}")
            return dest_path

        # Get download URL from vault
        download_url = driver.execute_script("""
            try {
                var inn = aras.newIOMInnovator();
                var fileItem = inn.getItemById('File', arguments[0]);
                if (fileItem.isError()) return 'error:' + fileItem.getErrorString();
                var vault = aras.getVault();
                var url = vault.getFileUrl(arguments[0]);
                return url || '';
            } catch(e) {
                return 'exception:' + e.message;
            }
        """, target["id"])

        if download_url and not download_url.startswith(('error:', 'exception:')):
            try:
                # Use Selenium to download the file
                import urllib.request
                opener = urllib.request.build_opener()
                # Get cookies from Selenium session
                cookies = driver.get_cookies()
                cookie_str = "; ".join(
                    f"{c['name']}={c['value']}" for c in cookies)
                opener.addheaders = [('Cookie', cookie_str)]
                resp = opener.open(download_url, timeout=60)
                with open(dest_path, 'wb') as out:
                    out.write(resp.read())
                print(f"  [onepdm] Downloaded: {target['name']}")
                return dest_path
            except Exception as e:
                print(f"  [onepdm] Download failed: {e}")

        # Fallback: use browser download via JavaScript
        try:
            driver.execute_script("""
                var a = document.createElement('a');
                a.href = arguments[0];
                a.download = arguments[1];
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
            """, download_url, target["name"])
            # Wait for download
            time.sleep(5)
            # Check default download folder
            dl_folder = os.path.join(os.path.expanduser("~"), "Downloads")
            dl_path = os.path.join(dl_folder, target["name"])
            if os.path.exists(dl_path):
                import shutil
                shutil.move(dl_path, dest_path)
                print(f"  [onepdm] Downloaded (browser): {target['name']}")
                return dest_path
        except Exception as e:
            print(f"  [onepdm] Browser download failed: {e}")

        return None

    # ---- Public API ----

    def open_part_page(self, part_number):
        """
        Open the OnePDM Part page and navigate to Get all files.
        Returns True if successful.
        """
        pn = part_number.upper().strip()

        with self._lock:
            # Reset per-request: allow new browser launch on each click
            self._user_closed = False

            if not self._ensure_browser():
                return False

            print(f"  [onepdm] {pn} ", end="", flush=True)
            t0 = time.time()

            try:
                # If already viewed, re-open item and navigate to Files
                if pn in self._part_id_cache:
                    driver = self._driver
                    driver.switch_to.default_content()
                    driver.execute_script(
                        "aras.uiShowItem('Part', arguments[0]);",
                        self._part_id_cache[pn]
                    )
                    self._navigate_and_expand_files()
                    self._bring_to_front()
                    dt = time.time() - t0
                    print(f"[{dt:.1f}s]")
                    return True

                t1 = time.time()
                ok = self._open_part_direct(pn)
                if ok:
                    self._navigate_and_expand_files()
                    self._bring_to_front()

                dt = time.time() - t0
                print(f"[{dt:.1f}s]")
                return ok
            except Exception as e:
                # Browser was closed mid-operation
                print(f" browser closed mid-operation", flush=True)
                self._driver = None
                self._ready = False
                self._user_closed = True
                return False

    def close(self):
        """Close the browser session."""
        if self._driver:
            try:
                self._driver.quit()
            except Exception:
                pass
            self._driver = None
            self._ready = False

    def _download_via_ui(self, pn, dest_dir):
        """Open Part → Files → expandAll → find PDF → download via vault URL.
        Returns local file path on success, None otherwise."""
        import os
        from selenium.webdriver.common.by import By
        driver = self._driver

        # Scrape the file grid for PDF names (files should already be expanded)
        files = self._find_file_urls(filename_pattern=r'\.pdf$')
        if not files:
            # Files tab may not be open yet — navigate to it
            self._navigate_and_expand_files()
            files = self._find_file_urls(filename_pattern=r'\.pdf$')
        if not files:
            return None

        print(f"  [onepdm] Found {len(files)} PDF(s) in file grid")

        # Prefer files with 'datasheet' or 'spec' in name
        target_name = None
        for f in files:
            name_lower = f["name"].lower()
            if 'datasheet' in name_lower or 'spec' in name_lower:
                target_name = f["name"]
                break
        if not target_name:
            target_name = files[0]["name"]

        os.makedirs(dest_dir, exist_ok=True)
        dest_path = os.path.join(dest_dir, target_name)
        if os.path.exists(dest_path):
            print(f"  [onepdm] Using cached: {target_name}")
            return dest_path

        # Try to get vault download URL via IOM File lookup
        driver.switch_to.default_content()
        iom_result = driver.execute_script("""
            try {
                var inn = aras.newIOMInnovator();
                var fileItem = inn.newItem("File", "get");
                fileItem.setProperty("filename", arguments[0]);
                fileItem = fileItem.apply();
                if (fileItem.isError()) return "";
                var fid = fileItem.getID();
                var serverBase = aras.getServerBaseURL();
                return serverBase + "vault/clonefile/" + fid;
            } catch(e) { return ""; }
        """, target_name)

        if iom_result:
            try:
                import urllib.request
                opener = urllib.request.build_opener()
                cookies = driver.get_cookies()
                cookie_str = "; ".join(
                    f"{c['name']}={c['value']}" for c in cookies)
                opener.addheaders = [('Cookie', cookie_str)]
                resp = opener.open(iom_result, timeout=60)
                data = resp.read()
                if len(data) > 100:  # sanity check it's not an error page
                    with open(dest_path, 'wb') as out:
                        out.write(data)
                    print(f"  [onepdm] Downloaded (vault): {target_name}")
                    return dest_path
                else:
                    print(f"  [onepdm] Vault returned tiny response ({len(data)}B)")
            except Exception as e:
                print(f"  [onepdm] Vault download failed: {e}")

        # Fallback: double-click the row in the grid to trigger browser download
        driver.switch_to.default_content()
        downloaded = False
        for f in driver.find_elements(By.TAG_NAME, "iframe"):
            try:
                if not f.is_displayed():
                    continue
                driver.switch_to.default_content()
                driver.switch_to.frame(f)
                dlgs = driver.find_elements(
                    By.CSS_SELECTOR, 'dialog.aras-dialog[open]')
                for dlg in dlgs:
                    inner = dlg.find_elements(By.TAG_NAME, "iframe")
                    if not inner:
                        continue
                    driver.switch_to.frame(inner[0])
                    clicked = driver.execute_script("""
                        var targetName = arguments[0].toLowerCase();
                        var rows = document.querySelectorAll('tr.aras-grid-row');
                        for (var i = 0; i < rows.length; i++) {
                            var cells = rows[i].querySelectorAll('td');
                            for (var j = 0; j < cells.length; j++) {
                                if (cells[j].textContent.trim().toLowerCase() === targetName) {
                                    var evt = new MouseEvent('dblclick', {bubbles: true});
                                    cells[j].dispatchEvent(evt);
                                    return true;
                                }
                            }
                        }
                        return false;
                    """, target_name)
                    if clicked:
                        downloaded = True
                    driver.switch_to.parent_frame()
                    if downloaded:
                        break
            except Exception:
                pass
            if downloaded:
                break
        driver.switch_to.default_content()

        if downloaded:
            dl_folder = os.path.join(os.path.expanduser("~"), "Downloads")
            for _ in range(30):
                time.sleep(1)
                dl_path = os.path.join(dl_folder, target_name)
                if os.path.exists(dl_path):
                    import shutil
                    shutil.move(dl_path, dest_path)
                    print(f"  [onepdm] Downloaded (browser): {target_name}")
                    return dest_path

        return None

    def download_datasheet(self, part_number, dest_dir):
        """Download the datasheet PDF for a part number.
        Uses UI Files tab (More → Files → expandAll) to find PDFs,
        then aras.downloadFile() to download via the vault client.
        Returns local file path on success, None otherwise."""
        pn = part_number.upper().strip()

        with self._lock:
            self._user_closed = False
            # Set browser download dir to dest_dir before launching
            import os
            os.makedirs(dest_dir, exist_ok=True)
            abs_dest = os.path.abspath(dest_dir)
            self._download_dir = abs_dest
            if not self._ensure_browser():
                return None

            # Ensure CDP download behavior points to our dir
            try:
                self._driver.execute_cdp_cmd("Page.setDownloadBehavior", {
                    "behavior": "allow",
                    "downloadPath": abs_dest
                })
            except Exception:
                pass

            print(f"  [onepdm] Downloading datasheet for {pn}...", flush=True)
            t0 = time.time()

            try:
                # Check cache first
                for fname in os.listdir(dest_dir):
                    if fname.lower().endswith('.pdf') and pn.lower() in fname.lower():
                        cached = os.path.join(dest_dir, fname)
                        print(f"  [onepdm] Using cached: {fname}")
                        dt = time.time() - t0
                        print(f"  [onepdm] Datasheet ready [{dt:.1f}s]")
                        return cached

                # Open part page and navigate to Files (expand all)
                if pn in self._part_id_cache:
                    self._driver.switch_to.default_content()
                    self._driver.execute_script(
                        "aras.uiShowItem('Part', arguments[0]);",
                        self._part_id_cache[pn]
                    )
                else:
                    ok = self._open_part_direct(pn)
                    if not ok:
                        dt = time.time() - t0
                        print(f"  [onepdm] Part not found [{dt:.1f}s]")
                        return None

                self._navigate_and_expand_files()

                # Find PDFs in the file grid
                pdf_files = self._find_file_urls(filename_pattern=r'\.pdf$')
                if not pdf_files:
                    dt = time.time() - t0
                    print(f"  [onepdm] No datasheet found [{dt:.1f}s]")
                    return None

                print(f"  [onepdm] Found {len(pdf_files)} PDF(s) in file grid")
                for i, f in enumerate(pdf_files):
                    print(f"    [{i}] {f['name']}")

                # Prefer files with 'datasheet', 'DS', or 'spec' in name
                target_name = None
                for f in pdf_files:
                    name_lower = f["name"].lower()
                    if 'datasheet' in name_lower or 'spec' in name_lower:
                        target_name = f["name"]
                        break
                if not target_name:
                    # Check for 'DS' or 'ds' as word boundary
                    import re as _re
                    for f in pdf_files:
                        if _re.search(r'(?:^|[_\-])ds(?:[_\-\.]|$)',
                                      f["name"], _re.IGNORECASE):
                            target_name = f["name"]
                            break
                if not target_name:
                    target_name = pdf_files[0]["name"]

                print(f"  [onepdm] Target: {target_name}")
                dest_path = os.path.join(abs_dest, target_name)
                if os.path.exists(dest_path):
                    print(f"  [onepdm] Using cached: {target_name}")
                    dt = time.time() - t0
                    print(f"  [onepdm] Datasheet ready [{dt:.1f}s]")
                    return dest_path

                # Download via aras.downloadFile() with proper IOM node
                self._driver.switch_to.default_content()
                dl_result = self._driver.execute_script("""
                    try {
                        var inn = aras.newIOMInnovator();
                        var fileItem = inn.newItem("File", "get");
                        fileItem.setProperty("filename", arguments[0]);
                        fileItem = fileItem.apply();
                        if (fileItem.isError()) return "FILE_ERROR:" + fileItem.getErrorString();
                        // Handle multiple results — use the first one
                        var count = fileItem.getItemCount();
                        var target = (count > 1) ? fileItem.getItemByIndex(0) : fileItem;
                        var fid = target.getID();
                        aras.downloadFile(target.node, arguments[0]);
                        return "OK:fid=" + fid + ",count=" + count;
                    } catch(e) { return "ERR:" + e.message; }
                """, target_name)
                print(f"  [onepdm] downloadFile result: {dl_result}")

                if dl_result and dl_result.startswith("OK"):
                    # Wait for download to complete — check both dest and Downloads
                    print(f"  [onepdm] Waiting for download in {abs_dest} or Downloads...")
                    for i in range(30):
                        time.sleep(1)
                        if os.path.exists(dest_path):
                            sz = os.path.getsize(dest_path)
                            if sz > 100:
                                time.sleep(1)
                                if os.path.getsize(dest_path) == sz:
                                    dt = time.time() - t0
                                    print(f"  [onepdm] Downloaded: {target_name}")
                                    print(f"  [onepdm] Datasheet ready [{dt:.1f}s]")
                                    return dest_path
                        dl_folder = os.path.join(
                            os.path.expanduser("~"), "Downloads")
                        dl_path = os.path.join(dl_folder, target_name)
                        if os.path.exists(dl_path):
                            import shutil
                            shutil.move(dl_path, dest_path)
                            dt = time.time() - t0
                            print(f"  [onepdm] Downloaded: {target_name}")
                            print(f"  [onepdm] Datasheet ready [{dt:.1f}s]")
                            return dest_path
                        # Also scan Downloads for any new PDF
                        if i == 15:
                            recent = []
                            for fn in os.listdir(dl_folder):
                                fp = os.path.join(dl_folder, fn)
                                if fn.lower().endswith('.pdf'):
                                    mt = os.path.getmtime(fp)
                                    if mt > t0:
                                        recent.append(fn)
                            if recent:
                                print(f"  [onepdm] Recent PDFs in Downloads: {recent}")
                    print(f"  [onepdm] Download timed out after 30s")

                # Fallback: try double-click download via UI grid
                result = self._download_via_ui(pn, dest_dir)
                if result:
                    dt = time.time() - t0
                    print(f"  [onepdm] Datasheet ready [{dt:.1f}s]")
                    return result

                dt = time.time() - t0
                print(f"  [onepdm] No datasheet found [{dt:.1f}s]")
                return None
            except Exception as e:
                print(f"  [onepdm] Download error: {e}")
                self._driver = None
                self._ready = False
                return None

    @property
    def is_ready(self):
        return self._ready


# Singleton
_client = None


def get_client(headless=False):
    """Get or create the singleton OnePDM client."""
    global _client
    if _client is None:
        _client = OnePdmClient(headless=headless)
    return _client

