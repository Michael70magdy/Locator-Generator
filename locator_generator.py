import csv
import re
import threading
import time
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import traceback

try:
    from bs4 import BeautifulSoup
except ImportError:
    BeautifulSoup = None

try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By as SelBy
    from selenium.webdriver.chrome.options import Options as ChromeOptions
    HAS_SELENIUM = True
except ImportError:
    HAS_SELENIUM = False

try:
    import openpyxl
    import openpyxl.utils as openpyxl_utils
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

TITLE = "Locator Generator"

USER_MANUAL_TEXT = """======================================================================
                LOCATOR GENERATOR  - USER MANUAL
======================================================================

Welcome to the Locator Generator user manual. This guide 
covers installation commands, core features, interactive browser tools, 
and best practices for generating robust and stable test locators.

----------------------------------------------------------------------
1. GETTING STARTED & INSTALLATION COMMANDS
----------------------------------------------------------------------

Prerequisites:
• Python 3.8 or higher installed on your system.

Required Packages Installation:
Open your Windows Command Prompt (cmd), PowerShell, or terminal and run 
the following command to install all necessary dependencies:

    pip install beautifulsoup4 selenium openpyxl lxml

Alternatively, if you use Python's explicit runner launcher:

    py -m pip install beautifulsoup4 selenium openpyxl lxml

Verification Command:
To double-check that everything installed correctly:

    python -c "import bs4, selenium, openpyxl, lxml; print('All dependencies installed successfully!')"

----------------------------------------------------------------------
2. DASHBOARD INTERFACE OVERVIEW
----------------------------------------------------------------------

The application features a clean dashboard split into key control sections:

• Live Browser & Noise Filtering Control:
  - Ignore Patterns: Specify comma-separated keywords or prefixes 
    (e.g., css-, sc-, ng-star-inserted) to automatically filter out 
    volatile framework classes.
  - Target URL: Enter the target web page URL for live testing or whole-page scanning.
  - Open Picker: Launches an interactive Chrome browser window to visually select elements.
• HTML Input Fragment: Paste raw HTML markup of a targeted web element here if you aren't using the live browser picker.
• Action Bar: Buttons to Generate Locators, Scan Whole Page, Load Example, Clear, Copy All Java, and Export to Excel.
• Candidate Locators & Stability Score Table: Displays generated XPath and CSS locators ranked by stability scores along with match counts and notes.

----------------------------------------------------------------------
3. CORE FEATURES & WORKFLOWS
----------------------------------------------------------------------

Workflow A: Using HTML Fragments (Static Analysis)
1. Click "Load Example" to populate a sample HTML form snippet.
2. Adjust or paste your own HTML element fragment into the text box.
3. Click "Generate Locators".
4. Review the candidate list in the results table, ordered from most stable (test attributes, unique IDs) to least stable (relative DOM paths).
5. Double-click any row or use the buttons below to copy the locator expression or its Java By statement.

Workflow B: Live Page Scanning
1. Enter your target URL in the Target URL field (e.g., https://example.com).
2. Click "Scan Whole Page". 
3. The app will launch a headless browser instance, parse interactive elements across the page, and populate the table with top locators.

Workflow C: Interactive Browser Picker
1. Enter a valid URL and click "Open Picker".
2. A Chrome browser window will open to your page. As you hover over elements, a red visual highlight box will appear around them.
3. Click on any element in the browser window. Its HTML fragment will be captured automatically, and the application will instantly generate and score its locators.
4. Use the Browser Tools (Back, Forward, Refresh) or toggle Pick Mode Active on/off if you need to navigate through login flows before picking an element.

----------------------------------------------------------------------
4. FILTERING UNWANTED NOISE (IGNORE PATTERNS)
----------------------------------------------------------------------

To keep your test automation scripts clean and maintainable, use the Ignore Patterns text box:
• Enter unwanted patterns separated by commas (e.g., ng-star-, volatile-, temp-).
• When generating or scanning, the application performs case-insensitive checks and automatically skips any candidate expressions containing these patterns.

----------------------------------------------------------------------
5. EXPORTING REPORTS
----------------------------------------------------------------------

When you have generated your locator suite:
1. Click "Export to Excel".
2. Choose your save destination and filename.
3. The tool generates a clean spreadsheet containing all generated metrics with standard gridlines and theme compatibility."""

def xp_literal(s):
    if "'" not in s: return "'" + s + "'"
    if '"' not in s: return '"' + s + '"'
    return "concat(" + ', "\'", '.join("'" + p + "'" for p in s.split("'")) + ")"

def esc(s):
    return s.replace("\\\\", "\\\\\\\\").replace('"', '\\\\"').replace("\\n", "\\\\n").replace("\\r", "\\\\r")

def cls_list(el):
    v = el.get("class", [])
    return v.split() if isinstance(v, str) else list(v)

def rank(score):
    return "Excellent" if score >= 90 else "Strong" if score >= 75 else "Fair" if score >= 55 else "Weak"

def generate_candidates(soup, el, ignored_patterns=None):
    if ignored_patterns is None:
        ignored_patterns = []

    tag = el.name.lower()
    candidates, seen = [], set()
    text = " ".join(el.stripped_strings)
    attrs = {}
    for k, v in el.attrs.items():
        attrs[k] = " ".join(v) if isinstance(v, list) else str(v)
    
    def add(kind, strategy, expr, score, note):
        key = (kind, expr)
        if not expr or key in seen: return
        # Check against ignored patterns
        if any(re.search(pat, expr, re.IGNORECASE) for pat in ignored_patterns):
            return
        seen.add(key)
        count = None
        if kind == "CSS":
            try: count = len(soup.select(expr))
            except Exception: pass
        elif kind == "XPath":
            try:
                from lxml import html
                count = len(html.fromstring(str(soup)).xpath(expr))
            except Exception: pass
        candidates.append(dict(kind=kind, strategy=strategy, expr=expr, score=score,
                              rank=rank(score), note=note, count=count))

    for key, score, note in [
        ("data-testid", 99, "Dedicated test attribute; highly stable."),
        ("data-test", 98, "Dedicated test attribute; highly stable."),
        ("data-qa", 97, "Dedicated QA attribute; highly stable."),
        ("id", 94, "Excellent if ID is stable and unique."),
        ("name", 91, "Semantic form attribute; stable."),
        ("aria-label", 88, "Accessible label; stable compared to classes."),
        ("placeholder", 81, "User-facing copy; may change."),
        ("title", 76, "Tooltips; may change."),
        ("role", 70, "Role attribute; broad without context."),
        ("type", 58, "Control type; shared across elements.")
    ]:
        if attrs.get(key):
            val = attrs[key]
            if any(re.search(pat, val, re.IGNORECASE) for pat in ignored_patterns): continue
            add("CSS", key, f'{tag}[{key}="{val.replace(chr(34), chr(92)+chr(34))}"]', score, note)
            add("XPath", key, f"//{tag}[@{key}={xp_literal(val)}]", score, note)
            if key == "id":
                if re.match(r"^[A-Za-z_][\\w-]*$", val):
                    add("CSS", "ID shorthand", "#" + val, 96, "Concise ID selector.")
                add("XPath", "ID shorthand", f"//*[@id={xp_literal(val)}]", 95, "Concise ID-based XPath.")

    classes = cls_list(el)
    safe = []
    for c in classes:
        if not re.match(r"^-?[_a-zA-Z][_a-zA-Z0-9-]*$", c): continue
        if any(re.search(pat, c, re.IGNORECASE) for pat in ignored_patterns): continue
        safe.append(c)
    
    if safe:
        selector = tag + "".join("." + c for c in safe)
        volatile = any(re.search(r"ng-star-inserted|ng-option-marked|active|selected|css-|sc-", c, re.IGNORECASE) for c in safe)
        add("CSS", "Tag + all classes", selector, 54 if volatile else 73,
            "Volatile framework classes detected." if volatile else "Class-based selector.")
        for c in safe[:4]:
            add("CSS", "Single class", "." + c, 45, "Single class token.")
            add("XPath", "Class token", f"//{tag}[contains(concat(' ', normalize-space(@class), ' '), {xp_literal(' ' + c + ' ')})]", 52, "Class token XPath match.")

    if attrs.get("type"):
        val = attrs["type"]
        if not any(re.search(pat, val, re.IGNORECASE) for pat in ignored_patterns):
            add("CSS", "Tag + type", f'{tag}[type="{val}"]', 58, "Tag and type combination.")
            add("XPath", "Tag + type", f"//{tag}[@type={xp_literal(val)}]", 58, "Tag and type combination.")

    if text:
        compact = re.sub(r"\\s+", " ", text).strip()
        if len(compact) <= 100:
            add("XPath", "Exact normalized text", f"//{tag}[normalize-space(.)={xp_literal(compact)}]", 78, "Exact text match.")
            if len(compact) >= 4:
                add("XPath", "Contains text", f"//{tag}[contains(normalize-space(.), {xp_literal(compact[:60])})]", 60, "Partial text match.")
            if tag == "a":
                add("XPath", "Link text", f"//a[normalize-space(.)={xp_literal(compact)}]", 79, "Anchor link text.")
            if safe and len(compact) <= 60:
                add("XPath", "Class + Text combination", f"//{tag}[contains(@class, {xp_literal(safe[0])}) and normalize-space(.)={xp_literal(compact)}]", 86, "High uniqueness class + text combo.")

    parent = el.parent
    if parent and getattr(parent, "name", None) and parent.name != "[document]":
        p_attrs = {}
        for pk, pv in parent.attrs.items():
            p_attrs[pk] = " ".join(pv) if isinstance(pv, list) else str(pv)
        for p_key in ("id", "data-testid", "data-test", "data-qa", "name", "class"):
            if p_val := p_attrs.get(p_key):
                p_val_clean = p_val.strip()
                if any(re.search(pat, p_val_clean, re.IGNORECASE) for pat in ignored_patterns): continue
                if p_key == "id":
                    add("CSS", "Parent ID + Target Tag", f"#{p_val_clean} > {tag}", 84, "Scoped under parent ID.")
                    add("XPath", "Parent ID + Target Tag", f"//*[@id={xp_literal(p_val_clean)}]//{tag}", 83, "Scoped under parent ID.")
                    if safe:
                        add("CSS", "Parent ID + Target Class", f"#{p_val_clean} .{safe[0]}", 87, "Parent ID + class combo.")
                    break
                elif p_key in ("data-testid", "data-test", "data-qa"):
                    add("XPath", f"Parent {p_key} + Target Tag", f"//*[@{p_key}={xp_literal(p_val_clean)}]//{tag}", 82, f"Scoped via parent {p_key}.")
                    break

    add("CSS", "Tag name", tag, 30, "Generic tag selector.")
    add("XPath", "Tag name", f"//{tag}", 30, "Generic tag XPath selector.")

    path, node = [], el
    while node and getattr(node, "name", None) and node.name != "[document]" and len(path) < 6:
        parent_node = node.parent
        siblings = parent_node.find_all(node.name, recursive=False) if parent_node and getattr(parent_node, "find_all", None) else [node]
        seg = node.name
        if len(siblings) > 1: seg += f"[{siblings.index(node)+1}]"
        path.insert(0, seg)
        node = parent_node
    if path:
        add("XPath", "Relative DOM path", "/" + "/".join(path), 35, "Brittle structural path.")

    candidates.sort(key=lambda x: (-x["score"], x["kind"], x["strategy"]))
    return candidates

class App:
    def __init__(self, root):
        self.root = root
        root.title(TITLE); root.geometry("1300x900"); root.minsize(1000, 720)
        self.rows = []
        self.picker_driver = None
        self.picker_active = False
        self.is_dark_mode = True
        
        self.themes = {
            "dark": {
                "bg": "#1E222B", "card_bg": "#262B36", "text": "#F0F2F5",
                "subtext": "#9DA5B4", "accent": "#D93838", "border": "#3A4150",
                "input_bg": "#171A21", "btn_bg": "#313846", "tree_heading": "#313846"
            },
            "light": {
                "bg": "#F4F6F9", "card_bg": "#FFFFFF", "text": "#2C3E50",
                "subtext": "#606F7B", "accent": "#C52828", "border": "#CBD5E0",
                "input_bg": "#FFFFFF", "btn_bg": "#E2E8F0", "tree_heading": "#E2E8F0"
            }
        }

        self.main_container = ttk.Frame(root, padding=14)
        self.main_container.pack(fill="both", expand=True)

        header_frame = ttk.Frame(self.main_container)
        header_frame.pack(fill="x", pady=(0, 10))
        
        title_box = ttk.Frame(header_frame)
        title_box.pack(side="left", anchor="w")
        self.title_label = ttk.Label(title_box, text="🎯Locator Generator", font=("Segoe UI", 16, "bold"))
        self.title_label.pack(anchor="w")
        self.sub_label = ttk.Label(title_box, text="Interactive QA automation assistant with live browser picking, smart scoring, and noise filtering.", font=("Segoe UI", 10))
        self.sub_label.pack(anchor="w", pady=(2, 0))

        self.theme_btn = ttk.Button(header_frame, text="☀️ Light Mode", command=self.toggle_theme)
        self.theme_btn.pack(side="right", anchor="e")

        self.notebook = ttk.Notebook(self.main_container)
        self.notebook.pack(fill="both", expand=True)

        self.home_tab = ttk.Frame(self.notebook, padding=10)
        self.about_tab = ttk.Frame(self.notebook, padding=10)

        self.notebook.add(self.home_tab, text="  🏠 Dashboard  ")
        self.notebook.add(self.about_tab, text="  ℹ️ About & User Manual  ")

        self.build_home_tab()
        self.build_about_tab()

        self.apply_theme()

    def build_home_tab(self):
        btn_state = "normal" if HAS_SELENIUM else "disabled"

        self.top_panel = ttk.LabelFrame(self.home_tab, text=" 🌐 Live Browser & Noise Filtering Control ", padding=10)
        self.top_panel.pack(fill="x", pady=(0, 10))
        
        ttk.Label(self.top_panel, text="Ignore Patterns:").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        self.ignore_entry = tk.Entry(self.top_panel, width=40, font=("Consolas", 10), relief="solid", bd=1)
        self.ignore_entry.insert(0, "css-, sc-, ng-star-inserted")
        self.ignore_entry.grid(row=0, column=1, sticky="w", padx=6, pady=4)

        ttk.Label(self.top_panel, text="Target URL:").grid(row=1, column=0, sticky="w", padx=6, pady=4)
        self.url_entry = tk.Entry(self.top_panel, width=50, font=("Consolas", 10), relief="solid", bd=1)
        self.url_entry.insert(0, "https://example.com")
        self.url_entry.grid(row=1, column=1, sticky="w", padx=6, pady=4)
        
        self.test_btn = ttk.Button(self.top_panel, text="🧪 Test Locator", command=self.test_live_url, state=btn_state)
        self.test_btn.grid(row=1, column=2, padx=6, pady=4)

        self.pick_btn = ttk.Button(self.top_panel, text="🔍 Open Picker", command=self.pick_element_interactive, state=btn_state, style="Accent.TButton")
        self.pick_btn.grid(row=1, column=3, padx=6, pady=4)

        self.close_browser_btn = ttk.Button(self.top_panel, text="❌ Close Browser", command=self.close_picker_browser, state="disabled")
        self.close_browser_btn.grid(row=1, column=4, padx=6, pady=4)

        nav_frame = ttk.Frame(self.top_panel)
        nav_frame.grid(row=2, column=0, columnspan=5, sticky="w", padx=6, pady=(8, 4))
        
        self.pick_mode_var = tk.BooleanVar(value=True)
        self.pick_mode_chk = ttk.Checkbutton(nav_frame, text="Pick Mode Active", variable=self.pick_mode_var, command=self.update_pick_mode, state="disabled")
        self.pick_mode_chk.pack(side="left", padx=(0, 20))

        ttk.Label(nav_frame, text="Browser Tools:").pack(side="left", padx=(0, 6))
        self.nav_back_btn = ttk.Button(nav_frame, text="◀ Back", command=self.browser_back, state="disabled")
        self.nav_back_btn.pack(side="left", padx=2)
        
        self.nav_fwd_btn = ttk.Button(nav_frame, text="Forward ▶", command=self.browser_forward, state="disabled")
        self.nav_fwd_btn.pack(side="left", padx=2)
        
        self.nav_refresh_btn = ttk.Button(nav_frame, text="🔄 Refresh", command=self.browser_refresh, state="disabled")
        self.nav_refresh_btn.pack(side="left", padx=2)

        if not HAS_SELENIUM:
            self.warning_lbl = ttk.Label(self.top_panel, text="⚠️ (Install 'selenium' package to enable live features)")
            self.warning_lbl.grid(row=0, column=2, columnspan=3, padx=6)

        frame = ttk.LabelFrame(self.home_tab, text=" 📝 HTML Input Fragment (or captured live element) ", padding=10)
        frame.pack(fill="x", pady=(0, 10))
        
        self.html = tk.Text(frame, height=5, wrap="none", font=("Consolas", 10), relief="solid", bd=1, highlightthickness=1)
        self.html.pack(side="left", fill="both", expand=True)
        sy = ttk.Scrollbar(frame, command=self.html.yview); sy.pack(side="right", fill="y"); self.html.configure(yscrollcommand=sy.set)
        self.html.insert("1.0", '<div>Paste Your Targeted Element Here</div>')
        
        bar = ttk.Frame(self.home_tab)
        bar.pack(fill="x", pady=(0, 8))
        ttk.Button(bar, text="⚡ Generate Locators", command=self.generate, style="Accent.TButton").pack(side="left")
        ttk.Button(bar, text="🌐 Scan Whole Page", command=self.scan_whole_page, state=btn_state).pack(side="left", padx=6)
        ttk.Button(bar, text="💡 Load Example", command=self.example).pack(side="left", padx=6)
        ttk.Button(bar, text="🗑 Clear", command=self.clear).pack(side="left")
        ttk.Button(bar, text="📋 Copy All Java", command=self.copy_all).pack(side="left", padx=(14, 0))
        ttk.Button(bar, text="📊 Export to Excel", command=self.export_excel).pack(side="left", padx=6)
        
        self.status = ttk.Label(bar, text="Ready"); self.status.pack(side="right")
        
        result = ttk.LabelFrame(self.home_tab, text=" 🎯 Candidate Locators & Stability Score ", padding=10)
        result.pack(fill="both", expand=True, pady=(0, 10))
        cols = ("kind", "strategy", "expr", "rank", "count", "note")
        self.tree = ttk.Treeview(result, columns=cols, show="headings", selectmode="browse")
        headings = {"kind":"Type", "strategy":"Strategy", "expr":"Locator Expression", "rank":"Stability Score", "count":"Matches", "note":"Notes"}
        widths = {"kind":68, "strategy":160, "expr":380, "rank":105, "count":95, "note":340}
        for c in cols: self.tree.heading(c, text=headings[c]); self.tree.column(c, width=widths[c], minwidth=55, anchor="w")
        y = ttk.Scrollbar(result, orient="vertical", command=self.tree.yview); x = ttk.Scrollbar(result, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=y.set, xscrollcommand=x.set)
        self.tree.grid(row=0, column=0, sticky="nsew"); y.grid(row=0, column=1, sticky="ns"); x.grid(row=1, column=0, sticky="ew")
        result.rowconfigure(0, weight=1); result.columnconfigure(0, weight=1)
        
        self.tree.bind("<<TreeviewSelect>>", self.preview)
        self.tree.bind("<Double-1>", lambda e: self.copy_expr())
        
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Copy Expression", command=self.copy_expr)
        self.context_menu.add_command(label="Copy Java By Statement", command=self.copy_java)

        def show_context_menu(event):
            item = self.tree.identify_row(event.y)
            if item:
                self.tree.selection_set(item)
                self.context_menu.post(event.x_root, event.y_root)

        self.tree.bind("<Button-3>", show_context_menu)
        self.tree.bind("<Button-2>", show_context_menu)

        bottom = ttk.Frame(self.home_tab)
        bottom.pack(fill="x", pady=(0, 4))
        ttk.Button(bottom, text="Copy Selected Expression", command=self.copy_expr).pack(side="left")
        ttk.Button(bottom, text="Copy Selected Java By", command=self.copy_java).pack(side="left", padx=6)
        self.preview_var = tk.StringVar(value="Select a row to preview Selenium Java code.")
        self.preview_lbl = ttk.Label(bottom, textvariable=self.preview_var)
        self.preview_lbl.pack(side="left", padx=10)

    def build_about_tab(self):
        header_frame = ttk.Frame(self.about_tab)
        header_frame.pack(fill="x", pady=(0, 10))
        
        ttk.Label(header_frame, text="📘 Project Overview & User Manual", font=("Segoe UI", 16, "bold")).pack(side="left", anchor="w")
        ttk.Button(header_frame, text="📋 Copy Manual to Clipboard", command=lambda: self.copy(USER_MANUAL_TEXT)).pack(side="right", anchor="e")

        manual_frame = ttk.LabelFrame(self.about_tab, text=" User Guide & Installation Documentation ", padding=10)
        manual_frame.pack(fill="both", expand=True)

        self.manual_text = tk.Text(manual_frame, wrap="word", font=("Consolas", 10), relief="solid", bd=1, highlightthickness=1)
        self.manual_text.pack(side="left", fill="both", expand=True)
        
        sy = ttk.Scrollbar(manual_frame, command=self.manual_text.yview)
        sy.pack(side="right", fill="y")
        self.manual_text.configure(yscrollcommand=sy.set)
        
        self.manual_text.insert("1.0", USER_MANUAL_TEXT)
        self.manual_text.config(state="disabled")

    def toggle_theme(self):
        self.is_dark_mode = not self.is_dark_mode
        self.apply_theme()

    def apply_theme(self):
        t = self.themes["dark" if self.is_dark_mode else "light"]
        self.root.configure(bg=t["bg"])
        
        style = ttk.Style()
        try: style.theme_use("clam")
        except Exception: pass

        style.configure(".", background=t["bg"], foreground=t["text"], font=("Segoe UI", 10))
        style.configure("TLabel", background=t["bg"], foreground=t["text"])
        style.configure("Header.TLabel", font=("Segoe UI", 16, "bold"), foreground=t["accent"], background=t["bg"])
        style.configure("SubHeader.TLabel", font=("Segoe UI", 10), foreground=t["subtext"], background=t["bg"])
        style.configure("TLabelframe", background=t["bg"], foreground=t["accent"], bordercolor=t["border"], borderwidth=2, relief="solid")
        style.configure("TLabelframe.Label", background=t["bg"], foreground=t["accent"], font=("Segoe UI", 10, "bold"))
        style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=6, background=t["btn_bg"], foreground=t["text"], relief="flat")
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), background=t["accent"], foreground="#FFFFFF", relief="flat")
        
        style.configure("TNotebook", background=t["bg"], borderwidth=0)
        style.configure("TNotebook.Tab", background=t["btn_bg"], foreground=t["text"], padding=[12, 6], font=("Segoe UI", 10, "bold"))
        style.map("TNotebook.Tab", background=[("selected", t["card_bg"])], foreground=[("selected", t["accent"])])

        style.configure("Treeview", background=t["card_bg"], foreground=t["text"], fieldbackground=t["card_bg"], rowheight=28, bordercolor=t["border"], borderwidth=1)
        style.configure("Treeview.Heading", background=t["tree_heading"], foreground=t["text"], font=("Segoe UI", 10, "bold"), relief="flat")
        style.map("Treeview", background=[("selected", t["accent"])], foreground=[("selected", "#FFFFFF")])

        self.ignore_entry.config(bg=t["input_bg"], fg=t["text"], insertbackground=t["text"], highlightbackground=t["border"], highlightcolor=t["accent"])
        self.url_entry.config(bg=t["input_bg"], fg=t["text"], insertbackground=t["text"], highlightbackground=t["border"], highlightcolor=t["accent"])
        self.html.config(bg=t["input_bg"], fg=t["text"], insertbackground=t["text"], highlightbackground=t["border"], highlightcolor=t["accent"])
        
        self.manual_text.config(state="normal")
        self.manual_text.config(bg=t["input_bg"], fg=t["text"], insertbackground=t["text"], highlightbackground=t["border"], highlightcolor=t["accent"])
        self.manual_text.config(state="disabled")

        self.context_menu.config(bg=t["card_bg"], fg=t["text"], activebackground=t["accent"], activeforeground="#FFFFFF")

        if self.is_dark_mode: self.theme_btn.config(text="☀️ Light Mode")
        else: self.theme_btn.config(text="🌙 Dark Mode")

        self.title_label.configure(style="Header.TLabel")
        self.sub_label.configure(style="SubHeader.TLabel")
        self.preview_lbl.configure(foreground=t["subtext"])

    def example(self):
        self.html.delete("1.0", "end")
        self.html.insert("1.0", '<div class="form-group">\n  <label for="username">Username</label>\n  <input id="username" name="userName" class="form-control" placeholder="Enter username" type="text">\n</div>')

    def clear(self):
        self.html.delete("1.0", "end")
        for i in self.tree.get_children(): self.tree.delete(i)
        self.rows = []; self.status.config(text="Cleared"); self.preview_var.set("Select a row to preview Selenium Java code.")

    def get_ignored_patterns(self):
        raw = self.ignore_entry.get().strip()
        if not raw: return []
        return [p.strip() for p in raw.split(",") if p.strip()]

    def generate(self):
        if BeautifulSoup is None:
            messagebox.showerror("Dependency missing", "Install BeautifulSoup with: py -m pip install beautifulsoup4"); return
        markup = self.html.get("1.0", "end").strip()
        if not markup:
            messagebox.showwarning("Missing HTML", "Paste a complete HTML element first."); return
        try:
            soup = BeautifulSoup(markup, "html.parser")
            els = soup.find_all(True)
            if not els: raise ValueError("No HTML elements found. Paste a complete element.")
            preferred = [e for e in els if any(e.get(k) for k in ("data-testid", "data-test", "data-qa", "id", "name", "aria-label", "placeholder"))]
            target = preferred[0] if preferred else els[-1]
            
            ignored = self.get_ignored_patterns()
            self.rows = generate_candidates(soup, target, ignored_patterns=ignored)
            
            for i in self.tree.get_children(): self.tree.delete(i)
            for i, c in enumerate(self.rows):
                count = str(c["count"]) if c["count"] is not None else "N/A"
                self.tree.insert("", "end", iid=str(i), values=(c["kind"], c["strategy"], c["expr"], f'{c["rank"]} ({c["score"]}/100)', count, c["note"]))
            self.status.config(text=f"Target: <{target.name}> · {len(self.rows)} candidates generated")
        except Exception as ex: messagebox.showerror("Could not parse HTML", str(ex))

    def scan_whole_page(self):
        if BeautifulSoup is None:
            messagebox.showerror("Dependency missing", "Install BeautifulSoup with: py -m pip install beautifulsoup4")
            return
        url = self.url_entry.get().strip()
        if not url:
            messagebox.showwarning("Missing URL", "Please enter a valid URL to scan.")
            return
        if not HAS_SELENIUM:
            messagebox.showerror("Selenium Missing", "Selenium is required to scan live pages.")
            return

        self.status.config(text="Scanning whole page elements via Selenium...")
        self.root.update_idletasks()

        def background_scan():
            driver = None
            try:
                options = ChromeOptions()
                options.add_argument("--headless")
                driver = webdriver.Chrome(options=options)
                driver.get(url)
                html_content = driver.page_source
                driver.quit()
                driver = None

                soup = BeautifulSoup(html_content, "html.parser")
                ignored = self.get_ignored_patterns()
                
                elements = soup.find_all(['input', 'button', 'a', 'select', 'textarea', 'form'])
                other_elements = soup.find_all(lambda tag: tag.has_attr('id') or tag.has_attr('data-testid') or tag.has_attr('data-qa'))
                
                all_els = list(dict.fromkeys(elements + other_elements))
                
                all_candidates = []
                for el in all_els[:150]:
                    cands = generate_candidates(soup, el, ignored_patterns=ignored)
                    if cands:
                        top_cand = cands[0]
                        top_cand['note'] = f"<{el.name}> | " + top_cand['note']
                        all_candidates.append(top_cand)

                self.root.after(0, lambda: self.populate_scan_results(all_candidates, url))
            except Exception as ex:
                err_msg = str(ex)
                self.root.after(0, lambda msg=err_msg: messagebox.showerror("Scan Error", msg))
                self.root.after(0, lambda: self.status.config(text="Page scan failed."))
            finally:
                if driver:
                    try: driver.quit()
                    except: pass

        threading.Thread(target=background_scan, daemon=True).start()

    def populate_scan_results(self, candidates, url):
        self.rows = candidates
        for i in self.tree.get_children(): 
            self.tree.delete(i)
        for i, c in enumerate(self.rows):
            count = str(c["count"]) if c["count"] is not None else "N/A"
            self.tree.insert("", "end", iid=str(i), values=(c["kind"], c["strategy"], c["expr"], f'{c["rank"]} ({c["score"]}/100)', count, c["note"]))
        self.status.config(text=f"Scanned {url}: Generated {len(self.rows)} top locators across the page.")

    def export_excel(self):
        if not self.rows:
            messagebox.showwarning("No Data", "No generated locators available to export.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx"), ("CSV File", "*.csv"), ("All Files", "*.*")],
            title="Export Generated Locators"
        )
        if not file_path:
            return

        try:
            headers = ["#", "Type", "Strategy", "Locator Expression", "Stability Rank", "Score Value", "Matches", "Java By Binding", "Notes"]
            data_rows = []
            for idx, c in enumerate(self.rows, 1):
                data_rows.append([
                    idx,
                    c["kind"],
                    c["strategy"],
                    c["expr"],
                    c["rank"],
                    c["score"],
                    c["count"] if c["count"] is not None else "N/A",
                    self.java(c),
                    c["note"]
                ])

            if file_path.endswith(".xlsx") and HAS_OPENPYXL:
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Locators"
                
                ws.views.sheetView[0].showGridLines = True

                ws.append(headers)
                for row_data in data_rows:
                    ws.append(row_data)

                for col in ws.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    col_letter = openpyxl_utils.get_column_letter(col[0].column)
                    ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

                wb.save(file_path)
            else:
                if file_path.endswith(".xlsx") and not HAS_OPENPYXL:
                    file_path = file_path.rsplit('.', 1)[0] + '.csv'
                    messagebox.showinfo("Export Notice", "openpyxl is not installed. Saving file as clean CSV instead.")

                with open(file_path, "w", newline="", encoding="utf-8-sig") as f:
                    writer = csv.writer(f)
                    writer.writerow(headers)
                    writer.writerows(data_rows)

            self.status.config(text=f"Exported {len(self.rows)} locators successfully!")
            messagebox.showinfo("Export Successful", f"Locators exported successfully to:\n{file_path}")
        except Exception as ex:
            messagebox.showerror("Export Failed", f"Could not export file:\n{str(ex)}")

    def pick_element_interactive(self):
        url = self.url_entry.get().strip()
        if not url:
            messagebox.showwarning("Missing URL", "Please enter a valid URL to pick from.")
            return
        
        if self.picker_driver:
            try:
                self.picker_driver.get(url)
                self.status.config(text="Browser navigated to new URL.")
                return
            except Exception:
                self.picker_driver = None

        self.status.config(text="Opening browser picker...")
        self.root.update_idletasks()

        def background_picker():
            try:
                options = ChromeOptions()
                self.picker_driver = webdriver.Chrome(options=options)
                self.picker_driver.get(url)
                self.picker_active = True
                
                initial_pick_state = "true" if self.pick_mode_var.get() else "false"

                self.root.after(0, lambda: self.close_browser_btn.config(state="normal"))
                self.root.after(0, lambda: self.pick_mode_chk.config(state="normal"))
                self.root.after(0, lambda: self.nav_back_btn.config(state="normal"))
                self.root.after(0, lambda: self.nav_fwd_btn.config(state="normal"))
                self.root.after(0, lambda: self.nav_refresh_btn.config(state="normal"))

                script = f"""
                window.__pickerEnabled = {initial_pick_state};
                window.__selectedElementHtml = null;
                let overlay = document.createElement('div');
                overlay.style.position = 'fixed';
                overlay.style.zIndex = '999999';
                overlay.style.border = '2px solid #D93838';
                overlay.style.background = 'rgba(217, 56, 56, 0.15)';
                overlay.style.pointerEvents = 'none';
                overlay.style.transition = 'all 0.05s ease';
                document.body.appendChild(overlay);

                document.addEventListener('mouseover', function(e) {{
                    if (!window.__pickerEnabled) {{
                        overlay.style.width = '0px';
                        overlay.style.height = '0px';
                        return;
                    }}
                    let rect = e.target.getBoundingClientRect();
                    overlay.style.top = rect.top + 'px';
                    overlay.style.left = rect.left + 'px';
                    overlay.style.width = rect.width + 'px';
                    overlay.style.height = rect.height + 'px';
                }}, true);

                document.addEventListener('click', function(e) {{
                    if (!window.__pickerEnabled) {{
                        return;
                    }}
                    e.preventDefault();
                    e.stopPropagation();
                    window.__selectedElementHtml = e.target.outerHTML;
                }}, true);
                """
                self.picker_driver.execute_script(script)
                
                while self.picker_active and self.picker_driver:
                    try:
                        self.picker_driver.execute_script("""
                            if (!window.__pickerInjected) {
                                window.__pickerInjected = true;
                                window.__selectedElementHtml = null;
                                let overlay = document.createElement('div');
                                overlay.style.position = 'fixed';
                                overlay.style.zIndex = '999999';
                                overlay.style.border = '2px solid #D93838';
                                overlay.style.background = 'rgba(217, 56, 56, 0.15)';
                                overlay.style.pointerEvents = 'none';
                                document.body.appendChild(overlay);
                                document.addEventListener('mouseover', function(e) {
                                    if (!window.__pickerEnabled) return;
                                    let rect = e.target.getBoundingClientRect();
                                    overlay.style.top = rect.top + 'px';
                                    overlay.style.left = rect.left + 'px';
                                    overlay.style.width = rect.width + 'px';
                                    overlay.style.height = rect.height + 'px';
                                }, true);
                                document.addEventListener('click', function(e) {
                                    if (!window.__pickerEnabled) return;
                                    e.preventDefault();
                                    e.stopPropagation();
                                    window.__selectedElementHtml = e.target.outerHTML;
                                }, true);
                            }
                        """)
                        
                        res = self.picker_driver.execute_script("return window.__selectedElementHtml;")
                        if res:
                            self.picker_driver.execute_script("window.__selectedElementHtml = null;")
                            self.root.after(0, lambda html=res: self.on_element_picked(html))
                    except Exception:
                        break
                    time.sleep(0.4)
            except Exception as ex:
                err_msg = str(ex)
                self.root.after(0, lambda msg=err_msg: messagebox.showerror("Picker Error", msg))
            finally:
                self.picker_active = False
                self.picker_driver = None
                self.root.after(0, lambda: self.close_browser_btn.config(state="disabled"))
                self.root.after(0, lambda: self.pick_mode_chk.config(state="disabled"))
                self.root.after(0, lambda: self.nav_back_btn.config(state="disabled"))
                self.root.after(0, lambda: self.nav_fwd_btn.config(state="disabled"))
                self.root.after(0, lambda: self.nav_refresh_btn.config(state="disabled"))

        threading.Thread(target=background_picker, daemon=True).start()

    def update_pick_mode(self):
        if self.picker_driver:
            val = "true" if self.pick_mode_var.get() else "false"
            try:
                self.picker_driver.execute_script(f"window.__pickerEnabled = {val};")
                state_str = "Enabled (Clicking captures elements)" if self.pick_mode_var.get() else "Disabled (Free browser navigation mode)"
                self.status.config(text=f"Pick Mode: {state_str}")
            except Exception:
                pass

    def browser_back(self):
        if self.picker_driver:
            try:
                self.picker_driver.back()
                self.picker_driver.execute_script("window.__pickerInjected = false;")
                self.status.config(text="Navigated back in browser.")
            except Exception as ex:
                messagebox.showerror("Navigation Error", str(ex))

    def browser_forward(self):
        if self.picker_driver:
            try:
                self.picker_driver.forward()
                self.picker_driver.execute_script("window.__pickerInjected = false;")
                self.status.config(text="Navigated forward in browser.")
            except Exception as ex:
                messagebox.showerror("Navigation Error", str(ex))

    def browser_refresh(self):
        if self.picker_driver:
            try:
                self.picker_driver.refresh()
                self.picker_driver.execute_script("window.__pickerInjected = false;")
                self.status.config(text="Refreshed browser page.")
            except Exception as ex:
                messagebox.showerror("Navigation Error", str(ex))

    def close_picker_browser(self):
        self.picker_active = False
        if self.picker_driver:
            try: self.picker_driver.quit()
            except Exception: pass
            self.picker_driver = None
        self.close_browser_btn.config(state="disabled")
        self.pick_mode_chk.config(state="disabled")
        self.nav_back_btn.config(state="disabled")
        self.nav_fwd_btn.config(state="disabled")
        self.nav_refresh_btn.config(state="disabled")
        self.status.config(text="Picker browser closed.")

    def on_element_picked(self, html_content):
        self.html.delete("1.0", "end")
        self.html.insert("1.0", str(html_content))
        self.status.config(text="Element clicked & captured! Locators generated.")
        self.generate()

    def test_live_url(self):
        c = self.selected()
        if not c:
            messagebox.showwarning("No Locator Selected", "Please select a locator row first.")
            return
        url = self.url_entry.get().strip()
        if not url:
            messagebox.showwarning("Missing URL", "Please enter a valid URL.")
            return
        
        self.status.config(text="Testing locator on live URL via Headless Chrome...")
        self.root.update_idletasks()
        
        driver = None
        try:
            options = ChromeOptions()
            options.add_argument("--headless")
            driver = webdriver.Chrome(options=options)
            driver.get(url)
            
            kind = c["kind"]
            expr = c["expr"]
            matches = 0
            if kind == "CSS":
                matches = len(driver.find_elements(SelBy.CSS_SELECTOR, expr))
            else:
                matches = len(driver.find_elements(SelBy.XPATH, expr))
                
            messagebox.showinfo("Live Test Result", f"URL: {url}\n\nLocator: {expr}\nLive Elements Found: {matches}")
            self.status.config(text=f"Live test complete: found {matches} matching element(s).")
        except Exception as ex:
            err_msg = str(ex)
            messagebox.showerror("Live Test Error", err_msg)
            self.status.config(text="Live test failed.")
        finally:
            if driver: driver.quit()

    def selected(self):
        s = self.tree.selection()
        return self.rows[int(s[0])] if s else None

    def java(self, c):
        expr = c["expr"]
        if c["kind"] == "XPath": return f'By.xpath("{esc(expr)}")'
        return f'By.cssSelector("{esc(expr)}")'

    def preview(self, _=None):
        c = self.selected()
        if c: self.preview_var.set("Java: " + self.java(c))

    def copy(self, text):
        self.root.clipboard_clear(); self.root.clipboard_append(text); self.root.update()
        self.status.config(text="Copied to clipboard")

    def copy_expr(self):
        c = self.selected()
        if c: self.copy(c["expr"])
        else: messagebox.showinfo("Select a locator", "Select a row first.")

    def copy_java(self):
        c = self.selected()
        if c: self.copy(self.java(c))
        else: messagebox.showinfo("Select a locator", "Select a row first.")

    def copy_all(self):
        if not self.rows: messagebox.showinfo("No locators", "Generate locators first."); return
        out = ["// Generated candidates; verify against the full page before use.", "import org.openqa.selenium.By;", ""]
        for i, c in enumerate(self.rows, 1):
            count = c["count"] if c["count"] is not None else "unknown"
            out += [f'// {c["kind"]} | {c["strategy"]} | {c["rank"]} ({c["score"]}/100) | fragment matches: {count}',
                    f'private final By locator{i} = {self.java(c)};', ""]
        self.copy("\\n".join(out))

def main():
    try:
        root = tk.Tk()
        try:
            if "vista" in ttk.Style().theme_names(): ttk.Style().theme_use("vista")
        except Exception: pass
        App(root)
        root.mainloop()
    except Exception as e:
        err = traceback.format_exc()
        print(err)
        root_err = tk.Tk()
        root_err.withdraw()
        messagebox.showerror("Startup Error", f"An error occurred while launching:\n\n{e}")

if __name__ == "__main__":
    main()
