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

TITLE = "Locator Generator Pro - Enterprise"

USER_MANUAL_TEXT = """======================================================================
              LOCATOR GENERATOR PRO - USER MANUAL
======================================================================

Welcome to the Locator Generator Pro user manual. This guide 
covers multi-selection workflows, Page Object Model (POM) generation, 
and smart auto-healing regression audits.

----------------------------------------------------------------------
1. GETTING STARTED & INSTALLATION COMMANDS
----------------------------------------------------------------------
Prerequisites: Python 3.8+
Required Packages:
    pip install beautifulsoup4 selenium openpyxl lxml

----------------------------------------------------------------------
2. PAGE OBJECT MODEL (POM) CLASS GENERATOR
----------------------------------------------------------------------
Navigate to the "🏗️ POM Export" tab:
• Select multiple rows on the Dashboard using Ctrl+Click or Shift+Click 
  and click "➕ Add Selected to POM".
• Variables are automatically given descriptive names while duplicates 
  are prevented.

----------------------------------------------------------------------
3. SMART AUTO-HEALING HEALTH AUDITOR
----------------------------------------------------------------------
Navigate to the "🔍 Health Auditor" tab:
• Send multiple selected locators straight from the Dashboard using 
  "🔍 Send Selected to Health Auditor".
• Run batch regression checks and view AI auto-healing recommendations.
"""

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
        root.title(TITLE); root.geometry("1400x950"); root.minsize(1000, 720)
        self.rows = []
        self.pom_elements = []
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
        self.title_label = ttk.Label(title_box, text="🎯 Locator Generator Pro - Enterprise", font=("Segoe UI", 16, "bold"))
        self.title_label.pack(anchor="w")
        self.sub_label = ttk.Label(title_box, text="Multi-selection batch workflow, Page Object Model classes, and smart auto-healing regression audit.", font=("Segoe UI", 10))
        self.sub_label.pack(anchor="w", pady=(2, 0))

        self.theme_btn = ttk.Button(header_frame, text="☀️ Light Mode", command=self.toggle_theme)
        self.theme_btn.pack(side="right", anchor="e")

        self.notebook = ttk.Notebook(self.main_container)
        self.notebook.pack(fill="both", expand=True)

        self.home_tab = ttk.Frame(self.notebook, padding=10)
        self.pom_tab = ttk.Frame(self.notebook, padding=10)
        self.audit_tab = ttk.Frame(self.notebook, padding=10)
        self.about_tab = ttk.Frame(self.notebook, padding=10)

        self.notebook.add(self.home_tab, text="  🏠 Dashboard  ")
        self.notebook.add(self.pom_tab, text="  🏗️ POM Export  ")
        self.notebook.add(self.audit_tab, text="  🔍 Health Auditor & Auto-Heal  ")
        self.notebook.add(self.about_tab, text="  ℹ️ About & User Manual  ")

        self.build_home_tab()
        self.build_pom_tab()
        self.build_audit_tab()
        self.build_about_tab()

        self.apply_theme()

    def build_home_tab(self):
        btn_state = "normal" if HAS_SELENIUM else "disabled"

        self.top_panel = ttk.LabelFrame(self.home_tab, text=" 🌐 Live Browser & Noise Filtering Control ", padding=10)
        self.top_panel.pack(fill="x", pady=(0, 10))
        
        ttk.Label(self.top_panel, text="Ignore Patterns:").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        self.ignore_entry = tk.Entry(self.top_panel, width=35, font=("Consolas", 10), relief="solid", bd=1)
        self.ignore_entry.insert(0, "css-, sc-, ng-star-inserted")
        self.ignore_entry.grid(row=0, column=1, sticky="w", padx=6, pady=4)

        ttk.Label(self.top_panel, text="Framework:").grid(row=0, column=2, sticky="e", padx=6, pady=4)
        self.framework_var = tk.StringVar(value="Java Selenium")
        self.framework_combo = ttk.Combobox(self.top_panel, textvariable=self.framework_var, values=["Java Selenium", "Python Selenium", "Playwright", "Cypress"], state="readonly", width=16)
        self.framework_combo.grid(row=0, column=3, columnspan=2, sticky="w", padx=6, pady=4)
        self.framework_combo.bind("<<ComboboxSelected>>", self.preview)

        ttk.Label(self.top_panel, text="Target URL:").grid(row=1, column=0, sticky="w", padx=6, pady=4)
        self.url_entry = tk.Entry(self.top_panel, width=45, font=("Consolas", 10), relief="solid", bd=1)
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

        frame = ttk.LabelFrame(self.home_tab, text=" 📝 HTML Input Fragment (or captured live element) ", padding=10)
        frame.pack(fill="x", pady=(0, 10))
        
        self.html = tk.Text(frame, height=5, wrap="none", font=("Consolas", 10), relief="solid", bd=1, highlightthickness=1)
        self.html.pack(side="left", fill="both", expand=True)
        sy = ttk.Scrollbar(frame, command=self.html.yview); sy.pack(side="right", fill="y"); self.html.configure(yscrollcommand=sy.set)
        self.html.string = self.html.insert("1.0", '<div>Paste Your Targeted Element Here</div>')
        
        bar = ttk.Frame(self.home_tab)
        bar.pack(fill="x", pady=(0, 8))
        ttk.Button(bar, text="⚡ Generate Locators", command=self.generate, style="Accent.TButton").pack(side="left")
        ttk.Button(bar, text="🌐 Scan Whole Page", command=self.scan_whole_page, state=btn_state).pack(side="left", padx=6)
        ttk.Button(bar, text="💡 Load Example", command=self.example).pack(side="left", padx=6)
        ttk.Button(bar, text="🗑 Clear", command=self.clear).pack(side="left")
        
        # Multi-selection action buttons
        ttk.Button(bar, text="➕ Add Selected to POM", command=self.add_selected_to_pom).pack(side="left", padx=(14, 0))
        ttk.Button(bar, text="🔍 Send Selected to Health Auditor", command=self.send_selected_to_auditor).pack(side="left", padx=6)
        ttk.Button(bar, text="📊 Export", command=self.export_excel).pack(side="left", padx=6)
        
        self.status = ttk.Label(bar, text="Ready"); self.status.pack(side="right")
        
        result = ttk.LabelFrame(self.home_tab, text=" 🎯 Candidate Locators & Stability Score (Multi-select enabled via Ctrl/Shift) ", padding=10)
        result.pack(fill="both", expand=True, pady=(0, 10))
        cols = ("kind", "strategy", "expr", "rank", "count", "note")
        
        # Changed selectmode to 'extended' to allow multi-selection!
        self.tree = ttk.Treeview(result, columns=cols, show="headings", selectmode="extended")
        
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
        self.context_menu.add_command(label="Copy Code Binding", command=self.copy_framework_code)
        self.context_menu.add_command(label="Add Selected to POM", command=self.add_selected_to_pom)
        self.context_menu.add_command(label="Send Selected to Auditor", command=self.send_selected_to_auditor)

        def show_context_menu(event):
            item = self.tree.identify_row(event.y)
            if item:
                if item not in self.tree.selection():
                    self.tree.selection_set(item)
                self.context_menu.post(event.x_root, event.y_root)

        self.tree.bind("<Button-3>", show_context_menu)
        self.tree.bind("<Button-2>", show_context_menu)

        bottom = ttk.Frame(self.home_tab)
        bottom.pack(fill="x", pady=(0, 4))
        ttk.Button(bottom, text="Copy Selected Expression", command=self.copy_expr).pack(side="left")
        ttk.Button(bottom, text="Copy Code Binding", command=self.copy_framework_code).pack(side="left", padx=6)
        self.preview_var = tk.StringVar(value="Select rows to preview automation code.")
        self.preview_lbl = ttk.Label(bottom, textvariable=self.preview_var)
        self.preview_lbl.pack(side="left", padx=10)

    def build_pom_tab(self):
        ttk.Label(self.pom_tab, text="🏗️ Page Object Model (POM) Class Generator", font=("Segoe UI", 16, "bold")).pack(anchor="w", pady=(0, 4))
        ttk.Label(self.pom_tab, text="Group captured locators and export clean, structured enterprise POM class boilerplates.", font=("Segoe UI", 10)).pack(anchor="w", pady=(0, 10))

        top_bar = ttk.Frame(self.pom_tab)
        top_bar.pack(fill="x", pady=(0, 10))

        ttk.Label(top_bar, text="Page Class Name:").pack(side="left", padx=(0, 6))
        self.pom_class_entry = tk.Entry(top_bar, width=25, font=("Consolas", 10), relief="solid", bd=1)
        self.pom_class_entry.insert(0, "LoginPage")
        self.pom_class_entry.pack(side="left", padx=(0, 15))

        ttk.Button(top_bar, text="🗑 Clear POM List", command=self.clear_pom).pack(side="left", padx=6)
        ttk.Button(top_bar, text="📋 Copy POM Class Code", command=self.copy_pom_class, style="Accent.TButton").pack(side="right", padx=6)

        pom_pane = ttk.Frame(self.pom_tab)
        pom_pane.pack(fill="both", expand=True, pady=(0, 10))

        left_frame = ttk.LabelFrame(pom_pane, text=" Added Elements Suite ", padding=8)
        left_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))

        pom_cols = ("varname", "kind", "expr")
        self.pom_tree = ttk.Treeview(left_frame, columns=pom_cols, show="headings", selectmode="browse")
        pom_headings = {"varname": "Descriptive Variable Name", "kind": "Type", "expr": "Expression"}
        pom_widths = {"varname": 180, "kind": 70, "expr": 245}
        for c in pom_cols:
            self.pom_tree.heading(c, text=pom_headings[c])
            self.pom_tree.column(c, width=pom_widths[c], minwidth=50, anchor="w")
        
        py_y = ttk.Scrollbar(left_frame, orient="vertical", command=self.pom_tree.yview)
        py_x = ttk.Scrollbar(left_frame, orient="horizontal", command=self.pom_tree.xview)
        self.pom_tree.configure(yscrollcommand=py_y.set, xscrollcommand=py_x.set)
        self.pom_tree.grid(row=0, column=0, sticky="nsew"); py_y.grid(row=0, column=1, sticky="ns"); py_x.grid(row=1, column=0, sticky="ew")
        left_frame.rowconfigure(0, weight=1); left_frame.columnconfigure(0, weight=1)

        right_frame = ttk.LabelFrame(pom_pane, text=" Generated Class Code Preview ", padding=8)
        right_frame.pack(side="right", fill="both", expand=True, padx=(5, 0))

        self.pom_code_text = tk.Text(right_frame, wrap="none", font=("Consolas", 10), relief="solid", bd=1)
        self.pom_code_text.pack(side="left", fill="both", expand=True)
        code_sy = ttk.Scrollbar(right_frame, command=self.pom_code_text.yview); code_sy.pack(side="right", fill="y")
        self.pom_code_text.configure(yscrollcommand=code_sy.set)

    def build_audit_tab(self):
        ttk.Label(self.audit_tab, text="🔍 Locator Health Auditor & Smart Auto-Healer", font=("Segoe UI", 16, "bold")).pack(anchor="w", pady=(0, 4))
        ttk.Label(self.audit_tab, text="Audits existing locators against a live URL. If a locator breaks, the AI auto-healer detects and suggests valid replacements.", font=("Segoe UI", 10)).pack(anchor="w", pady=(0, 10))

        ctrl_frame = ttk.Frame(self.audit_tab)
        ctrl_frame.pack(fill="x", pady=(0, 10))

        ttk.Label(ctrl_frame, text="Audit Target URL:").pack(side="left", padx=(0, 6))
        self.audit_url_entry = tk.Entry(ctrl_frame, width=45, font=("Consolas", 10), relief="solid", bd=1)
        self.audit_url_entry.insert(0, "https://example.com")
        self.audit_url_entry.pack(side="left", padx=(0, 10))

        audit_btn_state = "normal" if HAS_SELENIUM else "disabled"
        self.run_audit_btn = ttk.Button(ctrl_frame, text="🚀 Run Audit & Auto-Heal", command=self.run_health_audit_and_heal, state=audit_btn_state, style="Accent.TButton")
        self.run_audit_btn.pack(side="left", padx=6)

        content_frame = ttk.Frame(self.audit_tab)
        content_frame.pack(fill="both", expand=True, pady=(0, 10))

        left_pane = ttk.LabelFrame(content_frame, text=" Paste or Send Locators to Audit (One per line) ", padding=8)
        left_pane.pack(side="left", fill="both", expand=True, padx=(0, 5))
        
        self.audit_input_text = tk.Text(left_pane, height=15, wrap="none", font=("Consolas", 10), relief="solid", bd=1)
        self.audit_input_text.pack(side="left", fill="both", expand=True)
        aud_sy = ttk.Scrollbar(left_pane, command=self.audit_input_text.yview); aud_sy.pack(side="right", fill="y")
        self.audit_input_text.configure(yscrollcommand=aud_sy.set)
        self.audit_input_text.insert("1.0", "//h1\n#nonexistent-broken-id\n//a")

        right_pane = ttk.LabelFrame(content_frame, text=" Audit Results & Healing Suggestions ", padding=8)
        right_pane.pack(side="right", fill="both", expand=True, padx=(5, 0))

        audit_cols = ("expr", "status", "suggestion")
        self.audit_tree = ttk.Treeview(right_pane, columns=audit_cols, show="headings", selectmode="browse")
        audit_headings = {"expr": "Original Locator", "status": "Health Status", "suggestion": "Auto-Healed / Suggested Replacement"}
        audit_widths = {"expr": 240, "status": 140, "suggestion": 270}
        for c in audit_cols: 
            self.audit_tree.heading(c, text=audit_headings[c])
            self.audit_tree.column(c, width=audit_widths[c], minwidth=50, anchor="w")
        
        aud_y = ttk.Scrollbar(right_pane, orient="vertical", command=self.audit_tree.yview)
        aud_x = ttk.Scrollbar(right_pane, orient="horizontal", command=self.audit_tree.xview)
        self.audit_tree.configure(yscrollcommand=aud_y.set, xscrollcommand=aud_x.set)
        self.audit_tree.grid(row=0, column=0, sticky="nsew"); aud_y.grid(row=0, column=1, sticky="ns"); aud_x.grid(row=1, column=0, sticky="ew")
        right_pane.rowconfigure(0, weight=1); right_pane.columnconfigure(0, weight=1)

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

        for widget in [self.ignore_entry, self.url_entry, self.html, self.audit_url_entry, self.audit_input_text, self.pom_class_entry, self.pom_code_text]:
            try:
                widget.config(bg=t["input_bg"], fg=t["text"], insertbackground=t["text"])
            except Exception:
                pass
        
        self.manual_text.config(state="normal")
        self.manual_text.config(bg=t["input_bg"], fg=t["text"], insertbackground=t["text"])
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
        self.rows = []; self.status.config(text="Cleared"); self.preview_var.set("Select rows to preview automation code.")

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

    def generate_descriptive_name(self, expr):
        clean_expr = expr.lower()
        base_name = "element"
        if "id=" in clean_expr or "#" in clean_expr:
            match = re.search(r'(?:id=["\']?|#)([a-zA-Z0-9\-_]+)', clean_expr)
            if match: base_name = match.group(1)
        elif "name=" in clean_expr:
            match = re.search(r'name=["\']?([a-zA-Z0-9\-_]+)', clean_expr)
            if match: base_name = match.group(1)
        elif "data-testid" in clean_expr or "data-qa" in clean_expr:
            match = re.search(r'data-(?:testid|qa)=["\']?([a-zA-Z0-9\-_]+)', clean_expr)
            if match: base_name = match.group(1)
        elif "placeholder=" in clean_expr:
            match = re.search(r'placeholder=["\']?([a-zA-Z0-9\-_ ]+)', clean_expr)
            if match: base_name = match.group(1)
        elif "button" in clean_expr:
            base_name = "button"
        elif "input" in clean_expr:
            base_name = "input"
        elif "a[" in clean_expr or "link" in clean_expr:
            base_name = "link"

        words = re.findall(r'[a-zA-Z0-9]+', base_name)
        if not words: return "uiElement"
        
        camel_case = words[0].lower() + "".join(w.capitalize() for w in words[1:])
        if "button" in clean_expr and not camel_case.lower().endswith("button"):
            camel_case += "Button"
        elif "input" in clean_expr and not camel_case.lower().endswith("input"):
            camel_case += "Input"
        elif "a[" in clean_expr and not camel_case.lower().endswith("link"):
            camel_case += "Link"

        return camel_case

    def add_selected_to_pom(self):
        """Adds all currently selected rows to the POM suite, checking for duplicates and naming them uniquely."""
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("No Rows Selected", "Please select one or more rows from the table first.")
            return

        added_count = 0
        duplicate_count = 0

        for item in selected_items:
            c = self.rows[int(item)]
            expr = c["expr"]

            # Check if expression already exists in POM suite
            exists = any(el["expr"] == expr for el in self.pom_elements)
            if exists:
                duplicate_count += 1
                continue

            var_name = self.generate_descriptive_name(expr)
            original_var = var_name
            counter = 1
            existing_names = {el["varname"] for el in self.pom_elements}
            while var_name in existing_names:
                var_name = f"{original_var}{counter}"
                counter += 1

            self.pom_elements.append({"varname": var_name, "kind": c["kind"], "expr": expr})
            self.pom_tree.insert("", "end", values=(var_name, c["kind"], expr))
            added_count += 1

        self.update_pom_code_preview()
        self.status.config(text=f"Added {added_count} element(s) to POM suite ({duplicate_count} duplicates skipped).")
        if added_count > 0:
            messagebox.showinfo("POM Update", f"Successfully added {added_count} element(s) to your POM suite! ({duplicate_count} skipped duplicates).")

    def send_selected_to_auditor(self):
        """Sends all currently selected rows' locator expressions into the Health Auditor input text box."""
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("No Rows Selected", "Please select one or more rows from the table first.")
            return

        expressions_to_add = []
        for item in selected_items:
            c = self.rows[int(item)]
            expressions_to_add.append(c["expr"])

        current_audit_text = self.audit_input_text.get("1.0", "end").strip()
        existing_lines = set(current_audit_text.split("\n")) if current_audit_text else set()

        new_added = 0
        for expr in expressions_to_add:
            if expr not in existing_lines:
                self.audit_input_text.insert("end", f"\n{expr}")
                existing_lines.add(expr)
                new_added += 1

        # Switch notebook tab to Health Auditor tab automatically
        self.notebook.select(self.audit_tab)
        self.status.config(text=f"Sent {new_added} locator(s) to Health Auditor.")

    def clear_pom(self):
        self.pom_elements = []
        for i in self.pom_tree.get_children():
            self.pom_tree.delete(i)
        self.update_pom_code_preview()
        self.status.config(text="POM suite cleared.")

    def update_pom_code_preview(self):
        class_name = self.pom_class_entry.get().strip() or "LoginPage"
        fw = self.framework_var.get()
        code_lines = []

        if fw == "Java Selenium":
            code_lines = [
                "import org.openqa.selenium.By;",
                "import org.openqa.selenium.WebDriver;",
                "import org.openqa.selenium.WebElement;",
                "import org.openqa.selenium.support.PageFactory;",
                "",
                f"public class {class_name} {{",
                "    private WebDriver driver;",
                ""
            ]
            for el in self.pom_elements:
                by_stmt = f'By.xpath("{esc(el["expr"])}")' if el["kind"] == "XPath" else f'By.cssSelector("{esc(el["expr"])}")'
                code_lines.append(f"    private final By {el['varname']} = {by_stmt};")
            
            code_lines += [
                "",
                f"    public {class_name}(WebDriver driver) {{",
                "        this.driver = driver;",
                "    }",
                "}"
            ]
        elif fw == "Python Selenium":
            code_lines = [
                "from selenium.webdriver.common.by import By",
                "",
                f"class {class_name}:",
                "    def __init__(self, driver):",
                "        self.driver = driver",
                ""
            ]
            for el in self.pom_elements:
                sel_type = "By.XPATH" if el["kind"] == "XPath" else "By.CSS_SELECTOR"
                code_lines.append(f'    {el["varname"]} = ({sel_type}, "{esc(el["expr"])}")')
            code_lines.append("")
        elif fw == "Playwright":
            code_lines = [
                f"class {class_name}:",
                "    def __init__(self, page):",
                "        self.page = page",
                ""
            ]
            for el in self.pom_elements:
                expr_val = f'xpath={el["expr"]}' if el["kind"] == "XPath" else el["expr"]
                code_lines.append(f'        self.{el["varname"]} = page.locator("{esc(expr_val)}")')
            code_lines.append("")
        elif fw == "Cypress":
            code_lines = [
                f"class {class_name} {{",
            ]
            for el in self.pom_elements:
                cmd = f'cy.xpath("{esc(el["expr"])}")' if el["kind"] == "XPath" else f'cy.get("{esc(el["expr"])}")'
                code_lines.append(f'    {el["varname"]} = () => {cmd};')
            code_lines.append("}")

        self.pom_code_text.delete("1.0", "end")
        self.pom_code_text.insert("1.0", "\n".join(code_lines))

    def copy_pom_class(self):
        code = self.pom_code_text.get("1.0", "end").strip()
        if not code:
            messagebox.showwarning("Empty POM", "Add elements to the POM suite first.")
            return
        self.copy(code)
        messagebox.showinfo("POM Copied", "POM Class code copied to clipboard successfully!")

    def run_health_audit_and_heal(self):
        url = self.audit_url_entry.get().strip()
        raw_locators = self.audit_input_text.get("1.0", "end").strip()
        
        if not url or not raw_locators:
            messagebox.showwarning("Missing Data", "Please enter both a Target URL and locators to audit.")
            return
        if not HAS_SELENIUM:
            messagebox.showerror("Selenium Missing", "Selenium is required for live audits.")
            return

        locators = [line.strip() for line in raw_locators.split("\n") if line.strip()]
        for i in self.audit_tree.get_children():
            self.audit_tree.delete(i)

        self.status.config(text="Auditing and auto-healing locators...")
        self.root.update_idletasks()

        def audit_and_heal_task():
            driver = None
            try:
                options = ChromeOptions()
                options.add_argument("--headless")
                driver = webdriver.Chrome(options=options)
                driver.get(url)
                page_source = driver.page_source
                soup = BeautifulSoup(page_source, "html.parser") if BeautifulSoup else None

                results = []
                for expr in locators:
                    matches = 0
                    try:
                        if expr.startswith("//") or expr.startswith("("):
                            matches = len(driver.find_elements(SelBy.XPATH, expr))
                        else:
                            matches = len(driver.find_elements(SelBy.CSS_SELECTOR, expr))
                    except Exception:
                        matches = 0

                    if matches > 0:
                        status = "✅ Healthy"
                        suggestion = "N/A (Working)"
                    else:
                        status = "❌ Broken / 0 Matches"
                        suggestion = "No fallback found"
                        if soup:
                            if "#" in expr:
                                clean_id = expr.split("#")[-1].split()[0]
                                alt_el = soup.find(id=re.compile(clean_id, re.IGNORECASE))
                                if alt_el and alt_el.get('id'):
                                    suggestion = f"//*[@id='{alt_el.get('id')}']"
                            elif "//" in expr:
                                tag_match = re.search(r"//([a-zA-Z0-9]+)", expr)
                                if tag_match:
                                    tag = tag_match.group(1)
                                    alt_el = soup.find(tag)
                                    if alt_el and alt_el.get('class'):
                                        suggestion = f"//{tag}[contains(@class, '{alt_el.get('class')[0]}')]"
                                    elif alt_el:
                                        suggestion = f"//{tag}"

                    results.append((expr, status, suggestion))

                self.root.after(0, lambda res=results: self.populate_audit_results(res))
            except Exception as ex:
                self.root.after(0, lambda msg=str(ex): messagebox.showerror("Audit Error", msg))
            finally:
                if driver:
                    try: driver.quit()
                    except: pass

        threading.Thread(target=audit_and_heal_task, daemon=True).start()

    def populate_audit_results(self, results):
        for expr, status, suggestion in results:
            self.audit_tree.insert("", "end", values=(expr, status, suggestion))
        self.status.config(text="Health audit & auto-healing complete.")

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

        fw = self.framework_var.get()
        try:
            headers = ["#", "Type", "Strategy", "Locator Expression", "Stability Rank", "Score Value", "Matches", f"Code Binding ({fw})", "Notes"]
            data_rows = []
            for idx, c in enumerate(self.rows, 1):
                data_rows.append([
                    idx, c["kind"], c["strategy"], c["expr"], c["rank"], c["score"],
                    c["count"] if c["count"] is not None else "N/A",
                    self.format_framework_code(c, fw), c["note"]
                ])

            if file_path.endswith(".xlsx") and HAS_OPENPYXL:
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Locators"
                ws.views.sheetView[0].showGridLines = True
                ws.append(headers)
                for row_data in data_rows: ws.append(row_data)

                for col in ws.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    col_letter = openpyxl_utils.get_column_letter(col[0].column)
                    ws.column_dimensions[col_letter].width = max(max_len + 3, 12)
                wb.save(file_path)
            else:
                with open(file_path, "w", newline="", encoding="utf-8-sig") as f:
                    writer = csv.writer(f)
                    writer.writerow(headers)
                    writer.writerows(data_rows)

            messagebox.showinfo("Export Successful", f"Locators exported successfully to:\n{file_path}")
        except Exception as ex:
            messagebox.showerror("Export Failed", str(ex))

    def pick_element_interactive(self):
        url = self.url_entry.get().strip()
        if not url:
            messagebox.showwarning("Missing URL", "Please enter a valid URL.")
            return
        
        if self.picker_driver:
            try:
                self.picker_driver.get(url)
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
                    if (!window.__pickerEnabled) {{ overlay.style.width = '0px'; overlay.style.height = '0px'; return; }}
                    let rect = e.target.getBoundingClientRect();
                    overlay.style.top = rect.top + 'px'; overlay.style.left = rect.left + 'px';
                    overlay.style.width = rect.width + 'px'; overlay.style.height = rect.height + 'px';
                }}, true);

                document.addEventListener('click', function(e) {{
                    if (!window.__pickerEnabled) return;
                    e.preventDefault(); e.stopPropagation();
                    window.__selectedElementHtml = e.target.outerHTML;
                }}, true);
                """
                self.picker_driver.execute_script(script)
                
                last_tracked_url = ""
                while self.picker_active and self.picker_driver:
                    try:
                        current_url = self.picker_driver.current_url
                        if current_url and current_url != last_tracked_url:
                            last_tracked_url = current_url
                            self.root.after(0, lambda u=current_url: self.update_url_input(u))

                        self.picker_driver.execute_script("""
                            if (!window.__pickerInjected) {
                                window.__pickerInjected = true;
                                window.__selectedElementHtml = null;
                                let overlay = document.createElement('div');
                                overlay.style.position = 'fixed'; overlay.style.zIndex = '999999';
                                overlay.style.border = '2px solid #D93838'; overlay.style.background = 'rgba(217, 56, 56, 0.15)';
                                overlay.style.pointerEvents = 'none'; document.body.appendChild(overlay);
                                document.addEventListener('mouseover', function(e) {
                                    if (!window.__pickerEnabled) return;
                                    let rect = e.target.getBoundingClientRect();
                                    overlay.style.top = rect.top + 'px'; overlay.style.left = rect.left + 'px';
                                    overlay.style.width = rect.width + 'px'; overlay.style.height = rect.height + 'px';
                                }, true);
                                document.addEventListener('click', function(e) {
                                    if (!window.__pickerEnabled) return;
                                    e.preventDefault(); e.stopPropagation();
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
                self.root.after(0, lambda msg=str(ex): messagebox.showerror("Picker Error", msg))
            finally:
                self.picker_active = False
                self.picker_driver = None
                self.root.after(0, lambda: self.close_browser_btn.config(state="disabled"))

        threading.Thread(target=background_picker, daemon=True).start()

    def update_url_input(self, url):
        self.url_entry.delete(0, tk.END)
        self.url_entry.insert(0, url)

    def update_pick_mode(self):
        if self.picker_driver:
            val = "true" if self.pick_mode_var.get() else "false"
            try: self.picker_driver.execute_script(f"window.__pickerEnabled = {val};")
            except Exception: pass

    def browser_back(self):
        if self.picker_driver:
            try: self.picker_driver.back(); self.picker_driver.execute_script("window.__pickerInjected = false;")
            except Exception as ex: messagebox.showerror("Error", str(ex))

    def browser_forward(self):
        if self.picker_driver:
            try: self.picker_driver.forward(); self.picker_driver.execute_script("window.__pickerInjected = false;")
            except Exception as ex: messagebox.showerror("Error", str(ex))

    def browser_refresh(self):
        if self.picker_driver:
            try: self.picker_driver.refresh(); self.picker_driver.execute_script("window.__pickerInjected = false;")
            except Exception as ex: messagebox.showerror("Error", str(ex))

    def close_picker_browser(self):
        self.picker_active = False
        if self.picker_driver:
            try: self.picker_driver.quit()
            except Exception: pass
            self.picker_driver = None
        self.close_browser_btn.config(state="disabled")
        self.status.config(text="Browser closed.")

    def on_element_picked(self, html_content):
        self.html.delete("1.0", "end")
        self.html.insert("1.0", str(html_content))
        self.status.config(text="Element captured! Locators generated.")
        self.generate()

    def test_live_url(self):
        c = self.selected()
        if not c:
            messagebox.showwarning("Select Row", "Please select a locator row first.")
            return
        url = self.url_entry.get().strip()
        driver = None
        try:
            options = ChromeOptions(); options.add_argument("--headless")
            driver = webdriver.Chrome(options=options); driver.get(url)
            matches = len(driver.find_elements(SelBy.CSS_SELECTOR, c["expr"])) if c["kind"] == "CSS" else len(driver.find_elements(SelBy.XPATH, c["expr"]))
            messagebox.showinfo("Live Test Result", f"Elements Found on Live URL: {matches}")
        except Exception as ex:
            messagebox.showerror("Error", str(ex))
        finally:
            if driver: driver.quit()

    def selected(self):
        s = self.tree.selection()
        return self.rows[int(s[0])] if s else None

    def format_framework_code(self, c, framework):
        expr = c["expr"]
        kind = c["kind"]
        if framework == "Java Selenium":
            return f'By.xpath("{esc(expr)}")' if kind == "XPath" else f'By.cssSelector("{esc(expr)}")'
        elif framework == "Python Selenium":
            return f'driver.find_element({"By.XPATH" if kind == "XPath" else "By.CSS_SELECTOR"}, "{esc(expr)}")'
        elif framework == "Playwright":
            return f'page.locator("xpath={esc(expr)}")' if kind == "XPath" else f'page.locator("{esc(expr)}")'
        elif framework == "Cypress":
            return f'cy.xpath("{esc(expr)}")' if kind == "XPath" else f'cy.get("{esc(expr)}")'
        return expr

    def preview(self, _=None):
        s = self.tree.selection()
        if not s: return
        c = self.rows[int(s[0])]
        fw = self.framework_var.get()
        self.preview_var.set(f"{fw}: " + self.format_framework_code(c, fw))
        self.update_pom_code_preview()

    def copy(self, text):
        self.root.clipboard_clear(); self.root.clipboard_append(text); self.root.update()

    def copy_expr(self):
        s = self.tree.selection()
        if not s:
            messagebox.showinfo("Select a row", "Select at least one row first.")
            return
        if len(s) == 1:
            self.copy(self.rows[int(s[0])]["expr"])
        else:
            expressions = "\n".join(self.rows[int(item)]["expr"] for item in s)
            self.copy(expressions)
            self.status.config(text=f"Copied {len(s)} expressions to clipboard.")

    def copy_framework_code(self):
        s = self.tree.selection()
        if not s:
            messagebox.showinfo("Select a row", "Select at least one row first.")
            return
        fw = self.framework_var.get()
        if len(s) == 1:
            self.copy(self.format_framework_code(self.rows[int(s[0])], fw))
        else:
            codes = "\n".join(self.format_framework_code(self.rows[int(item)], fw) for item in s)
            self.copy(codes)
            self.status.config(text=f"Copied {len(s)} code bindings to clipboard.")

def main():
    try:
        root = tk.Tk()
        App(root)
        root.mainloop()
    except Exception as e:
        traceback.print_exc()
        input("\n[ERROR] Application crashed. Press Enter to exit...")

if __name__ == "__main__":
    main()
