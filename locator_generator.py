import re
import threading
import time
import tkinter as tk
from tkinter import ttk, messagebox
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

TITLE = "Selenium Locator Generator - NextGen Pro"

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
            if any(re.search(pat, val) for pat in ignored_patterns): continue
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
        if any(re.search(pat, c) for pat in ignored_patterns): continue
        safe.append(c)
    
    if safe:
        selector = tag + "".join("." + c for c in safe)
        volatile = any(re.search(r"ng-star-inserted|ng-option-marked|active|selected|css-|sc-", c) for c in safe)
        add("CSS", "Tag + all classes", selector, 54 if volatile else 73,
            "Volatile framework classes detected." if volatile else "Class-based selector.")
        for c in safe[:4]:
            add("CSS", "Single class", "." + c, 45, "Single class token.")
            add("XPath", "Class token", f"//{tag}[contains(concat(' ', normalize-space(@class), ' '), {xp_literal(' ' + c + ' ')})]", 52, "Class token XPath match.")

    if attrs.get("type"):
        val = attrs["type"]
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
        root.title(TITLE); root.geometry("1480x940"); root.minsize(1100, 750)
        self.rows = []
        self.picker_driver = None
        self.picker_active = False
        self.is_dark_mode = True
        
        # Next-Gen SaaS Color System (Deep Midnight / Slate Indigo Theme)
        self.themes = {
            "dark": {
                "bg": "#090D16", "sidebar_bg": "#05080E", "card_bg": "#111827",
                "text": "#F9FAFB", "subtext": "#9CA3AF", "accent": "#6366F1",
                "accent_hover": "#4F46E5", "border": "#1F2937", "input_bg": "#030712", 
                "tree_heading": "#1F2937", "success": "#10B981"
            },
            "light": {
                "bg": "#F3F4F6", "sidebar_bg": "#E5E7EB", "card_bg": "#FFFFFF",
                "text": "#111827", "subtext": "#4B5563", "accent": "#4F46E5",
                "accent_hover": "#4338CA", "border": "#D1D5DB", "input_bg": "#F9FAFB", 
                "tree_heading": "#E5E7EB", "success": "#059669"
            }
        }

        self.root.configure(bg=self.themes["dark"]["bg"])
        
        # Modern Floating Window Split Layout
        self.sidebar_frame = tk.Frame(root, width=240, bd=0)
        self.sidebar_frame.pack(side="left", fill="y")
        self.sidebar_frame.pack_propagate(False)

        self.content_container = tk.Frame(root, bd=0)
        self.content_container.pack(side="right", fill="both", expand=True)

        self.build_sidebar()
        self.build_home_view()
        self.build_about_view()
        
        self.home_view.pack(fill="both", expand=True, padx=20, pady=20)
        self.about_view.pack_forget()
        
        self.apply_theme()

    def build_sidebar(self):
        logo_area = tk.Frame(self.sidebar_frame, pady=32)
        logo_area.pack(fill="x")
        self.logo_lbl = tk.Label(logo_area, text="⚡ LocatorPro", font=("Segoe UI", 16, "bold"))
        self.logo_lbl.pack(padx=20, anchor="w")

        nav_container = tk.Frame(self.sidebar_frame, pady=10)
        nav_container.pack(fill="x", padx=12)

        self.btn_home = tk.Button(nav_container, text="   🏠   Dashboard", anchor="w", font=("Segoe UI", 10, "bold"), bd=0, relief="flat", padx=14, pady=12, command=lambda: self.show_view("home"), cursor="hand2")
        self.btn_home.pack(fill="x", pady=4)

        self.btn_about = tk.Button(nav_container, text="   ℹ️   Documentation", anchor="w", font=("Segoe UI", 10, "bold"), bd=0, relief="flat", padx=14, pady=12, command=lambda: self.show_view("about"), cursor="hand2")
        self.btn_about.pack(fill="x", pady=4)

        # Bottom info card in sidebar
        self.sidebar_footer = tk.Frame(self.sidebar_frame, pady=20)
        self.sidebar_footer.pack(side="bottom", fill="x", padx=16)
        self.version_lbl = tk.Label(self.sidebar_footer, text="v4.0.0 Stable\nEngine: Selenium + BS4", font=("Segoe UI", 8), justify="left")
        self.version_lbl.pack(anchor="w")

    def show_view(self, view_name):
        t = self.themes["dark" if self.is_dark_mode else "light"]
        if view_name == "home":
            self.about_view.pack_forget()
            self.home_view.pack(fill="both", expand=True, padx=20, pady=20)
            self.btn_home.config(bg=t["accent"], fg="#FFFFFF")
            self.btn_about.config(bg=t["sidebar_bg"], fg=t["subtext"])
        else:
            self.home_view.pack_forget()
            self.about_view.pack(fill="both", expand=True, padx=20, pady=20)
            self.btn_about.config(bg=t["accent"], fg="#FFFFFF")
            self.btn_home.config(bg=t["sidebar_bg"], fg=t["subtext"])

    def build_home_view(self):
        self.home_view = ttk.Frame(self.content_container)
        
        # Header Area
        header_frame = ttk.Frame(self.home_view)
        header_frame.pack(fill="x", pady=(0, 16))
        
        title_box = ttk.Frame(header_frame)
        title_box.pack(side="left", anchor="w")
        
        self.main_title = ttk.Label(title_box, text="Test Automation Control Hub", font=("Segoe UI", 20, "bold"))
        self.main_title.pack(anchor="w")
        self.main_sub = ttk.Label(title_box, text="Generate robust, high-scoring locators with instant browser inspection.", font=("Segoe UI", 10))
        self.main_sub.pack(anchor="w", pady=(3, 0))

        self.theme_btn = ttk.Button(header_frame, text="☀️ Light Mode", command=self.toggle_theme)
        self.theme_btn.pack(side="right", anchor="e")

        # Control Card Panel
        top_panel = ttk.LabelFrame(self.home_view, text=" 🌐 Live Environment & Noise Filtering ", padding=16)
        top_panel.pack(fill="x", pady=(0, 14))
        
        grid_frame = ttk.Frame(top_panel)
        grid_frame.pack(fill="x", pady=(0, 12))

        ttk.Label(grid_frame, text="Ignore Patterns:").grid(row=0, column=0, sticky="w", padx=(0, 12), pady=8)
        self.ignore_entry = tk.Entry(grid_frame, width=36, font=("Consolas", 10), relief="solid", bd=1)
        self.ignore_entry.insert(0, "css-, sc-, ng-star-inserted")
        self.ignore_entry.grid(row=0, column=1, sticky="w", padx=(0, 24), pady=8)

        ttk.Label(grid_frame, text="Target URL:").grid(row=0, column=2, sticky="w", padx=(0, 12), pady=8)
        self.url_entry = tk.Entry(grid_frame, width=42, font=("Consolas", 10), relief="solid", bd=1)
        self.url_entry.insert(0, "https://example.com")
        self.url_entry.grid(row=0, column=3, sticky="w", padx=(0, 18), pady=8)
        
        btn_state = "normal" if HAS_SELENIUM else "disabled"
        self.test_btn = ttk.Button(grid_frame, text="🧪 Test Locator", command=self.test_live_url, state=btn_state)
        self.test_btn.grid(row=0, column=4, padx=6, pady=8)

        self.pick_btn = ttk.Button(grid_frame, text="🔍 Open Picker", command=self.pick_element_interactive, state=btn_state, style="Accent.TButton")
        self.pick_btn.grid(row=0, column=5, padx=6, pady=8)

        self.close_browser_btn = ttk.Button(grid_frame, text="❌ Close Browser", command=self.close_picker_browser, state="disabled")
        self.close_browser_btn.grid(row=0, column=6, padx=6, pady=8)

        nav_frame = ttk.Frame(top_panel)
        nav_frame.pack(fill="x", anchor="w", pady=(6, 0))
        
        self.pick_mode_var = tk.BooleanVar(value=True)
        self.pick_mode_chk = ttk.Checkbutton(nav_frame, text="Pick Mode Active", variable=self.pick_mode_var, command=self.update_pick_mode, state="disabled")
        self.pick_mode_chk.pack(side="left", padx=(0, 24))

        ttk.Label(nav_frame, text="Browser Tools:").pack(side="left", padx=(0, 10))
        self.nav_back_btn = ttk.Button(nav_frame, text="◀ Back", command=self.browser_back, state="disabled")
        self.nav_back_btn.pack(side="left", padx=3)
        
        self.nav_fwd_btn = ttk.Button(nav_frame, text="Forward ▶", command=self.browser_forward, state="disabled")
        self.nav_fwd_btn.pack(side="left", padx=3)
        
        self.nav_refresh_btn = ttk.Button(nav_frame, text="🔄 Refresh", command=self.browser_refresh, state="disabled")
        self.nav_refresh_btn.pack(side="left", padx=3)

        # HTML Input Panel
        html_panel = ttk.LabelFrame(self.home_view, text=" 💻 DOM Element Source Inspector ", padding=16)
        html_panel.pack(fill="x", pady=(0, 14))
        
        self.html = tk.Text(html_panel, height=5, wrap="none", font=("Consolas", 10), relief="solid", bd=1, highlightthickness=1)
        self.html.pack(side="left", fill="both", expand=True)
        sy = ttk.Scrollbar(html_panel, command=self.html.yview); sy.pack(side="right", fill="y"); self.html.configure(yscrollcommand=sy.set)
        self.html.insert("1.0", '<div>Paste Your Targeted Element Here</div>')
        
        # Action Bar
        bar = ttk.Frame(self.home_view)
        bar.pack(fill="x", pady=(0, 14))
        ttk.Button(bar, text="⚡ Generate Locators", command=self.generate, style="Accent.TButton").pack(side="left")
        ttk.Button(bar, text="💡 Load Example", command=self.example).pack(side="left", padx=10)
        ttk.Button(bar, text="🗑 Clear", command=self.clear).pack(side="left")
        ttk.Button(bar, text="📋 Copy All Java", command=self.copy_all).pack(side="left", padx=(16, 0))
        
        status_box = ttk.Frame(bar)
        status_box.pack(side="right")
        self.status = ttk.Label(status_box, text="System Ready"); self.status.pack(side="left")
        
        # Result Treeview Panel
        result_panel = ttk.LabelFrame(self.home_view, text=" 🎯 Candidate Locators & Stability Confidence Matrix ", padding=16)
        result_panel.pack(fill="both", expand=True, pady=(0, 8))
        
        cols = ("kind", "strategy", "expr", "rank", "count", "note")
        self.tree = ttk.Treeview(result_panel, columns=cols, show="headings", selectmode="browse")
        headings = {"kind":"Type", "strategy":"Strategy", "expr":"Locator Expression", "rank":"Stability Score", "count":"Matches", "note":"Notes"}
        widths = {"kind":80, "strategy":170, "expr":410, "rank":120, "count":95, "note":340}
        for c in cols: self.tree.heading(c, text=headings[c]); self.tree.column(c, width=widths[c], minwidth=55, anchor="w")
        
        y = ttk.Scrollbar(result_panel, orient="vertical", command=self.tree.yview)
        x = ttk.Scrollbar(result_panel, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=y.set, xscrollcommand=x.set)
        self.tree.grid(row=0, column=0, sticky="nsew"); y.grid(row=0, column=1, sticky="ns"); x.grid(row=1, column=0, sticky="ew")
        result_panel.rowconfigure(0, weight=1); result_panel.columnconfigure(0, weight=1)
        
        self.tree.bind("<<TreeviewSelect>>", self.preview)
        self.tree.bind("<Double-1>", lambda e: self.copy_expr())
        
        self.context_menu = tk.Menu(self.home_view, tearoff=0)
        self.context_menu.add_command(label="Copy Expression", command=self.copy_expr)
        self.context_menu.add_command(label="Copy Java By Statement", command=self.copy_java)

        def show_context_menu(event):
            item = self.tree.identify_row(event.y)
            if item:
                self.tree.selection_set(item)
                self.context_menu.post(event.x_root, event.y_root)

        self.tree.bind("<Button-3>", show_context_menu)
        self.tree.bind("<Button-2>", show_context_menu)

        bottom = ttk.Frame(self.home_view)
        bottom.pack(fill="x", pady=(6, 0))
        ttk.Button(bottom, text="Copy Selected Expression", command=self.copy_expr).pack(side="left")
        ttk.Button(bottom, text="Copy Selected Java By", command=self.copy_java).pack(side="left", padx=10)
        self.preview_var = tk.StringVar(value="Select any candidate row above to inspect its Selenium Java binding statement.")
        self.preview_lbl = ttk.Label(bottom, textvariable=self.preview_var)
        self.preview_lbl.pack(side="left", padx=14)

    def build_about_view(self):
        self.about_view = ttk.Frame(self.content_container)
        
        ttk.Label(self.about_view, text="LocatorPro Documentation & Architecture", font=("Segoe UI", 22, "bold")).pack(anchor="w", pady=(0, 12))
        ttk.Label(self.about_view, text="Next-Generation Enterprise QA Automation Toolset", font=("Segoe UI", 11, "italic")).pack(anchor="w", pady=(0, 24))

        desc = (
            "Selenium Locator Generator Pro is engineered to solve modern web testing bottlenecks. "
            "By inspecting DOM trees in real-time or via live browser integration, it evaluates unique testing attributes "
            "(such as data-testid, data-qa, and stable IDs) to produce resilient locators that resist frequent UI refactoring.\n\n"
            "Key Technical Capabilities:\n"
            "• Heuristic Stability Scoring: Ranks candidates dynamically based on uniqueness and maintenance footprint.\n"
            "• Live Chrome Inspector: Highlights hovered elements with automated boundary bounding boxes.\n"
            "• Smart Noise Filter: Automatically discards volatile CSS framework hashes (Angular, React, Tailwind).\n"
            "• Enterprise Ready: Instant Java Selenium By-statement compilation."
        )
        
        about_box = ttk.LabelFrame(self.about_view, text=" System Architecture ", padding=24)
        about_box.pack(fill="x", pady=(0, 24))
        
        lbl = ttk.Label(about_box, text=desc, font=("Segoe UI", 11), wraplength=950, justify="left")
        lbl.pack(anchor="w")

        tech_stack = "Tech Stack: Python 3.8+ · Tkinter Modern UI Engine · BeautifulSoup4 · Selenium WebDriver · lxml Parser"
        ttk.Label(self.about_view, text=tech_stack, font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=10)

    def toggle_theme(self):
        self.is_dark_mode = not self.is_dark_mode
        self.apply_theme()

    def apply_theme(self):
        t = self.themes["dark" if self.is_dark_mode else "light"]
        
        self.root.configure(bg=t["bg"])
        self.sidebar_frame.configure(bg=t["sidebar_bg"])
        self.content_container.configure(bg=t["bg"])
        
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure(".", background=t["bg"], foreground=t["text"], font=("Segoe UI", 10))
        style.configure("TLabel", background=t["bg"], foreground=t["text"])
        style.configure("TLabelframe", background=t["bg"], foreground=t["accent"], bordercolor=t["border"], borderwidth=2, relief="solid")
        style.configure("TLabelframe.Label", background=t["bg"], foreground=t["accent"], font=("Segoe UI", 10, "bold"))
        style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=7, background=t["card_bg"], foreground=t["text"], relief="flat")
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), background=t["accent"], foreground="#FFFFFF", relief="flat")
        style.configure("Treeview", background=t["card_bg"], foreground=t["text"], fieldbackground=t["card_bg"], rowheight=30, bordercolor=t["border"], borderwidth=1)
        style.configure("Treeview.Heading", background=t["tree_heading"], foreground=t["text"], font=("Segoe UI", 10, "bold"), relief="flat")
        style.map("Treeview", background=[("selected", t["accent"])], foreground=[("selected", "#FFFFFF")])

        is_home_active = str(self.home_view.winfo_manager()) != ''
        self.btn_home.config(bg=t["accent"] if is_home_active else t["sidebar_bg"], fg="#FFFFFF" if is_home_active else t["subtext"])
        self.btn_about.config(bg=t["sidebar_bg"] if is_home_active else t["accent"], fg=t["subtext"] if is_home_active else "#FFFFFF")
        
        self.logo_lbl.config(bg=t["sidebar_bg"], fg=t["text"])
        self.version_lbl.config(bg=t["sidebar_bg"], fg=t["subtext"])
        self.sidebar_footer.config(bg=t["sidebar_bg"])

        self.ignore_entry.config(bg=t["input_bg"], fg=t["text"], insertbackground=t["text"], highlightbackground=t["border"])
        self.url_entry.config(bg=t["input_bg"], fg=t["text"], insertbackground=t["text"], highlightbackground=t["border"])
        self.html.config(bg=t["input_bg"], fg=t["text"], insertbackground=t["text"], highlightbackground=t["border"])
        self.context_menu.config(bg=t["card_bg"], fg=t["text"], activebackground=t["accent"], activeforeground="#FFFFFF")

        if self.is_dark_mode:
            self.theme_btn.config(text="☀️ Light Mode")
        else:
            self.theme_btn.config(text="🌙 Dark Mode")

        self.main_title.configure(foreground=t["text"])
        self.main_sub.configure(foreground=t["subtext"])
        self.preview_lbl.configure(foreground=t["subtext"])

    def example(self):
        self.html.delete("1.0", "end")
        self.html.insert("1.0", '<div class="form-group">\n  <label for="username">Username</label>\n  <input id="username" name="userName" class="form-control" placeholder="Enter username" type="text">\n</div>')

    def clear(self):
        self.html.delete("1.0", "end")
        for i in self.tree.get_children(): self.tree.delete(i)
        self.rows = []; self.status.config(text="Cleared"); self.preview_var.set("Select any candidate row above to inspect its Selenium Java binding statement.")

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
                overlay.style.border = '2px solid #6366F1';
                overlay.style.background = 'rgba(99, 102, 241, 0.15)';
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
                                overlay.style.border = '2px solid #6366F1';
                                overlay.style.background = 'rgba(99, 102, 241, 0.15)';
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
            try:
                self.picker_driver.quit()
            except Exception:
                pass
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
        if c: self.preview_var.set("Java Binding: " + self.java(c))

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