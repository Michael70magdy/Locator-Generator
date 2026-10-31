# 🎯 Locator Generator (PySide6 Edition)

[![Python Version](https://img.shields.io/badge/python-3.9%2B-blue.svg)](https://www.python.org/)
[![GUI Framework](https://img.shields.io/badge/GUI-PySide6%20%2F%20Qt6-brightgreen.svg)](https://wiki.qt.io/Qt_for_Python)
[![Test Frameworks](https://img.shields.io/badge/Frameworks-Selenium%20%7C%20Playwright%20%7C%20Cypress-orange.svg)](#cross-framework-code-generation)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

An enterprise-grade, non-blocking desktop studio built with **PySide6** for automation engineers, QA leads, and SDETs. **Locator Generator** accelerates test script development by generating resilient, ranked, and stability-scored CSS and XPath selectors across multiple web test frameworks (**Java Selenium**, **Python Selenium**, **Playwright**, and **Cypress**).

The application combines embedded visual element picking (`QWebEngineView`), raw DOM snapshot heuristic analysis, live browser validation with multi-driver fallback (Chrome/Edge), an automated Page Object Model (POM) builder, and dynamic Health Auditing for self-healing test suites.

---

## 📋 Table of Contents

- [Key Architecture & Features](#-key-architecture--features)
- [System Requirements & Dependencies](#-system-requirements--dependencies)
- [Installation & Quick Start](#-installation--quick-start)
- [Detailed Workspace Guide](#-detailed-workspace-guide)
  - [1. Element Picker & HTML Engine](#1-element-picker--html-engine)
  - [2. Page Objects Builder](#2-page-objects-builder)
  - [3. Live Health Audit & Healing](#3-live-health-audit--healing)
  - [4. Help, Diagnostics & Guide](#4-help-diagnostics--guide)
- [Cross-Framework Code Generation](#-cross-framework-code-generation)
- [Technical Design & Concurrency Model](#-technical-design--concurrency-model)
- [UI/UX Highlights & Theme Engine](#-uiux-highlights--theme-engine)
- [State Persistence (QSettings)](#-state-persistence-qsettings)
- [Troubleshooting & Diagnostic Logging](#-troubleshooting--diagnostic-logging)
- [Contributing](#-contributing)
- [License](#-license)

---

## 🚀 Key Architecture & Features

* **Visual Embedded Browser Picker:** Point-and-click element capture powered by PySide6 `QWebEngineView` with real-time DOM hover overlays and Javascript event delegation.
* **Smart Heuristic Scoring Engine:** Ranks candidate selectors on a 1–100 scale using attributes like `data-testid`, `data-qa`, `id`, `name`, semantic attributes, unique text, and scoped parent relationships.
* **Multi-Framework Output:** Dynamically generates bindings for **Java Selenium**, **Python Selenium**, **Playwright**, and **Cypress** simultaneously.
* **Non-Blocking Multithreaded Engine:** All network calls, Selenium page drives, and DOM health audits execute off the main thread using worker threads (`QThread` + `Worker`), keeping the GUI fluid at 60 FPS.
* **Browser Driver Fallback:** Intelligent headless runner that attempts Google Chrome first via `selenium.webdriver`, smoothly falling back to Microsoft Edge if Chrome is uninstalled or broken.
* **Page Object Model Generator:** Converts isolated capture bundles into ready-to-use POM class files with camelCase naming heuristics and deduplication logic.
* **Heuristic Health Auditor:** Evaluates active selector lists against live web endpoints, flagging missing or syntax-broken selectors while calculating unverified fallback alternatives.
* **Enterprise Excel & CSV Exporter:** Generates formatted `.xlsx` workbooks (via `openpyxl`) featuring styled header rows, custom status highlights, auto-filters, frozen panes, and strict write verification.

---

## 🛠 System Requirements & Dependencies

### Core Requirements
* **Python:** `3.9` or higher
* **Operating System:** Windows 10/11, macOS, or Linux

### Python Package Dependencies

| Library | Purpose | Mandatory? |
| :--- | :--- | :--- |
| `PySide6` | Qt6 GUI framework & `QWebEngineView` embedded browser | **Required** |
| `beautifulsoup4` | DOM tree parsing & heuristic candidate candidate ranking | **Required** |
| `selenium` | Background URL scanning, live selector verification & audits | Recommended |
| `lxml` | High-performance XPath matching and evaluation | Recommended |
| `openpyxl` | Formatted Excel (`.xlsx`) export generation | Optional |

---

## 📦 Installation & Quick Start

1. **Clone the Repository:**
   ```bash
   git clone https://github.com/your-org/locator-generator-pyside6.git
   cd locator-generator-pyside6
   ```

2. **Create and Activate a Virtual Environment:**
   * **Windows:**
     ```cmd
     python -m venv venv
     venv\Scripts\activate
     ```
   * **macOS / Linux:**
     ```bash
     python3 -m venv venv
     source venv/bin/activate
     ```

3. **Install Core and Recommended Dependencies:**
   ```bash
   pip install PySide6 beautifulsoup4 selenium openpyxl lxml
   ```

4. **Launch the Application:**
   ```bash
   python locator_generator_pyside6_2.py
   ```

---

## 🖥 Detailed Workspace Guide

### 1. Element Picker & HTML Engine

The central hub for locator generation, supporting both live web pages and static HTML snippets.

```
┌─────────────────────────────────────────┬─────────────────────────────────────────┐
│ Embedded Browser / HTML Input           │ Candidate Locators & Code Preview       │
│                                         │                                         │
│ • Live web rendering (QWebEngineView)   │ • Ranked candidate table (0-100 Score)  │
│ • Hover overlay & click delegation      │ • Multi-attribute ignore patterns       │
│ • Static HTML fallback paste editor     │ • Framework selector (Java/Py/PW/Cy)    │
└─────────────────────────────────────────┴─────────────────────────────────────────┘
```

* **Embedded Web Navigation:** Enter any valid HTTP/HTTPS URL and browse inline. Toggle **Click to capture** to highlight DOM elements with a blue target overlay. Clicking any button, link, form, or input creates an isolated capture bundle.
* **Static HTML Fallback Editor:** When operating without WebEngine or offline, paste raw HTML fragments (e.g., `<button data-testid="submit-btn">Submit</button>`) into the editor to process target candidates instantly.
* **Exclusion & Ignore Rules:** Use the persistent **Ignore patterns** checkable menu to filter out fragile strategies (e.g., generic class names, tag names, relative DOM paths) or selector kinds (CSS vs. XPath) prior to scanning.
* **Live Test Verification:** Validate selected candidate locators directly against live web pages using the headless Selenium driver.

---

### 2. Page Objects Builder

Turn captured element bundles into clean, structured Page Object Model class definitions.

* **Automatic Identifier Normalization:** Automatically converts element types, IDs, text content, and `data-testid` attributes into standard `camelCase` variable names.
* **Variable Name Customization:** Inline table controls allow you to select any element, rename its target variable (with real-time regex identifier validation), and re-render the target class.
* **Multi-Framework Rendering:** Instantly convert the exact same Page Object definition into:
  * **Java Selenium:** `private final By loginButton = By.xpath("//button[@id='login']");`
  * **Python Selenium:** `loginButton = (By.XPATH, "//button[@id='login']")`
  * **Playwright:** `self.loginButton = page.locator("xpath=//button[@id='login']")`
  * **Cypress:** `loginButton = () => cy.xpath("//button[@id='login']");`

---

### 3. Live Health Audit & Healing

Maintain existing test suites by detecting broken locators and suggesting smart alternatives.

1. **Paste Existing Locators:** Paste raw CSS/XPath selectors or complete framework line declarations (e.g., `By.id("username")`, `login = (By.CSS_SELECTOR, "#user")`).
2. **Execute Live Verification:** The background engine visits the target URL, querying each element against the live DOM.
3. **Audit Status Breakdown:**
   * **Found:** Element verified and present.
   * **0 matches:** Element missing or locator broken due to DOM changes.
   * **Evaluation error:** Invalid syntax or unparsable expression.
4. **Unverified Fallback Engine:** For `0 matches` results, the auditor analyzes the page source for elements with similar tags, classes, parent containers, or normalized text, producing ranked suggestions with clear matching notes.

---

### 4. Help, Diagnostics & Guide

* **Integrated Manual:** Embedded HTML documentation detailing usage strategies, dependency flags, and quality scoring rules.
* **One-Click Diagnostic Export:** Generates a compressed `.zip` archive containing system details (Python version, PySide6 build, Qt core version, runtime platform) and app log files for rapid troubleshooting.

---

## 💻 Cross-Framework Code Generation

The application normalizes attributes into framework-compliant locator constructs. Below is an example of how a single target candidate is represented across all supported frameworks:

```python
# Target Expression: //input[@data-testid='user-email']
```

| Framework | Formatted Output Binding |
| :--- | :--- |
| **Java Selenium** | `By.xpath("//input[@data-testid='user-email']")` |
| **Python Selenium** | `driver.find_element(By.XPATH, "//input[@data-testid='user-email']")` |
| **Playwright** | `page.locator("xpath=//input[@data-testid='user-email']")` |
| **Cypress** | `cy.xpath("//input[@data-testid='user-email']")` |

---

## ⚡ Technical Design & Concurrency Model

### Non-Blocking Qt Event Loop
To avoid application freezes during browser automation, long-running operational jobs (e.g., live URL scans, multi-locator verification, audit runs) run on worker objects moved to dedicated `QThread` instances:

```
[Main Qt GUI Thread] ──(Signal: Run Task)──► [Worker Object on QThread]
         │                                               │
  (UI Remains Responsive)                       (Executes Selenium/DOM Job)
         │                                               │
[UI Updates / Displays Data] ◄──(Signal: Emit Results)───┘
```

### Automatic Thread Cleanup Sequence
Threads and workers are safely discarded post-execution via Qt signal-slot bindings:
```python
worker.finished.connect(thread.quit)
worker.finished.connect(worker.deleteLater)
thread.finished.connect(thread.deleteLater)
```

---

## 🎨 UI/UX Highlights & Theme Engine

* **Dynamic Dark/Light Modes:** Modern stylesheet styling tuned for long test-authoring sessions.
* **Proportional Responsive UI Scale:** Built-in UI zoom controls (`A-` / `A+` or `Ctrl+-` / `Ctrl++`) scale all fonts, buttons, table row dimensions, and margins dynamically from **50% to 150%**.
* **Responsive Splitter Layouts:** Multi-pane workspace splitters automatically adjust orientation (horizontal vs. vertical) based on window aspect ratio and breakpoint thresholds.
* **Persistent Check Menus:** Custom `PersistentCheckMenu` allows users to select multiple strategies/types without the dropdown closing after every click.

---

## 💾 State Persistence (QSettings)

Application configuration is preserved using standard OS registry/config backends via PySide6 `QSettings`:

```
Windows Registry : HKCU\Software\Locator Generator\Locator Generator
macOS / Linux    : ~/.config/Locator Generator/Locator Generator.conf
```

### Preserved Preferences Across Sessions
- Window dimensions, position, and maximized state.
- Selected theme preference (Dark vs. Light).
- Interface scale factor (Zoom percentage).
- Active test framework choice (Java, Python, Playwright, Cypress).
- Drawer state (Open vs. Collapsed) and custom pane splitter ratios.

---

## 🔍 Troubleshooting & Diagnostic Logging

Logs are automatically stored in a per-user, rotating application log directory:
* **Windows:** `%LOCALAPPDATA%\Locator Generator\logs\app.log`
* **macOS/Linux:** `~/.local/share/Locator Generator/logs/app.log` (or system temporary directory)

### Common Scenarios

#### 1. Embedded Browser Displays a Blank Screen
* **Cause:** `PySide6.QtWebEngineWidgets` is missing or unsupported on your Python build.
* **Fix:** Reinstall PySide6 with WebEngine support:
  ```bash
  pip install --force-reinstall PySide6
  ```

#### 2. Live Scanning or Audit Fails Immediately
* **Cause:** Selenium cannot locate a valid Google Chrome or Microsoft Edge binary.
* **Fix:** Ensure either Chrome or Edge is installed. Selenium Manager will automatically retrieve the required driver binary if an internet connection is present.

---

## 🤝 Contributing

Contributions are welcome! Follow these steps to submit improvements:

1. **Fork the Repository** on GitHub.
2. **Create a Feature Branch:**
   ```bash
   git checkout -b feature/amazing-feature
   ```
3. **Commit Your Changes:**
   ```bash
   git commit -m "Add amazing feature"
   ```
4. **Push to the Branch:**
   ```bash
   git push origin feature/amazing-feature
   ```
5. **Open a Pull Request**.

---

## 📄 License

Distributed under the **MIT License**. See `LICENSE` for more information.
