======================================================================
               SELENIUM LOCATOR GENERATOR PRO
             Complete Installation & User Guide
======================================================================

A professional-grade desktop application built with Python and Tkinter, 
designed for QA automation engineers to generate, test, and export stable, 
scored CSS and XPath locators using live browser inspection.


----------------------------------------------------------------------
1. FULL SYSTEM & PACKAGE INSTALLATION (DEVELOPMENT CYCLE)
----------------------------------------------------------------------
Throughout the development and setup of this tool, the following components 
and Python libraries were installed and configured on the PC:

* Python 3.8+ (Core runtime environment)
* Google Chrome (Target browser for live inspection and testing)
* Python Packages installed via terminal/CMD:
  - beautifulsoup4: Used to parse outerHTML strings and analyze DOM trees.
  - selenium: Powers the automated browser session, element picking, and live testing.
  - lxml: Provides advanced HTML/XML processing engine for precise XPath match counting.


----------------------------------------------------------------------
2. TERMINAL COMMANDS USED (CMD / POWERSHELL)
----------------------------------------------------------------------
Keep these commands handy for setting up and running your project:

A. Install Required Dependencies:
   py -m pip install beautifulsoup4 selenium lxml
   *(Or if Python is directly added to your system PATH: pip install beautifulsoup4 selenium lxml)*

B. Navigate to Project Directory:
   cd path\to\your\project_folder

C. Run the Application:
   python locator_generator.py
   *(Or using the Windows Python launcher: py locator_generator.py)*


----------------------------------------------------------------------
3. USER GUIDE & FEATURE WALKTHROUGH
----------------------------------------------------------------------

A. MANUAL HTML PARSING
   * Paste any raw HTML fragment into the "HTML Input Fragment" text box.
   * Click "⚡ Generate Locators" to instantly analyze the element and view 
     scored candidates in the table below.

B. LIVE BROWSER ELEMENT PICKING
   * Enter your target website URL in the "Target URL" box.
   * Click "🔍 Open Picker". This launches a live Chrome browser session 
     connected to your URL.
   * Hover over any element on the webpage to see a visual highlight border, 
     and click it to instantly capture its HTML and generate its locators.

C. PICK MODE VS. FREE NAVIGATION
   * "Pick Mode Active" checkbox:
     - CHECKED: Hovering highlights elements in red and clicking captures them.
     - UNCHECKED: Element picking is paused, allowing you to click links, 
       buttons, and navigate your web app freely without accidental capture.
   * Browser Tools (◀ Back, Forward ▶, 🔄 Refresh): Use these built-in 
     controls to navigate through your web application pages seamlessly 
     without closing the session.

D. STABILITY SCORING & NOISE FILTERING
   * Locators are automatically ranked (Excellent, Strong, Fair, Weak) 
     based on best automation practices (prioritizing `data-testid`, IDs, and semantic attributes).
   * Customize "Ignore Patterns" (comma-separated) to automatically filter out 
     volatile framework classes (such as Angular `ng-star-inserted`, React hashes, or styled-components).

E. TESTING & EXPORTING
   * Test Locator: Select a row in the table, enter your URL, and click 
     "🧪 Test Locator" to verify matching element counts via headless Chrome.
   * Copying: Double-click any row or right-click to copy individual expressions 
     or formatted Java Selenium `By` statements.
   * Bulk Export: Click "📋 Copy All Java" to copy all generated candidates 
     formatted into a clean Java class snippet ready for your test automation framework.