# Locator Generator Pro

A Python desktop utility for QA engineers and test automation developers
to generate, compare, organize, export, and audit CSS/XPath locators. It
combines HTML-fragment analysis with optional Selenium-powered live-page
workflows and framework-aware code previews.

> **Important:** Locator scores and auto-heal suggestions are
> heuristics. Validate every selector and generated code against your
> application before using it in a test suite.

## Contents

-   [Overview](#overview)
-   [Key Features](#key-features)
-   [Supported Frameworks](#supported-frameworks)
-   [How It Works](#how-it-works)
-   [Installation](#installation)
-   [Run the Application](#run-the-application)
-   [Quick Start](#quick-start)
-   [Live Browser Features](#live-browser-features)
-   [POM Export](#pom-export)
-   [Locator Health Auditor](#locator-health-auditor)
-   [Export Formats](#export-formats)
-   [Dependencies and Optional
    Capabilities](#dependencies-and-optional-capabilities)
-   [Limitations](#limitations)
-   [Recommended Workflow](#recommended-workflow)
-   [Responsible Use](#responsible-use)

## Overview

Locator Generator Pro helps you move from a DOM element to candidate
selectors and automation bindings. It provides a dashboard for locator
generation and review, a Page Object Model (POM) export workspace, a
locator health auditor with fallback suggestions, and an in-app manual.

## Key Features

### 1. Locator Candidate Generation

-   Analyze a pasted HTML fragment.
-   Generate CSS and XPath candidates from common element attributes,
    text, classes, parent context, and DOM structure.
-   Rank candidates using a heuristic score from 0--100, with labels:
    **Excellent**, **Strong**, **Fair**, and **Weak**.
-   Display strategy, expression, estimated match count when available,
    and explanatory notes.
-   Configure comma-separated ignore patterns to filter unwanted
    candidate expressions.

The generator gives preference to signals such as `data-testid`,
`data-test`, `data-qa`, `id`, `name`, and `aria-label`. Class-based,
generic tag, and structural-path selectors generally receive lower
scores. These are recommendations, not guarantees of selector quality.

### 2. Live Page Scanning and Element Picker

When Selenium is installed and Chrome can be launched, the app can:

-   Scan a target URL and produce top locator candidates for selected
    page elements.
-   Open an interactive browser picker that highlights hovered elements
    and captures a clicked element's HTML.
-   Test a selected locator against a live URL.
-   Provide browser back, forward, refresh, pick-mode, and close
    controls.

Live-page support depends on the target site's behavior and the local
browser/Selenium environment.

### 3. Framework-Aware Code Bindings

Select a framework to preview or copy locator code for:

-   **Java Selenium**
-   **Python Selenium**
-   **Playwright**
-   **Cypress**

The tool formats the generated CSS/XPath expression for the selected
framework. Review the result in context; generated bindings do not
replace a complete test or framework setup.

### 4. Page Object Model (POM) Export

-   Add multiple selected locator candidates to a POM workspace.
-   Generate descriptive variable names from locator expressions.
-   Skip duplicate expressions and avoid duplicate variable names.
-   Set the target page class name.
-   Preview and copy the generated class definition.

The generated POM code is a starting point. Adapt imports, visibility,
element access, waits, and project conventions as needed.

### 5. Locator Health Auditor & Fallback Suggestions

-   Audit a list of locators against a live URL.
-   Report whether a locator matches elements or returns zero matches.
-   Identify invalid locator expressions as failed checks.
-   Suggest a simple fallback in some cases when the original locator
    appears broken.

**Auto-healing is heuristic and limited.** Suggestions may be
unavailable or incorrect; manually inspect and re-test them before
replacing a production locator. The feature does not use an external AI
service.

### 6. Export and Productivity

-   Export results to `.xlsx` when `openpyxl` is available.
-   Export to CSV as an alternative.
-   Include locator type, strategy, expression, stability rank/score,
    match count, framework code binding, and notes.
-   Copy one or multiple expressions or framework bindings.
-   Switch between dark and light themes.
-   Access the user manual from the **Manual & Docs** tab.

## Supported Frameworks

  Framework         Generated binding style
  ----------------- -----------------------------------------
  Java Selenium     `By.cssSelector(...)` / `By.xpath(...)`
  Python Selenium   `driver.find_element(...)`
  Playwright        `page.locator(...)`
  Cypress           `cy.get(...)` / `cy.xpath(...)`

The tool generates bindings; it does not install or configure these
automation frameworks in your project.

## How It Works

1.  **Provide an element:** paste HTML into the Dashboard, use the
    built-in example, or capture an element from a live page.
2.  **Generate candidates:** the app extracts attributes and contextual
    signals and creates candidate CSS/XPath expressions.
3.  **Review scores and notes:** use the ranking to prioritize
    candidates, then validate uniqueness and semantic correctness.
4.  **Use the output:** copy a selector or framework binding, add
    candidates to POM Export, or send them to the Health Auditor.
5.  **Export results:** save the generated candidate list as Excel or
    CSV for review or documentation.

## Installation

### Requirements

-   Python 3.8 or later
-   Tkinter support for the desktop GUI
-   Google Chrome for Selenium-powered live browser features
-   Compatible Selenium/Chrome setup for live scanning, picking,
    testing, and auditing

### Install dependencies

From a terminal in the project directory:

``` bash
py -m pip install beautifulsoup4 selenium openpyxl lxml
```

On systems where `py` is not available, use:

``` bash
python -m pip install beautifulsoup4 selenium openpyxl lxml
```

`lxml` is used where available for XPath candidate match counting. The
application can still attempt HTML parsing through BeautifulSoup's
built-in parser.

## Run the Application

Save the Python source file in your project directory, then run it. For
example, if the file is named `locator_generator.py`:

``` bash
py locator_generator.py
```

Or:

``` bash
python locator_generator.py
```

If your filename contains parentheses, quote it in the terminal when
needed, for example:

``` bash
py "locator_generator(1).py"
```

## Quick Start

### Generate locators from an HTML snippet

1.  Launch the app and open **Dashboard**.
2.  Click **Load Example**, or paste the target element's HTML into the
    HTML DOM input.
3.  Click **Generate Locators**.
4.  Review candidate type, strategy, score, match count, and notes.
5.  Select a row to preview its framework binding.
6.  Use **Copy Selected Expression** or **Copy Code Binding** to reuse
    the output.

Example HTML:

``` html
<div class="form-group">
  <label for="username">Username</label>
  <input id="username" name="userName"
         class="form-control"
         placeholder="Enter username"
         type="text">
</div>
```

### Scan a live page

1.  Enter a URL in **Target URL**.
2.  Click **Scan Whole Page** to collect top candidates for relevant
    page elements, or **Open Picker** to select a specific element
    visually.
3.  Review and validate the results before adding them to your tests.

### Change framework output

Use the **Framework Target** dropdown and select Java Selenium, Python
Selenium, Playwright, or Cypress. Select a candidate to view the
formatted binding, then copy it.

## Live Browser Features

The live browser controls use Selenium with Chrome. Depending on your
installed Selenium version and environment, Selenium Manager may help
locate a compatible driver automatically. If Chrome cannot launch, check
Chrome installation, network/driver restrictions, and Selenium setup.

Live features may be affected by authentication, redirects, dynamic
rendering, frames, shadow DOM, browser security restrictions, or
application-specific behavior.

## POM Export

1.  Select one or more locator rows in the Dashboard (use Ctrl-click or
    Shift-click for multiple selections).
2.  Click **Add Selected to POM**.
3.  Open **POM Export**.
4.  Set a class name such as `LoginPage`.
5.  Review the generated class preview.
6.  Click **Copy POM Class Code**.

Duplicate expressions are skipped, and duplicate variable names receive
numeric suffixes.

## Locator Health Auditor

1.  Open **Health Auditor & Auto-Heal**.
2.  Enter the audit target URL.
3.  Paste one locator per line, or send selected candidates from the
    Dashboard.
4.  Click **Run Audit & Auto-Heal**.
5.  Review each locator's health status and any suggested replacement.

The current implementation treats expressions beginning with `//` or `(`
as XPath and other expressions as CSS selectors. Ensure your locator
syntax matches this convention when pasting locators into the auditor.

## Export Formats

Use **Export** on the Dashboard:

-   **Excel (`.xlsx`)** --- requires `openpyxl`.
-   **CSV (`.csv`)** --- suitable for spreadsheets, version control, or
    further processing.

The export contains the candidate locator, type, strategy, score,
estimated matches, framework-specific binding, and notes.

## Dependencies and Optional Capabilities

  -----------------------------------------------------------------------
  Package                             Purpose
  ----------------------------------- -----------------------------------
  `beautifulsoup4`                    Parses pasted HTML and page source

  `selenium`                          Live browser scanning, element
                                      picker, locator testing, and
                                      auditing

  `openpyxl`                          Excel workbook export

  `lxml`                              XPath candidate match counting when
                                      available

  `tkinter`                           Desktop GUI (typically bundled with
                                      Python; packaging varies by OS)
  -----------------------------------------------------------------------

The app checks for several dependencies at startup. Without
BeautifulSoup, HTML-based generation is unavailable. Without Selenium,
live browser features are disabled. Without `openpyxl`, export can use
CSV instead.

## Limitations

-   Scores are heuristic and cannot prove a selector will remain stable
    after application changes.
-   A match count greater than zero does not prove the selector
    identifies the intended element; a selector matching multiple
    elements may be ambiguous.
-   Live-page scanning processes a bounded set of relevant elements and
    may not cover every element on complex pages.
-   The picker captures an element's outer HTML; it does not
    automatically solve every iframe, shadow DOM, or dynamic-content
    scenario.
-   Fallback suggestions are limited and should not be treated as
    guaranteed repairs.
-   Framework bindings and POM code may need adaptation to match your
    project's conventions and dependencies.
-   Browser startup and compatibility depend on the local
    Chrome/Selenium environment.

## Recommended Workflow

1.  Prefer stable, semantic attributes such as dedicated test IDs where
    available.
2.  Check that the locator uniquely identifies the intended element.
3.  Test the locator against the target application and relevant states.
4.  Use POM Export to organize reusable locators, then review the
    generated code.
5.  Run the Health Auditor after meaningful UI changes, and inspect
    every fallback suggestion.
6.  Keep important locators covered by automated tests and maintain them
    with code review.

## Responsible Use

Use live browser features only against applications and environments you
are authorized to access and test. Do not use the tool to bypass access
controls or interact with systems without permission.

## License

No license is specified in the provided source file. Add a `LICENSE`
file and update this section if you intend to distribute the project
publicly.
