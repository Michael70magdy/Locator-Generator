# 🎯 Locator Generator

A Python desktop utility for generating, comparing, validating, and exporting CSS and XPath locator candidates for Selenium-based UI automation.

Locator Generator combines static HTML-fragment analysis with optional live Chrome workflows. It ranks candidate selectors using a transparent heuristic, filters unwanted patterns, supports interactive element picking, and can export results for review or reuse in Selenium Java tests.

> **Important:** Scores are heuristic guidance, not a guarantee of uniqueness or long-term stability. Always review and validate locators against the actual rendered page.

## ✨ Features

- **CSS and XPath candidate generation** from element attributes, text, classes, parent context, and DOM structure.
- **Heuristic stability scores** from 0–100 with labels: Excellent, Strong, Fair, and Weak.
- **Strategy explanations** to help compare test attributes, IDs, semantic attributes, text, classes, and structural paths.
- **Match counts** where the current HTML parser can evaluate a candidate.
- **Ignore-pattern filtering** to suppress selectors containing volatile tokens such as `css-`, `sc-`, or `ng-star-inserted`.
- **HTML fragment mode** for pasted markup or a sample element.
- **Whole-page scanning** using Selenium with headless Chrome.
- **Interactive Chrome picker** with hover highlighting and click-to-capture HTML.
- **Live locator testing** against a target URL.
- **Selenium Java output** using `By.cssSelector(...)` and `By.xpath(...)`.
- **Copy selected / copy all** helpers for moving candidates into test code.
- **Excel (`.xlsx`) and CSV export** with strategy, score, match count, Java binding, and notes.
- **Dark and light themes**.
- **Built-in About & User Manual** covering setup and workflows.

## 🧰 Technology stack

- Python 3.8+
- Tkinter — desktop GUI
- BeautifulSoup 4 — HTML parsing
- Selenium — live Chrome browser automation
- lxml — XPath evaluation against parsed HTML
- openpyxl — Excel workbook export

Some live-browser and export capabilities depend on optional packages being installed.

## 📦 Installation

### 1. Prerequisites

- Python 3.8 or newer
- Google Chrome installed for live browser workflows
- Network access to the target page for live scanning/testing
- A working Selenium-compatible Chrome/WebDriver setup. Selenium Manager may manage the driver in supported environments.

### 2. Install dependencies

On Windows, open Command Prompt or PowerShell in your project directory and run:

```bash
py -m pip install beautifulsoup4 selenium openpyxl lxml
```

Alternatively:

```bash
python -m pip install beautifulsoup4 selenium openpyxl lxml
```

### 3. Run the application

Save the script in your project directory (for example, as `locator_generator.py`) and run:

```bash
py locator_generator.py
```

If your downloaded script has a different filename, use that filename in the command.

## 🖥️ How to use it

### Workflow A — Generate locators from an HTML fragment

1. Launch the application.
2. In **HTML Input Fragment**, paste the HTML element markup.
3. Optionally click **Load Example** to inspect a sample form field.
4. Configure **Ignore Patterns** if you want to filter volatile selector strings. Use comma-separated patterns.
5. Click **Generate Locators**.
6. Compare the candidate locator expression, strategy, score, match count, and note.
7. Select a candidate to preview the corresponding Selenium Java `By` expression.
8. Copy the expression or Java binding, or export the candidate list.

The current implementation parses the fragment and chooses a preferred element based on available attributes such as `data-testid`, `data-test`, `data-qa`, `id`, `name`, `aria-label`, and `placeholder`. If none are present, it falls back to a parsed element. For best results, provide markup that clearly represents the element you want to inspect.

### Workflow B — Scan a live page

1. Enter the target page URL.
2. Click **Scan Whole Page**.
3. The tool opens Chrome in headless mode through Selenium, retrieves the page source, and generates top candidate locators for a bounded set of common controls and selected identified/test-attributed elements.
4. Review the results and use live testing or export as needed.

This is a bounded scan, not a guarantee that every element on a complex page will be included. Dynamic content may require additional handling.

### Workflow C — Use the interactive browser picker

1. Enter the target URL and click **Open Picker**.
2. In the Chrome window, move the pointer over elements to see the highlight overlay.
3. With **Pick Mode Active** enabled, click the target element.
4. The tool captures the element's outer HTML and generates locator candidates.
5. Use **Back**, **Forward**, and **Refresh** to navigate when needed.
6. Disable pick mode to interact with the page without capturing clicks, then enable it again when ready.
7. Click **Close Browser** when finished.

The picker intercepts clicks while pick mode is active so that an element can be captured. Use it only on pages where you are authorized to test.

### Workflow D — Test a locator against a live URL

1. Generate or select a candidate locator.
2. Enter the URL where the locator should be checked.
3. Click **Test Locator**.
4. Review the number of matching elements reported by Selenium.

A successful query only confirms that the expression can be evaluated on that page at that time. Confirm that it identifies the intended element, especially when the count is greater than one.

### Workflow E — Copy or export results

- **Copy Selected Expression** copies the CSS or XPath expression.
- **Copy Selected Java By** copies a Java Selenium binding.
- **Copy All Java** copies the current candidate set as Java `By` declarations with comments.
- **Export to Excel** opens a save dialog for `.xlsx` or `.csv`.

The export contains the locator type, strategy, expression, rank, score, match count, Java binding, and notes. If `openpyxl` is unavailable, Excel export may fall back to CSV.

## 📊 Understanding the scores

Scores prioritize strategies using a fixed heuristic in the code. Examples include:

| Strategy / attribute | Example score | Notes |
|---|---:|---|
| `data-testid` | 99 | Dedicated test attribute |
| `data-test` | 98 | Dedicated test attribute |
| `data-qa` | 97 | QA-oriented attribute |
| ID selector shorthand | 96 | Concise ID selector |
| XPath using `id` | 95 | ID-based XPath |
| `id` attribute | 94 | Good if stable and unique |
| `name` | 91 | Semantic form attribute |
| `aria-label` | 88 | Accessibility-oriented attribute |
| Class + text combination | 86 | Contextual candidate |
| `placeholder` | 81 | Can change with UI copy |
| Exact normalized text | 78 | Sensitive to text changes |
| Tag + all classes | 73 | Depends on class stability |
| Generic tag / DOM path | 30–35 | Often less specific or more brittle |

The application labels scores as:

- **Excellent:** 90–100
- **Strong:** 75–89
- **Fair:** 55–74
- **Weak:** below 55

These are code-defined ranking thresholds. They are not measured probabilities and do not guarantee that a locator will remain stable after an application update.

## 🧹 Ignore patterns

The default patterns include:

```text
css-, sc-, ng-star-inserted
```

Enter multiple patterns separated by commas. The implementation treats them as regular expressions and checks them case-insensitively. Use simple tokens unless you intentionally need regular-expression behavior. Invalid regular expressions may cause candidate generation to fail.

Ignore patterns can help remove known noisy class names, but do not filter out meaningful attributes without checking the impact on locator quality.

## 🧪 Example Selenium Java output

CSS selector:

```java
By.cssSelector("[data-testid=\"login-button\"]")
```

XPath:

```java
By.xpath("//button[@id='login-button']")
```

Use the generated expression in your page object or test, then verify it against the actual page and your framework conventions.

## 📁 Project layout

```text
your-project/
├── locator_generator.py
└── README.md
```

The application is currently implemented as a single Python script. The exact filename may differ in your repository.

## ⚠️ Limitations and notes

- **Heuristic scoring:** scores are strategy-based, not learned from historical DOM changes.
- **Static HTML context:** fragment analysis cannot fully reproduce the rendered application, runtime state, or dynamically injected elements.
- **Live scanning scope:** the current whole-page scan is bounded to a subset of common controls and elements with selected attributes.
- **Browser compatibility:** live features require Selenium and a working Chrome/WebDriver setup.
- **Match counts:** counts from the parsed fragment are not equivalent to counts on the complete live page.
- **Modern browser features:** shadow DOM, complex frames, and highly dynamic applications may need additional locator strategies or manual inspection.
- **Regex patterns:** ignore patterns are treated as regular expressions.
- **Security and authorization:** use the tool only on pages you are permitted to inspect or test. Avoid placing credentials or sensitive data in exported reports.

## 🛠️ Troubleshooting

### Live buttons are disabled or unavailable
Install Selenium:

```bash
py -m pip install selenium
```

### HTML parsing is unavailable
Install BeautifulSoup:

```bash
py -m pip install beautifulsoup4
```

### XPath match counts are unavailable
Install `lxml`:

```bash
py -m pip install lxml
```

### Excel export does not create an `.xlsx` file
Install `openpyxl`:

```bash
py -m pip install openpyxl
```

### Chrome fails to start
- Confirm Google Chrome is installed.
- Update Selenium.
- Check that your environment permits Chrome/WebDriver to launch.
- Review the error shown by the application; browser policies or network restrictions may affect live workflows.

## 💡 Best practices

1. Prefer dedicated test attributes such as `data-testid` when the application team supports them.
2. Prefer stable, unique semantic attributes over long structural paths.
3. Check the locator against the full page, not only a copied fragment.
4. Verify that the locator points to the intended element and review match counts.
5. Avoid relying on classes that change with styling, build output, or UI state.
6. Keep locators in Page Objects or another consistent abstraction in your automation framework.
7. Revalidate important locators when the application UI changes.

## 🎯 Intended audience

- QA engineers and software testers
- Selenium automation engineers and SDETs
- QA learners practicing XPath and CSS selectors
- Teams that want a lightweight desktop helper for locator discovery and documentation

## 📄 License

No license is specified in the supplied project source. Add a `LICENSE` file and update this section if you intend to publish the project for reuse.

---

**Locator Generator** — discover candidate locators, validate them, and move them into your Selenium workflow with less manual effort.
