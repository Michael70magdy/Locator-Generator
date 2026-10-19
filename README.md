# ⚡ LocatorPro — Selenium Locator Generator

**A Python-powered desktop application for generating, analyzing, and testing CSS selectors and XPath locators for Selenium WebDriver automation.**

LocatorPro helps QA engineers and automation testers discover potential web element locators, compare their stability scores, and generate Java Selenium `By` statements. It provides both an HTML-based locator generation workflow and an interactive Chrome element picker.

---

## 📌 Table of Contents

- [Overview](#-overview)
- [Key Features](#-key-features)
- [Application Preview](#-application-preview)
- [Technology Stack](#-technology-stack)
- [Requirements](#-requirements)
- [Installation](#-installation)
- [How to Use](#-how-to-use)
- [Locator Scoring System](#-locator-scoring-system)
- [Generated Selenium Java Code](#-generated-selenium-java-code)
- [Ignore Patterns](#-ignore-patterns)
- [Project Structure](#-project-structure)
- [Troubleshooting](#-troubleshooting)
- [Limitations](#-limitations)
- [Future Improvements](#-future-improvements)
- [Author](#-author)
- [License](#-license)

---

## 🔎 Overview

LocatorPro is a desktop utility designed to simplify one of the most common tasks in web test automation: identifying reliable element locators.

Writing and maintaining locators manually can be time-consuming, especially when working with dynamic web applications, complex DOM structures, or frequently changing CSS classes.

LocatorPro analyzes HTML elements and generates multiple CSS selector and XPath candidates using attributes, element text, class names, parent elements, and DOM structure.

Each candidate receives a heuristic stability score to help testers prioritize potentially reliable locators. Users can review the generated expressions, inspect the matching information available from the supplied HTML, and copy Selenium Java statements directly into their automation frameworks.

The application also integrates with Selenium WebDriver to support interactive element selection and live locator testing in Google Chrome.

### 🎯 Project Goals

- Reduce the time required to identify web element locators.
- Generate alternative CSS selectors and XPath expressions.
- Encourage the use of stable, maintainable locator strategies.
- Simplify the transfer of generated locators into Selenium Java projects.
- Support practical QA automation workflows through a desktop GUI.

---

## ✨ Key Features

### 1. 🧭 Intelligent Locator Generation

Generate multiple locator candidates from an HTML element, including:

- CSS selectors based on HTML attributes.
- XPath expressions based on HTML attributes.
- ID-based selectors.
- Name, placeholder, title, role, and accessibility-label selectors.
- Data-test attribute selectors, including `data-testid`, `data-test`, and `data-qa`.
- Class-based CSS and XPath selectors.
- Exact and partial text-based XPath expressions.
- Parent-scoped selectors.
- Relative DOM path expressions.
- Generic tag-name selectors.

### 2. 📊 Locator Stability Scoring

Each generated candidate receives a score from 0 to 100 and a corresponding classification.

| Classification | Score |
|---|---:|
| Excellent | 90–100 |
| Strong | 75–89 |
| Fair | 55–74 |
| Weak | Below 55 |

The scoring system uses predefined heuristics based on locator strategy and attribute characteristics. It helps users compare candidates, but it does not guarantee that a locator will remain stable or unique on the full website.

### 3. 🌐 Interactive Chrome Element Picker

Use the integrated browser picker to inspect a website interactively.

Features include:

- Open a target URL in Chrome.
- Highlight elements when hovering over them.
- Capture an element by clicking it.
- Automatically load the selected element's HTML into the application.
- Navigate backward and forward.
- Refresh the browser.
- Enable or disable element-picking mode.
- Close the picker browser when finished.

The interactive picker requires Selenium WebDriver and a working Chrome installation.

### 4. 🧪 Live Locator Testing

Test a selected CSS or XPath locator against a live website using headless Chrome.

The application:

1. Opens the configured URL.
2. Loads the page using Selenium WebDriver.
3. Searches for elements matching the selected locator.
4. Reports the number of matching elements.
5. Displays the test result.

A locator matching exactly one element is often a useful starting point, but the tester must also verify that it identifies the intended element.

### 5. ☕ Selenium Java Code Generation

Copy generated locators in Selenium Java `By` format.

Supported output examples:

```java
By.id("username");
By.cssSelector("input[name=\"userName\"]");
By.xpath("//input[@id='username']");
```

The actual generated output depends on the selected locator candidate.

### 6. 📋 Clipboard Integration

Copy individual expressions, Selenium Java statements, or a complete collection of generated locators.

Available actions include:

- Copy Selected Expression.
- Copy Selected Java By.
- Copy All Java.
- Double-click a result row to copy its locator expression.

### 7. 🎨 Modern Desktop Interface

The application includes:

- Dark and light themes.
- Sidebar navigation.
- Dashboard and documentation views.
- HTML source editor.
- Locator results table.
- Stability score and match-count columns.
- Selected-locator Java preview.
- Context-menu actions.
- Status messages and error dialogs.

### 8. 🧹 Dynamic Class Noise Filtering

Configure comma-separated regular-expression patterns to exclude matching attribute values and CSS class tokens when generating candidates.

This can help reduce reliance on potentially unstable framework-generated classes.

---

## 🖥️ Application Preview

The application provides two primary views.

### Dashboard

The dashboard contains the HTML editor, browser controls, generation actions, and locator results.

Suggested screenshots to add to this repository:

- `screenshots/dashboard.png` — Main application dashboard.
- `screenshots/dark-mode.png` — Dark theme.
- `screenshots/light-mode.png` — Light theme.
- `screenshots/about-documentation.png` — About and usage instructions.
- `screenshots/locator-results.png` — Generated CSS and XPath candidates.

To display a screenshot after uploading it to the repository, use:

```markdown
![LocatorPro Dashboard](screenshots/dashboard.png)
```

These are suggested screenshot paths; add the actual screenshots before using the image links.

---

## 🛠️ Technology Stack

| Technology | Purpose |
|---|---|
| Python | Application logic and locator generation |
| Tkinter | Desktop graphical user interface |
| ttk | Themed widgets and result tables |
| BeautifulSoup4 | HTML parsing and DOM element analysis |
| Selenium WebDriver | Browser automation and live locator testing |
| Google Chrome | Interactive picker and headless testing browser |
| lxml | Optional XPath match counting against parsed HTML |

---

## 📋 Requirements

Before running LocatorPro, ensure the following are available:

- Python 3.8 or newer.
- Google Chrome for browser-based functionality.
- A working Selenium-compatible Chrome WebDriver setup.
- BeautifulSoup4 for HTML parsing.
- Selenium for browser interaction.
- lxml for optional XPath counting against the supplied HTML.

**Note:** Selenium 4.6 and newer generally includes Selenium Manager, which can help manage compatible browser drivers automatically. Browser and driver compatibility may still affect execution.

---

## 🚀 Installation

### Step 1: Clone the Repository

```bash
git clone https://github.com/YOUR_USERNAME/YOUR_REPOSITORY.git
```

Replace `YOUR_USERNAME` and `YOUR_REPOSITORY` with your GitHub details.

### Step 2: Navigate to the Project Directory

```bash
cd YOUR_REPOSITORY
```

### Step 3: Install Dependencies

```bash
python -m pip install beautifulsoup4 selenium lxml
```

On Windows, you can also use:

```bash
py -m pip install beautifulsoup4 selenium lxml
```

### Step 4: Launch the Application

If your main Python file is named `locator_generator_enhanced.py`:

```bash
python locator_generator_enhanced.py
```

If you use the original filename:

```bash
python locator_generator.py
```

Run the filename that exists in your project directory.

---

## 📖 How to Use

### Method 1: Generate Locators from HTML

This method is useful when you already have the HTML of the element you want to automate.

**Step 1 — Launch LocatorPro**

Open the application and navigate to the Dashboard.

**Step 2 — Prepare the HTML**

Paste the target element's HTML into the DOM Element Source Inspector.

Example:

```html
<div class="form-group">
    <label for="username">Username</label>
    <input
        id="username"
        name="userName"
        class="form-control"
        placeholder="Enter username"
        type="text"
    >
</div>
```

**Step 3 — Generate Candidates**

Click **Generate Locators**.

The application parses the HTML and generates locator candidates for the selected target element.

**Step 4 — Review the Results**

Inspect the following columns:

- **Type:** CSS or XPath.
- **Strategy:** The method used to construct the locator.
- **Locator Expression:** The actual selector or XPath.
- **Stability Score:** The heuristic score and classification.
- **HTML Matches:** The number of matches in the supplied HTML fragment, when available.
- **Notes:** An explanation of the locator strategy.

**Step 5 — Copy a Locator**

Select a candidate and use the copy actions to transfer the expression or Selenium Java statement to your automation project.

### Method 2: Use the Interactive Element Picker

This method helps you capture an element directly from a website.

1. Enter the target website URL.
2. Click **Open Element Picker**.
3. Wait for Chrome to open the page.
4. Move your mouse over elements to see the highlight overlay.
5. Ensure picking mode is enabled.
6. Click the desired element.
7. The selected element's HTML is transferred to LocatorPro.
8. Review the generated candidates.
9. Select and copy the locator you want to use.

Use **Back**, **Forward**, and **Refresh** to navigate the browser. Disable picking mode when you need to interact with the page without capturing elements.

### Method 3: Test a Locator on a Live Website

1. Enter the target website URL.
2. Generate candidates or select an existing candidate.
3. Click the desired locator row.
4. Click **Test Selected Locator**.
5. Review the number of matching elements reported by the application.

A locator returning zero matches may be invalid for the current page, may depend on a different DOM state, or may require a different browsing context. Multiple matches may indicate that a more specific locator is needed.

### Method 4: Export Selenium Java Statements

Select a candidate and click **Copy Selected Java By**.

Example:

```java
By.xpath("//input[@id='username']");
```

For a complete list, click **Copy All Java**. The generated output includes comments describing the candidate strategies, scores, and fragment match counts, followed by Java `By` declarations.

You can then adapt the declarations to your Page Object Model (POM) classes or existing Selenium test framework.

---

## 📈 Locator Scoring System

LocatorPro prioritizes locator candidates according to predefined heuristic scores.

Examples of the scoring rules include:

| Locator Strategy | Score | Classification |
|---|---:|---|
| `data-testid` attribute | 99 | Excellent |
| `data-test` attribute | 98 | Excellent |
| `data-qa` attribute | 97 | Excellent |
| Stable `id` attribute | 94 | Excellent |
| `name` attribute | 91 | Excellent |
| `aria-label` attribute | 88 | Strong |
| `placeholder` attribute | 81 | Strong |
| Exact normalized text | 78 | Strong |
| `title` attribute | 76 | Strong |
| Tag plus all classes | 73 | Fair |
| Generic tag selector | 30 | Weak |

These values represent the application's configured heuristic scores, not measured probabilities of locator reliability.

### Recommended Locator Selection Strategy

When possible, prioritize:

1. Dedicated test attributes such as `data-testid`.
2. Stable and unique IDs.
3. Semantic attributes such as `name`.
4. Accessible labels and other meaningful attributes.
5. Text-based or parent-scoped selectors when appropriate.
6. Class-based selectors when the classes are stable.
7. Long structural DOM paths only when better alternatives are unavailable.

Always validate the locator against the full page and confirm that it identifies the correct element.

---

## 💻 Generated Selenium Java Code

Example HTML:

```html
<input
    id="username"
    name="userName"
    data-testid="login-username"
    placeholder="Enter username"
    type="text"
/>
```

Possible generated candidates include:

```java
By.cssSelector("input[data-testid=\"login-username\"]");

By.xpath("//input[@data-testid='login-username']");

By.id("username");

By.cssSelector("input[name=\"userName\"]");

By.xpath("//input[@placeholder='Enter username']");
```

Example usage in a Selenium Java test:

```java
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class LocatorExample {

    public static void main(String[] args) {
        WebDriver driver = new ChromeDriver();

        try {
            driver.get("https://example.com");

            WebElement element = driver.findElement(
                By.cssSelector("input[data-testid=\"login-username\"]")
            );

            element.sendKeys("test-user");

        } finally {
            driver.quit();
        }
    }
}
```

The sample demonstrates how a generated locator can be used in Selenium. The example URL and element are illustrative; they are not guaranteed to exist on the same page.

---

## 🧹 Ignore Patterns

The Ignore Patterns field accepts comma-separated regular expressions.

Default patterns:

```text
css-, sc-, ng-star-inserted
```

These patterns can be used to filter values associated with dynamically generated or potentially unstable CSS classes and attributes.

For example:

```text
css-, sc-, ng-star-inserted, active, selected
```

Use these filters carefully. They may exclude legitimate and useful locator values if the patterns are too broad.

---

## 📁 Project Structure

A minimal repository structure can look like this:

```text
LocatorPro/
│
├── locator_generator_enhanced.py
├── README.md
├── requirements.txt
├── .gitignore
│
└── screenshots/
    ├── dashboard.png
    ├── dark-mode.png
    ├── light-mode.png
    └── about-documentation.png
```

The screenshot directory is optional. Include it only if you add actual application screenshots.

### Example `requirements.txt`

```text
beautifulsoup4
selenium
lxml
```

Create this file in the repository root if you want to install dependencies using:

```bash
python -m pip install -r requirements.txt
```

---

## 🛠️ Troubleshooting

### BeautifulSoup Is Missing

Error or message:

```text
Install BeautifulSoup with: py -m pip install beautifulsoup4
```

Solution:

```bash
python -m pip install beautifulsoup4
```

### Selenium Is Not Available

Install Selenium:

```bash
python -m pip install selenium
```

Restart the application afterward.

### Chrome Does Not Open

Check that:

- Google Chrome is installed.
- Selenium is installed correctly.
- Your network and browser policies allow access to the target site.
- Selenium Manager can obtain a compatible driver, or a compatible ChromeDriver is available.

### The Live Test Finds Zero Elements

Check that:

- The URL is correct.
- The selected locator is valid for the live page.
- The element is present in the DOM when the test runs.
- The locator is not being evaluated inside an iframe or shadow DOM context.
- The element does not require additional navigation or interaction before appearing.

### The Generated Locator Matches Multiple Elements

Try a more specific locator using a stable test attribute, ID, name, accessible label, or suitable parent context.

### XPath Match Count Is Unavailable

The application uses `lxml` when available to count XPath matches in the supplied HTML fragment. Install it with:

```bash
python -m pip install lxml
```

Live XPath testing is handled separately by Selenium.

---

## ⚠️ Limitations

- Locator scores are heuristic recommendations, not guarantees of stability.
- HTML-fragment match counts do not necessarily represent uniqueness on the full webpage.
- Live testing uses the current page state and does not automatically handle every dynamic loading scenario.
- Iframes and shadow DOM may require additional handling.
- Chrome, Selenium, and driver compatibility can affect browser-based features.
- The application generates candidate locators; it does not automatically rewrite an existing test suite.
- Generated Java statements may need adaptation to your project's naming conventions, Page Object Model, and framework structure.

---

## 🔮 Future Improvements

Potential enhancements for future versions include:

- [ ] Export locator candidates to CSV and JSON.
- [ ] Generate ready-to-use Page Object Model classes.
- [ ] Add explicit wait and dynamic element handling recommendations.
- [ ] Add a locator uniqueness indicator based on the full live DOM.
- [ ] Add locator history and saved projects.
- [ ] Add additional browser support.
- [ ] Improve HTML editor syntax highlighting.
- [ ] Add unit tests and automated CI checks.
- [ ] Package the application as a standalone Windows executable.
- [ ] Add locator comparison and maintenance recommendations.

---

## 👨‍💻 Author

**Michael Magdy Fekry**

Software Testing and QA Automation enthusiast focused on test automation, Selenium WebDriver, API testing, mobile testing, and performance testing.

### Technical Interests

- Manual and Automation Testing
- Selenium WebDriver with Java
- Python and Test Automation
- Page Object Model (POM)
- API Testing
- Mobile Application Testing
- Performance Testing
- QA Tools and Framework Development

---

## 📄 License

No license has been specified for this project yet.

If you intend to make the project open source, consider adding an appropriate license file, such as the MIT License, after deciding how others may use, modify, and distribute your work.

---

**If LocatorPro helps your testing workflow, consider giving the repository a ⭐ on GitHub.**

*Built to make Selenium locator discovery faster, clearer, and easier to maintain.*
