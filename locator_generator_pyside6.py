"""Locator Generator — PySide6 desktop edition.

Install: pip install PySide6 beautifulsoup4 selenium openpyxl lxml
Run:     python locator_generator_pyside6.py

The locator-generation engine below is preserved from the original Tkinter project.
"""
import csv
import queue
import re
import sys
import time
from pathlib import Path
from urllib.parse import urlparse

from PySide6.QtCore import Qt, QObject, QThread, Signal, QPoint
from PySide6.QtGui import QFont, QKeySequence, QShortcut, QPainter, QPolygon, QPalette
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QFrame, QLabel, QPushButton,
    QVBoxLayout, QHBoxLayout, QGridLayout, QLineEdit, QPlainTextEdit,
    QTableWidget, QTableWidgetItem, QAbstractItemView, QHeaderView,
    QSplitter, QStackedWidget, QComboBox, QCheckBox, QToolButton,
    QProgressBar, QFileDialog, QMessageBox, QGroupBox, QScrollArea,
    QSizePolicy, QTextBrowser
)

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
    from openpyxl.utils import get_column_letter
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

TITLE = "Locator Generator"
FRAMEWORKS = ["Java Selenium", "Python Selenium", "Playwright", "Cypress"]
EXAMPLE_HTML = '<div class="form-group">\n  <label for="username">Username</label>\n  <input id="username" name="userName" class="form-control" placeholder="Enter username" type="text">\n</div>'

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



def format_framework_code(candidate, framework):
    """Original four framework bindings, unchanged."""
    expr, kind = candidate['expr'], candidate['kind']
    if framework == 'Java Selenium':
        return f'By.xpath("{esc(expr)}")' if kind == 'XPath' else f'By.cssSelector("{esc(expr)}")'
    if framework == 'Python Selenium':
        by = 'By.XPATH' if kind == 'XPath' else 'By.CSS_SELECTOR'
        return f'driver.find_element({by}, "{esc(expr)}")'
    if framework == 'Playwright':
        return f'page.locator("xpath={esc(expr)}")' if kind == 'XPath' else f'page.locator("{esc(expr)}")'
    if framework == 'Cypress':
        return f'cy.xpath("{esc(expr)}")' if kind == 'XPath' else f'cy.get("{esc(expr)}")'
    return expr


def generate_descriptive_name(expr):
    """Preserved variable-naming heuristics from the Tkinter app."""
    clean_expr = expr.lower()
    base_name = 'element'
    if 'id=' in clean_expr or '#' in clean_expr:
        match = re.search(r'(?:id=["\']?|#)([a-zA-Z0-9\-_]+)', clean_expr)
        if match: base_name = match.group(1)
    elif 'name=' in clean_expr:
        match = re.search(r'name=["\']?([a-zA-Z0-9\-_]+)', clean_expr)
        if match: base_name = match.group(1)
    elif 'data-testid' in clean_expr or 'data-qa' in clean_expr:
        match = re.search(r'data-(?:testid|qa)=["\']?([a-zA-Z0-9\-_]+)', clean_expr)
        if match: base_name = match.group(1)
    elif 'placeholder=' in clean_expr:
        match = re.search(r'placeholder=["\']?([a-zA-Z0-9\-_ ]+)', clean_expr)
        if match: base_name = match.group(1)
    elif 'button' in clean_expr:
        base_name = 'button'
    elif 'input' in clean_expr:
        base_name = 'input'
    elif 'a[' in clean_expr or 'link' in clean_expr:
        base_name = 'link'
    words = re.findall(r'[a-zA-Z0-9]+', base_name)
    if not words: return 'uiElement'
    camel_case = words[0].lower() + ''.join(w.capitalize() for w in words[1:])
    if 'button' in clean_expr and not camel_case.lower().endswith('button'):
        camel_case += 'Button'
    elif 'input' in clean_expr and not camel_case.lower().endswith('input'):
        camel_case += 'Input'
    elif 'a[' in clean_expr and not camel_case.lower().endswith('link'):
        camel_case += 'Link'
    return camel_case


def generate_pom_code(elements, class_name, framework):
    """Keep original POM formats; syntax is the same as the Tkinter edition."""
    if framework == 'Java Selenium':
        lines = [
            'import org.openqa.selenium.By;',
            'import org.openqa.selenium.WebDriver;',
            'import org.openqa.selenium.WebElement;',
            'import org.openqa.selenium.support.PageFactory;',
            '', f'public class {class_name} {{',
            '    private WebDriver driver;', ''
        ]
        for el in elements:
            by_stmt = f'By.xpath("{esc(el["expr"])}")' if el['kind'] == 'XPath' else f'By.cssSelector("{esc(el["expr"])}")'
            lines.append(f'    private final By {el["varname"]} = {by_stmt};')
        lines += ['', f'    public {class_name}(WebDriver driver) {{',
                  '        this.driver = driver;', '    }', '}']
    elif framework == 'Python Selenium':
        lines = ['from selenium.webdriver.common.by import By', '',
                 f'class {class_name}:', '    def __init__(self, driver):',
                 '        self.driver = driver', '']
        for el in elements:
            kind = 'By.XPATH' if el['kind'] == 'XPath' else 'By.CSS_SELECTOR'
            lines.append(f'    {el["varname"]} = ({kind}, "{esc(el["expr"])}")')
        lines.append('')
    elif framework == 'Playwright':
        lines = [f'class {class_name}:', '    def __init__(self, page):',
                 '        self.page = page', '']
        for el in elements:
            value = 'xpath=' + el['expr'] if el['kind'] == 'XPath' else el['expr']
            lines.append(f'        self.{el["varname"]} = page.locator("{esc(value)}")')
        lines.append('')
    else:  # Cypress
        lines = [f'class {class_name} {{']
        for el in elements:
            command = f'cy.xpath("{esc(el["expr"])}")' if el['kind'] == 'XPath' else f'cy.get("{esc(el["expr"])}")'
            lines.append(f'    {el["varname"]} = () => {command};')
        lines.append('}')
    return '\n'.join(lines)


def selenium_driver(headless=True):
    if not HAS_SELENIUM:
        raise RuntimeError('Selenium is not installed. Run: pip install selenium')
    options = ChromeOptions()
    if headless:
        options.add_argument('--headless=new')
    return webdriver.Chrome(options=options)


def validate_http_url(value):
    """Return a normalized HTTP(S) URL or raise a readable input error."""
    value = (value or '').strip()
    if not value:
        raise ValueError('Enter a target URL.')
    if '://' not in value:
        value = 'https://' + value
    parsed = urlparse(value)
    if parsed.scheme not in ('http', 'https') or not parsed.netloc:
        raise ValueError('Enter a valid HTTP or HTTPS URL.')
    return value


def scan_page_job(url, ignored):
    if BeautifulSoup is None:
        raise RuntimeError('BeautifulSoup is not installed.')
    url = validate_http_url(url)
    driver = selenium_driver()
    try:
        driver.get(url)
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        elements = soup.find_all(['input', 'button', 'a', 'select', 'textarea', 'form'])
        others = soup.find_all(lambda tag: tag.has_attr('id') or tag.has_attr('data-testid') or tag.has_attr('data-qa'))
        all_elements = list(dict.fromkeys(elements + others))
        candidates = []
        for element in all_elements[:150]:
            matches = generate_candidates(soup, element, ignored_patterns=ignored)
            if matches:
                top = matches[0].copy()
                top['note'] = f'<{element.name}> | ' + top['note']
                candidates.append(top)
        # Counts here are DOM snapshot counts, NOT Selenium-verified results.
        return {'rows': candidates, 'url': url, 'elements': len(all_elements)}
    finally:
        driver.quit()


def test_live_job(url, candidate):
    url = validate_http_url(url)
    driver = selenium_driver()
    try:
        driver.get(url)
        by = SelBy.XPATH if candidate['kind'] == 'XPath' else SelBy.CSS_SELECTOR
        return len(driver.find_elements(by, candidate['expr']))
    finally:
        driver.quit()


def suggest_locator_fallbacks(soup, expr, limit=3):
    """Suggest ranked selectors for likely related DOM elements; never claim live validation."""
    if soup is None:
        return []
    hints = {'tag': None, 'id': None, 'classes': [], 'attrs': {}, 'text': None}
    xpath_tag = re.search(r'//([A-Za-z][\w:-]*)', expr)
    css_tag = re.match(r'^\s*([A-Za-z][\w:-]*)', expr)
    tag_match = xpath_tag or css_tag
    if tag_match:
        hints['tag'] = tag_match.group(1).lower()

    id_match = re.search(r'#([A-Za-z_][\w-]*)|@id\s*=\s*([\'\"])(.*?)\2', expr)
    if id_match:
        hints['id'] = id_match.group(1) or id_match.group(3)

    hints['classes'] = re.findall(r'\.([A-Za-z_][\w-]*)', expr)
    class_match = re.search(r'@class\s*=\s*([\'\"])(.*?)\1|contains\s*\(\s*@class\s*,\s*([\'\"])(.*?)\3', expr)
    if class_match:
        value = class_match.group(2) or class_match.group(4) or ''
        hints['classes'].extend(value.split())
    hints['classes'] = list(dict.fromkeys(hints['classes']))

    for name, quote, value in re.findall(r'\[([\w:-]+)\s*=\s*([\'\"])(.*?)\2\]', expr):
        hints['attrs'][name.lower()] = value
    for name, quote, value in re.findall(r'@([\w:-]+)\s*=\s*([\'\"])(.*?)\2', expr):
        if name.lower() != 'id' and name.lower() != 'class':
            hints['attrs'][name.lower()] = value
    text_match = re.search(r'(?:normalize-space\s*\(\s*\.\s*\)|text\s*\(\s*\)|\.)\s*=\s*([\'\"])(.*?)\1|contains\s*\([^,]+,\s*([\'\"])(.*?)\3', expr)
    if text_match:
        hints['text'] = (text_match.group(2) or text_match.group(4) or '').strip()

    if not any((hints['tag'], hints['id'], hints['classes'], hints['attrs'], hints['text'])):
        return []

    scored_elements = []
    for element in soup.find_all(True):
        score, reasons = 0, []
        if hints['tag']:
            if element.name.lower() != hints['tag']:
                continue
            score += 3
            reasons.append(f'tag <{hints["tag"]}>')
        if hints['id']:
            actual = str(element.get('id', ''))
            if actual == hints['id']:
                score += 12
                reasons.append('matching id')
            elif hints['id'].lower() in actual.lower() or actual.lower() in hints['id'].lower():
                score += 5
                reasons.append('similar id')
        actual_classes = set(cls_list(element))
        shared_classes = actual_classes.intersection(hints['classes'])
        if shared_classes:
            score += 4 * len(shared_classes)
            reasons.append('matching class ' + ', '.join(sorted(shared_classes)))
        for name, value in hints['attrs'].items():
            actual = element.get(name)
            if actual is not None:
                actual = ' '.join(actual) if isinstance(actual, list) else str(actual)
                if actual == value:
                    score += 8
                    reasons.append(f'matching {name}')
                elif value.lower() in actual.lower() or actual.lower() in value.lower():
                    score += 3
                    reasons.append(f'similar {name}')
        if hints['text']:
            actual_text = ' '.join(element.stripped_strings)
            if actual_text == hints['text']:
                score += 8
                reasons.append('matching text')
            elif hints['text'].lower() in actual_text.lower() or actual_text.lower() in hints['text'].lower():
                score += 3
                reasons.append('similar text')
        if score:
            scored_elements.append((score, element, reasons))

    scored_elements.sort(key=lambda item: (-item[0], item[1].name))
    suggestions, seen = [], {expr}
    for evidence_score, element, reasons in scored_elements[:30]:
        for candidate in generate_candidates(soup, element):
            selector = candidate['expr']
            if selector in seen:
                continue
            seen.add(selector)
            dom_count = candidate.get('count')
            if dom_count != 1:
                continue
            suggestions.append((evidence_score, candidate['score'], selector,
                                f'DOM snapshot: 1 match; evidence: {", ".join(reasons)}; locator rank: {candidate["rank"]}. Unverified live.'))
    suggestions.sort(key=lambda item: (-item[0], -item[1], item[2]))
    return [(selector, note) for _, _, selector, note in suggestions[:limit]]


def audit_job(url, locators):
    url = validate_http_url(url)
    driver = selenium_driver()
    try:
        driver.get(url)
        soup = BeautifulSoup(driver.page_source, 'html.parser') if BeautifulSoup else None
        results, suggestions = [], []
        for expr in locators:
            by = SelBy.XPATH if expr.startswith(('//', '(', './/', '/', 'id(')) else SelBy.CSS_SELECTOR
            try:
                matches = len(driver.find_elements(by, expr))
                status = 'Found' if matches else '0 matches'
                results.append((expr, str(matches), status))
            except Exception as error:
                # Never claim a broken locator when the underlying call could not evaluate it.
                results.append((expr, '—', 'Evaluation error'))
                suggestions.append((expr, 'Cannot evaluate selector', str(error).split('\n')[0][:160]))
                continue
            if matches:
                continue
            fallbacks = suggest_locator_fallbacks(soup, expr)
            if fallbacks:
                suggestions.extend((expr, selector, note) for selector, note in fallbacks)
            else:
                suggestions.append((expr, 'No likely fallback found',
                                    'The page snapshot did not contain a sufficiently similar, unique element.'))
        return {'results': results, 'suggestions': suggestions, 'url': url}
    finally:
        driver.quit()


# Browser automation runs exclusively in its own worker thread. GUI events only
# enqueue commands; they never call into Selenium from the Qt GUI thread.
PICKER_SCRIPT = r"""
if (!window.__locatorPickerInstalled) {
  window.__locatorPickerInstalled = true;
  window.__selectedElementHtml = null;
  const overlay = document.createElement('div');
  overlay.style.cssText = 'position:fixed;z-index:2147483646;border:2px solid #2563eb;background:rgba(37,99,235,.12);pointer-events:none;box-sizing:border-box';
  document.documentElement.appendChild(overlay);
  window.__locatorPickerOverlay = overlay;
  document.addEventListener('mouseover', function(e) {
    if (!window.__pickerEnabled) { overlay.style.display = 'none'; return; }
    overlay.style.display = 'block';
    const rect = e.target.getBoundingClientRect();
    overlay.style.top = rect.top+'px'; overlay.style.left = rect.left+'px';
    overlay.style.width = rect.width+'px'; overlay.style.height = rect.height+'px';
  }, true);
  document.addEventListener('click', function(e) {
    if (!window.__pickerEnabled) return;
    e.preventDefault(); e.stopPropagation();
    window.__selectedElementHtml = e.target.outerHTML;
  }, true);
}
window.__pickerEnabled = arguments[0];
"""


class Worker(QObject):
    result = Signal(object)
    error = Signal(str)
    finished = Signal()

    def __init__(self, func):
        super().__init__()
        self.func = func

    def run(self):
        try:
            self.result.emit(self.func())
        except Exception as exc:
            self.error.emit(str(exc))
        finally:
            self.finished.emit()


class PickerWorker(QObject):
    picked = Signal(str)
    navigated = Signal(str)
    error = Signal(str)
    finished = Signal()
    opened = Signal()

    def __init__(self, url, commands, enabled):
        super().__init__()
        self.url = url
        self.commands = commands
        self.enabled = enabled

    def run(self):
        driver = None
        try:
            driver = selenium_driver(headless=False)
            driver.get(self.url)
            self.opened.emit()
            last_url = None
            running = True
            while running:
                while True:
                    try:
                        command, value = self.commands.get_nowait()
                    except queue.Empty:
                        break
                    if command == 'close':
                        running = False
                        break
                    if command == 'enable': self.enabled = bool(value)
                    elif command == 'goto': driver.get(value)
                    elif command == 'back': driver.back()
                    elif command == 'forward': driver.forward()
                    elif command == 'refresh': driver.refresh()
                if not running:
                    break
                url = driver.current_url
                if url != last_url:
                    self.navigated.emit(url)
                    last_url = url
                driver.execute_script(PICKER_SCRIPT, self.enabled)
                selected = driver.execute_script('var x = window.__selectedElementHtml; window.__selectedElementHtml = null; return x;')
                if selected:
                    self.picked.emit(selected)
                time.sleep(0.35)
        except Exception as exc:
            self.error.emit(str(exc))
        finally:
            if driver:
                try: driver.quit()
                except Exception: pass
            self.finished.emit()


def label(text, secondary=False):
    obj = QLabel(text)
    if secondary:
        obj.setObjectName('muted')
    return obj


def panel(title, content=None):
    frame = QFrame()
    frame.setObjectName('panel')
    layout = QVBoxLayout(frame)
    layout.setContentsMargins(14, 12, 14, 12)
    layout.setSpacing(10)
    if title:
        name = QLabel(title)
        name.setObjectName('sectionTitle')
        layout.addWidget(name)
    if content:
        layout.addWidget(content, 1)
    return frame, layout


def table(headers, widths=None):
    view = QTableWidget(0, len(headers))
    view.setHorizontalHeaderLabels(headers)
    view.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
    view.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
    view.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
    view.setAlternatingRowColors(True)
    view.verticalHeader().setVisible(False)
    view.setWordWrap(False)
    view.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
    view.horizontalHeader().setStretchLastSection(True)
    view.horizontalHeader().setDefaultAlignment(Qt.AlignmentFlag.AlignLeft)
    for index, width in enumerate(widths or []):
        view.setColumnWidth(index, width)
    return view


def put_rows(view, rows):
    view.setUpdatesEnabled(False)
    try:
        view.setRowCount(0)
        for row in rows:
            at = view.rowCount()
            view.insertRow(at)
            for col, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                if col == 0:
                    item.setToolTip(str(value))
                view.setItem(at, col, item)
    finally:
        view.setUpdatesEnabled(True)


class ArrowComboBox(QComboBox):
    """Combo box with a clearly drawn down arrow and the standard popup list."""
    def paintEvent(self, event):
        super().paintEvent(event)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(self.palette().color(QPalette.ColorRole.ButtonText))
        center_x = self.width() - 15
        center_y = self.height() // 2
        arrow = QPolygon([
            QPoint(center_x - 4, center_y - 2),
            QPoint(center_x + 4, center_y - 2),
            QPoint(center_x, center_y + 3),
        ])
        painter.drawPolygon(arrow)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(TITLE + ' | PySide6')
        self.resize(1360, 850)
        self.setMinimumSize(1000, 720)
        self.rows = []
        self.pom_elements = []
        self.count_context = 'HTML fragment'
        self.dark = True
        self.busy = False
        self.worker_threads = []  # Strong references until Qt threads finish
        self.picker_commands = None
        self.picker_thread = None
        self.picker_open = False
        self._close_when_idle = False
        self.framework = FRAMEWORKS[0]
        self.nav_buttons = []
        self._setup_ui()
        self.apply_theme()
        self.switch_workspace(0)
        self.set_status('Ready. Paste HTML, pick an element, or scan a page.')

    def _setup_ui(self):
        base = QWidget()
        base.setObjectName('root')
        self.setCentralWidget(base)
        root = QVBoxLayout(base)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        header = QFrame()
        header.setObjectName('header')
        head = QHBoxLayout(header)
        head.setContentsMargins(22, 12, 22, 12)
        title_box = QVBoxLayout()
        title_box.setSpacing(2)
        title = QLabel('Locator Generator')
        title.setObjectName('appTitle')
        title_box.addWidget(title)
        title_box.addWidget(label('QA automation workspace  /  CSS · XPath · Page Objects · Audits', True))
        head.addLayout(title_box, 1)
        self.theme_button = QPushButton('Light theme')
        self.theme_button.setObjectName('secondaryButton')
        self.theme_button.clicked.connect(self.toggle_theme)
        head.addWidget(self.theme_button)
        root.addWidget(header)

        middle = QHBoxLayout()
        middle.setSpacing(0)
        root.addLayout(middle, 1)
        sidebar = QFrame()
        sidebar.setObjectName('sidebar')
        sidebar.setFixedWidth(194)
        side = QVBoxLayout(sidebar)
        side.setContentsMargins(12, 20, 12, 16)
        side.setSpacing(7)
        side.addWidget(label('WORKSPACES', True))
        for i, title in enumerate(['Generate Locators', 'Page Objects', 'Health Audit', 'Help']):
            b = QPushButton(title)
            b.setObjectName('navButton')
            b.setCheckable(True)
            b.setMinimumHeight(40)
            b.clicked.connect(lambda checked=False, index=i: self.switch_workspace(index))
            self.nav_buttons.append(b)
            side.addWidget(b)
        side.addStretch(1)
        side.addWidget(label('PySide6 desktop edition', True))
        middle.addWidget(sidebar)

        self.workspaces = QStackedWidget()
        self.workspaces.setObjectName('workspaces')
        middle.addWidget(self.workspaces, 1)
        self.build_generate_workspace()
        self.build_pom_workspace()
        self.build_audit_workspace()
        self.build_help_workspace()

        footer = QFrame()
        footer.setObjectName('footer')
        fl = QHBoxLayout(footer)
        fl.setContentsMargins(18, 7, 18, 7)
        self.status = label('', True)
        fl.addWidget(self.status, 1)
        self.progress = QProgressBar()
        self.progress.setFixedWidth(170)
        self.progress.setFixedHeight(9)
        self.progress.setRange(0, 0)
        self.progress.hide()
        fl.addWidget(self.progress)
        root.addWidget(footer)
        self.shortcuts = []
        for index in range(4):
            shortcut = QShortcut(QKeySequence(f'Ctrl+{index + 1}'), self)
            shortcut.activated.connect(lambda index=index: self.switch_workspace(index))
            self.shortcuts.append(shortcut)

    def workspace(self, title, subtitle):
        container = QWidget()
        outer = QVBoxLayout(container)
        outer.setContentsMargins(18, 16, 18, 16)
        outer.setSpacing(12)
        h = QLabel(title)
        h.setObjectName('pageTitle')
        outer.addWidget(h)
        outer.addWidget(label(subtitle, True))
        self.workspaces.addWidget(container)
        return outer

    def build_generate_workspace(self):
        outer = self.workspace('Generate Locators', 'Choose an input mode, generate ranked selectors, then export or reuse them.')
        primary_split = QSplitter(Qt.Orientation.Horizontal)
        primary_split.setChildrenCollapsible(False)
        outer.addWidget(primary_split, 1)

        input_panel, il = panel('INPUT')
        input_panel.setMinimumWidth(300)
        primary_split.addWidget(input_panel)
        self.mode_buttons = []
        modebar = QHBoxLayout()
        modebar.setSpacing(4)
        for i, text in enumerate(['Paste HTML', 'Pick Element', 'Scan Page']):
            b = QPushButton(text)
            b.setObjectName('modeButton')
            b.setCheckable(True)
            b.clicked.connect(lambda checked=False, index=i: self.switch_mode(index))
            self.mode_buttons.append(b)
            modebar.addWidget(b)
        il.addLayout(modebar)
        self.modes = QStackedWidget()
        il.addWidget(self.modes, 1)

        paste = QWidget()
        pv = QVBoxLayout(paste)
        pv.setContentsMargins(0, 5, 0, 0)
        pv.setSpacing(9)
        pv.addWidget(label('Paste a target HTML element or a small HTML fragment.', True))
        self.html = QPlainTextEdit()
        self.html.setObjectName('codeArea')
        self.html.setPlaceholderText('<button data-testid="login">Log in</button>')
        self.html.setPlainText(EXAMPLE_HTML)
        pv.addWidget(self.html, 1)
        html_actions = QHBoxLayout()
        example = QPushButton('Load example')
        example.clicked.connect(lambda: self.html.setPlainText(EXAMPLE_HTML))
        html_actions.addWidget(example)
        clear = QPushButton('Clear input')
        clear.clicked.connect(self.clear_html)
        html_actions.addWidget(clear)
        html_actions.addStretch(1)
        pv.addLayout(html_actions)
        self.modes.addWidget(paste)

        pick = QWidget()
        pl = QVBoxLayout(pick)
        pl.setContentsMargins(0, 5, 0, 0)
        pl.setSpacing(12)
        pl.addWidget(label('Open Chrome, hover an element, and click to capture its HTML.', True))
        pl.addWidget(label('Target URL'))
        self.url_entry = QLineEdit('https://example.com')
        pl.addWidget(self.url_entry)
        self.url_entry.editingFinished.connect(lambda: self.live_test_url.setText(self.url_entry.text().strip()))
        pick_actions = QHBoxLayout()
        self.pick_button = QPushButton('Open picker')
        self.pick_button.setObjectName('primaryButton')
        self.pick_button.clicked.connect(self.pick_element_interactive)
        pick_actions.addWidget(self.pick_button)
        self.close_picker_button = QPushButton('Close browser')
        self.close_picker_button.clicked.connect(self.close_picker_browser)
        self.close_picker_button.setEnabled(False)
        pick_actions.addWidget(self.close_picker_button)
        pl.addLayout(pick_actions)
        self.pick_mode = QCheckBox('Picking enabled')
        self.pick_mode.setChecked(True)
        self.pick_mode.toggled.connect(self.update_pick_mode)
        self.pick_mode.setEnabled(False)
        pl.addWidget(self.pick_mode)
        nav = QHBoxLayout()
        self.nav_back = QPushButton('Back')
        self.nav_forward = QPushButton('Forward')
        self.nav_refresh = QPushButton('Refresh')
        for btn, name in [(self.nav_back, 'back'), (self.nav_forward, 'forward'), (self.nav_refresh, 'refresh')]:
            btn.setEnabled(False)
            btn.clicked.connect(lambda checked=False, command=name: self.picker_command(command))
            nav.addWidget(btn)
        pl.addLayout(nav)
        pl.addWidget(label('Captured HTML is moved to Paste HTML and locators are generated automatically.', True))
        pl.addStretch(1)
        self.modes.addWidget(pick)

        scan = QWidget()
        sl = QVBoxLayout(scan)
        sl.setContentsMargins(0, 5, 0, 0)
        sl.setSpacing(12)
        sl.addWidget(label('Scan buttons, inputs, links, forms and tagged elements from a live page.', True))
        sl.addWidget(label('Target URL'))
        self.scan_url = QLineEdit('https://example.com')
        sl.addWidget(self.scan_url)
        self.scan_url.editingFinished.connect(lambda: self.live_test_url.setText(self.scan_url.text().strip()))
        self.scan_button = QPushButton('Scan whole page')
        self.scan_button.setObjectName('primaryButton')
        self.scan_button.clicked.connect(self.scan_whole_page)
        sl.addWidget(self.scan_button)
        sl.addWidget(label('Counts come from the captured page-source snapshot, not live Selenium checks.', True))
        sl.addStretch(1)
        self.modes.addWidget(scan)

        self.advanced = QToolButton()
        self.advanced.setText('Advanced options  ▸')
        self.advanced.setCheckable(True)
        self.advanced.setObjectName('linkButton')
        self.advanced.toggled.connect(self.toggle_advanced)
        il.addWidget(self.advanced)
        self.advanced_panel = QFrame()
        adv = QVBoxLayout(self.advanced_panel)
        adv.setContentsMargins(0, 0, 0, 0)
        adv.addWidget(label('Ignore patterns (comma-separated regular expressions)'))
        self.ignore_entry = QLineEdit('css-, sc-, ng-star-inserted')
        adv.addWidget(self.ignore_entry)
        self.advanced_panel.hide()
        il.addWidget(self.advanced_panel)
        self.generate_button = QPushButton('Generate Locators')
        self.generate_button.setObjectName('primaryButton')
        self.generate_button.setMinimumHeight(39)
        self.generate_button.clicked.connect(self.generate)
        il.addWidget(self.generate_button)

        right = QSplitter(Qt.Orientation.Vertical)
        right.setChildrenCollapsible(False)
        primary_split.addWidget(right)
        primary_split.setSizes([395, 790])

        results_panel, rl = panel('CANDIDATE LOCATORS')
        result_desc = label('Fragment matches are calculated from supplied HTML, or from the scanned DOM snapshot.', True)
        result_desc.setWordWrap(True)
        rl.addWidget(result_desc)
        self.results_caption = label('No locators generated yet.', True)
        rl.addWidget(self.results_caption)
        self.tree = table(['Select', 'Selector', 'Type', 'Strategy', 'Score', 'Source matches'], [55, 385, 72, 172, 98, 120])
        self.tree.itemSelectionChanged.connect(self.preview)
        self.tree.itemChanged.connect(self.on_candidate_check_changed)
        self.tree.itemDoubleClicked.connect(lambda *_: self.copy_expr())
        rl.addWidget(self.tree, 1)
        self.select_all_candidates = QCheckBox('Select all generated selectors')
        self.select_all_candidates.toggled.connect(self.set_all_candidates_checked)
        rl.addWidget(self.select_all_candidates)
        live_row = QHBoxLayout()
        live_row.addWidget(label('Live test URL'))
        self.live_test_url = QLineEdit('https://example.com')
        self.live_test_url.setPlaceholderText('https://your-test-environment.example')
        live_row.addWidget(self.live_test_url, 1)
        rl.addLayout(live_row)
        actions = QHBoxLayout()
        for text, callback in [
            ('Copy selectors', self.copy_expr),
            ('Add to Page Objects', self.add_selected_to_pom),
            ('Send to Health Audit', self.send_selected_to_auditor),
            ('Test live', self.test_live_url),
            ('Export', self.export_excel),
        ]:
            b = QPushButton(text)
            b.clicked.connect(callback)
            actions.addWidget(b)
            if text == 'Test live':
                self.test_button = b
                self.test_button.setEnabled(HAS_SELENIUM)
        rl.addLayout(actions)
        right.addWidget(results_panel)

        lower_split = QSplitter(Qt.Orientation.Horizontal)
        lower_split.setChildrenCollapsible(False)
        notes_panel, nl = panel('SELECTED LOCATOR')
        self.selected_note = QPlainTextEdit()
        self.selected_note.setReadOnly(True)
        self.selected_note.setPlaceholderText('Select a locator to see its stability notes and match scope.')
        nl.addWidget(self.selected_note, 1)
        lower_split.addWidget(notes_panel)
        code_panel, cl = panel('CODE PREVIEW')
        options = QHBoxLayout()
        options.addWidget(label('Framework'))
        self.framework_combo = ArrowComboBox()
        self.framework_combo.addItems(FRAMEWORKS)
        self.framework_combo.currentTextChanged.connect(self.change_framework)
        options.addWidget(self.framework_combo, 1)
        self.copy_code_button = QPushButton('Copy code')
        self.copy_code_button.clicked.connect(self.copy_framework_code)
        options.addWidget(self.copy_code_button)
        cl.addLayout(options)
        self.code_preview = QPlainTextEdit()
        self.code_preview.setObjectName('codeArea')
        self.code_preview.setReadOnly(True)
        self.code_preview.setPlaceholderText('Select one or more rows to preview framework bindings.')
        cl.addWidget(self.code_preview, 1)
        lower_split.addWidget(code_panel)
        lower_split.setSizes([300, 530])
        right.addWidget(lower_split)
        right.setSizes([390, 210])
        self.switch_mode(0)

    def build_pom_workspace(self):
        outer = self.workspace('Page Objects', 'Collect selected locators and generate class code using your preferred framework.')
        bar, bl = panel('PAGE OBJECT SETTINGS')
        controls = QHBoxLayout()
        controls.addWidget(label('Class name'))
        self.pom_class_entry = QLineEdit('LoginPage')
        self.pom_class_entry.setMaximumWidth(235)
        self.pom_class_entry.textChanged.connect(self.update_pom_code_preview)
        controls.addWidget(self.pom_class_entry)
        controls.addWidget(label('Framework'))
        self.pom_framework = ArrowComboBox()
        self.pom_framework.addItems(FRAMEWORKS)
        self.pom_framework.currentTextChanged.connect(self.change_framework)
        controls.addWidget(self.pom_framework)
        controls.addStretch(1)
        clear = QPushButton('Clear elements')
        clear.clicked.connect(self.clear_pom)
        controls.addWidget(clear)
        bl.addLayout(controls)
        outer.addWidget(bar)

        split = QSplitter(Qt.Orientation.Horizontal)
        split.setChildrenCollapsible(False)
        outer.addWidget(split, 1)
        elems, el = panel('SELECTED ELEMENTS')
        self.pom_count = label('No elements added. Select rows from Generate Locators.', True)
        el.addWidget(self.pom_count)
        self.pom_tree = table(['Variable', 'Type', 'Selector'], [145, 72, 300])
        self.pom_tree.itemSelectionChanged.connect(self.load_selected_pom_variable)
        el.addWidget(self.pom_tree, 1)
        variable_row = QHBoxLayout()
        variable_row.addWidget(label('Variable name'))
        self.pom_variable_entry = QLineEdit()
        self.pom_variable_entry.setPlaceholderText('Select an element, then enter a variable name')
        self.pom_variable_entry.returnPressed.connect(self.rename_selected_pom_variable)
        variable_row.addWidget(self.pom_variable_entry, 1)
        self.rename_pom_button = QPushButton('Set name')
        self.rename_pom_button.clicked.connect(self.rename_selected_pom_variable)
        variable_row.addWidget(self.rename_pom_button)
        el.addLayout(variable_row)
        split.addWidget(elems)
        preview_panel, pl = panel('GENERATED CLASS CODE')
        self.pom_code_text = QPlainTextEdit()
        self.pom_code_text.setObjectName('codeArea')
        self.pom_code_text.setReadOnly(True)
        pl.addWidget(self.pom_code_text, 1)
        self.copy_pom_button = QPushButton('Copy POM class code')
        self.copy_pom_button.setObjectName('primaryButton')
        self.copy_pom_button.clicked.connect(self.copy_pom_class)
        pl.addWidget(self.copy_pom_button)
        split.addWidget(preview_panel)
        split.setSizes([430, 710])
        self.update_pom_code_preview()

    def build_audit_workspace(self):
        outer = self.workspace('Health Audit', 'Check selectors against a live page; review tentative fallback suggestions separately.')
        ctrl, cl = panel('AUDIT TARGET')
        bar = QHBoxLayout()
        bar.addWidget(label('Target URL'))
        self.audit_url_entry = QLineEdit('https://example.com')
        bar.addWidget(self.audit_url_entry, 1)
        self.audit_button = QPushButton('Run Health Audit')
        self.audit_button.setObjectName('primaryButton')
        self.audit_button.clicked.connect(self.run_health_audit_and_heal)
        self.audit_button.setEnabled(HAS_SELENIUM)
        bar.addWidget(self.audit_button)
        cl.addLayout(bar)
        outer.addWidget(ctrl)
        main_split = QSplitter(Qt.Orientation.Horizontal)
        main_split.setChildrenCollapsible(False)
        outer.addWidget(main_split, 1)
        inputs, il = panel('LOCATORS TO AUDIT')
        il.addWidget(label('Enter one CSS or XPath locator per line.', True))
        self.audit_input_text = QPlainTextEdit()
        self.audit_input_text.setObjectName('codeArea')
        self.audit_input_text.setPlaceholderText('//h1\n#submit-button\ninput[name="email"]')
        il.addWidget(self.audit_input_text, 1)
        main_split.addWidget(inputs)
        rhs = QSplitter(Qt.Orientation.Vertical)
        rhs.setChildrenCollapsible(False)
        result, rl = panel('LIVE VALIDATION RESULTS')
        rl.addWidget(label('Only the original locators are checked with Selenium.', True))
        self.audit_tree = table(['Locator', 'Matches', 'Status'], [390, 90, 180])
        rl.addWidget(self.audit_tree, 1)
        rhs.addWidget(result)
        fallbacks, fl = panel('UNVERIFIED FALLBACK SUGGESTIONS')
        desc = label('Ranked alternatives based on matching tag, attributes, classes, or text in the captured page snapshot. Suggestions are not live-validated.', True)
        desc.setWordWrap(True)
        fl.addWidget(desc)
        self.suggestion_tree = table(['Original', 'Suggested alternative', 'Notes'], [285, 300, 270])
        fl.addWidget(self.suggestion_tree, 1)
        rhs.addWidget(fallbacks)
        rhs.setSizes([280, 240])
        main_split.addWidget(rhs)
        main_split.setSizes([350, 800])

    def build_help_workspace(self):
        outer = self.workspace('Help & Guide', 'A practical guide to generating, selecting, validating, and exporting locators.')
        contents, body = panel('LOCATOR GENERATOR  ·  USER GUIDE')
        guide = QTextBrowser()
        guide.setReadOnly(True)
        guide.setOpenExternalLinks(False)
        guide.setObjectName('guideBrowser')
        self.help_browser = guide
        guide.setHtml('''
<html><body style="font-family:Segoe UI,Arial,sans-serif; font-size:12pt;">
<h2 style="color:#3b82f6; margin-bottom:4px;">Start here</h2>
<p>Paste a small HTML fragment, use the browser picker, or scan a live page. Generate Locators creates ranked CSS and XPath options for the target element.</p>
<p><b>Install:</b> <code>pip install PySide6 beautifulsoup4 selenium openpyxl lxml</code><br>
<b>Run:</b> <code>python locator_generator_pyside6.py</code></p>
<hr>
<h3 style="color:#3b82f6;">1 · Generate Locators</h3>
<ul>
 <li><b>Paste HTML:</b> Enter an element or fragment. The app chooses the first element with a useful identifying attribute; if none has one, it uses the final element in the fragment.</li>
 <li><b>Pick Element:</b> Open Chrome, hover over an element, and click it to capture its HTML. Use the picker controls to navigate, refresh, or pause picking.</li>
 <li><b>Scan Page:</b> Enter a URL to inspect buttons, inputs, links, forms, and elements with test IDs. Scan results come from a page-source snapshot.</li>
 <li><b>Select locators:</b> Tick the checkbox beside each selector, or use <b>Select all generated selectors</b>. Copy, add to Page Objects, send to Health Audit, and Test live use the checked selectors.</li>
 <li><b>Inspect and copy:</b> Select a checked locator to read its strategy, score, stability rank, and match count. Choose a framework to preview or copy its code binding.</li>
 <li><b>Advanced options:</b> Ignore selectors whose expression matches one of the comma-separated regular expressions.</li>
</ul>
<p><i>Fragment and scan counts are calculated from captured HTML. They do not prove that a selector works on the live page.</i></p>
<hr>
<h3 style="color:#3b82f6;">2 · Page Objects</h3>
<ul>
 <li>Check locators on the Generate Locators page, then choose <b>Add to Page Objects</b>. Duplicate expressions are skipped and new variable names are made unique automatically.</li>
 <li>On the Page Objects page, select one element, enter a variable name, and click <b>Set name</b> (or press Enter). Names must start with a letter or underscore and contain only letters, digits, or underscores; names already in use are rejected.</li>
 <li>Edit the class name to refresh the code. The framework selector is a bordered dropdown with an arrow; changing it updates both code previews.</li>
 <li>Use <b>Clear elements</b> to empty the Page Object list, or <b>Copy POM class code</b> to copy the generated class.</li>
</ul>
<hr>
<h3 style="color:#3b82f6;">3 · Health Audit</h3>
<ul>
 <li>Enter a page URL and one CSS or XPath locator per line, then run the audit. Selenium checks each original locator against the live page.</li>
 <li><b>Found</b>, <b>0 matches</b>, and <b>Evaluation error</b> are different results. An evaluation error means Selenium could not evaluate the selector.</li>
 <li>For zero-match locators, the app may show up to three ranked alternatives based on matching tags, IDs, classes, attributes, or text in the captured page snapshot. Notes explain the match evidence and snapshot count.</li>
 <li>Fallbacks are suggestions only. They are not tested against the live page and do not replace or repair the original selector.</li>
</ul>
<hr>
<h3 style="color:#3b82f6;">4 · Export</h3>
<ul>
 <li><b>Excel (.xlsx):</b> Exports the generated locator list with a clean white cell background, readable text, wrapped long expressions, filters, and a frozen header row.</li>
 <li><b>CSV (.csv):</b> Exports the same columns as plain data. CSV does not store colors, filters, or column formatting.</li>
 <li>Exports include locator type and strategy, expression, stability rank and score, match count, framework binding, notes, and match source.</li>
</ul>
<hr>
<h3 style="color:#3b82f6;">Frameworks and display</h3>
<p>The framework dropdown supports <b>Java Selenium</b>, <b>Python Selenium</b>, <b>Playwright</b>, and <b>Cypress</b>. Use the header button to switch between dark and light app themes. The Help page follows the active app theme.</p>
<h3 style="color:#3b82f6;">Keyboard and dependencies</h3>
<ul>
 <li><b>Ctrl+1 … Ctrl+4:</b> Switch between Generate Locators, Page Objects, Health Audit, and Help.</li>
 <li><b>Tab / Shift+Tab:</b> Move between controls.</li>
 <li><b>PySide6:</b> Required for the desktop app. <b>BeautifulSoup:</b> Required for HTML parsing. <b>Selenium and Chrome:</b> Required for the picker and live browser features. <b>openpyxl:</b> Required for Excel export; CSV works without it. <b>lxml:</b> Improves XPath match counts.</li>
</ul>
<p><i>Scores and fallback rankings are heuristics, not guarantees. Some framework snippets may need project-specific imports or syntax adjustments.</i></p>
</body></html>
''')
        body.addWidget(guide, 1)
        copy = QPushButton('Copy help text')
        copy.clicked.connect(lambda: self.copy(guide.toPlainText(), 'Help copied.'))
        body.addWidget(copy)
        outer.addWidget(contents, 1)

    def switch_workspace(self, index):
        self.workspaces.setCurrentIndex(index)
        for number, button in enumerate(self.nav_buttons):
            button.setChecked(number == index)

    def switch_mode(self, index):
        self.modes.setCurrentIndex(index)
        for number, button in enumerate(self.mode_buttons):
            button.setChecked(number == index)
        self.generate_button.setVisible(index != 2)
        self.scan_button.setEnabled(not self.busy and HAS_SELENIUM)
        self.pick_button.setEnabled(not self.busy and HAS_SELENIUM and not self.picker_open)
        if not HAS_SELENIUM and index in (1, 2):
            self.set_status('Selenium is required for browser picking and live scans.')

    def toggle_advanced(self, checked):
        self.advanced_panel.setVisible(checked)
        self.advanced.setText('Advanced options  ▾' if checked else 'Advanced options  ▸')

    def set_status(self, message):
        self.status.setText(message)

    def popup(self, title, message):
        QMessageBox.warning(self, title, message)

    def get_ignored_patterns(self):
        patterns = [p.strip() for p in self.ignore_entry.text().split(',') if p.strip()]
        for pattern in patterns:
            re.compile(pattern)
        return patterns

    def generate(self):
        if BeautifulSoup is None:
            self.popup('Dependency missing', 'Install BeautifulSoup: pip install beautifulsoup4')
            return
        markup = self.html.toPlainText().strip()
        if not markup:
            self.set_status('Paste HTML before generating locators.')
            self.html.setFocus()
            return
        try:
            ignored = self.get_ignored_patterns()
            soup = BeautifulSoup(markup, 'html.parser')
            elements = soup.find_all(True)
            if not elements:
                raise ValueError('No HTML elements found. Paste a complete element.')
            preferred = [e for e in elements if any(e.get(k) for k in (
                'data-testid', 'data-test', 'data-qa', 'id', 'name', 'aria-label', 'placeholder'))]
            target = preferred[0] if preferred else elements[-1]
            self.rows = generate_candidates(soup, target, ignored_patterns=ignored)
            self.count_context = 'HTML fragment'
            self.show_candidates(f'Target <{target.name}> · {len(self.rows)} candidates')
        except Exception as exc:
            self.popup('Generation error', str(exc))

    def show_candidates(self, summary):
        self.put_candidate_rows()
        self.results_caption.setText(summary + f'   |   Match source: {self.count_context}')
        self.preview()
        self.set_status(summary)

    def clear_html(self):
        self.html.clear()
        self.rows = []
        put_rows(self.tree, [])
        self.select_all_candidates.setChecked(False)
        self.results_caption.setText('No locators generated yet.')
        self.preview()
        self.set_status('HTML and locator results cleared.')

    def selected_indices(self):
        return sorted({index.row() for index in self.tree.selectedIndexes()})

    def selected_candidates(self):
        return [self.rows[i] for i in range(min(len(self.rows), self.tree.rowCount()))
                if self.tree.item(i, 0) is not None
                and self.tree.item(i, 0).checkState() == Qt.CheckState.Checked]

    def put_candidate_rows(self):
        self.tree.blockSignals(True)
        try:
            self.tree.setRowCount(0)
            for candidate in self.rows:
                row = self.tree.rowCount()
                self.tree.insertRow(row)
                check = QTableWidgetItem('')
                check.setFlags(check.flags() | Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled)
                check.setCheckState(Qt.CheckState.Unchecked)
                self.tree.setItem(row, 0, check)
                values = (candidate['expr'], candidate['kind'], candidate['strategy'],
                          f'{candidate["rank"]} · {candidate["score"]}/100',
                          str(candidate['count']) if candidate['count'] is not None else 'N/A')
                for column, value in enumerate(values, 1):
                    self.tree.setItem(row, column, QTableWidgetItem(str(value)))
        finally:
            self.tree.blockSignals(False)
        self.select_all_candidates.blockSignals(True)
        self.select_all_candidates.setChecked(False)
        self.select_all_candidates.blockSignals(False)

    def set_all_candidates_checked(self, checked):
        self.tree.blockSignals(True)
        try:
            for row in range(self.tree.rowCount()):
                item = self.tree.item(row, 0)
                if item is not None:
                    item.setCheckState(Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)
        finally:
            self.tree.blockSignals(False)
        self.preview()

    def on_candidate_check_changed(self, item):
        if item.column() != 0:
            return
        all_checked = bool(self.rows) and all(
            self.tree.item(row, 0) is not None
            and self.tree.item(row, 0).checkState() == Qt.CheckState.Checked
            for row in range(min(len(self.rows), self.tree.rowCount()))
        )
        self.select_all_candidates.blockSignals(True)
        self.select_all_candidates.setChecked(all_checked)
        self.select_all_candidates.blockSignals(False)
        self.preview()

    def preview(self):
        candidates = self.selected_candidates()
        if not candidates:
            self.selected_note.setPlainText('Select a locator to inspect its strategy and notes.')
            self.code_preview.setPlainText('Select one or more rows to preview framework code.')
            return
        c = candidates[0]
        count = c['count'] if c['count'] is not None else 'N/A'
        self.selected_note.setPlainText(
            f'Locator: {c["expr"]}\n'
            f'Type: {c["kind"]}   Strategy: {c["strategy"]}\n'
            f'Score: {c["score"]}/100 ({c["rank"]})\n'
            f'{self.count_context} matches: {count}\n\n'
            f'{c["note"]}\n\n'
            'These counts do not establish live Selenium verification.'
        )
        self.code_preview.setPlainText('\n'.join(format_framework_code(item, self.framework) for item in candidates))

    def change_framework(self, value):
        if value not in FRAMEWORKS:
            return
        self.framework = value
        for combo in (self.framework_combo, self.pom_framework):
            if combo.currentText() != value:
                was_blocked = combo.blockSignals(True)
                combo.setCurrentText(value)
                combo.blockSignals(was_blocked)
        self.preview()
        self.update_pom_code_preview()

    def copy(self, content, status):
        QApplication.clipboard().setText(content)
        self.set_status(status)

    def copy_expr(self):
        rows = self.selected_candidates()
        if not rows:
            self.set_status('Select one or more locator rows to copy.')
            return
        self.copy('\n'.join(c['expr'] for c in rows), f'Copied {len(rows)} selector(s).')

    def copy_framework_code(self):
        rows = self.selected_candidates()
        if not rows:
            self.set_status('Select one or more locator rows to copy code.')
            return
        self.copy('\n'.join(format_framework_code(c, self.framework) for c in rows),
                  f'Copied {len(rows)} {self.framework} binding(s).')

    def add_selected_to_pom(self):
        candidates = self.selected_candidates()
        if not candidates:
            self.set_status('Select one or more rows to add to Page Objects.')
            return
        added = duplicates = 0
        for c in candidates:
            expr = c['expr']
            if any(e['expr'] == expr for e in self.pom_elements):
                duplicates += 1
                continue
            original = generate_descriptive_name(expr)
            variable = original
            counter = 1
            while variable in {e['varname'] for e in self.pom_elements}:
                variable = f'{original}{counter}'
                counter += 1
            self.pom_elements.append({'varname': variable, 'kind': c['kind'], 'expr': expr})
            added += 1
        put_rows(self.pom_tree, [(e['varname'], e['kind'], e['expr']) for e in self.pom_elements])
        self.pom_count.setText(f'{len(self.pom_elements)} elements in Page Objects')
        self.update_pom_code_preview()
        self.set_status(f'Added {added} element(s) to Page Objects; skipped {duplicates} duplicate(s).')

    def load_selected_pom_variable(self):
        rows = sorted({index.row() for index in self.pom_tree.selectedIndexes()})
        if len(rows) == 1 and rows[0] < len(self.pom_elements):
            self.pom_variable_entry.setText(self.pom_elements[rows[0]]['varname'])
            self.rename_pom_button.setEnabled(True)
        else:
            self.rename_pom_button.setEnabled(False)
            if not rows:
                self.pom_variable_entry.clear()

    def rename_selected_pom_variable(self):
        rows = sorted({index.row() for index in self.pom_tree.selectedIndexes()})
        if len(rows) != 1 or rows[0] >= len(self.pom_elements):
            self.set_status('Select one Page Object element to name.')
            return
        new_name = self.pom_variable_entry.text().strip()
        if not re.fullmatch(r'[A-Za-z_]\w*', new_name):
            self.set_status('Use a valid identifier: start with a letter or underscore, then use letters, digits, or underscores.')
            return
        index = rows[0]
        if any(i != index and item['varname'] == new_name for i, item in enumerate(self.pom_elements)):
            self.set_status('That variable name is already used in this Page Object.')
            return
        self.pom_elements[index]['varname'] = new_name
        put_rows(self.pom_tree, [(e['varname'], e['kind'], e['expr']) for e in self.pom_elements])
        self.pom_tree.selectRow(index)
        self.update_pom_code_preview()
        self.set_status(f'Variable renamed to {new_name}.')

    def update_pom_code_preview(self):
        if not hasattr(self, 'pom_code_text'):
            return
        class_name = self.pom_class_entry.text().strip() or 'LoginPage'
        self.pom_code_text.setPlainText(generate_pom_code(self.pom_elements, class_name, self.framework))

    def clear_pom(self):
        self.pom_elements.clear()
        put_rows(self.pom_tree, [])
        self.pom_count.setText('No elements added. Select rows from Generate Locators.')
        self.update_pom_code_preview()
        self.set_status('Page Object element list cleared.')

    def copy_pom_class(self):
        if not self.pom_elements:
            self.set_status('Add elements to Page Objects first.')
            return
        self.copy(self.pom_code_text.toPlainText(), 'Page Object class copied.')

    def send_selected_to_auditor(self):
        candidates = self.selected_candidates()
        if not candidates:
            self.set_status('Select one or more locator rows to audit.')
            return
        existing = self.audit_input_text.toPlainText().splitlines()
        all_lines = {line.strip() for line in existing if line.strip()}
        added = 0
        for candidate in candidates:
            expr = candidate['expr']
            if expr not in all_lines:
                existing.append(expr)
                all_lines.add(expr)
                added += 1
        self.audit_input_text.setPlainText('\n'.join(existing).strip())
        self.switch_workspace(2)
        self.set_status(f'Sent {added} selector(s) to Health Audit.')

    def export_excel(self):
        if not self.rows:
            self.set_status('Generate locators before exporting.')
            return
        dest, selected_filter = QFileDialog.getSaveFileName(
            self, 'Export generated locators', 'Locators.xlsx',
            'Excel Workbook (*.xlsx);;CSV File (*.csv)')
        if not dest:
            return
        extension = Path(dest).suffix.lower()
        if extension not in ('.xlsx', '.csv'):
            dest += '.csv' if 'CSV' in selected_filter else '.xlsx'
            extension = Path(dest).suffix.lower()
        if extension == '.xlsx' and not HAS_OPENPYXL:
            self.popup('Dependency missing', 'Excel export requires openpyxl. Install: pip install openpyxl')
            return
        headers = ['#', 'Type', 'Strategy', 'Locator Expression', 'Stability Rank',
                   'Score Value', 'Matches', f'Code Binding ({self.framework})', 'Notes', 'Match Source']
        data = [[i, c['kind'], c['strategy'], c['expr'], c['rank'], c['score'],
                 c['count'] if c['count'] is not None else 'N/A',
                 format_framework_code(c, self.framework), c['note'], self.count_context]
                for i, c in enumerate(self.rows, 1)]
        try:
            if extension == '.xlsx':
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = 'Locators'
                ws.sheet_view.showGridLines = False
                ws.append(headers)
                for row in data: ws.append(row)
                # Explicit cell colors keep the export readable regardless of Excel's
                # dark-mode or system theme settings.
                header_fill = PatternFill('solid', fgColor='FFFFFFFF')
                odd_fill = PatternFill('solid', fgColor='FFFFFFFF')
                even_fill = PatternFill('solid', fgColor='FFFFFFFF')
                thin_gray = Side(style='thin', color='FFD9E2F3')
                border = Border(bottom=thin_gray)
                for cell in ws[1]:
                    cell.fill = header_fill
                    cell.font = Font(name='Calibri', size=11, bold=True, color='FF1F2937')
                    cell.alignment = Alignment(vertical='center', horizontal='left', wrap_text=True)
                ws.row_dimensions[1].height = 30

                widths = [7, 12, 27, 52, 17, 13, 12, 62, 48, 24]
                for column, width in enumerate(widths, 1):
                    ws.column_dimensions[get_column_letter(column)].width = width
                wrap_columns = {3, 4, 8, 9, 10}
                for row_index in range(2, ws.max_row + 1):
                    fill = odd_fill if row_index % 2 == 0 else even_fill
                    max_lines = 1
                    for column_index in range(1, ws.max_column + 1):
                        cell = ws.cell(row_index, column_index)
                        cell.fill = fill
                        cell.font = Font(name='Calibri', size=10, color='FF1F2937')
                        cell.border = border
                        cell.alignment = Alignment(
                            vertical='top',
                            horizontal='center' if column_index in (1, 6, 7) else 'left',
                            wrap_text=column_index in wrap_columns,
                        )
                        if column_index in wrap_columns:
                            max_lines = max(max_lines, min(8, (len(str(cell.value or '')) // widths[column_index - 1]) + 1))
                    ws.row_dimensions[row_index].height = min(96, max(24, 15 * max_lines))

                ws.freeze_panes = 'A2'
                ws.auto_filter.ref = ws.dimensions
                ws.sheet_properties.pageSetUpPr.fitToPage = True
                ws.page_setup.orientation = 'landscape'
                ws.page_setup.fitToWidth = 1
                ws.page_setup.fitToHeight = 0
                wb.save(dest)
            else:
                with open(dest, 'w', newline='', encoding='utf-8-sig') as handle:
                    writer = csv.writer(handle)
                    writer.writerow(headers)
                    writer.writerows(data)
            self.set_status(f'Exported {len(self.rows)} locators to {dest}.')
        except Exception as exc:
            self.popup('Export failed', str(exc))

    def set_busy(self, active, message=None):
        self.busy = active
        self.progress.setVisible(active)
        for btn in (self.generate_button, self.scan_button, self.audit_button, self.test_button):
            btn.setEnabled(not active)
        self.scan_button.setEnabled(not active and HAS_SELENIUM)
        self.test_button.setEnabled(not active and HAS_SELENIUM)
        self.audit_button.setEnabled(not active and HAS_SELENIUM)
        self.pick_button.setEnabled(not active and HAS_SELENIUM and not self.picker_open)
        if message:
            self.set_status(message)

    def launch_job(self, task, success, label_text):
        if self.busy:
            return
        self.set_busy(True, label_text)
        thread = QThread(self)
        worker = Worker(task)
        worker.moveToThread(thread)
        record = (thread, worker)
        self.worker_threads.append(record)
        thread.started.connect(worker.run)
        worker.result.connect(success)
        worker.error.connect(lambda msg: self.popup('Operation failed', msg))
        worker.error.connect(lambda msg: self.set_status('Operation failed: ' + msg[:120]))
        worker.finished.connect(lambda: self.set_busy(False))
        worker.finished.connect(thread.quit)
        worker.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        thread.finished.connect(lambda: self.worker_threads.remove(record) if record in self.worker_threads else None)
        thread.start()

    def scan_whole_page(self):
        if not HAS_SELENIUM:
            self.set_status('Install Selenium to scan pages.')
            return
        try:
            url = validate_http_url(self.scan_url.text())
        except ValueError as exc:
            self.set_status(str(exc))
            return
        if not url:
            self.set_status('Enter a URL to scan.')
            return
        try:
            ignored = self.get_ignored_patterns()
        except re.error as exc:
            self.popup('Invalid ignore pattern', str(exc))
            return
        self.live_test_url.setText(url)
        self.launch_job(lambda: scan_page_job(url, ignored), self.populate_scan_results,
                        'Scanning page source with Selenium…')

    def populate_scan_results(self, payload):
        self.rows = payload['rows']
        self.count_context = 'Page-source snapshot'
        self.show_candidates(f'Scanned {payload["url"]} · {len(self.rows)} top locators')

    def test_live_url(self):
        if not HAS_SELENIUM:
            self.set_status('Install Selenium to test selectors.')
            return
        candidates = self.selected_candidates()
        if not candidates:
            self.set_status('Select a locator to test against a live URL.')
            return
        try:
            url = validate_http_url(self.live_test_url.text())
        except ValueError as exc:
            self.set_status(str(exc))
            return
        if not url:
            self.set_status('Enter a target URL in Pick Element or Scan Page.')
            return
        c = candidates[0].copy()
        self.launch_job(lambda: test_live_job(url, c),
                        lambda count: self.set_status(f'Live Selenium test: {count} match(es) for {c["expr"]} on {url}.'),
                        'Testing selected selector against live URL…')

    def run_health_audit_and_heal(self):
        if not HAS_SELENIUM:
            self.set_status('Install Selenium to run live Health Audit.')
            return
        try:
            url = validate_http_url(self.audit_url_entry.text())
        except ValueError as exc:
            self.set_status(str(exc))
            return
        locators = [line.strip() for line in self.audit_input_text.toPlainText().splitlines() if line.strip()]
        if not url or not locators:
            self.set_status('Enter an audit URL and at least one locator.')
            return
        put_rows(self.audit_tree, [])
        put_rows(self.suggestion_tree, [])
        self.launch_job(lambda: audit_job(url, locators), self.populate_audit_results,
                        f'Running {len(locators)} live locator checks…')

    def populate_audit_results(self, payload):
        put_rows(self.audit_tree, payload['results'])
        put_rows(self.suggestion_tree, payload['suggestions'])
        found = sum(status == 'Found' for _, _, status in payload['results'])
        self.set_status(f'Health Audit finished: {found}/{len(payload["results"])} locators found on live URL; '
                        f'{len(payload["suggestions"])} unverified suggestions/notes.')

    def pick_element_interactive(self):
        if not HAS_SELENIUM:
            self.set_status('Install Selenium to open the browser picker.')
            return
        if self.picker_thread or self.busy:
            return
        try:
            url = validate_http_url(self.url_entry.text())
        except ValueError as exc:
            self.set_status(str(exc))
            return
        self.picker_commands = queue.Queue()
        self.picker_thread = QThread(self)
        self.picker_worker = PickerWorker(url, self.picker_commands, self.pick_mode.isChecked())
        self.picker_worker.moveToThread(self.picker_thread)
        self.picker_thread.started.connect(self.picker_worker.run)
        self.picker_worker.opened.connect(self.on_picker_open)
        self.picker_worker.picked.connect(self.on_element_picked)
        self.picker_worker.navigated.connect(self.on_picker_navigated)
        self.picker_worker.error.connect(lambda msg: self.popup('Picker error', msg))
        self.picker_worker.error.connect(lambda msg: self.set_status('Picker error: ' + msg[:120]))
        self.picker_worker.finished.connect(self.on_picker_closed)
        self.picker_worker.finished.connect(self.picker_thread.quit)
        self.picker_worker.finished.connect(self.picker_worker.deleteLater)
        self.picker_thread.finished.connect(self.picker_thread.deleteLater)
        self.picker_thread.finished.connect(self.reset_picker_refs)
        self.pick_button.setEnabled(False)
        self.set_status('Opening Chrome picker…')
        self.picker_thread.start()

    def on_picker_open(self):
        self.picker_open = True
        self.close_picker_button.setEnabled(True)
        self.pick_mode.setEnabled(True)
        for button in (self.nav_back, self.nav_forward, self.nav_refresh):
            button.setEnabled(True)
        self.set_status('Chrome picker ready. Click a page element to capture it.')

    def on_picker_navigated(self, url):
        self.url_entry.setText(url)
        self.scan_url.setText(url)
        self.live_test_url.setText(url)

    def update_pick_mode(self, checked):
        self.picker_command('enable', checked)

    def picker_command(self, command, value=None):
        if self.picker_commands:
            self.picker_commands.put((command, value))

    def close_picker_browser(self):
        self.picker_command('close')
        self.set_status('Closing browser picker…')
        self.close_picker_button.setEnabled(False)

    def on_picker_closed(self):
        self.picker_open = False
        self.close_picker_button.setEnabled(False)
        self.pick_mode.setEnabled(False)
        for button in (self.nav_back, self.nav_forward, self.nav_refresh):
            button.setEnabled(False)
        self.set_status('Browser picker closed.')

    def reset_picker_refs(self):
        self.picker_thread = None
        self.picker_worker = None
        self.picker_commands = None
        self.pick_button.setEnabled(HAS_SELENIUM and not self.busy)

    def on_element_picked(self, html_content):
        self.html.setPlainText(str(html_content))
        self.switch_workspace(0)
        self.switch_mode(0)
        self.generate()
        self.set_status(f'Browser element captured; {len(self.rows)} locator candidates generated.')

    def toggle_theme(self):
        self.dark = not self.dark
        self.apply_theme()

    def apply_theme(self):
        dark = self.dark
        bg = '#0f172a' if dark else '#f4f7fb'
        surface = '#151e30' if dark else '#ffffff'
        sidebar = '#111b2d' if dark else '#eaf0f8'
        muted = '#9aa9c0' if dark else '#5e6d83'
        fg = '#e9eef9' if dark else '#1d2c43'
        border = '#344258' if dark else '#d9e2ef'
        field = '#101a2b' if dark else '#f9fbfe'
        hover = '#233553' if dark else '#ddeafa'
        table_alt = '#1a263b' if dark else '#f6f9fe'
        self.theme_button.setText('Light theme' if dark else 'Dark theme')
        self.setStyleSheet(f'''
            QWidget {{ background: {bg}; color: {fg}; font-family: "Segoe UI", "Arial"; font-size: 12px; }}
            QFrame#header {{ background: {surface}; border-bottom: 1px solid {border}; }}
            QFrame#sidebar {{ background: {sidebar}; border-right: 1px solid {border}; }}
            QFrame#footer {{ background: {surface}; border-top: 1px solid {border}; }}
            QFrame#panel {{ background: {surface}; border: 1px solid {border}; border-radius: 10px; }}
            QLabel {{ background: transparent; }}
            QLabel#appTitle {{ font-size: 18px; font-weight: 700; }}
            QLabel#pageTitle {{ font-size: 20px; font-weight: 700; }}
            QLabel#sectionTitle {{ font-size: 11px; font-weight: 700; letter-spacing: 1px; color: {muted}; }}
            QLabel#muted {{ color: {muted}; font-size: 11px; }}
            QPushButton {{ background: {surface}; border: 1px solid {border}; border-radius: 6px;
                          padding: 8px 11px; font-weight: 500; }}
            QPushButton:hover {{ background: {hover}; border-color: #3b82f6; }}
            QPushButton:focus {{ border: 2px solid #60a5fa; }}
            QPushButton:disabled {{ color: {muted}; border-color: {border}; }}
            QPushButton#primaryButton {{ background: #2563eb; color: white; border-color: #2563eb; font-weight: 700; }}
            QPushButton#primaryButton:hover {{ background: #1d4ed8; }}
            QPushButton#navButton {{ background: transparent; text-align: left; border: none; padding: 10px 13px; }}
            QPushButton#navButton:hover {{ background: {hover}; }}
            QPushButton#navButton:checked {{ background: #2563eb; color: white; }}
            QPushButton#modeButton {{ padding: 7px 5px; }}
            QPushButton#modeButton:checked {{ background: #2563eb; color: white; border-color: #2563eb; }}
            QLineEdit, QPlainTextEdit, QComboBox, QTableWidget {{
                background: {field}; color: {fg}; border: 1px solid {border}; border-radius: 6px;
                selection-background-color: #2563eb; selection-color: white; padding: 5px;
            }}
            QPlainTextEdit#codeArea {{ font-family: Consolas, "Cascadia Code", monospace; font-size: 12px; }}
            QTextBrowser#guideBrowser {{ background: {field}; color: {fg}; border: 1px solid {border};
                                        border-radius: 7px; padding: 10px; selection-background-color: #2563eb;
                                        selection-color: white; }}
            QPlainTextEdit:focus, QLineEdit:focus, QComboBox:focus, QTableWidget:focus {{ border: 2px solid #60a5fa; }}
            QTableWidget {{ alternate-background-color: {table_alt}; gridline-color: {border}; }}
            QHeaderView::section {{ background: {sidebar}; color: {fg}; border: none;
                                   border-bottom: 1px solid {border}; padding: 9px 6px; font-weight: 600; }}
            QTableWidget::item:selected {{ background: #2563eb; color: white; }}
            QComboBox::drop-down {{ border: none; width: 28px; background: transparent;
                                    subcontrol-origin: padding; subcontrol-position: right center; }}
            QComboBox::down-arrow {{ image: none; }}
            QComboBox QAbstractItemView {{ background: {surface}; color: {fg};
                border: 1px solid {border}; border-radius: 7px; padding: 5px;
                outline: none; selection-background-color: #2563eb; selection-color: white; }}
            QComboBox QAbstractItemView::item {{ min-height: 28px; padding: 5px 10px; border-radius: 4px; }}
            QComboBox QAbstractItemView::item:hover {{ background: {hover}; }}
            QToolButton#linkButton {{ background: transparent; color: #60a5fa; border: none; text-align: left; padding: 4px; }}
            QCheckBox {{ spacing: 8px; }}
            QSplitter::handle {{ background: {border}; }}
            QSplitter::handle:horizontal {{ width: 5px; }}
            QSplitter::handle:vertical {{ height: 5px; }}
            QProgressBar {{ background: {border}; border: none; border-radius: 4px; }}
            QProgressBar::chunk {{ background: #2563eb; border-radius: 4px; }}
            QScrollBar:vertical, QScrollBar:horizontal {{ background: {field}; }}
            QScrollBar::handle:vertical, QScrollBar::handle:horizontal {{ background: {border}; border-radius: 4px; }}
        ''')
        if hasattr(self, 'help_browser'):
            self.help_browser.document().setDefaultStyleSheet(
                f'body {{ color: {fg}; font-family: "Segoe UI", Arial, sans-serif; }} '
                f'code {{ color: {fg}; background-color: {surface}; }} '
                f'hr {{ color: {border}; }}'
            )

    def closeEvent(self, event):
        if self.picker_commands:
            self.picker_commands.put(('close', None))
        if self.picker_thread and self.picker_thread.isRunning():
            # Avoid destroying a live QThread. Continue after browser shutdown.
            self.set_status('Close the browser picker before exiting the app.')
            if not self._close_when_idle:
                self._close_when_idle = True
                self.picker_thread.finished.connect(self._retry_close_when_idle)
            event.ignore()
            return
        if any(thread.isRunning() for thread, _ in self.worker_threads):
            self.set_status('A browser operation is running. Close after it finishes.')
            if not self._close_when_idle:
                self._close_when_idle = True
                for thread, worker in list(self.worker_threads):
                    thread.finished.connect(self._retry_close_when_idle)
            event.ignore()
            return
        event.accept()

    def _retry_close_when_idle(self):
        self._close_when_idle = False
        self.close()


def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == '__main__':
    sys.exit(main())
