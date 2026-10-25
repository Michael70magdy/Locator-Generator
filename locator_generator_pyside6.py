"""Locator Generator — PySide6 desktop edition.

Install: pip install PySide6 beautifulsoup4 selenium openpyxl lxml
Run:     python locator_generator_pyside6.py

The locator-generation engine below is preserved from the original Tkinter project.
"""
import csv
import ast
import queue
import re
import sys
import time
from pathlib import Path
from urllib.parse import urlparse

from PySide6.QtCore import Qt, QObject, QThread, Signal, QPoint
from PySide6.QtGui import QFont, QKeySequence, QShortcut, QPainter, QPolygon, QPalette, QAction
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QFrame, QLabel, QPushButton,
    QVBoxLayout, QHBoxLayout, QLineEdit, QPlainTextEdit,
    QTableWidget, QTableWidgetItem, QAbstractItemView, QHeaderView,
    QSplitter, QStackedWidget, QComboBox, QCheckBox, QToolButton, QMenu,
    QProgressBar, QFileDialog, QMessageBox, QGroupBox, QScrollArea,
    QSizePolicy, QTextBrowser
)

try:
    from bs4 import BeautifulSoup
except ImportError:
    BeautifulSoup = None
try:
    from selenium.webdriver.common.by import By as SelBy
    from selenium.webdriver.chrome.options import Options as ChromeOptions
    from selenium.webdriver.chrome.webdriver import WebDriver as ChromeWebDriver
    from selenium.webdriver.edge.options import Options as EdgeOptions
    from selenium.webdriver.edge.webdriver import WebDriver as EdgeWebDriver
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
LOCATOR_STRATEGIES = [
    'data-testid', 'data-test', 'data-qa', 'id', 'ID shorthand', 'name',
    'aria-label', 'placeholder', 'title', 'role', 'type',
    'Tag + all classes', 'Single class', 'Class token', 'Tag + type',
    'Exact normalized text', 'Contains text', 'Link text', 'Class + Text combination',
    'Parent ID + Target Tag', 'Parent ID + Target Class',
    'Parent data-testid + Target Tag', 'Parent data-test + Target Tag',
    'Parent data-qa + Target Tag', 'Tag name', 'Relative DOM path',
]
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


def exclude_locator_candidates(candidates, excluded_types=(), excluded_strategies=()):
    types = {value.upper() for value in excluded_types}
    strategies = set(excluded_strategies)
    return [candidate for candidate in candidates
            if candidate.get('kind', '').upper() not in types
            and candidate.get('strategy') not in strategies]


class PersistentCheckMenu(QMenu):
    """Checkable dropdown that stays open while its items are selected."""
    clear_requested = Signal()

    def mouseReleaseEvent(self, event):
        action = self.actionAt(event.pos())
        if event.button() == Qt.MouseButton.LeftButton and action and action.isEnabled():
            if action.isCheckable():
                action.setChecked(not action.isChecked())
                self.setActiveAction(action)
                event.accept()
                return
            if action.text() == 'Clear selections':
                self.clear_requested.emit()
                event.accept()
                return
        super().mouseReleaseEvent(event)

    def keyPressEvent(self, event):
        if event.key() in (Qt.Key.Key_Space, Qt.Key.Key_Return, Qt.Key.Key_Enter):
            action = self.activeAction()
            if action and action.isEnabled():
                if action.isCheckable():
                    action.setChecked(not action.isChecked())
                    self.setActiveAction(action)
                    event.accept()
                    return
                if action.text() == 'Clear selections':
                    self.clear_requested.emit()
                    event.accept()
                    return
        super().keyPressEvent(event)

def generate_candidates(soup, el):
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
        # A selector that does not resolve to the target element is not useful
        # to the user. Keep unknown counts (for optional parser support), but
        # never present a confirmed zero-match selector as a candidate.
        if count == 0:
            return
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
    return _start_supported_browser(headless=headless)[0]


def _driver_error_summary(error):
    lines = [line.strip() for line in str(error).splitlines() if line.strip()]
    return ' '.join(lines[:3])[:420] or type(error).__name__


def _start_supported_browser(headless):
    """Start Chrome, falling back to Microsoft Edge if Chrome is unavailable."""
    if not HAS_SELENIUM:
        raise RuntimeError('Selenium is missing from this build.')

    errors = []
    for browser_name, driver_class, options_class in (
            ('Google Chrome', ChromeWebDriver, ChromeOptions),
            ('Microsoft Edge', EdgeWebDriver, EdgeOptions)):
        options = options_class()
        options.add_argument('--no-first-run')
        options.add_argument('--no-default-browser-check')
        if headless:
            options.add_argument('--headless=new')
        try:
            return driver_class(options=options), browser_name
        except Exception as exc:
            errors.append(f'{browser_name}: {_driver_error_summary(exc)}')

    detail = '\n'.join(errors)
    raise RuntimeError(
        'Could not start Google Chrome or Microsoft Edge. Install or repair one of these browsers, '
        'then try again. Selenium Manager may need an internet connection to download the matching browser driver.\n\n'
        + detail
    )


def selenium_picker_driver():
    """Open the preferred interactive browser for the element picker."""
    return _start_supported_browser(headless=False)


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


def scan_page_job(url, excluded_types=(), excluded_strategies=()):
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
        seen_locators = set()
        for element in all_elements[:150]:
            matches = exclude_locator_candidates(
                generate_candidates(soup, element),
                excluded_types, excluded_strategies,
            )
            if matches:
                top = matches[0].copy()
                locator_key = (top['kind'], top['expr'])
                if locator_key in seen_locators:
                    continue
                seen_locators.add(locator_key)
                top['note'] = f'<{element.name}> | ' + top['note']
                candidates.append(top)
        # Counts here are DOM snapshot counts, NOT Selenium-verified results.
        return {'rows': candidates, 'url': url, 'elements': len(all_elements)}
    finally:
        driver.quit()


def test_live_job(url, candidates):
    url = validate_http_url(url)
    driver = selenium_driver()
    try:
        driver.get(url)
        results = []
        for candidate in candidates:
            by = SelBy.XPATH if candidate['kind'] == 'XPath' else SelBy.CSS_SELECTOR
            try:
                matches = len(driver.find_elements(by, candidate['expr']))
                results.append((candidate['expr'], matches, None))
            except Exception as exc:
                results.append((candidate['expr'], None, str(exc).splitlines()[0][:200]))
        return {'results': results, 'url': url}
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


def extract_audit_locators(source):
    """Extract selectors from plain selectors or common Selenium declarations.

    Accepts forms such as ``By login = By.id("login")``,
    ``login = (By.CSS_SELECTOR, "#login")`` and direct CSS/XPath lines.
    Selenium's ID/name/class/tag strategies are converted to CSS selectors so
    the audit runner can consistently evaluate CSS and XPath.
    """
    if not source:
        return []

    def decode(value):
        try:
            return str(ast.literal_eval(value))
        except (ValueError, SyntaxError):
            return value[1:-1] if len(value) >= 2 else value

    def normalize(strategy, value):
        strategy = strategy.upper().split('.')[-1]
        if strategy in ('XPATH', 'LINK_TEXT', 'PARTIAL_LINK_TEXT'):
            if strategy == 'LINK_TEXT':
                return f"//a[normalize-space(.)={xp_literal(value)}]"
            if strategy == 'PARTIAL_LINK_TEXT':
                return f"//a[contains(normalize-space(.), {xp_literal(value)})]"
            return value
        if strategy in ('ID',):
            return '#' + value if re.fullmatch(r'[A-Za-z_][\w-]*', value) else f'[id="{value}"]'
        if strategy == 'NAME':
            return f'[name="{value}"]'
        if strategy == 'CLASS_NAME':
            return '.' + value if re.fullmatch(r'-?[_a-zA-Z][_a-zA-Z0-9-]*', value) else f'[class~="{value}"]'
        if strategy == 'TAG_NAME':
            return value
        return value

    extracted = []
    for raw_line in source.splitlines():
        line = raw_line.strip().rstrip(';').strip()
        if not line or line == '//' or line.startswith(('// ', '//\t', '# ', '#\t')):
            continue

        # Java: By.variable = By.id("..."); Python calls/tuples: By.ID, "..."
        by_call = re.search(r'By\.(xpath|cssSelector|id|name|className|tagName|linkText|partialLinkText)\s*\(\s*((?:"(?:\\.|[^"\\])*"|\'(?:\\.|[^\'\\])*\'))\s*\)', line, re.I)
        if by_call:
            # Java Selenium uses camelCase names, while normalize() expects
            # Selenium's underscore-separated strategy constants.
            strategy = {
                'cssselector': 'CSS_SELECTOR',
                'classname': 'CLASS_NAME',
                'tagname': 'TAG_NAME',
                'linktext': 'LINK_TEXT',
                'partiallinktext': 'PARTIAL_LINK_TEXT',
            }.get(by_call.group(1).lower(), by_call.group(1))
            extracted.append(normalize(strategy, decode(by_call.group(2))))
            continue

        tuple_call = re.search(r'By\.(XPATH|CSS_SELECTOR|ID|NAME|CLASS_NAME|TAG_NAME|LINK_TEXT|PARTIAL_LINK_TEXT)\s*,\s*((?:"(?:\\.|[^"\\])*"|\'(?:\\.|[^\'\\])*\'))', line, re.I)
        if tuple_call:
            extracted.append(normalize(tuple_call.group(1), decode(tuple_call.group(2))))
            continue

        # Locator APIs commonly wrap the selector in a quoted call argument.
        wrapped = re.search(r'(?:page\.locator|cy\.get|cy\.xpath)\s*\(\s*((?:"(?:\\.|[^"\\])*"|\'(?:\\.|[^\'\\])*\'))', line, re.I)
        if wrapped:
            value = decode(wrapped.group(1))
            extracted.append(value[6:] if value.startswith('xpath=') else value)
            continue

        # Direct selectors may also be assigned to variables. Extract the
        # quoted right-hand-side only when it looks like a CSS/XPath selector.
        assignment = re.search(r'=\s*((?:"(?:\\.|[^"\\])*"|\'(?:\\.|[^\'\\])*\'))\s*$', line)
        candidate = decode(assignment.group(1)) if assignment else line
        if candidate.startswith('xpath='):
            candidate = candidate[6:]
        if candidate.startswith(('//', './/', '/', '(', 'id(')) or candidate.startswith(('#', '.', '[', '*')):
            extracted.append(candidate)
        elif re.fullmatch(r'[A-Za-z][\w-]*', candidate):
            extracted.append(candidate)

    # Preserve order while dropping duplicates and blank values.
    return list(dict.fromkeys(value.strip() for value in extracted if value.strip()))


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
    opened = Signal(str)

    def __init__(self, url, commands, enabled):
        super().__init__()
        self.url = url
        self.commands = commands
        self.enabled = enabled

    def run(self):
        driver = None
        try:
            driver, browser_name = selenium_picker_driver()
            driver.get(self.url)
            self.opened.emit(browser_name)
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
            message = str(exc)
            # Closing Chrome by its own window raises a WebDriver session
            # error on the next poll. Treat that as a normal picker shutdown.
            closed_browser = any(token in message.lower() for token in (
                'no such window', 'invalid session id', 'chrome not reachable',
                'disconnected', 'target window already closed', 'web view not found',
                'max retries exceeded', 'connection refused', 'failed to establish a new connection',
            ))
            if not closed_browser:
                self.error.emit(_driver_error_summary(exc) if not isinstance(exc, RuntimeError) else message)
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
    job_error = Signal(str)
    job_finished = Signal()
    picker_error = Signal(str)

    def __init__(self):
        super().__init__()
        self.setWindowTitle(TITLE + ' | PySide6')
        screen = QApplication.primaryScreen()
        available = screen.availableGeometry() if screen else None
        initial_width = min(1360, max(560, available.width() - 32)) if available else 1200
        initial_height = min(850, max(440, available.height() - 48)) if available else 760
        self.resize(initial_width, initial_height)
        self.setMinimumSize(min(680, initial_width), min(520, initial_height))
        self.ui_scale = 1.0
        self._responsive_splitters = []
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
        self.job_error.connect(self.on_job_error)
        self.job_finished.connect(self.on_job_finished)
        self.picker_error.connect(self.on_picker_error)
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
        self.header_subtitle = label('QA automation workspace  /  CSS · XPath · Page Objects · Audits', True)
        title_box.addWidget(self.header_subtitle)
        head.addLayout(title_box, 1)
        self.responsive_nav = ArrowComboBox()
        self.responsive_nav.addItems(['Generate Locators', 'Page Objects', 'Health Audit', 'Help'])
        self.responsive_nav.setMinimumWidth(150)
        self.responsive_nav.currentIndexChanged.connect(self.switch_workspace)
        self.responsive_nav.hide()
        head.addWidget(self.responsive_nav)
        self.zoom_out_button = QPushButton('A−')
        self.zoom_out_button.setObjectName('zoomButton')
        self.zoom_out_button.setToolTip('Zoom out (Ctrl+-)')
        self.zoom_out_button.clicked.connect(lambda: self.change_zoom(-0.1))
        head.addWidget(self.zoom_out_button)
        self.zoom_label = label('100%', True)
        self.zoom_label.setMinimumWidth(38)
        self.zoom_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        head.addWidget(self.zoom_label)
        self.zoom_in_button = QPushButton('A+')
        self.zoom_in_button.setObjectName('zoomButton')
        self.zoom_in_button.setToolTip('Zoom in (Ctrl++)')
        self.zoom_in_button.clicked.connect(lambda: self.change_zoom(0.1))
        head.addWidget(self.zoom_in_button)
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
        self.sidebar = sidebar
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
        zoom_in = QShortcut(QKeySequence('Ctrl++'), self)
        zoom_in.activated.connect(lambda: self.change_zoom(0.1))
        zoom_out = QShortcut(QKeySequence('Ctrl+-'), self)
        zoom_out.activated.connect(lambda: self.change_zoom(-0.1))
        zoom_reset = QShortcut(QKeySequence('Ctrl+0'), self)
        zoom_reset.activated.connect(self.reset_zoom)
        self.shortcuts.extend((zoom_in, zoom_out, zoom_reset))

    def register_responsive_splitter(self, splitter, breakpoint=1080, ratio=0.44):
        splitter.setChildrenCollapsible(False)
        splitter.setProperty('firstPaneRatio', ratio)
        self._responsive_splitters.append((splitter, breakpoint))

    def change_zoom(self, delta):
        self.ui_scale = min(1.5, max(0.8, round(self.ui_scale + delta, 2)))
        self.apply_theme()

    def reset_zoom(self):
        self.ui_scale = 1.0
        self.apply_theme()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        width = event.size().width()
        compact = width < 980 * self.ui_scale
        if hasattr(self, 'sidebar'):
            self.sidebar.setVisible(not compact)
        if hasattr(self, 'responsive_nav'):
            self.responsive_nav.setVisible(compact)
            if compact and self.responsive_nav.currentIndex() != self.workspaces.currentIndex():
                self.responsive_nav.blockSignals(True)
                self.responsive_nav.setCurrentIndex(self.workspaces.currentIndex())
                self.responsive_nav.blockSignals(False)
        if hasattr(self, 'header_subtitle'):
            self.header_subtitle.setVisible(width >= 1050 * self.ui_scale)
        for splitter, breakpoint in getattr(self, '_responsive_splitters', []):
            wanted = Qt.Orientation.Vertical if width < breakpoint * self.ui_scale else Qt.Orientation.Horizontal
            if splitter.orientation() != wanted:
                splitter.setOrientation(wanted)
                extent = splitter.height() if wanted == Qt.Orientation.Vertical else splitter.width()
                ratio = float(splitter.property('firstPaneRatio') or 0.44)
                splitter.setSizes([max(1, round(extent * ratio)), max(1, round(extent * (1 - ratio)))])

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
        self.register_responsive_splitter(primary_split, 1220, ratio=0.30)
        outer.addWidget(primary_split, 1)

        input_panel, il = panel('INPUT')
        input_panel.setMinimumWidth(210)
        primary_split.addWidget(input_panel)
        self.mode_buttons = []
        modebar = QHBoxLayout()
        modebar.setSpacing(4)
        for i, text in enumerate(['Paste / Pick', 'Scan Page']):
            b = QPushButton(text)
            b.setObjectName('modeButton')
            b.setMinimumHeight(38)
            b.setToolTip([
                'Paste HTML or use the browser picker to capture an element.',
                'Read eligible elements from a live page URL.',
            ][i])
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
        self.html.setPlaceholderText('Paste an HTML element here, for example:\n<button data-testid="login">Log in</button>')
        self.html.setPlainText(EXAMPLE_HTML)
        pv.addWidget(self.html, 1)
        html_actions = QHBoxLayout()
        paste_clipboard = QPushButton('Paste from clipboard')
        paste_clipboard.setToolTip('Insert HTML text from the system clipboard into the editor.')
        paste_clipboard.clicked.connect(lambda: self.html.setPlainText(QApplication.clipboard().text()))
        html_actions.addWidget(paste_clipboard)
        example = QPushButton('Load example')
        example.setToolTip('Replace the editor contents with a sample HTML element.')
        example.clicked.connect(lambda: self.html.setPlainText(EXAMPLE_HTML))
        html_actions.addWidget(example)
        clear = QPushButton('Clear input')
        clear.setToolTip('Clear the HTML editor and current generated selectors.')
        clear.clicked.connect(self.clear_html)
        html_actions.addWidget(clear)
        html_actions.addStretch(1)
        pv.addLayout(html_actions)
        self.modes.addWidget(paste)

        pick = QFrame()
        pick.setObjectName('pickerCard')
        pl = QVBoxLayout(pick)
        pl.setContentsMargins(10, 8, 10, 8)
        pl.setSpacing(7)
        pl.addWidget(label('OR PICK FROM A BROWSER', True))
        pl.addWidget(label('Open Chrome or Microsoft Edge, hover over an element, and click it to capture its HTML.', True))
        self.url_entry = QLineEdit('https://example.com')
        self.url_entry.setPlaceholderText('https://page-to-pick-from.example')
        pl.addWidget(self.url_entry)
        self.url_entry.editingFinished.connect(lambda: self.live_test_url.setText(self.url_entry.text().strip()))
        pick_actions = QHBoxLayout()
        self.pick_button = QPushButton('Open picker')
        self.pick_button.setObjectName('primaryButton')
        self.pick_button.setToolTip('Launch Chrome or fall back to Microsoft Edge for the locator picker.')
        self.pick_button.clicked.connect(self.pick_element_interactive)
        pick_actions.addWidget(self.pick_button)
        self.close_picker_button = QPushButton('Close browser')
        self.close_picker_button.setToolTip('Shut down the picker browser cleanly.')
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
        pl.addWidget(label('The captured element appears in the HTML editor above.', True))
        pv.addWidget(pick)

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
        scan_summary, scan_summary_layout = panel('SCAN SNAPSHOT')
        self.scan_summary = label('No page scanned yet.\n\nA scan collects buttons, inputs, links, forms, and elements with test attributes. Results appear in Candidate Locators.', True)
        self.scan_summary.setWordWrap(True)
        scan_summary_layout.addWidget(self.scan_summary)
        scan_summary.setMinimumHeight(155)
        sl.addWidget(scan_summary, 1)
        scan_note = label('Snapshot match counts are estimates from captured HTML. Use Test live to verify selectors in the browser.', True)
        scan_note.setWordWrap(True)
        sl.addWidget(scan_note)
        sl.addStretch(1)
        self.modes.addWidget(scan)

        self.generate_button = QPushButton('Generate Locators')
        self.generate_button.setObjectName('primaryButton')
        self.generate_button.setMinimumHeight(39)
        self.generate_button.clicked.connect(self.generate)
        il.addWidget(self.generate_button)

        right = QSplitter(Qt.Orientation.Vertical)
        right.setChildrenCollapsible(False)
        primary_split.addWidget(right)
        primary_split.setSizes([300, 885])

        results_panel, rl = panel('CANDIDATE LOCATORS')
        result_desc = label('Fragment matches are calculated from supplied HTML, or from the scanned DOM snapshot.', True)
        result_desc.setWordWrap(True)
        rl.addWidget(result_desc)
        self.exclusion_actions = []
        self.exclusion_menu = PersistentCheckMenu(self)
        self.exclusion_menu.setMinimumWidth(290)
        type_heading = QAction('LOCATOR TYPES', self.exclusion_menu)
        type_heading.setEnabled(False)
        type_font = type_heading.font()
        type_font.setBold(True)
        type_heading.setFont(type_font)
        self.exclusion_menu.addAction(type_heading)
        for title, option_type, value in [('CSS selectors', 'type', 'CSS'), ('XPath selectors', 'type', 'XPath')]:
            action = QAction(title, self.exclusion_menu)
            action.setCheckable(True)
            action.setData((option_type, value))
            action.toggled.connect(self.update_exclusion_label)
            self.exclusion_menu.addAction(action)
            self.exclusion_actions.append((action, option_type, value, title))
        self.exclusion_menu.addSeparator()
        strategy_heading = QAction('LOCATOR STRATEGIES', self.exclusion_menu)
        strategy_heading.setEnabled(False)
        strategy_font = strategy_heading.font()
        strategy_font.setBold(True)
        strategy_heading.setFont(strategy_font)
        self.exclusion_menu.addAction(strategy_heading)
        for strategy in LOCATOR_STRATEGIES:
            action = QAction(strategy, self.exclusion_menu)
            action.setCheckable(True)
            action.setData(('strategy', strategy))
            action.toggled.connect(self.update_exclusion_label)
            self.exclusion_menu.addAction(action)
            self.exclusion_actions.append((action, 'strategy', strategy, strategy))
        self.exclusion_menu.addSeparator()
        clear_exclusions_action = QAction('Clear selections', self.exclusion_menu)
        self.exclusion_menu.addAction(clear_exclusions_action)
        self.exclusion_menu.clear_requested.connect(self.clear_exclusions)
        self.exclude_button = QToolButton()
        self.exclude_button.setObjectName('excludeButton')
        self.exclude_button.setText('Ignore patterns  ▾')
        self.exclude_button.setToolTip('Choose locator types or strategies to omit from the next generated or scanned results. The menu stays open while you choose multiple items.')
        self.exclude_button.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        self.exclude_button.setMenu(self.exclusion_menu)
        rl.addWidget(self.exclude_button, 0, Qt.AlignmentFlag.AlignLeft)
        self.excluded_summary = label('Ignored: none', True)
        self.excluded_summary.setWordWrap(True)
        rl.addWidget(self.excluded_summary)
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
        self.register_responsive_splitter(lower_split, 900, ratio=0.34)
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
        lower_split.setSizes([285, 545])
        right.addWidget(lower_split)
        right.setSizes([480, 300])
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
        self.register_responsive_splitter(split, 1180)
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
        self.register_responsive_splitter(main_split, 1180)
        outer.addWidget(main_split, 1)
        inputs, il = panel('LOCATORS TO AUDIT')
        il.addWidget(label('Paste selectors or Selenium locator declarations, one per line.', True))
        self.audit_input_text = QPlainTextEdit()
        self.audit_input_text.setObjectName('codeArea')
        self.audit_input_text.setPlaceholderText(
            'CSS/XPath selectors or Selenium declarations:\n'
            '//h1\n#submit-button\nBy loginButton = By.id("login")\n'
            'submit = (By.CSS_SELECTOR, "button[type=submit]")'
        )
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
 <li><b>Paste / Pick:</b> Paste an element or fragment into the editor, or open the browser picker below it. The app chooses the first element with a useful identifying attribute; if none has one, it uses the final element in the fragment. Picking an element replaces the editor contents and generates its locators.</li>
 <li><b>Scan Page:</b> Enter a URL to inspect buttons, inputs, links, forms, and elements with test IDs. Scan results come from a page-source snapshot.</li>
 <li><b>Select locators:</b> Tick the checkbox beside each selector, or use <b>Select all generated selectors</b>. Copy, add to Page Objects, send to Health Audit, and Test live use the checked selectors.</li>
 <li><b>Inspect and copy:</b> Select a checked locator to read its strategy, score, stability rank, and match count. Choose a framework to preview or copy its code binding.</li>
 <li><b>Ignore patterns:</b> Choose CSS, XPath, or one or more locator strategies. The grouped checklist stays open while you choose multiple items; active exclusions remain visible above the results and apply to HTML generation and page scans.</li>
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
 <li><b>Zoom:</b> Use the A− and A+ buttons or Ctrl+- / Ctrl++ to resize the interface; Ctrl+0 restores 100%. Split views and navigation adapt to narrower windows.</li>
 <li><b>Tab / Shift+Tab:</b> Move between controls.</li>
 <li><b>PySide6:</b> Required for the desktop app. <b>BeautifulSoup:</b> Required for HTML parsing. <b>Selenium and Chrome or Edge:</b> Required for the picker and live browser features; Selenium Manager may need internet access to set up the matching driver. <b>openpyxl:</b> Required for Excel export; CSV works without it. <b>lxml:</b> Improves XPath match counts.</li>
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
        if hasattr(self, 'responsive_nav') and self.responsive_nav.currentIndex() != index:
            self.responsive_nav.blockSignals(True)
            self.responsive_nav.setCurrentIndex(index)
            self.responsive_nav.blockSignals(False)

    def switch_mode(self, index):
        self.modes.setCurrentIndex(index)
        for number, button in enumerate(self.mode_buttons):
            button.setChecked(number == index)
        self.generate_button.setVisible(index == 0)
        self.scan_button.setEnabled(not self.busy and HAS_SELENIUM)
        self.pick_button.setEnabled(not self.busy and HAS_SELENIUM and not self.picker_open)
        if not HAS_SELENIUM and index == 1:
            self.set_status('Selenium is required for browser picking and live scans.')

    def set_status(self, message):
        self.status.setText(message)

    def popup(self, title, message):
        QMessageBox.warning(self, title, message)

    def get_excluded_locator_options(self):
        excluded_types, excluded_strategies = [], []
        for action, option_type, value, _title in self.exclusion_actions:
            if not action.isChecked():
                continue
            (excluded_types if option_type == 'type' else excluded_strategies).append(value)
        return excluded_types, excluded_strategies

    def update_exclusion_label(self, *_):
        selected = [title for action, _kind, _value, title in self.exclusion_actions
                    if action.isChecked()]
        count = len(selected)
        label_text = f'Ignore patterns ({count})' if count else 'Ignore patterns'
        self.exclude_button.setText(label_text + '  ▾')
        self.excluded_summary.setText('Ignored: ' + (', '.join(selected) if selected else 'none'))

    def clear_exclusions(self):
        for action, _kind, _value, _title in self.exclusion_actions:
            action.setChecked(False)
        self.update_exclusion_label()

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
            soup = BeautifulSoup(markup, 'html.parser')
            elements = soup.find_all(True)
            if not elements:
                raise ValueError('No HTML elements found. Paste a complete element.')
            preferred = [e for e in elements if any(e.get(k) for k in (
                'data-testid', 'data-test', 'data-qa', 'id', 'name', 'aria-label', 'placeholder'))]
            target = preferred[0] if preferred else elements[-1]
            excluded_types, excluded_strategies = self.get_excluded_locator_options()
            self.rows = exclude_locator_candidates(
                generate_candidates(soup, target),
                excluded_types, excluded_strategies,
            )
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
        worker.error.connect(self.job_error)
        worker.finished.connect(self.job_finished)
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
        self.live_test_url.setText(url)
        excluded_types, excluded_strategies = self.get_excluded_locator_options()
        self.launch_job(lambda: scan_page_job(url, excluded_types, excluded_strategies), self.populate_scan_results,
                        'Scanning page source with Selenium…')

    def populate_scan_results(self, payload):
        self.rows = payload['rows']
        self.count_context = 'Page-source snapshot'
        self.scan_summary.setText(
            f'Last scanned page:\n{payload["url"]}\n\n'
            f'Generated {len(self.rows)} top locator(s). Their match counts come from the captured page source; '
            'use Test live to verify them against the loaded page.'
        )
        self.show_candidates(f'Scanned {payload["url"]} · {len(self.rows)} top locators')

    def test_live_url(self):
        if not HAS_SELENIUM:
            self.set_status('Install Selenium to test selectors.')
            return
        candidates = self.selected_candidates()
        if not candidates:
            self.popup('Select locators', 'Check one or more locator rows before running a live test.')
            return
        try:
            url = validate_http_url(self.live_test_url.text())
        except ValueError as exc:
            self.set_status(str(exc))
            return
        if not url:
            self.set_status('Enter a target URL in the browser picker or Scan Page.')
            return
        selected = [candidate.copy() for candidate in candidates]
        self.launch_job(
            lambda: test_live_job(url, selected),
            self.show_live_test_results,
            f'Testing {len(selected)} selected locator(s) against live URL…',
        )

    def show_live_test_results(self, payload):
        results = payload['results']
        url = payload['url']
        found = sum(matches is not None and matches > 0 for _, matches, error in results if error is None)
        no_matches = sum(matches == 0 for _, matches, error in results if error is None)
        errors = sum(error is not None for _, _, error in results)
        summary = (f'Live test complete for {len(results)} locator(s).\n'
                   f'Found: {found}   |   No matches: {no_matches}   |   Errors: {errors}')
        lines = []
        for expr, matches, error in results:
            state = f'{matches} match(es)' if error is None else f'Evaluation error: {error}'
            lines.append(f'{expr}\n  {state}')
        dialog = QMessageBox(self)
        dialog.setIcon(QMessageBox.Icon.Information if errors == 0 else QMessageBox.Icon.Warning)
        dialog.setWindowTitle('Live locator test results')
        dialog.setText(summary)
        dialog.setInformativeText(url)
        dialog.setDetailedText('\n\n'.join(lines))
        dialog.setStandardButtons(QMessageBox.StandardButton.Ok)
        dialog.exec()
        self.set_status(f'Live locator test finished: {found} found, {no_matches} with no matches, {errors} errors.')

    def on_job_error(self, message):
        self.popup('Operation failed', message)
        self.set_status('Operation failed: ' + message[:120])

    def on_job_finished(self):
        self.set_busy(False)

    def on_picker_error(self, message):
        self.popup('Picker error', message)
        self.set_status('Picker error: ' + message[:120])

    def run_health_audit_and_heal(self):
        if not HAS_SELENIUM:
            self.set_status('Install Selenium to run live Health Audit.')
            return
        try:
            url = validate_http_url(self.audit_url_entry.text())
        except ValueError as exc:
            self.set_status(str(exc))
            return
        locators = extract_audit_locators(self.audit_input_text.toPlainText())
        if not url or not locators:
            self.set_status('Enter an audit URL and at least one selector or Selenium locator declaration.')
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
        self.picker_worker.error.connect(self.picker_error)
        self.picker_worker.finished.connect(self.on_picker_closed)
        self.picker_worker.finished.connect(self.picker_thread.quit)
        self.picker_worker.finished.connect(self.picker_worker.deleteLater)
        self.picker_thread.finished.connect(self.picker_thread.deleteLater)
        self.picker_thread.finished.connect(self.reset_picker_refs)
        self.pick_button.setEnabled(False)
        self.set_status('Opening browser picker…')
        self.picker_thread.start()

    def on_picker_open(self, browser_name='Google Chrome'):
        self.picker_open = True
        self.close_picker_button.setEnabled(True)
        self.pick_mode.setEnabled(True)
        for button in (self.nav_back, self.nav_forward, self.nav_refresh):
            button.setEnabled(True)
        self.set_status(f'{browser_name} picker ready. Click a page element to capture it.')

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
        scale = self.ui_scale
        px = lambda value: max(1, round(value * scale))
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
        self.zoom_label.setText(f'{round(scale * 100)}%')
        self.progress.setFixedWidth(px(170))
        self.progress.setFixedHeight(px(9))
        self.sidebar.setFixedWidth(px(194))
        for button in self.nav_buttons:
            button.setMinimumHeight(px(40))
        for button in getattr(self, 'mode_buttons', []):
            button.setMinimumHeight(px(38))
        if hasattr(self, 'generate_button'):
            self.generate_button.setMinimumHeight(px(39))
        self.setStyleSheet(f'''
            QWidget {{ background: {bg}; color: {fg}; font-family: "Segoe UI", "Arial"; font-size: {px(12)}px; }}
            QFrame#header {{ background: {surface}; border-bottom: 1px solid {border}; }}
            QFrame#sidebar {{ background: {sidebar}; border-right: 1px solid {border}; }}
            QFrame#footer {{ background: {surface}; border-top: 1px solid {border}; }}
            QFrame#panel {{ background: {surface}; border: 1px solid {border}; border-radius: 10px; }}
            QFrame#pickerCard {{ background: {sidebar}; border: 1px solid {border}; border-radius: 8px; }}
            QLabel {{ background: transparent; }}
            QLabel#appTitle {{ font-size: {px(18)}px; font-weight: 700; }}
            QLabel#pageTitle {{ font-size: {px(20)}px; font-weight: 700; }}
            QLabel#sectionTitle {{ font-size: {px(11)}px; font-weight: 700; letter-spacing: {px(1)}px; color: {muted}; }}
            QLabel#muted {{ color: {muted}; font-size: {px(11)}px; }}
            QPushButton {{ background: {surface}; border: 1px solid {border}; border-radius: 6px;
                          padding: {px(8)}px {px(11)}px; font-weight: 500; }}
            QPushButton:hover {{ background: {hover}; border-color: #3b82f6; }}
            QPushButton:focus {{ border: 2px solid #60a5fa; }}
            QPushButton:disabled {{ color: {muted}; border-color: {border}; }}
            QPushButton#primaryButton {{ background: #2563eb; color: white; border-color: #2563eb; font-weight: 700; }}
            QPushButton#primaryButton:hover {{ background: #1d4ed8; }}
            QPushButton#navButton {{ background: transparent; text-align: left; border: none; padding: {px(10)}px {px(13)}px; }}
            QPushButton#navButton:hover {{ background: {hover}; }}
            QPushButton#navButton:checked {{ background: #2563eb; color: white; }}
            QPushButton#modeButton {{ padding: {px(7)}px {px(5)}px; }}
            QPushButton#modeButton:checked {{ background: #2563eb; color: white; border-color: #2563eb; }}
            QPushButton#zoomButton {{ padding: {px(5)}px {px(8)}px; min-width: {px(24)}px; }}
            QLineEdit, QPlainTextEdit, QComboBox, QTableWidget {{
                background: {field}; color: {fg}; border: 1px solid {border}; border-radius: 6px;
                selection-background-color: #2563eb; selection-color: white; padding: {px(5)}px;
            }}
            QPlainTextEdit#codeArea {{ font-family: Consolas, "Cascadia Code", monospace; font-size: {px(12)}px; }}
            QTextBrowser#guideBrowser {{ background: {field}; color: {fg}; border: 1px solid {border};
                                        border-radius: 7px; padding: {px(10)}px; selection-background-color: #2563eb;
                                        selection-color: white; }}
            QPlainTextEdit:focus, QLineEdit:focus, QComboBox:focus, QTableWidget:focus {{ border: 2px solid #60a5fa; }}
            QTableWidget {{ alternate-background-color: {table_alt}; gridline-color: {border}; }}
            QHeaderView::section {{ background: {sidebar}; color: {fg}; border: none;
                                   border-bottom: 1px solid {border}; padding: {px(9)}px {px(6)}px; font-weight: 600; }}
            QTableWidget::item:selected {{ background: #2563eb; color: white; }}
            QComboBox::drop-down {{ border: none; width: 28px; background: transparent;
                                    subcontrol-origin: padding; subcontrol-position: right center; }}
            QComboBox::down-arrow {{ image: none; }}
            QComboBox QAbstractItemView {{ background: {surface}; color: {fg};
                border: 1px solid {border}; border-radius: 7px; padding: 5px;
                outline: none; selection-background-color: #2563eb; selection-color: white; }}
            QComboBox QAbstractItemView::item {{ min-height: {px(28)}px; padding: {px(5)}px {px(10)}px; border-radius: 4px; }}
            QComboBox QAbstractItemView::item:hover {{ background: {hover}; }}
            QToolButton#excludeButton {{ background: {surface}; color: {fg}; border: 1px solid {border};
                                         border-radius: 6px; padding: {px(7)}px {px(10)}px; }}
            QToolButton#excludeButton:hover {{ background: {hover}; border-color: #3b82f6; }}
            QMenu {{ background: {surface}; color: {fg}; border: 1px solid {border}; padding: {px(5)}px; }}
            QMenu::item {{ padding: {px(6)}px {px(26)}px {px(6)}px {px(8)}px; border-radius: 4px; }}
            QMenu::item:selected {{ background: {hover}; }}
            QMenu::item:disabled {{ color: {muted}; padding-top: {px(8)}px; padding-bottom: {px(4)}px; }}
            QMenu::separator {{ height: 1px; background: {border}; margin: {px(4)}px {px(6)}px; }}
            QMenu::indicator {{ width: 14px; height: 14px; }}
            QCheckBox {{ spacing: {px(8)}px; }}
            QSplitter::handle {{ background: {border}; }}
            QSplitter::handle:horizontal {{ width: {px(5)}px; }}
            QSplitter::handle:vertical {{ height: {px(5)}px; }}
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
