import os
import re
import sys
import shutil
import difflib
import tempfile
import fitz  # PyMuPDF
from docx import Document
import tkinter as tk
from tkinter import filedialog as fd

try:
    from docx2pdf import convert as docx2pdf_convert
except Exception:
    docx2pdf_convert = None

try:
    import win32com.client as win32
except Exception:
    win32 = None

RADIO_HANDLING = 'fallback'  # options: 'fallback' | 'skip' | 'strict'
PLACEHOLDER_PATTERN = re.compile(
    r'(?:\{\{)?\s*('
    r'(?:textbox|multilinetextfield|radiobutton|checkbox|combobox|'
    r'listbox|pushbutton|submitbutton|resetbutton|imagebutton|datefield|'
    r'timefield|datetimefield|signaturefield|barcodefield|qrcodefield|'
    r'pdf417field|code128field|numericfield|decimalfield|currencyfield|'
    r'percentfield|emailfield|phonefield)'
    r':[A-Za-z0-9_\-]+(?:\|[^{}|]+)*)\s*(?:\}\})?',
    re.IGNORECASE,
)
# Option keys we try to recover even if Word mangled the separator
OPTION_KEYWORDS = (
    'value',
    'default',
    'options',
    'format',
    'tooltip',
    'width',
    'height',
    'rowheight',
    'cellwidth',
    'columnwidth',
    'label',
    'url',
    'data',
    'calculation',
    'validation',
    'multi',
    'checked',
    'required',
    'readonly',
)
NUMERIC_OPTION_KEYS = {
    'rowheight',
    'cellwidth',
    'columnwidth',
    'width',
    'height',
    'required-border-width',
    'default-border-width',
}
TEXT_OPTION_KEYS = {'value', 'default', 'label', 'tooltip', 'data', 'url'}
SCRIPT_OPTION_KEYS = {'calculation', 'validation'}

# Characters Word often injects while wrapping long tokens. They break placeholder
# detection if we do not strip them out beforehand.
_INVISIBLE_CHAR_CODES = (
    0x00AD,  # soft hyphen (discretionary hyphen)
    0x200B,  # zero-width space
    0x200C,  # zero-width non-joiner
    0x200D,  # zero-width joiner
    0x200E,  # left-to-right mark
    0x200F,  # right-to-left mark
    0x2060,  # word joiner
    0xFEFF,  # zero-width no-break space / BOM
)
_INVISIBLE_CHAR_MAP = {code: None for code in _INVISIBLE_CHAR_CODES}


def strip_invisible(text):
    """
    Remove zero-width and hyphenation characters that Word may inject when a
    placeholder wraps mid-word. These characters are not whitespace, so regex
    matching would otherwise fail.
    """
    if not text:
        return text
    return text.translate(_INVISIBLE_CHAR_MAP)


def render_docx_to_pdf(docx_path):
    """
    Convert a DOCX file to a temporary PDF so we can continue with form detection.
    Uses docx2pdf when available, otherwise falls back to a Word COM automation.
    Returns (pdf_path, cleanup_dir)
    """
    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"DOCX not found: {docx_path}")

    temp_dir = tempfile.mkdtemp(prefix="pdfconv_docx_")
    temp_pdf = os.path.join(temp_dir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")

    try:
        if docx2pdf_convert:
            docx2pdf_convert(docx_path, temp_pdf)
        elif win32:
            word = None
            try:
                word = win32.DispatchEx("Word.Application")
                word.Visible = False
                doc = word.Documents.Open(docx_path)
                doc.SaveAs(temp_pdf, FileFormat=17)  # 17 -> wdFormatPDF
                doc.Close(False)
            finally:
                if word is not None:
                    word.Quit()
        else:
            raise RuntimeError(
                "DOCX input requires either docx2pdf (recommended) or Microsoft Word with pywin32 installed."
            )
    except Exception:
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise

    return temp_pdf, temp_dir

# Use local fallbacks for constants instead of mutating fitz
BTN_PUSH = getattr(fitz, 'PDF_BTN_TYPE_PUSHBUTTON', 0)
BTN_SUBMIT = getattr(fitz, 'PDF_BTN_TYPE_SUBMIT', 1)
BTN_RESET = getattr(fitz, 'PDF_BTN_TYPE_RESET', 2)

FIELD_IS_READONLY = getattr(fitz, 'PDF_FIELD_IS_READONLY', 1)
CH_FIELD_IS_MULTISELECT = getattr(fitz, 'PDF_CH_FIELD_IS_MULTISELECT', (1 << 21))
FIELD_IS_REQUIRED = getattr(fitz, 'PDF_FIELD_IS_REQUIRED', (1 << 1))
SIG_FLAG_DIGITAL = getattr(fitz, 'PDF_SIG_FLAG_DIGITAL', 1)

# Track created radio parents per document so we can link radio options into groups.
RADIO_PARENTS = {}
# Keep richer metadata so we can repair broken groups or mismatched On-states
RADIO_GROUPS = {}
def _normalize_export_value(value):
    """
    Convert raw placeholder radio/checkbox values into Acrobat-friendly tokens.
    Acrobat prefers alphanumeric identifiers (e.g., Yes, No, Option1).
    """
    if not value:
        return 'Yes'
    cleaned = re.sub(r'[^0-9A-Za-z]+', '', value)
    if not cleaned:
        return 'Yes'
    # Title-case keeps readability while remaining deterministic
    return cleaned[0].upper() + cleaned[1:]


def _apply_radio_parent(field_name, widget):
    """
    Ensure that radio widgets sharing the same field_name point at a common parent.
    PyMuPDF models radio groups by linking members to the first widget's xref.
    """
    if widget is None:
        return
    if field_name not in RADIO_GROUPS:
        RADIO_GROUPS[field_name] = {'parent': None, 'widgets': []}
    group = RADIO_GROUPS[field_name]
    group['widgets'].append(widget)
    if group['parent'] is None and getattr(widget, 'xref', None):
        group['parent'] = widget.xref
    parent = group['parent']
    if parent is None:
        return
    # First widget uses itself as parent; all others re-point to the stored parent.
    try:
        widget.rb_parent = parent
        widget.update()
    except Exception:
        pass


def _ensure_radio_state(widget, export_value):
    """
    Ensure the radio widget exposes an On-state matching the requested export value.
    Without this Acrobat will refuse to toggle specific options (e.g., 'No').
    """
    if widget is None:
        return
    normalized = _normalize_export_value(export_value)
    try:
        widget.button_caption = normalized
        widget.update()
    except Exception:
        pass


def _build_length_guard(limit):
    """
    Acrobat does not honor max_len on multiline fields, so we inject a JS guard.
    """
    if not limit:
        return None
    guard = f"""
if (!event.willCommit) {{
    var selection = event.selEnd - event.selStart;
    var replacement = event.change ? event.change.length : 0;
    var projected = event.value.length - selection + replacement;
    if (projected > {limit}) {{
        event.rc = false;
    }}
}}
""".strip()
    return guard

# Default border colors and widths (RGB tuples and floats)
REQUIRED_BORDER_COLOR = (1.0, 0.0, 0.0)
DEFAULT_BORDER_COLOR = (0.0, 0.5, 1.0)
REQUIRED_BORDER_WIDTH = 1.0
DEFAULT_BORDER_WIDTH = 0.6

def clean_placeholder_text(text):
    """
    Clean spaces in placeholder strings to handle Word conversion artifacts.
    """
    text = strip_invisible(text)
    # First, fix spaced brackets
    text = re.sub(r'\{\s*\{', '{{', text)
    text = re.sub(r'\}\s*\}', '}}', text)
    
    # Then clean inside placeholders
    def clean_inside(match):
        inside = match.group(1)
        # Remove all whitespace (spaces, newlines, etc.)
        inside = re.sub(r'\s+', '', inside)
        # Remove spaces around colons and pipes (though spaces are removed, keep for consistency)
        inside = re.sub(r'\s*:\s*', ':', inside)
        inside = re.sub(r'\s*\|\s*', '|', inside)
        return '{{' + inside + '}}'
    
    # Split merged placeholders like '}}{{' so they become separate tokens
    text = text.replace('}}{{', '}} {{')
    text = re.sub(r'\{\{(.*?)\}\}', clean_inside, text)
    return text


def normalize_placeholder_token(token):
    """
    Repair placeholders that lost their pipe separators or picked up stray
    punctuation during the DOCX -> PDF trip.
    """
    if not token:
        return token
    token = strip_invisible(token)
    token = token.strip().strip('{}[]()')
    token = re.sub(r'\s+', '', token)

    def _prefix_option(match, key):
        start = match.start()
        prev = match.string[start - 1] if start > 0 else ''
        normalized = f"{key}:"
        if prev in ('|', ':'):
            return normalized
        if start == 0:
            return normalized
        return f"|{normalized}"

    for key in OPTION_KEYWORDS:
        pattern = re.compile(rf'(?i){key}:')
        token = pattern.sub(lambda m, _key=key: _prefix_option(m, _key), token)

    # Trim doubled separators that may have been introduced
    token = re.sub(r'\|{2,}', '|', token)
    # Remove trailing punctuation that survived the cleanup
    token = token.strip('|').strip()
    return token


def normalize_option_key(key):
    """
    Normalize option keys even if Word injected stray characters (e.g., 'Ivalue').
    """
    if not key:
        return None
    base = strip_invisible(key).lower().strip()
    if base in OPTION_KEYWORDS:
        return base
    base = re.sub(r'^[^a-z0-9]+', '', base)
    base = re.sub(r'[^a-z0-9]+$', '', base)
    if base in OPTION_KEYWORDS:
        return base
    for prefix in ('i', 'l', '1', 'j'):
        if base.startswith(prefix) and base[len(prefix):] in OPTION_KEYWORDS:
            return base[len(prefix):]
    match = difflib.get_close_matches(base, OPTION_KEYWORDS, n=1, cutoff=0.72)
    if match:
        return match[0]
    return base

def parse_placeholder(placeholder):
    """
    Parse the placeholder string like 'textbox:firstname|required'
    Returns: field_type_str, field_name, is_required, is_readonly, options_dict
    """
    placeholder = normalize_placeholder_token(placeholder)
    parts = placeholder.split(':')
    if len(parts) < 2:
        raise ValueError(f"Invalid placeholder format: {placeholder}")
    
    field_type_str = parts[0].lower()
    rest = ':'.join(parts[1:])
    match_name = re.match(r'([A-Za-z0-9_\-]+)(.*)', rest)
    if not match_name:
        raise ValueError(f"Invalid placeholder format: {placeholder}")
    field_name = match_name.group(1)
    tail = match_name.group(2)
    tail = tail or ''
    if tail and not tail.startswith('|'):
        tail = '|' + tail
    raw_subparts = [field_name] + [p.strip() for p in tail.split('|') if p.strip()]
    subparts = []
    for sub in raw_subparts:
        cleaned = sub.strip().strip(')}]')
        cleaned = cleaned.strip()
        if cleaned:
            subparts.append(cleaned)
    
    # Remove whitespace (including newlines) that may be introduced by Word line breaks
    field_name = re.sub(r'\s+', '', field_name)
    subparts[0] = field_name
    is_required = 'required' in [s.lower() for s in subparts]
    is_readonly = 'readonly' in [s.lower() for s in subparts]
    
    # Parse additional options like options:a,b,c
    options_dict = {}
    for sub in subparts[1:]:
        if ':' in sub:
            key, value = sub.split(':', 1)
            key = normalize_option_key(key.strip().strip(')}]'))
            if not key:
                continue
            value = value.strip().strip(')}]')
            if key == 'options':
                clean_value = re.sub(r'\s+', '', value)
            elif key in NUMERIC_OPTION_KEYS:
                numeric_value = value.replace(',', '.')
                clean_value = re.sub(r'\s+', '', numeric_value)
            elif key in TEXT_OPTION_KEYS:
                clean_value = ' '.join(value.split())
            elif key in SCRIPT_OPTION_KEYS:
                clean_value = value.strip()
            else:
                clean_value = re.sub(r'\s+', '', value)
            options_dict[key] = clean_value
    
    return field_type_str, field_name, is_required, is_readonly, options_dict

def get_field_type(field_type_str):
    """
    Map string to PyMuPDF widget type
    Supports common Adobe PDF form fields.
    """
    type_map = {
        # Text Fields
        'textfield': fitz.PDF_WIDGET_TYPE_TEXT,
        'textbox': fitz.PDF_WIDGET_TYPE_TEXT,
        'multilinetextfield': fitz.PDF_WIDGET_TYPE_TEXT,
        'passwordfield': fitz.PDF_WIDGET_TYPE_TEXT,
        'numericfield': fitz.PDF_WIDGET_TYPE_TEXT,
        'decimalfield': fitz.PDF_WIDGET_TYPE_TEXT,
        'numberfield': fitz.PDF_WIDGET_TYPE_TEXT,
        'currencyfield': fitz.PDF_WIDGET_TYPE_TEXT,
        'percentfield': fitz.PDF_WIDGET_TYPE_TEXT,
        'datefield': fitz.PDF_WIDGET_TYPE_TEXT,
        'timefield': fitz.PDF_WIDGET_TYPE_TEXT,
        'datetimefield': fitz.PDF_WIDGET_TYPE_TEXT,
        'emailfield': fitz.PDF_WIDGET_TYPE_TEXT,
        'phonefield': fitz.PDF_WIDGET_TYPE_TEXT,
        'richtextfield': fitz.PDF_WIDGET_TYPE_TEXT,
        'calculatedfield': fitz.PDF_WIDGET_TYPE_TEXT,
        'validationfield': fitz.PDF_WIDGET_TYPE_TEXT,
        'hiddenfield': fitz.PDF_WIDGET_TYPE_TEXT,
        'readonlyfield': fitz.PDF_WIDGET_TYPE_TEXT,
        'requiredfieldattribute': fitz.PDF_WIDGET_TYPE_TEXT,
        'tooltipfieldattribute': fitz.PDF_WIDGET_TYPE_TEXT,
        
        # Choice Fields
        'checkbox': fitz.PDF_WIDGET_TYPE_CHECKBOX,
        'radiobutton': fitz.PDF_WIDGET_TYPE_RADIOBUTTON,
        'combobox': fitz.PDF_WIDGET_TYPE_COMBOBOX,
        'listbox': fitz.PDF_WIDGET_TYPE_LISTBOX,
        'dropdownlist': fitz.PDF_WIDGET_TYPE_COMBOBOX,
        
        # Button Fields
        'pushbutton': fitz.PDF_WIDGET_TYPE_BUTTON,
        'submitbutton': fitz.PDF_WIDGET_TYPE_BUTTON,
        'resetbutton': fitz.PDF_WIDGET_TYPE_BUTTON,
        'imagebutton': fitz.PDF_WIDGET_TYPE_BUTTON,
        'imagefield': fitz.PDF_WIDGET_TYPE_BUTTON,
        
        # Signature Fields
        'signaturefield': fitz.PDF_WIDGET_TYPE_SIGNATURE,
        'digitalsignaturefield': fitz.PDF_WIDGET_TYPE_SIGNATURE,
        
        # Other Fields (mapped to text or button)
        'fileattachmentfield': fitz.PDF_WIDGET_TYPE_BUTTON,
        'barcodefield': fitz.PDF_WIDGET_TYPE_TEXT,
        'qrcodefield': fitz.PDF_WIDGET_TYPE_TEXT,
        'pdf417field': fitz.PDF_WIDGET_TYPE_TEXT,
        'code128field': fitz.PDF_WIDGET_TYPE_TEXT,
        
        # Annotations (not form fields, but mapped for compatibility)
        'annotationwidget': fitz.PDF_WIDGET_TYPE_TEXT,
        'freetextannotation': fitz.PDF_WIDGET_TYPE_TEXT,
        'inkannotation': fitz.PDF_WIDGET_TYPE_TEXT,
        'stampannotation': fitz.PDF_WIDGET_TYPE_TEXT,
        'popupannotation': fitz.PDF_WIDGET_TYPE_TEXT,
        'soundannotation': fitz.PDF_WIDGET_TYPE_TEXT,
        'movieannotation': fitz.PDF_WIDGET_TYPE_TEXT,
        'screenannotation': fitz.PDF_WIDGET_TYPE_TEXT,
        'lineannotation': fitz.PDF_WIDGET_TYPE_TEXT,
        'squareannotation': fitz.PDF_WIDGET_TYPE_TEXT,
        'circleannotation': fitz.PDF_WIDGET_TYPE_TEXT,
        'polygonannotation': fitz.PDF_WIDGET_TYPE_TEXT,
        'polylineannotation': fitz.PDF_WIDGET_TYPE_TEXT,
        'fileattachmentannotation': fitz.PDF_WIDGET_TYPE_TEXT,
        'widgetannotation': fitz.PDF_WIDGET_TYPE_TEXT,
        
        # XFA Fields (not fully supported, mapped to closest)
        'subformfield': fitz.PDF_WIDGET_TYPE_TEXT,
        'drawfield': fitz.PDF_WIDGET_TYPE_TEXT,
        'decimalfield': fitz.PDF_WIDGET_TYPE_TEXT,  # XFA
        'numericupdownfield': fitz.PDF_WIDGET_TYPE_TEXT,
        'barcodefield': fitz.PDF_WIDGET_TYPE_TEXT,  # XFA
        'validationgroup': fitz.PDF_WIDGET_TYPE_TEXT,
        
        # Legacy
        'date': fitz.PDF_WIDGET_TYPE_TEXT,
    }
    return type_map.get(field_type_str, fitz.PDF_WIDGET_TYPE_TEXT)

def get_font_info(page, rect, search_text=None):
    """
    Extract font name, size, and color for the placeholder text.
    Returns: font_name, font_size, text_color (as tuple)
    """
    text_dict = page.get_text("dict")
    for block in text_dict.get('blocks', []):
        for line in block.get('lines', []):
            for span in line.get('spans', []):
                span_text = span.get('text', '').strip()
                span_bbox = fitz.Rect(span['bbox'])
                if span_bbox.intersects(rect) and (search_text is None or search_text in span_text):
                    font_name = span.get('font', 'Helvetica')
                    font_size = span.get('size', 12.0)
                    color = span.get('color', 0)
                    r = (color >> 16) & 0xFF
                    g = (color >> 8) & 0xFF
                    b = color & 0xFF
                    text_color = (r / 255.0, g / 255.0, b / 255.0)
                    return font_name, font_size, text_color
    return 'Helvetica', 12.0, (0, 0, 0)

def get_cell_dimensions(cell_rect):
    """
    Extract actual cell dimensions from the table cell rectangle.
    Returns: cell_width, cell_height (row_height)
    """
    cell_width = cell_rect.width
    cell_height = cell_rect.height
    return cell_width, cell_height

def _union_rects(rects, padding=1.5):
    """
    Combine multiple rectangles into one, expanded slightly by padding.
    """
    if not rects:
        return None
    x0 = min(r.x0 for r in rects)
    y0 = min(r.y0 for r in rects)
    x1 = max(r.x1 for r in rects)
    y1 = max(r.y1 for r in rects)
    return fitz.Rect(x0 - padding, y0 - padding, x1 + padding, y1 + padding)

def _center_square(rect):
    """
    Return a square rect centered within the given rectangle.
    """
    side = min(rect.width, rect.height)
    cx = rect.x0 + rect.width / 2.0
    cy = rect.y0 + rect.height / 2.0
    half = side / 2.0
    return fitz.Rect(cx - half, cy - half, cx + half, cy + half)


def _snap_rect(rect, field_type):
    """
    Provide consistent alignment for small controls so they land in the visual center
    of the placeholder (important for Acrobat's modern rendering).
    """
    if rect is None:
        return rect
    if field_type in (fitz.PDF_WIDGET_TYPE_CHECKBOX, fitz.PDF_WIDGET_TYPE_RADIOBUTTON):
        return _center_square(rect)
    return rect

def estimate_max_length(rect, font_size, multiline=False):
    """
    Estimate a safe max character length for a text field based on geometry.
    """
    if font_size <= 0 or rect.width <= 0:
        return None
    usable_width = max(rect.width - 2, font_size)
    chars_per_line = max(1, int(usable_width / (font_size * 0.55)))
    if multiline:
        lines = max(1, int(rect.height / (font_size * 1.3)))
        chars_per_line *= lines
    return chars_per_line

def build_page_text_index(page):
    """
    Build a whitespace-free string for the page along with span ownership metadata.
    Now also tracks character positions within each span for precise redaction.
    """
    text_dict = page.get_text("dict")
    spans = []
    combined_parts = []
    char_to_span = []
    char_to_pos_in_span = []  # Track position of each char within its span's clean text
    for block in text_dict.get('blocks', []):
        for line in block.get('lines', []):
            for span in line.get('spans', []):
                text = strip_invisible(span.get('text', ''))
                if not text:
                    continue
                clean_text = re.sub(r'\s+', '', text)
                if not clean_text:
                    continue
                rect = fitz.Rect(span['bbox'])
                original_text = span.get('text', '')
                spans.append({'rect': rect, 'text': original_text, 'clean_text': clean_text})
                span_idx = len(spans) - 1
                for pos_in_clean, char in enumerate(clean_text):
                    combined_parts.append(char)
                    char_to_span.append(span_idx)
                    char_to_pos_in_span.append(pos_in_clean)
    return {
        'spans': spans,
        'page_clean': ''.join(combined_parts),
        'char_to_span': char_to_span,
        'char_to_pos_in_span': char_to_pos_in_span,
    }

def _compute_partial_span_rect(span_info, start_pos, end_pos):
    """
    Compute a sub-rectangle within a span for characters from start_pos to end_pos
    (positions in the clean_text). Uses proportional width estimation with safety margins.
    """
    rect = span_info['rect']
    clean_text = span_info.get('clean_text', '')
    total_chars = len(clean_text)
    
    if total_chars == 0:
        return rect
    
    # If the entire span is covered, return the full rect
    if start_pos == 0 and end_pos >= total_chars:
        return rect
    
    # Estimate character width (proportional)
    char_width = rect.width / total_chars
    
    # Calculate the sub-rectangle with safety margins to avoid cutting into adjacent text
    # Add a small inward margin (about half a character width) at boundaries that aren't at span edges
    margin = char_width * 0.5
    
    x0 = rect.x0 + (start_pos * char_width)
    x1 = rect.x0 + (end_pos * char_width)
    
    # Only apply margin if not at the start of the span
    if start_pos > 0:
        x0 += margin
    
    # Only apply margin if not at the end of the span
    if end_pos < total_chars:
        x1 -= margin
    
    # Ensure we don't create an invalid rectangle
    if x1 <= x0:
        x1 = x0 + char_width
    
    return fitz.Rect(x0, rect.y0, x1, rect.y1)


def _union_rects_no_padding(rects):
    """
    Combine multiple rectangles into one without adding padding.
    """
    if not rects:
        return None
    x0 = min(r.x0 for r in rects)
    y0 = min(r.y0 for r in rects)
    x1 = max(r.x1 for r in rects)
    y1 = max(r.y1 for r in rects)
    return fitz.Rect(x0, y0, x1, y1)


def iter_placeholders(page):
    """
    Yield (placeholder_tag, detection_rect, redact_rects) tuples detected from the page text index.
    detection_rect is the bounding box for field placement (with padding).
    redact_rects is a list of precise rectangles to redact (only the placeholder text).
    """
    index = build_page_text_index(page)
    page_clean = index['page_clean']
    if not page_clean:
        return []
    char_to_span = index['char_to_span']
    char_to_pos_in_span = index['char_to_pos_in_span']
    spans = index['spans']
    detections = []
    for match in PLACEHOLDER_PATTERN.finditer(page_clean):
        placeholder = normalize_placeholder_token(match.group(1))
        start, end = match.span()
        
        # Build precise redaction rectangles for each span segment
        redact_rects = []
        # Group consecutive characters by span
        span_ranges = {}  # span_idx -> (min_pos, max_pos) in that span's clean_text
        for i in range(start, end):
            if i >= len(char_to_span):
                continue
            span_idx = char_to_span[i]
            pos_in_span = char_to_pos_in_span[i]
            if span_idx not in span_ranges:
                span_ranges[span_idx] = [pos_in_span, pos_in_span + 1]
            else:
                span_ranges[span_idx][0] = min(span_ranges[span_idx][0], pos_in_span)
                span_ranges[span_idx][1] = max(span_ranges[span_idx][1], pos_in_span + 1)
        
        # Compute precise sub-rectangles for each span
        for span_idx, (min_pos, max_pos) in span_ranges.items():
            if 0 <= span_idx < len(spans):
                span_info = spans[span_idx]
                partial_rect = _compute_partial_span_rect(span_info, min_pos, max_pos)
                redact_rects.append(partial_rect)
        
        # Detection rect for field placement uses union with padding
        detection_rect = _union_rects(redact_rects) if redact_rects else None
        
        if detection_rect:
            detections.append((placeholder, detection_rect, redact_rects))
    return detections

def add_form_field(page, rect, field_type, field_name, is_required, is_readonly, options_dict, field_type_str, font_name, font_size, text_color, found_in_table=False, radio_handling=None):
    """
    Add a widget (form field) to the page at the given rect.
    """
    try:
        widget = fitz.Widget()
    except Exception as e:
        print(f"add_form_field: failed to create Widget object: {e}")
        raise
    if rect is None:
        print(f"No rect available for field '{field_name}', skipping")
        return
    radiomode = radio_handling or RADIO_HANDLING
    radio_export_value = None
    if field_type == fitz.PDF_WIDGET_TYPE_RADIOBUTTON:
        raw_radio_value = options_dict.get('value') or options_dict.get('default') or field_name
        radio_export_value = _normalize_export_value(raw_radio_value)
        options_dict['value'] = radio_export_value
    widget.field_name = field_name
    widget.field_type = field_type
    # Ensure text_color/text_font properties exist on the widget to satisfy
    # PyMuPDF's internal validation code; set a safe default (black) where needed.
    try:
        if not getattr(widget, 'text_color', None):
            widget.text_color = text_color if text_color else (0, 0, 0)
    except Exception:
        pass
    # Proceed with widget setup
    # Debug: print type info for signature troubleshooting
    if field_type_str in ('signaturefield', 'digitalsignaturefield') or field_type == fitz.PDF_WIDGET_TYPE_SIGNATURE:
        try:
            print(f"add_form_field: creating SIGNATURE field: name={field_name}, field_type={field_type}, field_type_str={field_type_str}")
        except Exception:
            pass
    
    # Remove red background to make fields invisible
    # widget.fill_color = (1, 0, 0)  # RGB red - commented out for invisible fields
    
    if is_required:
        widget.field_flags |= FIELD_IS_REQUIRED
    
    if is_readonly:
        widget.field_flags |= FIELD_IS_READONLY
    
    # Adjust rect for specific types
    if field_type == fitz.PDF_WIDGET_TYPE_CHECKBOX:
        widget.rect = _snap_rect(rect, field_type)
        # If a desired export value/default is provided, set it as the button caption
        provided_val = options_dict.get('value') or options_dict.get('default')
        if provided_val:
            try:
                widget.button_caption = provided_val
            except Exception:
                # Some PyMuPDF versions may not support button_caption on checkboxes
                pass
        # Acrobat commonly expects the export value 'Yes' for checked state. If the user
        # did not provide an export value and a default is requested, set the export
        # caption to 'Yes' (and set the field default/value in later steps).
        if not provided_val and 'default' in options_dict:
            try:
                widget.button_caption = 'Yes'
            except Exception:
                pass
    elif field_type == fitz.PDF_WIDGET_TYPE_RADIOBUTTON:
        if radiomode == 'skip':
            print(f"Skipping radio field '{field_name}' (mode=skip)")
            return
        widget.rect = _snap_rect(rect, field_type)
        existing_parent = RADIO_PARENTS.get(field_name)
        if existing_parent:
            widget.rb_parent = existing_parent
    else:
        widget.rect = _snap_rect(rect, field_type)
    
        # Apply font formatting for text-based fields (do not apply to signature widgets)
        if field_type in (fitz.PDF_WIDGET_TYPE_TEXT, fitz.PDF_WIDGET_TYPE_COMBOBOX, fitz.PDF_WIDGET_TYPE_LISTBOX):
            widget.text_font = font_name
            widget.text_fontsize = font_size
            widget.text_color = text_color
    
    # Handle text field variants
    if field_type == fitz.PDF_WIDGET_TYPE_TEXT:
        # Multiline only for specific types
        if field_type_str in ['multilinetextfield', 'richtextfield']:
            widget.field_flags |= fitz.PDF_TX_FIELD_IS_MULTILINE
        
        # Rich text
        if field_type_str in ['richtextfield']:
            widget.field_flags |= fitz.PDF_TX_FIELD_IS_RICH_TEXT
        
        # Password
        if field_type_str in ['passwordfield']:
            widget.field_flags |= fitz.PDF_TX_FIELD_IS_PASSWORD
        
        # Hidden
        if field_type_str in ['hiddenfield']:
            widget.field_flags |= fitz.PDF_FIELD_IS_HIDDEN
        
        # Prevent text overflow/scrolling when field is full
        widget.field_flags |= 0x800000  # PDF_TX_FIELD_IS_DONOTSCROLL
        
        widget.text_margin = (0, 0, 0, 0)
        
        # For single-line fields, adjust height to prevent vertical centering
        if not (widget.field_flags & fitz.PDF_TX_FIELD_IS_MULTILINE):
            widget.rect.y1 = widget.rect.y0 + widget.text_fontsize * 1.5
        
        multiline = bool(widget.field_flags & fitz.PDF_TX_FIELD_IS_MULTILINE)
        max_len = estimate_max_length(widget.rect, widget.text_fontsize or font_size, multiline)
        guard_script = None
        if max_len:
            try:
                widget.max_len = max_len
            except Exception:
                pass
            guard_script = _build_length_guard(max_len)
        
        # Set default value if provided
        if 'default' in options_dict:
            widget.field_value = options_dict['default']
        
        # Format scripts for special fields or format option
        if 'format' in options_dict:
            fmt = options_dict['format'].lower()
            if fmt == 'date':
                widget.script_format = 'AFDate_FormatEx("yyyy-mm-dd");'
                widget.script = 'AFDate_KeystrokeEx("yyyy-mm-dd");'
            elif fmt == 'number':
                widget.script_format = 'AFNumber_Format(0, 0, 0, 0, "", false);'
                widget.script = 'AFNumber_Keystroke(0, 0, 0, 0, "", false);'
            elif fmt == 'currency':
                widget.script_format = 'AFNumber_Format(2, 0, 0, 0, "$", false);'
                widget.script = 'AFNumber_Keystroke(2, 0, 0, 0, "$", false);'
            elif fmt == 'percent':
                widget.script_format = 'AFPercent_Format(2, 0, 0, 0, "", false);'
                widget.script = 'AFPercent_Keystroke(2, 0, 0, 0, "", false);'
            elif fmt == 'email':
                widget.script = 'event.rc = /^\\S+@\\S+\\.\\S+$/.test(event.value) || event.value == "";'
            elif fmt == 'phone':
                widget.script = 'event.rc = /^\\d{10}$/.test(event.value.replace(/\\D/g, "")) || event.value == "";'
        elif field_type_str in ['datefield', 'date']:
            widget.script_format = 'AFDate_FormatEx("yyyy-mm-dd");'
            widget.script = 'AFDate_KeystrokeEx("yyyy-mm-dd");'
            # Add a simple on-entry script that makes Acrobat recognize it as a date field
            try:
                widget.script = 'AFDate_KeystrokeEx("yyyy-mm-dd");'
            except Exception:
                pass
        elif field_type_str == 'timefield':
            widget.script_format = 'AFTime_FormatEx("HH:MM");'
            widget.script = 'AFTime_KeystrokeEx("HH:MM");'
        elif field_type_str == 'datetimefield':
            widget.script_format = 'AFDate_FormatEx("yyyy-mm-dd HH:MM");'
            widget.script = 'AFDate_KeystrokeEx("yyyy-mm-dd HH:MM");'
        elif field_type_str in ['numericfield', 'numberfield']:
            widget.script_format = 'AFNumber_Format(0, 0, 0, 0, "", false);'
            widget.script = 'AFNumber_Keystroke(0, 0, 0, 0, "", false);'
        elif field_type_str == 'decimalfield':
            widget.script_format = 'AFNumber_Format(2, 0, 0, 0, "", false);'
            widget.script = 'AFNumber_Keystroke(2, 0, 0, 0, "", false);'
        elif field_type_str == 'currencyfield':
            widget.script_format = 'AFNumber_Format(2, 0, 0, 0, "$", false);'
            widget.script = 'AFNumber_Keystroke(2, 0, 0, 0, "$", false);'
        elif field_type_str == 'percentfield':
            widget.script_format = 'AFPercent_Format(2, 0, 0, 0, "", false);'
            widget.script = 'AFPercent_Keystroke(2, 0, 0, 0, "", false);'
        elif field_type_str == 'emailfield':
            # Basic email validation script
            widget.script = 'event.rc = /^\\S+@\\S+\\.\\S+$/.test(event.value) || event.value == "";'
        elif field_type_str == 'phonefield':
            # Basic phone validation (adjust as needed)
            widget.script = 'event.rc = /^\\d{10}$/.test(event.value.replace(/\\D/g, "")) || event.value == "";'
        
        # Calculated field
        if field_type_str == 'calculatedfield' and 'calculation' in options_dict:
            widget.script = options_dict['calculation']
        
        # Validation field
        if field_type_str == 'validationfield' and 'validation' in options_dict:
            widget.script = options_dict['validation']
        
        # Barcode field
        if field_type_str in ['barcodefield', 'qrcodefield', 'pdf417field', 'code128field']:
            if 'data' in options_dict:
                widget.field_value = options_dict['data']
        if guard_script:
            existing_script = getattr(widget, 'script', '') or ''
            widget.script = f"{existing_script}\n{guard_script}".strip() if existing_script else guard_script
    
    # For choice fields (combobox, listbox), add options if provided
    if field_type in (fitz.PDF_WIDGET_TYPE_COMBOBOX, fitz.PDF_WIDGET_TYPE_LISTBOX):
        if 'options' in options_dict:
            options = [opt.strip() for opt in options_dict['options'].split(',')]
            widget.choice_values = options
            if options and 'default' not in options_dict:
                widget.field_value = options[0]  # Default to first option
        if field_type == fitz.PDF_WIDGET_TYPE_LISTBOX and 'multi' in [k.lower() for k in options_dict.keys()]:
            widget.field_flags |= CH_FIELD_IS_MULTISELECT
        if 'default' in options_dict:
            widget.field_value = options_dict['default']
    
    # For radio buttons, set button_caption if provided (export value)
    if field_type == fitz.PDF_WIDGET_TYPE_RADIOBUTTON:
        caption = radio_export_value or 'Yes'
        try:
            widget.button_caption = caption
        except Exception:
            pass
    
    # For buttons, set button type and caption
    if field_type == fitz.PDF_WIDGET_TYPE_BUTTON:
        if field_type_str in ['pushbutton', 'imagebutton']:
            widget.button_type = BTN_PUSH
        elif field_type_str == 'submitbutton':
            widget.button_type = BTN_SUBMIT
        elif field_type_str == 'resetbutton':
            widget.button_type = BTN_RESET
        if 'label' in options_dict:
            widget.button_caption = options_dict['label']
        elif field_type_str == 'pushbutton':
            widget.button_caption = 'Button'
        elif field_type_str == 'submitbutton':
            widget.button_caption = 'Submit'
        elif field_type_str == 'resetbutton':
            widget.button_caption = 'Reset'
        if 'url' in options_dict and field_type_str == 'submitbutton':
            widget.submit_url = options_dict['url']
    
        # Tooltip
        if 'tooltip' in options_dict:
            widget.field_tooltip = options_dict['tooltip']

    # Signature fields: ensure widget is a signature widget and avoid adding text properties
    if field_type == fitz.PDF_WIDGET_TYPE_SIGNATURE:
        # Ensure signature widgets stay interactive even when marked as required
        try:
            widget.field_flags &= ~FIELD_IS_READONLY
        except Exception:
            pass
        # Ensure no text font or color gets applied to signature widgets
        try:
            # Reset text-specific properties to safe defaults (don't remove properties,
            # because PyMuPDF validation expects them to exist). Signature fields should
            # not display text, but we leave a default text color to avoid validation errors.
            if hasattr(widget, 'text_font'):
                widget.text_font = None
            if hasattr(widget, 'text_fontsize'):
                widget.text_fontsize = None
            if hasattr(widget, 'text_color'):
                widget.text_color = (0, 0, 0)
        except Exception:
            pass
        # Signature widgets should not have a default value or choice values
        try:
            widget.field_value = None
        except Exception:
            pass
        # Add a tooltip to help Acrobat identify the field as a signature box
        if 'tooltip' not in options_dict:
            try:
                widget.field_tooltip = 'Sign here'
            except Exception:
                pass
        try:
            widget.sig_flags = (getattr(widget, 'sig_flags', 0) or 0) | SIG_FLAG_DIGITAL
        except Exception:
            pass
    
        # Add custom dimensions if specified in tag (overrides source dimensions)
    if any(k in options_dict for k in ['rowheight', 'cellwidth', 'columnwidth', 'width', 'height']):
        print(f"Overriding dimensions for {field_name}: original rect {widget.rect}, options {options_dict}")
        new_width = widget.rect.width
        if 'cellwidth' in options_dict:
            new_width = float(options_dict['cellwidth'])
        elif 'columnwidth' in options_dict:
            new_width = float(options_dict['columnwidth'])
        elif 'width' in options_dict:
            new_width = float(options_dict['width'])
        
        new_height = widget.rect.height
        if 'rowheight' in options_dict:
            new_height = float(options_dict['rowheight'])
        elif 'height' in options_dict:
            new_height = float(options_dict['height'])
        
        widget.rect = fitz.Rect(widget.rect.x0, widget.rect.y0, widget.rect.x0 + new_width, widget.rect.y0 + new_height)
        print(f"New rect: {widget.rect}")

        # Set border colors & widths compatible with Acrobat: red for required, blue for other fields
    try:
        if is_required:
            widget.border_color = REQUIRED_BORDER_COLOR
            widget.border_width = REQUIRED_BORDER_WIDTH
        else:
            widget.border_color = DEFAULT_BORDER_COLOR
            widget.border_width = DEFAULT_BORDER_WIDTH
        # consistent border style
        widget.border_style = 'solid'
    except Exception:
        # Not all widgets/versions support these attributes; skip silently
        pass
    
    # Add widget with graceful radio fallback handling
    try:
        page.add_widget(widget)
        # Attempt to track the live widget and clear text attributes for signature widgets
        live_widget = None
        try:
            for w in page.widgets():
                try:
                    if w.field_name == field_name and fitz.Rect(w.rect) == fitz.Rect(widget.rect):
                        live_widget = w
                        break
                except Exception:
                    continue
        except Exception:
            live_widget = None
        if live_widget is not None and field_type == fitz.PDF_WIDGET_TYPE_SIGNATURE:
            try:
                # Clear any text-specific properties; leave text_color as a default tuple to satisfy validation
                try:
                    live_widget.text_font = None
                except Exception:
                    pass
                try:
                    live_widget.text_fontsize = None
                except Exception:
                    pass
                try:
                    live_widget.text_color = (0, 0, 0)
                except Exception:
                    pass
                try:
                    live_widget.update()
                except Exception:
                    pass
            except Exception:
                pass
    except Exception as e:
        if field_type == fitz.PDF_WIDGET_TYPE_RADIOBUTTON:
            if radiomode == 'strict':
                raise
            if radiomode == 'fallback':
                print(f"Falling back to checkbox for radio field '{field_name}' ({e})")
                field_type = fitz.PDF_WIDGET_TYPE_CHECKBOX
                widget.field_type = fitz.PDF_WIDGET_TYPE_CHECKBOX
                widget.rect = _snap_rect(rect, widget.field_type)
                if hasattr(widget, 'button_caption'):
                    widget.field_value = widget.button_caption
                try:
                    page.add_widget(widget)
                except Exception as e2:
                    print(f"Fallback checkbox failed for {field_name}: {e2}")
                    return
            else:
                print(f"Skipping radio field '{field_name}' due to error: {e}")
                return
        else:
            print(f"Skipping widget {field_name} due to error: {e}")
            import traceback as _tb
            _tb.print_exc()
            return
    # After adding, find the live widget instance and ensure radio parents are registered
    live_widget = None
    try:
        for w in page.widgets():
            try:
                if w.field_name == field_name and fitz.Rect(w.rect) == fitz.Rect(widget.rect):
                    live_widget = w
                    break
            except Exception:
                continue
    except Exception:
        live_widget = None

    if field_type == fitz.PDF_WIDGET_TYPE_RADIOBUTTON:
        try:
            print(f"Radio created (live_widget) name={field_name}, rb_parent={getattr(live_widget,'rb_parent', None)}, xref={getattr(live_widget,'xref', None)}")
        except Exception:
            pass
        # If we just created the first radio option, register it as the parent (xref)
        try:
            parent = RADIO_PARENTS.get(field_name)
            if parent is None and live_widget is not None and hasattr(live_widget, 'xref'):
                RADIO_PARENTS[field_name] = live_widget.xref
                parent = RADIO_PARENTS[field_name]
            if parent is not None and live_widget is not None:
                try:
                    live_widget.rb_parent = parent
                except Exception:
                    pass
                try:
                    live_widget.update()
                except Exception:
                    pass
        except Exception:
            pass
        _apply_radio_parent(field_name, live_widget)
        _ensure_radio_state(live_widget, radio_export_value or options_dict.get('value') or 'Yes')

    # If added successfully, try to ensure appropriate visual appearance and default value.
    try:
        # For choice fields (combobox/listbox): ensure a default value is set
        if field_type in (fitz.PDF_WIDGET_TYPE_COMBOBOX, fitz.PDF_WIDGET_TYPE_LISTBOX) and 'default' in options_dict:
            widget.field_value = options_dict['default']

        # For checkboxes and radio buttons, ensure on/off states exist and set field_value accordingly
        if field_type == fitz.PDF_WIDGET_TYPE_CHECKBOX:
            # Try to set checked state if 'checked' or 'value' is provided, otherwise leave 'Off'
            try:
                states = list(widget.button_states())
            except Exception:
                states = []
            # Find a sensible on-state (prefer any state that's not 'Off' or 'Normal')
            on_state = None
            for s in states:
                sl = str(s).lower()
                if sl not in ('off', 'normal'):
                    on_state = s
                    break

            # Helper to normalize names for comparison (strip slashes, underscores, case)
            def norm(name):
                if name is None:
                    return None
                return str(name).strip().lower().lstrip('/').replace('_', '')

            provided_val = options_dict.get('value') or options_dict.get('default')
            if provided_val:
                provided_norm = norm(provided_val)
                matched = None
                for s in states:
                    if norm(s) == provided_norm:
                        matched = s
                        break
                if matched:
                    widget.field_value = matched
                elif on_state:
                    widget.field_value = on_state
                else:
                    print(f"Checkbox '{field_name}': provided default '{provided_val}' did not match any states {states}; attempting raw assignment")
                    try:
                        widget.field_value = provided_val
                    except Exception:
                        pass
            else:
                # If no provided value but there's a default option, use the first on_state
                if on_state:
                    widget.field_value = on_state
            # Debug output to help diagnose mismatches in tests
            try:
                print(f"Checkbox '{field_name}' states={states} -> set field_value={widget.field_value}")
            except Exception:
                pass
            widget.update()

        if field_type == fitz.PDF_WIDGET_TYPE_RADIOBUTTON:
            # Button caption should supply an export value
            states = list(widget.button_states())
            on_state = options_dict.get('value') or (states[0] if states else None)
            if on_state and on_state in states:
                widget.field_value = on_state
            # Ensure it's grouped (if parent is known)
            if widget.rb_parent is not None:
                widget.update()
            else:
                # Update anyway to generate an appearance
                widget.update()
    except Exception as e:
        print(f"Error finalizing widget '{field_name}': {e}")
    # End widget setup

def convert_docx_to_fillable_pdf(pdf_path, output_pdf_path, radio_handling=None):
    """
    Processes a PDF file by detecting placeholders like {{textbox:firstname}}
    and replacing them with PDF form fields that automatically size to table cell dimensions.
    """
    doc = fitz.open(pdf_path)
    # Ask viewers to regenerate appearances if needed; this improves display in browser PDF viewers
    try:
        doc.need_appearances = True
    except Exception:
        pass
    # Clear radio parents for this run so we don't mix docs
    RADIO_PARENTS.clear()
    RADIO_GROUPS.clear()
    # Map of explicit defaults provided via placeholders: field_name -> default_value
    explicit_defaults = {}
    # Track field name usage across the document to make duplicates unique
    field_name_counts = {}
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        tables = list(page.find_tables(strategy="lines"))
        detections = list(iter_placeholders(page))
        print(f"Page {page_num}: Found placeholders: {[ph for ph, _, _ in detections]}")
        
        for ph, detection_rect, redact_rects in detections:
            expanded_rect = fitz.Rect(
                detection_rect.x0 - 18,
                detection_rect.y0 - 12,
                detection_rect.x1 + 18,
                detection_rect.y1 + 12,
            )
            
            cell_rect = detection_rect
            found_in_table = False
            max_area = 0
            
            for table in tables:
                for row in table.cells:
                    cell_bbox = fitz.Rect(row[0], row[1], row[2], row[3])
                    intersection = cell_bbox & expanded_rect
                    if intersection:
                        area = intersection.get_area()
                        if area > max_area:
                            max_area = area
                            cell_rect = cell_bbox
                            found_in_table = True
            
            field_type_str, field_name, is_required, is_readonly, options_dict = parse_placeholder(ph)
            
            # Make field names unique by appending page number and occurrence count
            base_field_name = field_name
            if base_field_name not in field_name_counts:
                field_name_counts[base_field_name] = 0
            field_name_counts[base_field_name] += 1
            occurrence = field_name_counts[base_field_name]
            # First occurrence keeps original name, subsequent get _P{page}_N{count}
            if occurrence > 1:
                field_name = f"{base_field_name}_P{page_num}_N{occurrence}"
            
            if 'default' in options_dict:
                explicit_defaults[field_name] = options_dict.get('default')
            
            if found_in_table and not any(k in options_dict for k in ['rowheight', 'cellwidth', 'columnwidth', 'width', 'height']):
                options_dict['cellwidth'] = str(cell_rect.width)
                options_dict['rowheight'] = str(cell_rect.height)
            
            field_type = get_field_type(field_type_str)
            font_name, font_size, text_color = get_font_info(page, cell_rect, ph)
            
            add_form_field(
                page,
                cell_rect,
                field_type,
                field_name,
                is_required,
                is_readonly,
                options_dict,
                field_type_str,
                font_name,
                font_size,
                text_color,
                found_in_table,
                radio_handling,
            )
            
            # Use precise redact_rects to only remove the placeholder text, not surrounding text
            for redact_rect in redact_rects:
                page.add_redact_annot(redact_rect)
        
        if detections:
            page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
    
    # Save the modified PDF
    # Ensure widget appearances are updated and saved in the PDF
    try:
        # Update each widget appearance where possible
        for p in doc:
            for w in list(p.widgets()):
                try:
                    # For checkboxes and radios, the field_value controls the visual state
                    if w.field_type in (fitz.PDF_WIDGET_TYPE_CHECKBOX, fitz.PDF_WIDGET_TYPE_RADIOBUTTON):
                        # If no value is set, attempt to default to the first available on-state
                        try:
                            states = list(w.button_states())
                        except Exception:
                            states = []
                        if (not w.field_value or w.field_value in ('Off', None, '')) and states:
                            # pick a non-Off state if available
                            for s in states:
                                if s.lower() not in ('off', 'normal'):
                                    w.field_value = s
                                    break
                        # Enforce explicit defaults
                        fname = getattr(w, 'field_name', None)
                        if fname and fname in explicit_defaults:
                            w.field_value = explicit_defaults[fname]
                    # For combo/list, ensure a value is set
                    if w.field_type in (fitz.PDF_WIDGET_TYPE_COMBOBOX, fitz.PDF_WIDGET_TYPE_LISTBOX) and not w.field_value:
                        vals = getattr(w, 'choice_values', None)
                        if vals:
                            w.field_value = vals[0]
                    w.update()
                except Exception:
                    # ignore widget-specific failures but continue
                    pass
    except Exception:
        pass

    doc.save(output_pdf_path, garbage=4, deflate=True)
    doc.close()


if __name__ == "__main__":
    import sys
    import os
    import argparse

    parser = argparse.ArgumentParser(description='Convert DOCX or PDF to a fillable PDF with placeholders.')
    parser.add_argument('input', nargs='?', help='Path to DOCX or PDF input file')
    parser.add_argument('-o', '--output', help='Output PDF path (optional). If omitted, saves next to input with _fillable appended')
    parser.add_argument('--radio-handling', choices=['fallback', 'skip', 'strict'], default='fallback', help='How to handle radio buttons when PyMuPDF cannot create them (default: fallback -> convert to checkbox)')
    parser.add_argument('--required-border-color', help='Comma-separated RGB values for required field border (0-1 or 0-255), e.g. 1,0,0 or 255,0,0')
    parser.add_argument('--default-border-color', help='Comma-separated RGB values for default field border (0-1 or 0-255), e.g. 0,0.5,1 or 0,128,255')
    parser.add_argument('--required-border-width', type=float, help='Border width (points) for required fields, e.g. 1.0')
    parser.add_argument('--default-border-width', type=float, help='Border width (points) for default fields, e.g. 0.6')
    args = parser.parse_args()

    input_path = args.input
    if not input_path:
        # Use file dialog to select file
        root = tk.Tk()
        root.withdraw()  # Hide the main window
        input_path = fd.askopenfilename(
            title="Select DOCX or PDF file",
            filetypes=[("DOCX files", "*.docx"), ("PDF files", "*.pdf"), ("All files", "*.*")]
        )
    if not input_path:
        print("No file selected.")
        sys.exit(1)
    
    if not os.path.exists(input_path):
        print(f"Error: File not found at {input_path}")
        sys.exit(1)
    
    temp_dir = None
    if input_path.lower().endswith('.docx'):
        try:
            pdf_path, temp_dir = render_docx_to_pdf(input_path)
            print(f"Converted DOCX to temporary PDF at {pdf_path}")
        except Exception as exc:
            print(f"Failed to convert DOCX to PDF: {exc}")
            sys.exit(1)
    else:
        pdf_path = input_path
    
    # Process the PDF
    # If the user passed an explicit output, use it. Otherwise, if the input was a DOCX,
    # save the output next to the original DOCX (not next to the temporary PDF). For
    # PDF input, save next to the PDF file.
    if args.output:
        output_pdf_path = args.output
    else:
        if input_path.lower().endswith('.docx'):
            # Save next to the DOCX that the user opened
            output_pdf_path = os.path.splitext(input_path)[0] + '_fillable2.pdf'
        else:
            # For PDF input, use the PDF path base
            output_pdf_path = os.path.splitext(pdf_path)[0] + '_fillable2.pdf'
    # Parse and update border options if provided
    def parse_color_arg(arg):
        if not arg:
            return None
        parts = [p.strip() for p in arg.split(',')]
        vals = []
        for p in parts:
            try:
                v = float(p)
                vals.append(v)
            except Exception:
                return None
        # If components seem to be in 0..255, normalize
        if any(v > 1.0 for v in vals):
            vals = [v / 255.0 for v in vals]
        return tuple(vals)

    rcol = parse_color_arg(args.required_border_color)
    dcol = parse_color_arg(args.default_border_color)
    # Apply parsed values to module-level defaults
    if rcol:
        REQUIRED_BORDER_COLOR = rcol
    if dcol:
        DEFAULT_BORDER_COLOR = dcol
    if args.required_border_width is not None:
        REQUIRED_BORDER_WIDTH = args.required_border_width
    if args.default_border_width is not None:
        DEFAULT_BORDER_WIDTH = args.default_border_width

    try:
        print(f"Processing {pdf_path} -> {output_pdf_path} (radio handling={args.radio_handling})")
        convert_docx_to_fillable_pdf(pdf_path, output_pdf_path, radio_handling=args.radio_handling)
        print(f"Output saved to: {output_pdf_path}")
        print(f"Output saved to: {output_pdf_path}")
    finally:
        pass
