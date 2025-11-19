import re
import fitz  # PyMuPDF
from docx import Document
from docx2pdf import convert

# Define button type constants if not present
if not hasattr(fitz, 'PDF_BTN_TYPE_PUSHBUTTON'):
    fitz.PDF_BTN_TYPE_PUSHBUTTON = 0
if not hasattr(fitz, 'PDF_BTN_TYPE_SUBMIT'):
    fitz.PDF_BTN_TYPE_SUBMIT = 1
if not hasattr(fitz, 'PDF_BTN_TYPE_RESET'):
    fitz.PDF_BTN_TYPE_RESET = 2

def parse_placeholder(placeholder):
    """
    Parse the placeholder string like 'textbox:firstname|required'
    Returns: field_type_str, field_name, is_required, options_dict
    """
    parts = placeholder.split(':')
    if len(parts) < 2:
        raise ValueError(f"Invalid placeholder format: {placeholder}")
    
    field_type_str = parts[0].lower()
    rest = ':'.join(parts[1:])
    subparts = [p.strip() for p in rest.split('|')]
    
    field_name = subparts[0]
    is_required = 'required' in [s.lower() for s in subparts]
    
    # Parse additional options like options:a,b,c
    options_dict = {}
    for sub in subparts[1:]:
        if ':' in sub:
            key, value = sub.split(':', 1)
            # Clean the value: remove trailing units like 'pt', replace comma with dot for decimals
            value = value.strip().rstrip('abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ').replace(',', '.')
            options_dict[key.lower().strip()] = value
    
    return field_type_str, field_name, is_required, options_dict

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

def get_font_info(page, rect, search_text):
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
                if span_bbox.intersects(rect) and search_text in span_text:
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

def add_form_field(page, rect, field_type, field_name, is_required, options_dict, field_type_str, font_name, font_size, text_color, found_in_table=False):
    """
    Add a widget (form field) to the page at the given rect.
    """
    widget = fitz.Widget()
    widget.field_name = field_name
    widget.field_type = field_type
    
    # Remove red background to make fields invisible
    # widget.fill_color = (1, 0, 0)  # RGB red - commented out for invisible fields
    
    if is_required or field_type_str in ['requiredfieldattribute']:
        widget.field_flags |= fitz.PDF_FIELD_IS_REQUIRED
    
    # Adjust rect for specific types
    if field_type == fitz.PDF_WIDGET_TYPE_CHECKBOX:
        side = min(rect.width, rect.height)
        widget.rect = fitz.Rect(rect.x0, rect.y0, rect.x0 + side, rect.y0 + side)
    elif field_type == fitz.PDF_WIDGET_TYPE_RADIOBUTTON:
        side = min(rect.width, rect.height)
        widget.rect = fitz.Rect(rect.x0, rect.y0, rect.x0 + side, rect.y0 + side)
    else:
        widget.rect = rect
    
    # Apply font formatting for text-based fields
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
        
        # Readonly
        if field_type_str in ['readonlyfield']:
            widget.field_flags |= fitz.PDF_FIELD_IS_READONLY
        
        # Prevent text overflow/scrolling when field is full (only for non-table fields)
        if not found_in_table:
            widget.field_flags |= 0x800000  # PDF_TX_FIELD_IS_DONOTSCROLL
        
        widget.text_margin = (0, 0, 0, 0)
        
        # For single-line fields, adjust height to prevent vertical centering
        if not (widget.field_flags & fitz.PDF_TX_FIELD_IS_MULTILINE):
            widget.rect.y1 = widget.rect.y0 + widget.text_fontsize * 1.5
        
        # Format scripts for special fields
        if field_type_str in ['datefield', 'date']:
            widget.script_format = 'AFDate_FormatEx("yyyy-mm-dd");'
            widget.script = 'AFDate_KeystrokeEx("yyyy-mm-dd");'
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
    
    # For choice fields (combobox, listbox), add options if provided
    if field_type in (fitz.PDF_WIDGET_TYPE_COMBOBOX, fitz.PDF_WIDGET_TYPE_LISTBOX):
        if 'options' in options_dict:
            options = [opt.strip() for opt in options_dict['options'].split(',')]
            widget.choice_values = options
            if options:
                widget.field_value = options[0]  # Default to first option
    
    # For radio buttons, set button_caption if provided
    if field_type == fitz.PDF_WIDGET_TYPE_RADIOBUTTON:
        if 'value' in options_dict:
            widget.button_caption = options_dict['value']
    
    # For buttons, set button type
    if 'url' in options_dict and field_type_str == 'submitbutton':
        widget.submit_url = options_dict['url']
    
    # Tooltip
    if field_type_str == 'tooltipfieldattribute' and 'tooltip' in options_dict:
        widget.field_tooltip = options_dict['tooltip']
    
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
    
    page.add_widget(widget)

def convert_docx_to_fillable_pdf(pdf_path, output_pdf_path):
    """
    Processes a PDF file by detecting placeholders like {{textbox:firstname}}
    and replacing them with PDF form fields that automatically size to table cell dimensions.
    """
    doc = fitz.open(pdf_path)
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        full_text = page.get_text()
        
        # Detect tables on the page
        tables = list(page.find_tables(strategy="lines"))
        
        # Find all unique placeholders by searching for {{
        placeholders = set()
        for inst in page.search_for('{{'):
            # Get text in a rect around the {{
            text_rect = fitz.Rect(inst.x0, inst.y0, inst.x0 + 300, inst.y0 + 50)  # Adjust size as needed
            text = page.get_textbox(text_rect)
            matches = re.findall(r'\{\{(.*?)\}\}', text)
            placeholders.update(matches)
        
        placeholders = list(placeholders)
        print(f"Page {page_num}: Found placeholders: {placeholders}")
        
        for ph in set(placeholders):
            search_text = f"{{{{{ph}}}}}"
            instances = page.search_for(search_text)
            
            for inst_rect in instances:
                # Expand the placeholder rect slightly for better intersection detection
                expanded_rect = fitz.Rect(inst_rect.x0 - 20, inst_rect.y0 - 20, inst_rect.x1 + 20, inst_rect.y1 + 20)
                
                if ph == 'textbox:fullname':
                    print(f"Expanded rect for {ph}: {expanded_rect}")
                
                # Find the table cell containing this placeholder
                cell_rect = inst_rect
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
                
                # Apply margin to avoid overlapping borders
                margin = 0  # points
                if found_in_table:
                    cell_rect = fitz.Rect(
                        cell_rect.x0 + margin, 
                        cell_rect.y0 + margin, 
                        cell_rect.x1 - margin, 
                        cell_rect.y1 - margin
                    )
                
                # Parse the placeholder
                field_type_str, field_name, is_required, options_dict = parse_placeholder(ph)
                
                # If dimensions not specified in tag and found in table, use cell dimensions
                if found_in_table and not any(k in options_dict for k in ['rowheight', 'cellwidth', 'columnwidth', 'width', 'height']):
                    options_dict['cellwidth'] = str(cell_rect.width)
                    options_dict['rowheight'] = str(cell_rect.height)
                
                field_type = get_field_type(field_type_str)
                
                # Get font info
                font_name, font_size, text_color = get_font_info(page, cell_rect, search_text)
                
                # Add the form field using the cell dimensions
                add_form_field(page, cell_rect, field_type, field_name, is_required, 
                             options_dict, field_type_str, font_name, font_size, text_color, found_in_table)
                
                # Remove the placeholder text
                redact_annot = page.add_redact_annot(inst_rect)
                page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
    
    # Save the modified PDF
    doc.save(output_pdf_path, garbage=4, deflate=True)
    doc.close()

if __name__ == "__main__":
    import sys
    import os
    if len(sys.argv) != 2:
        print("Usage: python pdfconv.py <docx_file>")
        sys.exit(1)
    docx_path = sys.argv[1]
    if not os.path.exists(docx_path):
        print(f"Error: DOCX file not found at {docx_path}")
        sys.exit(1)
    # Convert DOCX to PDF
    pdf_path = os.path.splitext(docx_path)[0] + '.pdf'
    print(f"Converting {docx_path} to {pdf_path}")
    convert(docx_path, pdf_path)
    # Process the PDF
    output_pdf_path = os.path.splitext(pdf_path)[0] + '_fillable.pdf'
    print(f"Processing {pdf_path}")
    convert_docx_to_fillable_pdf(pdf_path, output_pdf_path)
    print(f"Output saved to: {output_pdf_path}")
