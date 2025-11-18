# PDF Converter - Fillable PDF Generator

This script converts Microsoft Word (.docx) documents to fillable PDF forms by replacing placeholder tags with interactive form fields that automatically size to table cell dimensions.

## Installation

1. Install Python 3.8 or higher
2. Install required packages:
   ```bash
   pip install PyMuPDF python-docx docx2pdf
   ```

## Usage

1. Create a Word document with placeholder tags
2. Run the script:
   ```bash
   python pdfconv.py
   ```
3. The script will generate `output.pdf` in the same directory as your DOCX file

## Placeholder Syntax

Placeholders use the format: `{{fieldtype:fieldname|option1|option2:value|...}}`

- `fieldtype`: The type of form field (see below)
- `fieldname`: Unique name for the field
- `options`: Optional modifiers (see Options section)

### Examples
- `{{textbox:name}}` - Basic text field
- `{{textbox:email|required}}` - Required text field
- `{{datefield:birthdate|width:100}}` - Date field with custom width

## Supported Field Types

### Text Fields

#### Basic Text Fields
- `{{textbox:name}}` - Single-line text input
- `{{textfield:name}}` - Same as textbox

#### Multiline Text
- `{{multilinetextfield:description}}` - Multi-line text area

#### Specialized Text Fields
- `{{passwordfield:password}}` - Masked password input
- `{{emailfield:email}}` - Email with validation
- `{{phonefield:phone}}` - Phone number with validation
- `{{datefield:birthdate}}` - Date picker with calendar
- `{{timefield:appointment}}` - Time input
- `{{datetimefield:timestamp}}` - Combined date and time
- `{{numericfield:age}}` - Numbers only
- `{{decimalfield:price}}` - Decimal numbers
- `{{currencyfield:salary}}` - Currency formatting ($)
- `{{percentfield:discount}}` - Percentage formatting (%)
- `{{richtextfield:notes}}` - Rich text formatting support

#### Advanced Text Fields
- `{{calculatedfield:total|calculation:quantity*price}}` - Auto-calculated field
- `{{validationfield:code|validation:/^\d{5}$/}}` - Custom validation
- `{{hiddenfield:userid}}` - Invisible field
- `{{readonlyfield:version}}` - Non-editable field

### Choice Fields

#### Checkboxes and Radio Buttons
- `{{checkbox:agree}}` - Single checkbox
- `{{radiobutton:gender|value:male}}` - Radio button (group by name)

#### Dropdowns and Lists
- `{{combobox:country|options:USA,Canada,Mexico}}` - Dropdown list
- `{{listbox:skills|options:Python,Java,JavaScript}}` - Multi-select list
- `{{dropdownlist:state|options:CA,NY,TX}}` - Same as combobox

### Button Fields
- `{{pushbutton:submit}}` - Clickable button
- `{{submitbutton:send|url:http://example.com}}` - Submit form to URL
- `{{resetbutton:clear}}` - Reset form button
- `{{imagebutton:upload}}` - Button with image capability
- `{{imagefield:photo}}` - Image upload field

### Signature Fields
- `{{signaturefield:signature}}` - Digital signature field
- `{{digitalsignaturefield:esign}}` - Electronic signature

### Other Fields
- `{{fileattachmentfield:document}}` - File attachment button
- `{{barcodefield:code}}` - Barcode field (text input)
- `{{qrcodefield:qrcode}}` - QR code field (text input)
- `{{pdf417field:pdf417}}` - PDF417 barcode (text input)
- `{{code128field:code128}}` - Code 128 barcode (text input)

### Annotation Fields (Mapped to Text)
- `{{freetextannotation:comment}}` - Free text annotation
- `{{stampannotation:stamp}}` - Stamp annotation
- `{{inkannotation:drawing}}` - Ink annotation
- `{{lineannotation:line}}` - Line annotation
- `{{squareannotation:box}}` - Square annotation
- `{{circleannotation:circle}}` - Circle annotation
- `{{polygonannotation:shape}}` - Polygon annotation
- `{{polylineannotation:path}}` - Polyline annotation
- `{{popupannotation:popup}}` - Popup annotation
- `{{soundannotation:audio}}` - Sound annotation
- `{{movieannotation:video}}` - Movie annotation
- `{{screenannotation:screen}}` - Screen annotation
- `{{fileattachmentannotation:file}}` - File attachment annotation

### XFA Fields (Limited Support)
- `{{subformfield:subform}}` - XFA subform
- `{{drawfield:draw}}` - XFA draw field
- `{{decimalfield:decimal}}` - XFA decimal field
- `{{numericupdownfield:spinner}}` - XFA numeric spinner
- `{{barcodefield:xfabarcode}}` - XFA barcode
- `{{validationgroup:group}}` - XFA validation group

## Options

### Common Options
- `required` - Makes field mandatory
- `width:value` - Custom field width (in points)
- `height:value` - Custom field height (in points)
- `cellwidth:value` - Override table cell width
- `rowheight:value` - Override table row height

### Field-Specific Options
- `options:value1,value2,value3` - For choice fields (combobox, listbox)
- `value:default` - Default value for radio buttons
- `url:address` - Submit URL for submit buttons
- `calculation:expression` - JavaScript calculation for calculated fields
- `validation:script` - JavaScript validation for validation fields
- `tooltip:text` - Help text for fields

### Examples with Options
```
{{textbox:name|required|width:200}}
{{combobox:country|options:USA,Canada,Mexico|required}}
{{datefield:birthdate|tooltip:Enter your date of birth}}
{{calculatedfield:total|calculation:price*quantity}}
{{submitbutton:submit|url:http://example.com/submit}}
```

## Automatic Sizing

The script automatically detects table structures in your PDF and sizes fields to fit exactly within table cells. If no table is detected, fields use the placeholder text dimensions.

## JavaScript Integration

Many field types include automatic JavaScript for formatting and validation:
- Date fields: Calendar popup and date formatting
- Number fields: Automatic number formatting
- Email/Phone: Input validation
- Calculated fields: Automatic computation

## Limitations

- XFA (XML Forms Architecture) fields have limited support
- Barcode fields are text inputs (actual barcode generation requires additional tools)
- Some annotation types are simplified to text fields
- Advanced button actions may require manual PDF editing

## Troubleshooting

- Ensure placeholders are exactly formatted: `{{type:name}}`
- Check that field names are unique
- For table fields, ensure placeholders are inside table cells
- Verify all required packages are installed

## Advanced Usage

For complex forms, you can combine multiple options:
```
{{multilinetextfield:comments|required|width:300|height:100|tooltip:Enter detailed comments}}
```

The generated PDF will be compatible with Adobe Acrobat and other PDF readers that support AcroForms.