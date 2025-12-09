# Dynamic Table Generation with docxtpl

## Overview
The `change.py` script uses **docxtpl** library with **Jinja2 syntax** to replace variables and generate dynamic tables in Word documents.

## JSON Data Structure
```json
{
    "EmpName": "John Doe",
    "Qualifications": [
        { "degree": "B.Tech CSE", "year": 2020 },
        { "degree": "M.Tech AIML", "year": 2023 }
    ]
}
```

## Template Syntax

### Simple Variables
Use double curly braces:
```
{{ EmpName }}
{{ DOB }}
```

### Dynamic Table Rows
Use `{%tr for/endfor %}` tags to create repeating rows:

**In your Word template, create a table like this:**

| Degree | Year |
|--------|------|
| {%tr for qual in Qualifications %} |
| {{ qual.degree }} | {{ qual.year }} |
| {%tr endfor %} |

**Important Notes:**
- Put `{%tr for qual in Qualifications %}` in a row BEFORE your data row
- Put `{%tr endfor %}` in a row AFTER your data row
- Use `{{ qual.degree }}` and `{{ qual.year }}` in the data row cells
- The `{%tr %}` tag tells docxtpl to repeat the entire table row

## Output
The template will automatically generate one row for each item in the Qualifications array:

| Degree | Year |
|--------|------|
| B.Tech CSE | 2020 |
| M.Tech AIML | 2023 |

## Usage
1. Edit `details.json` with your data
2. Create or edit `Template.docx` with Jinja2 syntax
3. Run: `python change.py`
4. Output saved to: `output.docx`

## Files
- `change.py` - Main processing script
- `details.json` - Data source
- `Template.docx` - Word template with Jinja2 tags
- `output.docx` - Generated output
- `create_dynamic_table_template.py` - Helper to create sample templates
