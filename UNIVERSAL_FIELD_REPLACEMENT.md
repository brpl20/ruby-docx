# Universal Field Replacement Feature

## Overview

This fork includes a robust universal field replacement system that solves the common issue of text fields being split across multiple XML text runs in Word documents. This makes template-based document generation reliable and production-ready.

## The Problem

Microsoft Word internally splits text into multiple `<w:r>` (text run) elements, even when the text appears continuous in the document. For example, the field `_company_name_` might be stored as:

```xml
<w:r><w:t>_</w:t></w:r>
<w:r><w:t>company_name</w:t></w:r>
<w:r><w:t>_</w:t></w:r>
```

This causes the standard `substitute` method to fail because it only searches within individual runs.

## The Solution

The new `replace_fields` method works at the paragraph level, reconstructing the full text before performing replacements. This ensures fields are replaced correctly regardless of how Word has split them internally.

## Features

✅ **Universal delimiter support** - Use any delimiter pattern (not hardcoded)  
✅ **Paragraph-level processing** - More reliable than run-level substitution  
✅ **Preserves formatting** - Maintains document structure and styling  
✅ **Table support** - Works with table cells automatically  
✅ **Production-tested** - Handles edge cases and complex documents  

## Usage

### Basic Usage

```ruby
require 'docx'

# Open a document with template fields
doc = Docx::Document.open('template.docx')

# Define your replacements
replacements = {
  'company_name' => 'ACME Corporation',
  'date' => '2024-01-15',
  'client_name' => 'John Doe',
  'amount' => '10,000.00'
}

# Replace all fields (using default _ delimiters)
doc.replace_fields(replacements)

# Save the result
doc.save('output.docx')
```

### Custom Delimiters

```ruby
# For templates using {{field}} pattern
doc.replace_fields(replacements, '{{', '}}')

# For templates using [field] pattern
doc.replace_fields(replacements, '[', ']')

# For templates using <<field>> pattern
doc.replace_fields(replacements, '<<', '>>')
```

### Working with Tables

```ruby
doc = Docx::Document.open('invoice_template.docx')

# Table fields are automatically handled
replacements = {
  'item_name' => 'Professional Services',
  'quantity' => '10',
  'unit_price' => '150.00',
  'total' => '1,500.00'
}

doc.replace_fields(replacements)
doc.save('invoice.docx')
```

### Multiple Partners/Items Example

```ruby
# For handling multiple items (e.g., multiple partners in a contract)
doc = Docx::Document.open('contract_template.docx')

# Process the document first
partners = [
  "John Doe, Lawyer, ID: 12345",
  "Jane Smith, Lawyer, ID: 67890",
  "Bob Johnson, Lawyer, ID: 11111"
]

replacements = {
  'company' => 'Legal Associates LLC',
  'date' => Date.today.to_s,
  'partners' => partners.join("; "), # Join multiple items
  'total_partners' => partners.size.to_s
}

doc.replace_fields(replacements)

# For tables with multiple rows, iterate and update
doc.tables.each do |table|
  # Check if this is the partners table
  if table.rows.first.cells.first.text.include?("Partner Name")
    partners.each_with_index do |partner, index|
      row = table.rows[index + 1] # Skip header
      if row
        row.cells[0].text = partner
        # Update other cells as needed
      end
    end
  end
end

doc.save('contract_filled.docx')
```

## API Reference

### Document#replace_fields

```ruby
replace_fields(replacements, start_delimiter = '_', end_delimiter = '_')
```

**Parameters:**
- `replacements` (Hash) - Keys are field names, values are replacement text
- `start_delimiter` (String) - Opening delimiter (default: '_')
- `end_delimiter` (String) - Closing delimiter (default: '_')

### Paragraph#replace_fields

```ruby
paragraph.replace_fields(replacements, start_delimiter = '_', end_delimiter = '_')
```

Same parameters as Document#replace_fields, but operates on a single paragraph.

### TableCell#text=

```ruby
cell.text = "New content"
```

Sets the text content of a table cell (updates first paragraph or creates one).

### TableCell#replace_fields

```ruby
cell.replace_fields(replacements, start_delimiter = '_', end_delimiter = '_')
```

Replaces fields in all paragraphs within a table cell.

## Template Best Practices

1. **Use clear, descriptive field names**: `_client_full_name_` instead of `_name_`
2. **Be consistent with delimiters**: Pick one pattern and stick to it
3. **Avoid special characters in field names**: Use only letters, numbers, and underscores
4. **Test your templates**: Word may add hidden formatting that affects fields

## Migration from Standard Substitute

If you're currently using the standard `substitute` method:

```ruby
# Old way (unreliable with split text runs)
doc.paragraphs.each do |p|
  p.each_text_run do |tr|
    tr.substitute('_name_', 'John Doe')
  end
end

# New way (reliable)
doc.replace_fields({'name' => 'John Doe'})
```

## Performance Considerations

The `replace_fields` method:
- Processes the entire paragraph text at once (more efficient than multiple substitutions)
- Only modifies paragraphs that contain matching fields
- Preserves document structure and formatting

## Troubleshooting

### Fields not being replaced

1. Check the exact field name in your template (including delimiters)
2. Ensure delimiters match what you're passing to `replace_fields`
3. Use a debugging tool to inspect the document:

```ruby
# Debug: See all text in document
doc.paragraphs.each do |p|
  puts p.text
end

# Debug: See all text in tables
doc.tables.each do |table|
  table.rows.each do |row|
    row.cells.each do |cell|
      puts cell.text
    end
  end
end
```

### Formatting issues after replacement

The `replace_fields` method preserves paragraph-level formatting but may lose some character-level formatting (bold, italic) within replaced text. If you need to preserve character formatting, consider using multiple smaller fields instead of one large field.

## Contributing

Found a bug or want to add a feature? Please open an issue or submit a pull request on GitHub.

## License

This enhancement maintains the same license as the original ruby-docx gem.