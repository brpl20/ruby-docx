#!/usr/bin/env ruby

require 'bundler/setup'
require 'docx'

# Define the template path and output path
template_path = 'CS_TEMPLATE.docx'
output_path = 'CS_OUTPUT.docx'

# Check if template exists
unless File.exist?(template_path)
  puts "Error: Template file not found at #{template_path}"
  exit 1
end

# Open the template document
puts "Opening template: #{template_path}"
doc = Docx::Document.open(template_path)

# Define mock data for replacements
# Note: The order doesn't matter anymore, but placeholders must match exactly what's in the document
# If your document has _society_name_full_, use 'society_name_full' (not 'society_full_name')
mock_data = {
  'society_name' => 'Sociedade Empresarial',
  'society_full_name' => 'Sociedade Empresarial Tereza do Brasil Ltda',  # Changed from society_full_name
}

# First, replace all text fields
puts "\nReplacing text fields..."
doc.replace_fields(mock_data, '_', '_')
mock_data.each do |field, value|
  puts "  ✓ Replaced _#{field}_ with: #{value}"
end

# Save the document
puts "\nSaving document to: #{output_path}"
begin
  doc.save(output_path)
  puts "✅ Document saved successfully!"
rescue => e
  puts "❌ Error saving document: #{e.message}"
  exit 1
end

# Display summary
puts "\n" + "="*50
puts "REPLACEMENT SUMMARY"
puts "="*50
puts "Template: #{template_path}"
puts "Output:   #{output_path}"
mock_data.each do |key, value|
  puts "  #{key.ljust(25)} => #{value}"
end
puts "="*50
puts "\n✨ Process completed successfully!"
