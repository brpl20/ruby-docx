#!/usr/bin/env ruby

require 'bundler/setup'
require 'docx'

template_path = 'examples/CS_TEMPLATE.docx'

unless File.exist?(template_path)
  puts "Error: Template file not found at #{template_path}"
  exit 1
end

doc = Docx::Document.open(template_path)

puts "=== BEFORE REPLACEMENT ==="
doc.paragraphs.each_with_index do |p, i|
  text = p.to_s
  if text.include?('_society_name') 
    puts "Paragraph #{i}: #{text}"
  end
end

# Use the CORRECT keys that match what's in the document
correct_data = {
  'society_name_full' => 'Sociedade Empresarial Completa Ltda',  # This must come before society_name
  'society_name' => 'Sociedade Empresarial'
}

doc.replace_fields(correct_data, '_', '_')

puts "\n=== AFTER REPLACEMENT ==="
doc.paragraphs.each_with_index do |p, i|
  text = p.to_s
  if text.include?('Sociedade') || text.include?('full')
    puts "Paragraph #{i}: #{text}"
  end
end

# Save to test output
doc.save('test_output.docx')
puts "\nSaved to test_output.docx"