#!/usr/bin/env ruby

require 'bundler/setup'
require 'docx'

# Define the template path
template_path = 'examples/CS_TEMPLATE.docx'

unless File.exist?(template_path)
  puts "Error: Template file not found at #{template_path}"
  exit 1
end

# Open the template document
puts "Opening template: #{template_path}"
doc = Docx::Document.open(template_path)

# Print all paragraphs to see what's in the document
puts "\n=== DOCUMENT CONTENT ==="
doc.paragraphs.each_with_index do |paragraph, index|
  text = paragraph.to_s
  next if text.strip.empty?
  
  puts "\nParagraph #{index + 1}:"
  puts "  Text: '#{text}'"
  
  # Check for our placeholders
  if text.include?('_society_name_') || text.include?('_society_full_name_')
    puts "  ⚠️  Contains placeholders!"
    
    # Show character by character for debugging
    puts "  Character breakdown:"
    text.chars.each_with_index do |char, i|
      puts "    [#{i}]: '#{char}' (code: #{char.ord})"
    end
  end
end

puts "\n=== APPLYING REPLACEMENTS ==="

# Define mock data for replacements
mock_data = {
  'society_name' => 'Sociedade Empresarial',
  'society_full_name' => 'Sociedade Empresarial Tereza do Brasil Ltda'
}

# Apply replacements with debug output
doc.paragraphs.each_with_index do |paragraph, index|
  original = paragraph.to_s
  next if original.strip.empty?
  
  if original.include?('_society_name_') || original.include?('_society_full_name_')
    puts "\nProcessing paragraph #{index + 1}:"
    puts "  Original: '#{original}'"
    
    # Manually apply the replacements to see what happens
    text = original.dup
    
    # Sort by full pattern length
    sorted = mock_data.sort_by { |k, _| -"_#{k}_".length }
    
    sorted.each do |field_name, replacement_value|
      pattern = "_#{field_name}_"
      if text.include?(pattern)
        puts "  Found '#{pattern}', replacing with '#{replacement_value}'"
        text = text.gsub(pattern, replacement_value)
        puts "  After replacement: '#{text}'"
      end
    end
  end
end

# Now do the actual replacement
doc.replace_fields(mock_data, '_', '_')

puts "\n=== AFTER REPLACEMENT ==="
doc.paragraphs.each_with_index do |paragraph, index|
  text = paragraph.to_s
  next if text.strip.empty?
  
  if text.include?('full_name_') || text.include?('Empresarial')
    puts "\nParagraph #{index + 1}:"
    puts "  Text: '#{text}'"
    
    if text.include?('full_name_')
      puts "  ❌ BUG: Still contains 'full_name_'!"
    end
  end
end