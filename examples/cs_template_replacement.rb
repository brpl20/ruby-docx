#!/usr/bin/env ruby

require 'bundler/setup'
require 'docx'

# Define the template path and output path
template_path = '/Users/brpl/code/prc_api/docx/CS-TEMPLATE.docx'
output_path = '/Users/brpl/code/prc_api/docx/CS-OUTPUT.docx'
logo_path = '/Users/brpl/code/prc_api/docx/logo.png'

# Check if template exists
unless File.exist?(template_path)
  puts "Error: Template file not found at #{template_path}"
  exit 1
end

# Open the template document
puts "Opening template: #{template_path}"
doc = Docx::Document.open(template_path)

# Define mock data for replacements
mock_data = {
  'parner_total_quotes' => '15',
  'partner_full_name' => 'João Silva Santos',
  'partner_qualification' => 'Senior Partner - MBA, CPA',
  'partner_sum' => 'R$ 450.000,00',
  'percentage' => '35%',
  'society_address' => 'Rua das Empresas, 1234 - Sala 500',
  'society_city' => 'São Paulo',
  'society_name' => 'Sociedade Empresarial Exemplo Ltda.',
  'society_quote_value' => 'R$ 50.000,00',
  'society_quotes' => '1.250',
  'society_state' => 'SP',
  'society_total_value' => 'R$ 1.250.000,00',
  'society_zip_code' => '01310-100',
  'sum_percentage' => '100%',
  'total_quotes' => '2.500'
}

# First, replace all text fields
puts "\nReplacing text fields..."
doc.replace_fields(mock_data, '_', '_')
mock_data.each do |field, value|
  puts "  ✓ Replaced _#{field}_ with: #{value}"
end

# Handle logo replacement if logo file exists
if File.exist?(logo_path)
  puts "\nHandling logo replacement..."

  # First, replace the logo placeholder with a unique marker
  # that won't get split across text runs
  logo_marker = "LOGO_INSERT_HERE_UNIQUE_MARKER_12345"
  doc.replace_fields({'society_logo' => logo_marker}, '_', '_')
  puts "  ✓ Logo placeholder replaced with marker"

  # Now find the paragraph with our marker and replace it
  logo_replaced = false
  doc.paragraphs.each_with_index do |paragraph, index|
    if paragraph.to_s.include?(logo_marker)
      puts "  Found logo marker in paragraph #{index + 1}"

      # Store the paragraph position info
      paragraph_node = paragraph.node
      parent_node = paragraph_node.parent

      # Clear the paragraph
      paragraph.text = ''

      # Add the image (it will go at the end of document)
      img_paragraph = doc.add_image(logo_path, width: 200, height: 80)
      puts "  ✓ Logo image created (200x80 px)"

      # Move the image paragraph to the right position
      img_node = img_paragraph.node
      paragraph_node.add_next_sibling(img_node)  # Insert after current paragraph
      paragraph_node.remove  # Remove the empty marker paragraph
      puts "  ✓ Logo positioned correctly"

      logo_replaced = true
      break
    end
  end

  unless logo_replaced
    puts "  ⚠ Logo marker not found after replacement"
  end
else
  puts "\n⚠ Warning: Logo file not found at #{logo_path}"
  puts "  Skipping logo replacement"
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
puts "Total text replacements: #{mock_data.keys.count}"
puts "Logo replaced: #{File.exist?(logo_path) ? 'Yes' : 'No'}"
puts "\nMock data used:"
mock_data.each do |key, value|
  puts "  #{key.ljust(25)} => #{value}"
end
puts "="*50
puts "\n✨ Process completed successfully!"
