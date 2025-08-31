#!/usr/bin/env ruby

# Test the issue with the actual placeholders from the document
placeholders_in_doc = ['_society_name_', '_society_name_full_']

# Mock data from cs_test.rb (wrong keys)
wrong_mock_data = {
  'society_name' => 'Sociedade Empresarial',
  'society_full_name' => 'Sociedade Empresarial Tereza do Brasil Ltda'
}

# Correct mock data (matching document)
correct_mock_data = {
  'society_name' => 'Sociedade Empresarial',
  'society_name_full' => 'Sociedade Empresarial Completa'
}

puts "=== THE PROBLEM ==="
puts "Document contains: #{placeholders_in_doc.inspect}"
puts "cs_test.rb tries to replace: #{wrong_mock_data.keys.map{|k| "_#{k}_"}.inspect}"
puts "This mismatch means _society_name_full_ won't be found!"

puts "\n=== EVEN WITH CORRECT KEYS ==="
text = "_society_name_full_"
puts "Original text: '#{text}'"

# Current sorting approach
sorted = correct_mock_data.sort_by { |k, _| -"_#{k}_".length }
puts "\nSorted replacements:"
sorted.each do |k, v|
  puts "  _#{k}_ (length: #{("_#{k}_").length})"
end

# Apply replacements
sorted.each do |field_name, replacement_value|
  pattern = "_#{field_name}_"
  if text.include?(pattern)
    puts "\nReplacing '#{pattern}' with '#{replacement_value}'"
    text = text.gsub(pattern, replacement_value)
    puts "Result: '#{text}'"
  end
end

puts "\n=== THE REAL ISSUE ==="
puts "The problem is that '_society_name_full_' contains '_society_name_' as a substring!"
puts "Even when sorted by length, '_society_name_full_' gets replaced first,"
puts "but then '_society_name_' is ALSO found in the original text because gsub"
puts "looks at the whole string, not just complete placeholders."

puts "\n=== THE SOLUTION ==="
puts "We need to find ALL placeholders first, then replace only complete ones!"