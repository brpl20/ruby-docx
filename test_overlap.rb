#!/usr/bin/env ruby

require 'bundler/setup'
require 'docx'

# Test the sorting and replacement logic
replacements = {
  'society_name' => 'Sociedade Empresarial',
  'society_full_name' => 'Sociedade Empresarial Tereza do Brasil Ltda'
}

start_delimiter = '_'
end_delimiter = '_'

# Show how sorting works
sorted_replacements = replacements.sort_by do |field_name, _| 
  -"#{start_delimiter}#{field_name}#{end_delimiter}".length
end

puts "Original order:"
replacements.each do |k, v|
  pattern = "#{start_delimiter}#{k}#{end_delimiter}"
  puts "  #{pattern} (length: #{pattern.length})"
end

puts "\nSorted order (by full pattern length):"
sorted_replacements.each do |k, v|
  pattern = "#{start_delimiter}#{k}#{end_delimiter}"
  puts "  #{pattern} (length: #{pattern.length})"
end

# Test replacement on sample text
test_text = "Company: _society_name_ and Full: _society_full_name_"
puts "\nOriginal text: #{test_text}"

result = test_text.dup
sorted_replacements.each do |field_name, replacement_value|
  field_pattern = "#{start_delimiter}#{field_name}#{end_delimiter}"
  puts "  Replacing '#{field_pattern}' with '#{replacement_value}'"
  result = result.gsub(field_pattern, replacement_value.to_s)
  puts "  Result: #{result}"
end

puts "\nFinal result: #{result}"