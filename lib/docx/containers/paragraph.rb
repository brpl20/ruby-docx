require 'docx/containers/text_run'
require 'docx/containers/container'

module Docx
  module Elements
    module Containers
      class Paragraph
        include Container
        include Elements::Element

        def self.tag
          'p'
        end


        # Child elements: pPr, r, fldSimple, hlink, subDoc
        # http://msdn.microsoft.com/en-us/library/office/ee364458(v=office.11).aspx
        def initialize(node, document_properties = {}, doc = nil)
          @node = node
          @properties_tag = 'pPr'
          @document_properties = document_properties
          @font_size = @document_properties[:font_size]
          @document = doc
        end

        # Set text of paragraph
        def text=(content)
          if text_runs.size == 1
            text_runs.first.text = content
          elsif text_runs.size == 0
            new_r = TextRun.create_within(self)
            new_r.text = content
          else
            text_runs.each {|r| r.node.remove }
            new_r = TextRun.create_within(self)
            new_r.text = content
          end
        end

        # Return text of paragraph
        def to_s
          text_runs.map(&:text).join('')
        end

        # Return paragraph as a <p></p> HTML fragment with formatting based on properties.
        def to_html
          html = ''
          text_runs.each do |text_run|
            html << text_run.to_html
          end
          styles = { 'font-size' => "#{font_size}pt" }
          styles['color'] = "##{font_color}" if font_color
          styles['text-align'] = alignment if alignment
          html_tag(:p, content: html, styles: styles)
        end


        # Array of text runs contained within paragraph
        def text_runs
          @node.xpath('w:r|w:hyperlink').map { |r_node| Containers::TextRun.new(r_node, @document_properties) }
        end

        # Iterate over each text run within a paragraph
        def each_text_run
          text_runs.each { |tr| yield(tr) }
        end

        # Add a new text run with specified formatting
        # @param content [String] the text content
        # @param options [Hash] formatting options
        # @option options [Boolean] :bold make text bold
        # @option options [Boolean] :italic make text italic
        # @option options [Boolean] :underline underline text
        # @option options [Integer] :size font size in points
        # @option options [String] :color font color in hex
        # @return [TextRun] the newly created text run
        def add_text(content, options = {})
          new_run = TextRun.create_within(self)
          new_run.text = content
          
          # Apply formatting
          new_run.bold! if options[:bold]
          new_run.italic! if options[:italic]
          new_run.underline! if options[:underline]
          new_run.font_size = options[:size] if options[:size]
          new_run.color = options[:color] if options[:color]
          
          new_run
        end

        # Add bold text to the paragraph
        def add_bold_text(content)
          add_text(content, bold: true)
        end

        # Add italic text to the paragraph
        def add_italic_text(content)
          add_text(content, italic: true)
        end

        # Universal field replacement that handles split text runs
        # @param replacements [Hash] field_name => replacement_value pairs
        # @param start_delimiter [String] opening delimiter (default: '_')
        # @param end_delimiter [String] closing delimiter (default: '_')
        def replace_fields(replacements, start_delimiter = '_', end_delimiter = '_')
          full_text = to_s
          original_text = full_text.dup
          
          # Create a mapping of actual placeholders to their replacements
          # Only include placeholders that actually exist in the text
          pattern_map = {}
          
          replacements.each do |field_name, replacement_value|
            field_pattern = "#{start_delimiter}#{field_name}#{end_delimiter}"
            # Only add to map if this exact pattern exists
            if full_text.include?(field_pattern)
              pattern_map[field_pattern] = replacement_value.to_s
            end
          end
          
          # Sort patterns by length (longest first) to avoid partial replacements
          # For example, this ensures _society_name_full_ is replaced before _society_name_
          sorted_patterns = pattern_map.keys.sort_by { |pattern| -pattern.length }
          
          # Replace each pattern in order
          sorted_patterns.each do |pattern|
            full_text = full_text.gsub(pattern, pattern_map[pattern])
          end
          
          # If text changed, update the paragraph
          if full_text != original_text
            self.text = full_text
          end
        end

        # Legacy method for backward compatibility - works only within individual runs
        def substitute(pattern, replacement)
          each_text_run { |tr| tr.substitute(pattern, replacement) }
        end

        def aligned_left?
          ['left', nil].include?(alignment)
        end

        def aligned_right?
          alignment == 'right'
        end

        def aligned_center?
          alignment == 'center'
        end

        def font_size
          size_attribute = @node.at_xpath('w:pPr//w:sz//@w:val')

          return @font_size unless size_attribute

          size_attribute.value.to_i / 2
        end

        def font_color
          color_tag = @node.xpath('w:r//w:rPr//w:color').first
          color_tag ? color_tag.attributes['val'].value : nil
        end

        def style
          return nil unless @document

          @document.style_name_of(style_id) ||
            @document.default_paragraph_style
        end

        def style_id
          style_property.get_attribute('w:val')
        end

        def style=(identifier)
          id = @document.styles_configuration.style_of(identifier).id

          style_property.set_attribute('w:val', id)
        end

        alias_method :style_id=, :style=
        alias_method :text, :to_s

        private

        def style_property
          properties&.at_xpath('w:pStyle') || properties&.add_child('<w:pStyle/>').first
        end

        # Returns the alignment if any, or nil if left
        def alignment
          @node.at_xpath('.//w:jc/@w:val')&.value
        end
      end
    end
  end
end
