require 'docx/containers/text_run'
require 'docx/containers/container'

module Docx
  module Elements
    module Containers
      class TableCell
        include Container
        include Elements::Element

        def self.tag
          'tc'
        end

        def initialize(node)
          @node = node
          @properties_tag = 'tcPr'
        end

        # Return text of paragraph's cell
        def to_s
          paragraphs.map(&:text).join('')
        end

        # Array of paragraphs contained within cell
        def paragraphs
          @node.xpath('w:p').map {|p_node| Containers::Paragraph.new(p_node) }
        end

        # Iterate over each text run within a paragraph's cell
        def each_paragraph
          paragraphs.each { |tr| yield(tr) }
        end
        
        # Set text content of the cell (updates first paragraph or creates one)
        def text=(content)
          if paragraphs.any?
            paragraphs.first.text = content
          else
            # Create a new paragraph if none exists
            p_node = Nokogiri::XML::Node.new('p', @node.document)
            p_node.namespace = @node.namespace
            @node.add_child(p_node)
            Containers::Paragraph.new(p_node).text = content
          end
        end
        
        # Replace fields in all paragraphs within the cell
        def replace_fields(replacements, start_delimiter = '_', end_delimiter = '_')
          paragraphs.each do |paragraph|
            paragraph.replace_fields(replacements, start_delimiter, end_delimiter)
          end
        end
        
        alias_method :text, :to_s
      end
    end
  end
end
