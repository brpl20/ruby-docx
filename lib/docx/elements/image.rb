require 'docx/elements/element'
require 'base64'
require 'securerandom'

module Docx
  module Elements
    class Image
      include Docx::Elements::Element
      
      DEFAULT_IMAGE_PPI = 72  # pixels per inch
      
      attr_reader :path, :width, :height, :ppi, :embed_id, :relationship_id
      
      def self.tag
        'drawing'
      end
      
      # Initialize a new image element
      # @param node [Nokogiri::XML::Node] the XML node
      # @param options [Hash] image options
      # @option options [String] :path path to the image file
      # @option options [Integer] :width image width in pixels
      # @option options [Integer] :height image height in pixels
      # @option options [Integer] :ppi pixels per inch (default: 72)
      def initialize(node, options = {})
        @node = node
        @path = options[:path]
        @width = options[:width]
        @height = options[:height]
        @ppi = options[:ppi] || DEFAULT_IMAGE_PPI
        @embed_id = options[:embed_id]
        @relationship_id = options[:relationship_id]
      end
      
      # Convert pixels to EMUs (English Metric Units)
      # Word uses EMUs internally for measurements
      def pixels_to_emus(pixels)
        inches = pixels.to_f / @ppi
        emus_per_inch = 914400
        (inches * emus_per_inch).to_i
      end
      
      # Get width in EMUs
      def width_emus
        pixels_to_emus(@width)
      end
      
      # Get height in EMUs
      def height_emus
        pixels_to_emus(@height)
      end
      
      # Create the XML structure for an inline image
      def to_xml
        drawing = Nokogiri::XML::Node.new('w:drawing', @node.document)
        
        # Create inline element
        inline = Nokogiri::XML::Node.new('wp:inline', @node.document)
        inline['distT'] = '0'
        inline['distB'] = '0'
        inline['distL'] = '0'
        inline['distR'] = '0'
        
        # Add extent (size)
        extent = Nokogiri::XML::Node.new('wp:extent', @node.document)
        extent['cx'] = width_emus.to_s
        extent['cy'] = height_emus.to_s
        inline.add_child(extent)
        
        # Add effect extent
        effect_extent = Nokogiri::XML::Node.new('wp:effectExtent', @node.document)
        effect_extent['l'] = '0'
        effect_extent['t'] = '0'
        effect_extent['r'] = '0'
        effect_extent['b'] = '0'
        inline.add_child(effect_extent)
        
        # Add docPr (document properties)
        doc_pr = Nokogiri::XML::Node.new('wp:docPr', @node.document)
        doc_pr['id'] = SecureRandom.random_number(1000000).to_s
        doc_pr['name'] = "Picture #{doc_pr['id']}"
        inline.add_child(doc_pr)
        
        # Add cNvGraphicFramePr
        cnv_graphic_frame_pr = Nokogiri::XML::Node.new('wp:cNvGraphicFramePr', @node.document)
        graphic_frame_locks = Nokogiri::XML::Node.new('a:graphicFrameLocks', @node.document)
        graphic_frame_locks['xmlns:a'] = 'http://schemas.openxmlformats.org/drawingml/2006/main'
        graphic_frame_locks['noChangeAspect'] = '1'
        cnv_graphic_frame_pr.add_child(graphic_frame_locks)
        inline.add_child(cnv_graphic_frame_pr)
        
        # Add graphic
        graphic = Nokogiri::XML::Node.new('a:graphic', @node.document)
        graphic['xmlns:a'] = 'http://schemas.openxmlformats.org/drawingml/2006/main'
        
        # Add graphicData
        graphic_data = Nokogiri::XML::Node.new('a:graphicData', @node.document)
        graphic_data['uri'] = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
        
        # Add picture
        pic = Nokogiri::XML::Node.new('pic:pic', @node.document)
        pic['xmlns:pic'] = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
        
        # Add nvPicPr (non-visual picture properties)
        nv_pic_pr = Nokogiri::XML::Node.new('pic:nvPicPr', @node.document)
        
        cnv_pr = Nokogiri::XML::Node.new('pic:cNvPr', @node.document)
        cnv_pr['id'] = doc_pr['id']
        cnv_pr['name'] = doc_pr['name']
        nv_pic_pr.add_child(cnv_pr)
        
        cnv_pic_pr = Nokogiri::XML::Node.new('pic:cNvPicPr', @node.document)
        nv_pic_pr.add_child(cnv_pic_pr)
        
        pic.add_child(nv_pic_pr)
        
        # Add blipFill
        blip_fill = Nokogiri::XML::Node.new('pic:blipFill', @node.document)
        
        blip = Nokogiri::XML::Node.new('a:blip', @node.document)
        blip['r:embed'] = @relationship_id
        blip['xmlns:r'] = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        blip_fill.add_child(blip)
        
        stretch = Nokogiri::XML::Node.new('a:stretch', @node.document)
        fill_rect = Nokogiri::XML::Node.new('a:fillRect', @node.document)
        stretch.add_child(fill_rect)
        blip_fill.add_child(stretch)
        
        pic.add_child(blip_fill)
        
        # Add spPr (shape properties)
        sp_pr = Nokogiri::XML::Node.new('pic:spPr', @node.document)
        
        xfrm = Nokogiri::XML::Node.new('a:xfrm', @node.document)
        off = Nokogiri::XML::Node.new('a:off', @node.document)
        off['x'] = '0'
        off['y'] = '0'
        xfrm.add_child(off)
        
        ext = Nokogiri::XML::Node.new('a:ext', @node.document)
        ext['cx'] = width_emus.to_s
        ext['cy'] = height_emus.to_s
        xfrm.add_child(ext)
        
        sp_pr.add_child(xfrm)
        
        prst_geom = Nokogiri::XML::Node.new('a:prstGeom', @node.document)
        prst_geom['prst'] = 'rect'
        av_lst = Nokogiri::XML::Node.new('a:avLst', @node.document)
        prst_geom.add_child(av_lst)
        sp_pr.add_child(prst_geom)
        
        pic.add_child(sp_pr)
        
        # Assemble the structure
        graphic_data.add_child(pic)
        graphic.add_child(graphic_data)
        inline.add_child(graphic)
        drawing.add_child(inline)
        
        drawing
      end
    end
  end
end