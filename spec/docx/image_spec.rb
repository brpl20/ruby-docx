require 'spec_helper'
require 'docx'
require 'docx/elements/image'
require 'tempfile'

describe 'Image Features' do
  let(:doc) { Docx::Document.open(fixture_path('basic.docx')) }
  let(:test_image_path) { fixture_path('replacement.png') }
  
  describe Docx::Elements::Image do
    let(:node) { Nokogiri::XML::Node.new('w:r', doc.doc) }
    let(:image) do
      Docx::Elements::Image.new(node, {
        path: test_image_path,
        width: 200,
        height: 150,
        ppi: 72,
        relationship_id: 'rId123'
      })
    end
    
    describe '#pixels_to_emus' do
      it 'converts pixels to EMUs correctly' do
        # 1 inch = 914400 EMUs, 72 pixels = 1 inch at 72 PPI
        expect(image.pixels_to_emus(72)).to eq 914400
      end
    end
    
    describe '#width_emus' do
      it 'returns width in EMUs' do
        expected = (200.0 / 72 * 914400).to_i
        expect(image.width_emus).to eq expected
      end
    end
    
    describe '#height_emus' do
      it 'returns height in EMUs' do
        expected = (150.0 / 72 * 914400).to_i
        expect(image.height_emus).to eq expected
      end
    end
    
    describe '#to_xml' do
      it 'generates valid drawing XML' do
        xml = image.to_xml
        expect(xml.name).to eq 'w:drawing'
        # Check for inline element by traversing children
        inline = xml.children.find { |c| c.name.end_with?('inline') }
        expect(inline).not_to be_nil
        # Look for a:blip element
        blip = xml.xpath('.//*[name()="a:blip"]').first
        expect(blip).not_to be_nil
        # Check the r:embed attribute directly
        expect(blip['r:embed']).to eq 'rId123'
      end
      
      it 'includes correct dimensions' do
        xml = image.to_xml
        # Find extent element in first child (inline)
        inline = xml.children.first
        extent = inline.children.find { |c| c.name.end_with?('extent') }
        expect(extent).not_to be_nil
        expect(extent['cx']).to eq image.width_emus.to_s
        expect(extent['cy']).to eq image.height_emus.to_s
      end
    end
  end
  
  describe 'Document#add_image' do
    context 'with valid image' do
      it 'adds an image to the document' do
        initial_para_count = doc.paragraphs.count
        para = doc.add_image(test_image_path, width: 300, height: 200)
        
        expect(doc.paragraphs.count).to eq initial_para_count + 1
        expect(para).to be_a(Docx::Elements::Containers::Paragraph)
      end
      
      it 'creates image with specified dimensions' do
        para = doc.add_image(test_image_path, width: 400, height: 300, ppi: 96)
        drawing = para.node.at_xpath('.//w:drawing')
        expect(drawing).not_to be_nil
      end
    end
    
    context 'with invalid parameters' do
      it 'raises error when image file does not exist' do
        expect {
          doc.add_image('nonexistent.png', width: 100, height: 100)
        }.to raise_error(ArgumentError, /Image file not found/)
      end
      
      it 'raises error when width is missing' do
        expect {
          doc.add_image(test_image_path, height: 100)
        }.to raise_error(ArgumentError, /Width is required/)
      end
      
      it 'raises error when height is missing' do
        expect {
          doc.add_image(test_image_path, width: 100)
        }.to raise_error(ArgumentError, /Height is required/)
      end
    end
    
    context 'document saving' do
      it 'includes image in saved document' do
        doc.add_image(test_image_path, width: 200, height: 150)
        
        # Create a temp file to save to
        temp_file = Tempfile.new(['test', '.docx'])
        begin
          doc.save(temp_file.path)
          
          # Verify the saved file contains the image
          saved_doc = Docx::Document.open(temp_file.path)
          expect(saved_doc).not_to be_nil
        ensure
          temp_file.close
          temp_file.unlink
        end
      end
    end
  end
  
  def fixture_path(filename)
    File.join(File.dirname(__FILE__), '..', 'fixtures', filename)
  end
end