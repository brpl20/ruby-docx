require 'spec_helper'
require 'docx'

describe 'Text Formatting Features' do
  let(:doc) { Docx::Document.open(fixture_path('basic.docx')) }
  
  describe 'TextRun formatting methods' do
    let(:paragraph) { doc.paragraphs.first }
    let(:text_run) { paragraph.text_runs.first }
    
    describe '#bold!' do
      it 'makes text bold' do
        text_run.bold!
        expect(text_run.bolded?).to be true
      end
      
      it 'adds bold property to XML' do
        text_run.bold!
        expect(text_run.node.at_xpath('.//w:b')).not_to be_nil
      end
    end
    
    describe '#italic!' do
      it 'makes text italic' do
        text_run.italic!
        expect(text_run.italicized?).to be true
      end
      
      it 'adds italic property to XML' do
        text_run.italic!
        expect(text_run.node.at_xpath('.//w:i')).not_to be_nil
      end
    end
    
    describe '#underline!' do
      it 'makes text underlined' do
        text_run.underline!
        expect(text_run.underlined?).to be true
      end
      
      it 'adds underline property to XML' do
        text_run.underline!
        expect(text_run.node.at_xpath('.//w:u')).not_to be_nil
      end
    end
    
    describe '#font_size=' do
      it 'sets font size' do
        text_run.font_size = 14
        size_attr = text_run.node.at_xpath('.//w:sz/@w:val')
        expect(size_attr&.value).to eq '28' # Word uses half-points
      end
    end
    
    describe '#color=' do
      it 'sets font color' do
        text_run.color = 'FF0000'
        color_attr = text_run.node.at_xpath('.//w:color/@w:val')
        expect(color_attr&.value).to eq 'FF0000'
      end
      
      it 'removes # from hex color' do
        text_run.color = '#FF0000'
        color_attr = text_run.node.at_xpath('.//w:color/@w:val')
        expect(color_attr&.value).to eq 'FF0000'
      end
    end
  end
  
  describe 'Paragraph formatting methods' do
    let(:paragraph) { doc.paragraphs.first }
    
    describe '#add_text' do
      it 'adds a new text run with content' do
        initial_count = paragraph.text_runs.count
        paragraph.add_text('New text')
        expect(paragraph.text_runs.count).to eq initial_count + 1
        expect(paragraph.text_runs.last.text).to eq 'New text'
      end
      
      it 'applies formatting options' do
        run = paragraph.add_text('Bold text', bold: true, italic: true, size: 16)
        expect(run.bolded?).to be true
        expect(run.italicized?).to be true
      end
    end
    
    describe '#add_bold_text' do
      it 'adds bold text to paragraph' do
        run = paragraph.add_bold_text('Bold content')
        expect(run.bolded?).to be true
        expect(run.text).to eq 'Bold content'
      end
    end
    
    describe '#add_italic_text' do
      it 'adds italic text to paragraph' do
        run = paragraph.add_italic_text('Italic content')
        expect(run.italicized?).to be true
        expect(run.text).to eq 'Italic content'
      end
    end
  end
  
  describe 'Document-level methods' do
    describe '#add_paragraph' do
      it 'adds a new paragraph to the document' do
        initial_count = doc.paragraphs.count
        doc.add_paragraph('New paragraph')
        expect(doc.paragraphs.count).to eq initial_count + 1
        expect(doc.paragraphs.last.to_s).to eq 'New paragraph'
      end
      
      it 'applies formatting to paragraph text' do
        para = doc.add_paragraph('Formatted text', bold: true, italic: true)
        run = para.text_runs.first
        expect(run.bolded?).to be true
        expect(run.italicized?).to be true
      end
    end
    
    describe '#add_bold_paragraph' do
      it 'adds a paragraph with bold text' do
        para = doc.add_bold_paragraph('Bold paragraph')
        expect(para.text_runs.first.bolded?).to be true
      end
    end
    
    describe '#add_italic_paragraph' do
      it 'adds a paragraph with italic text' do
        para = doc.add_italic_paragraph('Italic paragraph')
        expect(para.text_runs.first.italicized?).to be true
      end
    end
  end
  
  def fixture_path(filename)
    File.join(File.dirname(__FILE__), '..', 'fixtures', filename)
  end
end