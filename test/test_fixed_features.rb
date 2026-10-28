require 'test_helper'

# These raised NameError or TypeError in 0.0.6.
class TestFixedFeatures < Minitest::Test
  def test_tab_inside_a_paragraph
    parts = render_parts { |document| document.p { |p| p.text 'a'; p.tab times: 2; p.text 'b' } }

    assert_equal 2, xml(parts['word/document.xml']).xpath('//w:p/w:r/w:tab', 'w' => W).size
  end

  def test_tab_as_its_own_paragraph
    parts = render_parts { |document| document.tab(page: true) }

    assert_equal 1, xml(parts['word/document.xml']).xpath('//w:body/w:p/w:r/w:tab', 'w' => W).size
  end

  def test_frame_shortcut_with_text
    parts = render_parts { |document| document.frame 'Framed', style: 'width:100pt;height:20pt', bold: true }
    doc = xml(parts['word/document.xml'])

    assert_equal 'width:100pt;height:20pt', doc.at_xpath('//v:rect', 'v' => V)['style']
    assert_equal 'Framed', doc.at_xpath('//v:textbox//w:t', 'v' => V, 'w' => W).text
    assert doc.at_xpath('//v:textbox//w:r/w:rPr/w:b', 'v' => V, 'w' => W), 'frame text options were not applied'
  end

  def test_frame_block_inside_a_paragraph
    parts = render_parts do |document|
      document.p do |p|
        p.frame(style: 'width:50pt') do |frame|
          frame.text 'Inside'
          frame.tab
          frame.br
        end
      end
    end
    doc = xml(parts['word/document.xml'])

    assert_equal 'Inside', doc.at_xpath('//v:textbox//w:t', 'v' => V, 'w' => W).text
    assert doc.at_xpath('//v:textbox//w:tab', 'v' => V, 'w' => W)
  end

  def test_header_page_numbers
    parts = render_parts do |document|
      document.add_header(Docxi::Word::Headers::Header.new { |h| h.page_numbers name: 'Page', align: 'center' })
    end
    header = xml(parts['word/header1.xml'])

    assert_includes header.xpath('//w:instrText', 'w' => W).text, 'PAGE'
    assert_equal 'Page   ', header.at_xpath('//w:sdtContent//w:t', 'w' => W).text
    assert_equal 'center', header.at_xpath('//w:sdtContent//w:jc', 'w' => W)['w:val']
  end
end
