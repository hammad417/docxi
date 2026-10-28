module SampleDocuments
  FILES = File.expand_path('../fixtures/files', __dir__)

  def self.image(name = 'red.png')
    File.open(File.join(FILES, name), 'rb')
  end

  def self.all
    { 'kitchen_sink' => kitchen_sink, 'blank_options' => blank_options }
  end

  # Uses every feature that works in 0.0.6, starting from the README example.
  def self.kitchen_sink
    word = Docxi::Document.new
    document = word.document

    header = Docxi::Word::Headers::Header.new do |h|
      logo = document.add_media(image)
      h.image logo, align: 'right', width: 100, height: 100
      h.text 'Header text', size: 9, color: '999999'
      h.br
      h.table(columns_width: [200, 200]) do |table|
        table.tr do |tr|
          tr.tc { |tc| tc.text 'Header cell 1' }
          tr.tc { |tc| tc.text 'Header cell 2' }
        end
      end
    end
    document.add_header(header)

    first_header = Docxi::Word::Headers::Header.new do |h|
      h.text 'First page header', align: 'center'
    end
    document.add_first_header(first_header)

    footer = Docxi::Word::Footers::Footer.new do |f|
      f.text 'Footer Text', size: 8, align: 'center'
      f.page_numbers
      f.page_numbers name: 'Page', align: 'center'
      f.br
      f.table(columns_width: [300]) do |table|
        table.tr { |tr| tr.tc { |tc| tc.text 'Footer cell' } }
      end
    end
    document.add_footer(footer)

    first_footer = Docxi::Word::Footers::Footer.new do |f|
      f.text 'First page footer'
    end
    document.add_first_footer(first_footer)

    document.p(size: 12, bold: true, align: 'center', color: '000000', fill: 'AAAAAA') do |p|
      p.text 'Title', size: 12, bold: true, align: 'center', color: '000000'
      p.br times: 3
    end

    document.p do |p|
      p.text 'Subtitle', size: 14, color: '568AB2', bold: true
      p.text ' with spaces ', space: true, italic: true, underline: true
      p.text 'Arial on yellow', font: 'Arial', fill: 'FFFF00', br: true
      p.text nil
    end

    document.br(page: true)

    document.text 'Table of Content', size: 18, color: 'D36C8A', bold: true
    document.table_of_content(size: 14, bold: true) do |toc|
      toc.add_row('Table', 1)
      toc.add_row('Images', 2, leader: 'hyphen', pos: 8000)
      toc.add_row('Appendix')
    end

    document.br(page: true)

    document.table(columns_width: [316, 160, 135]) do |table|
      table.tr(fill: 'FBD4B4', bold: true, align: 'center') do |tr|
        tr.tc(borders: { top: 1, right: nil, bottom: 1, left: 1 }) { |tc| tc.text 'Header 1' }
        tr.tc(borders: { top: 1, right: nil, bottom: 1, left: nil }) { |tc| tc.text 'Header 2' }
        tr.tc(borders: { top: 1, right: 1, bottom: 1, left: nil }) { |tc| tc.text 'Header 3' }
      end
      table.tr(align: 'center') do |tr|
        tr.tc(rowspan: true, borders: { top: 1, right: 1, bottom: 1, left: 1 }, valign: 'center') do |tc|
          tc.text 'Rowspan', align: 'left'
        end
        tr.tc(borders: { top: 1, right: 1, bottom: 1, left: 1 }) { |tc| tc.text 'Cell Value 2' }
        tr.tc(borders: { top: 1, right: 1, bottom: 1, left: 1 }) { |tc| tc.text 'Cell Value 3' }
      end
      table.tr(fill: 'D9D9D9', bold: true, align: 'center') do |tr|
        tr.tc(borders: { top: 1, right: 1, bottom: 1, left: 1 }, merged: true)
        tr.tc(colspan: 2, valign: 'bottom', borders: { top: 1, right: 1, bottom: 1, left: 1 }) do |tc|
          tc.text 'Colspan'
        end
      end
    end

    document.table(columns_width: [316, 160, 135], height: 268, iwidth: 108) do |table|
      table.tr(fill: 'FBD4B4', bold: true, align: 'center') do |tr|
        tr.tc { |tc| tc.text 'Header 1' }
        tr.tc { |tc| tc.text 'Header 2' }
        tr.tc { |tc| tc.text 'Header 3' }
      end
    end

    cell_media = document.add_media(image)
    document.table(style: 'TableGrid', width: 5000, align: 'center', padding: true,
                   borders: { top: 4, left: nil, bottom: 4, right: nil, insideH: 2, insideV: nil }) do |table|
      table.tr(height: 400) do |tr|
        tr.tc(valign: 'center') do |tc|
          tc.p(align: 'right') { |p| p.text 'Paragraph in a cell' }
          tc.image cell_media, width: 20, height: 20
          tc.page_numbers
        end
      end
    end

    document.p do |p|
      p.text 'Image Example', size: 16, bold: true, italic: true, underline: true, br: true
      media = document.add_media(image)
      p.image media, style: { width: '200pt', height: '200pt' }, title: 'Red square'
      p.image media, style: { width: '50pt', height: '50pt' }, fill_color: '#00FF00', opacity: '50%', wrap: true
      p.image media, pH: 1000, pV: 2000, wrap: true, width: 50, height: 50
      hyper = document.add_hyperlink('http://google.com')
      p.hyperlink 'Go to Google', 'http://google.com', hyper
    end

    document.p(bottom: true) { |p| p.text 'Example text' }

    document.p(fill: 'd7d7d7', align: 'center', right: -1000, left: -1000) do |p|
      p.text 'Example Text 2'
      p.br
    end

    document.p(ul: true) { |p| p.text 'Bullet item' }

    word
  end

  # Option values that used to go through ActiveSupport's blank?/present?:
  # empty and whitespace-only strings, nil, false and 0.
  def self.blank_options
    word = Docxi::Document.new(top: 0, right: 0, left: 0, bottom: 0, header: '', footer: " \t")
    document = word.document

    document.table(columns_width: [100, 100, 100]) do |table|
      table.tr(height: '') do |tr|
        tr.tc(fill: '') { |tc| tc.text 'empty fill' }
        tr.tc(fill: "  \t") { |tc| tc.text 'whitespace fill' }
        tr.tc(fill: nil) { |tc| tc.text 'nil fill' }
      end
      table.tr(height: 300) do |tr|
        tr.tc(fill: 'ABCDEF') { |tc| tc.text 'with fill' }
        tr.tc { |tc| tc.text 'no fill option' }
        tr.tc(fill: false) { |tc| tc.text 'false fill' }
      end
      table.tr(height: 0) do |tr|
        tr.tc(fill: []) { |tc| tc.text 'empty array fill' }
      end
    end

    media = document.add_media(image)
    document.p do |p|
      p.image media, style: { width: '10pt', height: '10pt' }, fill_color: '#000000', opacity: ''
      p.image media, style: { width: '10pt', height: '10pt' }, fill_color: '#000000', opacity: '  '
      p.image media, style: { width: '10pt', height: '10pt' }, fill_color: '#000000', opacity: 0
    end

    hyper = document.add_hyperlink('https://example.com/a?b=c&d=e')
    document.p { |p| p.hyperlink 'Escaped link', 'https://example.com/a?b=c&d=e', hyper }

    word
  end
end
