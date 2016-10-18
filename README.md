Docxi gem provides ability to create docx documents using Ruby.

## Features

1. Paragraphs
2. Text formatting
3. Header and Footer (limited)
4. Tables (limited)
5. Images
6. Table of Content
7. Page Numbering

At the moment, the gem supports a limited number of functions, but the development of the gem is still in progress, and perhaps functionality will be expanded in the future. First of all, try to add documentation for the code and tests.
## NOTE
 Use Latest version of Gem. Old Version has bugs.
## Usage

```ruby
# default attributes (top: 720,right: 1440,left: 1440,bottom: 720,footer: 200, header: 190)
word = Docxi::Document.new 
# or change margins
# word = Docxi::Document.new(top: 0,right: 0,left: 0,bottom: 0)

document = word.document

# Creates Header
header = Docxi::Word::Headers::Header.new do |h|
  # Inserts image into header
  header_logo = document.add_media File.open('logo.png')
  h.image header_logo, align: 'right', width: 100, height: 100
end

# Adds header
document.add_header(header)

# Creates Footer
footer = Docxi::Word::Footers::Footer.new do |f|
  # Adds footer text
  f.text "Footer Text", size: 8, align: 'center'
  # Adds page numbers
  f.page_numbers
end
# Adds footer
document.add_footer(footer)

# Creates paragraph
document.p(size: 12, bold: true, align: "center", color: "000000", fill: "AAAAAA") do |p|
  p.text "Title", size: 12, bold: true, align: "center", color: "000000"
  # Adds 3 break lines
  p.br times: 3
end

document.p do |p|
  # Adds text with specific styles
  p.text "Subtitle", size: 14, color: '568AB2', bold: true
end

# Adds break page
document.br(page: true)


# Adds TOC
document.text "Table of Content", size: 18, color: 'D36C8A', bold: true
document.table_of_content(size: 14, bold: true) do |table|
  # Adds TOC element
  table.add_row("Table", 1)
end

# Adds break page
document.br(page: true)

# Adds table
document.table( columns_width: [ 316, 160, 135 ] ) do |table|
  # Adds row with specific styles
  table.tr(fill: "FBD4B4", bold: true, align: 'center' ) do |tr|
    # Adds cells
    tr.tc(borders: { top: 1, right: nil, bottom: 1, left: 1 }) do |tc|
      tc.text "Header 1"
    end
    tr.tc(borders: { top: 1, right: nil, bottom: 1, left: nil }) do |tc|
      tc.text "Header 2"
    end
    tr.tc(borders: { top: 1, right: 1, bottom: 1, left: nil }) do |tc|
      tc.text "Header 3"
    end
  end

  # Adds another row
  table.tr(align: 'center') do |tr|
    # Adds rowspan
    tr.tc(rowspan: true, borders: { top: 1, right: 1, bottom: 1, left: 1 }, valign: 'center') do |tc|
      tc.text "Rowspan", align: 'left'
    end
    tr.tc(borders: { top: 1, right: 1, bottom: 1, left: 1 }) do |tc|
      tc.text "Cell Value 2"
    end
    tr.tc(borders: { top: 1, right: 1, bottom: 1, left: 1 }) do |tc|
      tc.text "Cell Value 3"
    end
  end

  table.tr(fill: "D9D9D9", bold: true, align: 'center' ) do |tr|
    tr.tc(borders: { top: 1, right: 1, bottom: 1, left: 1 }, merged: true) # Use with rowspan to mark cell as merged
    # Adds colspan
    tr.tc(colspan: 2, valign: 'bottom', borders: { top: 1, right: 1, bottom: 1, left: 1 }) do |tc|
      tc.text "Colspan"
    end
  end
end

# Adds table with row height and column indentation
document.table( columns_width: [ 316, 160, 135 ], height: 268, iwidth: 108) do |table|
  # Adds row with specific styles
  table.tr(fill: "FBD4B4", bold: true, align: 'center' ) do |tr|
    # Adds cells
    tr.tc(borders: { top: 1, right: nil, bottom: 1, left: 1 }) do |tc|
      tc.text "Header 1"
    end
    tr.tc(borders: { top: 1, right: nil, bottom: 1, left: nil }) do |tc|
      tc.text "Header 2"
    end
    tr.tc(borders: { top: 1, right: 1, bottom: 1, left: nil }) do |tc|
      tc.text "Header 3"
    end
  end
end


# Adds image
document.p do |p|
  p.text "Image Example", size: 16, bold: true, italic: true, underline: true, br: true

  # Adds image to container
  media = document.add_media File.open('image.png')
  # Or add image URL
  media = document.add_media image_url 
  # Insert media into paragraph
  p.image media, style: { width: '200pt', height: '200pt' }

  # Add Hyperlink
  hyper = document.add_hyperlink("http://google.com")
  p.hyperlink "Go to Google", "http://google.com", hyper

end

# Add Underline 
document.p(bottom: true) do |p|
  p.text "Example text"
end

# Set position of Paragraph
document.p(fill: "d7d7d7",align: "center", right: -1000, left: -1000) do |p| 
  p.text "Example Text 2"
  p.br
end

# Creates Word Document
io = word.render
File.open('doc_1.docx', 'wb'){|f| f.write(io.string) }
```