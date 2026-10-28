require 'nokogiri'
require 'zip'

require 'docxi/version'
require 'docxi/util'
require 'docxi/relationships'
require 'docxi/content_types'
require 'docxi/properties'
require 'docxi/word'
require 'docxi/document'

module Docxi

  def self.to_xml(document)
    document.to_xml(:save_with => Nokogiri::XML::Node::SaveOptions::AS_XML).gsub("\n", "\r\n").strip
  end

  # The size is set up front because rubyzip 3 otherwise adds Zip64 extra
  # fields to every entry, and Google Docs and older LibreOffice reject those.
  def self.write_entry(zip, name, data)
    entry = Zip::Entry.new('', name)
    entry.size = data.bytesize
    zip.put_next_entry(entry)
    zip.write(data)
  end

  def self.test
    word = Docxi::Document.new

    word.document.break
    word.document.break
    word.document.break

    file = word.render
    File.open('test.docx', 'wb') do |f|
      f.write(file.string)
    end
  end

end
