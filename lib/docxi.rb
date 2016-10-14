require 'nokogiri'
require 'zip/zip'

require 'docxi/relationships'
require 'docxi/content_types'
require 'docxi/properties'
require 'docxi/word'
require 'docxi/document'

module Docxi

  def self.to_xml(document)
    document.to_xml(:save_with => Nokogiri::XML::Node::SaveOptions::AS_XML).gsub("\n", "\r\n").strip
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
