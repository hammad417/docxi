#encoding: utf-8

module Docxi
  module Word
    class Hyperlinks
      class Hyperlink

        attr_accessor :id, :sequence
        def initialize(id,options={})
          @options = options
          @id = id

          if block_given?
            yield self
          else

          end
        end

        def set_sequence(sequence)
          @sequence = sequence
          self
        end

        def content_type
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
        end

        def target
          "document.xml"
        end

        def render(zip)
          zip.put_next_entry("word/#{target}")
          zip.write(Docxi.to_xml(document))
          if !@relationships.empty?
            zip.put_next_entry("word/_rels/#{target}.rels")
            zip.write(Docxi.to_xml(relationships))
          end
        end

        def hyperlink(text,link)
          text = Hyperlink.new(text, link, id)
          @content << text
        end

        class Hyperlink
          attr_accessor :text, :link
          def initialize(text, link )
            @text = text
            @link = link
          end

          def content_type
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
          end

          def render(xml)
            xml['w'].hyperlink('w:id' => '1') do 
              xml['w'].r do
                xml['w'].pPr do
                  xml['w'].rStyle( 'w:val' => "InternetLink")
                end
                xml['w'].t @text
              end
            end
          end
        end


        private
        def document
          Nokogiri::XML::Builder.with(Nokogiri::XML('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')) do |xml|
            xml.hdr( 'xmlns:wpc' => "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas", 'xmlns:mc' => "http://schemas.openxmlformats.org/markup-compatibility/2006", 'xmlns:o' => "urn:schemas-microsoft-com:office:office", 'xmlns:r' => "http://schemas.openxmlformats.org/officeDocument/2006/relationships", 'xmlns:m' => "http://schemas.openxmlformats.org/officeDocument/2006/math", 'xmlns:v' => "urn:schemas-microsoft-com:vml", 'xmlns:wp14' => "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing", 'xmlns:wp' => "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing", 'xmlns:w10' => "urn:schemas-microsoft-com:office:word", 'xmlns:w' => "http://schemas.openxmlformats.org/wordprocessingml/2006/main", 'xmlns:w14' => "http://schemas.microsoft.com/office/word/2010/wordml", 'xmlns:wpg' => "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup", 'xmlns:wpi' => "http://schemas.microsoft.com/office/word/2010/wordprocessingInk", 'xmlns:wne' => "http://schemas.microsoft.com/office/word/2006/wordml", 'xmlns:wps' => "http://schemas.microsoft.com/office/word/2010/wordprocessingShape", 'mc:Ignorable' => "w14 wp14" ) do
              xml.parent.namespace = xml.parent.namespace_definitions.find{|ns| ns.prefix == "w" }
              @content.each do |element|
                element.render(xml)
              end
            end
          end
        end

        def relationships
          Nokogiri::XML::Builder.with(Nokogiri::XML('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')) do |xml|
            xml.Relationships(xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships') do
              @relationships.each do |relationship|
                xml.Relationship('Id' => relationship.sequence, 'Type' => relationship.content_type, 'Target' => relationship.target)
              end
            end
          end
        end

      end
    end
  end
end
