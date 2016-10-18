#encoding: utf-8

module Docxi
  module Word
    class Headers
      class Header

        attr_accessor :id, :sequence, :options, :content, :relationships
        def initialize(options={})
          @options = options
          @content = []
          @relationships = []

          if block_given?
            yield self
          else

          end
        end

        def content_type
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
        end

        def target
          "header#{id}.xml"
        end

        def render(zip)
          zip.put_next_entry("word/#{target}")
          zip.write(Docxi.to_xml(document))

          if !@relationships.empty?
            zip.put_next_entry("word/_rels/#{target}.rels")
            zip.write(Docxi.to_xml(relationships))
          end
        end

        def image(image, options={})
          img = Image.new(image, options)
          @content << img
          @relationships << image
          img
        end

        def text(text, options={})
          options = @options.merge(options)
          element = Docxi::Word::Contents::Paragraph.new(options) do |p|
            p.text(text)
          end
          @content << element
          element
        end

        def page_numbers
          numbers = PageNumbers.new
          @content << numbers
          numbers
        end

        def br(options={})
          br = Docxi::Word::Contents::Break.new(options)
          @content << br
          br
        end

        def table(options={}, &block)
          table = Docxi::Word::Contents::Table.new(options, &block)
          @content << table
          table
        end

        class Image
          attr_accessor :media, :options
          def initialize(media, options={})
            @media = media
            # @media.file.rewind
            @options = options
          end

          def width
            @options[:width]
          end

          def height
            @options[:height]
          end

          def render(xml)
            xml['w'].p do
              xml['w'].pPr do
                xml['w'].jc( 'w:val' => @options[:align] ) if @options[:align]
              end
              xml['w'].r do
                xml['w'].rPr do
                  xml['w'].noProof
                end
                xml['w'].drawing do
                  xml['wp'].anchor( "behindDoc" => options[:behind] ,"distT" => 0,"distB" => 0, "distL" => 0, "distR" => 0 ,"simplePos" => 0 ,"locked" => 0, "layoutInCell" => 1, "allowOverlap" => 1, "relativeHeight" => 4) do
                    xml['wp'].simplePos('x'=> 0 , 'y' => 0 )
                    xml['wp'].positionH("relativeFrom" => "column") do
                      xml['wp'].posOffset options[:pH] 
                    end
                    xml['wp'].positionV("relativeFrom" => "paragraph") do
                      xml['wp'].posOffset options[:pV] 
                    end
                    xml['wp'].extent( 'cx' => ( options[:width] * options[:height] * 14.92 ).to_i, 'cy' => ( options[:width] * options[:height] * 14.92 ).to_i ) if options[:width] && options[:height]
                    xml['wp'].effectExtent( 'l' => 0, 't' => 0, 'r' => 0, 'b' => 0 )
                    if options[:wrap] == "none"
                      xml['wp'].wrapNone
                    else
                      xml['wp'].wrapSquare("wrapText"=> "bothSides")
                    end
                    xml['wp'].docPr( 'id' => 1, 'name'=> "Image", 'descr' => "image")
                    xml['wp'].cNvGraphicFramePr do
                      xml.graphicFrameLocks( 'xmlns:a' => "http://schemas.openxmlformats.org/drawingml/2006/main", 'noChangeAspect' => "1" ) do
                        xml.parent.namespace = xml.parent.namespace_definitions.find{|ns| ns.prefix == "a" }
                      end
                    end
                    xml.graphic( 'xmlns:a' => "http://schemas.openxmlformats.org/drawingml/2006/main" ) do
                      xml.parent.namespace = xml.parent.namespace_definitions.find{|ns| ns.prefix == "a" }
                      xml['a'].graphicData( 'uri' => "http://schemas.openxmlformats.org/drawingml/2006/picture") do
                        xml.pic( 'xmlns:pic' => "http://schemas.openxmlformats.org/drawingml/2006/picture") do
                          xml.parent.namespace = xml.parent.namespace_definitions.find{|ns| ns.prefix == "pic" }
                          xml['pic'].nvPicPr do
                            xml['pic'].cNvPr( 'id' => 0, 'name' => "Image" )
                            xml['pic'].cNvPicPr 
                          end
                          xml['pic'].blipFill do
                            xml['a'].blip( 'r:embed' => @media.sequence ) do
                              xml['a'].extLst
                            end
                            xml['a'].stretch do
                              xml['a'].fillRect
                            end
                          end
                          xml['pic'].spPr( 'bwMode' => "auto" ) do
                            xml['a'].xfrm do
                              xml['a'].off( 'x' => 0, 'y' => 0 )
                              xml['a'].ext( 'cx' => ( options[:width] * options[:height] * 14.92 ).to_i, 'cy' => ( options[:width] * options[:height] * 14.92).to_i ) if options[:width] && options[:height]
                            end
                            xml['a'].prstGeom( 'prst' => "rect" ) do
                              xml['a'].avLst
                            end
                            # xml['a'].noFill
                            # xml['a'].ln('w' => "9525") do
                            #   xml['a'].noFill
                            #   xml['a'].miter('lim'=>"800000")
                            #   xml['a'].headEnd
                            #   xml['a'].tailEnd
                            # end
                          end
                        end
                      end
                    end
                  end
                end
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
