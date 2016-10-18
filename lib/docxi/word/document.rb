require_relative 'hyperlinks/hyperlink.rb'
require_relative 'hyperlink.rb'
module Docxi
  module Word
    class Document

      include Helpers

      attr_accessor :options, :fonts, :setings, :styles, :numbering, :effects, :web_settings, :themes, :relationships, :footnotes, :endnotes, :footers, :headers, :hyperlinks

      def initialize(options)
        @options = options
        @content = []
      end

      def add_media(file, options={})
        media = Medias::Media.new(medias.sequence, file, options)
        rel = relationships.add(media.content_type, media.target)
        media.set_sequence(rel.id)
        medias.add(media)
      end

      def add_footer(footer, options={})
        footer.id = footers.sequence
        rel = relationships.add(footer.content_type, footer.target)
        footer.sequence = rel.id
        footers.add(footer)
        @footer = footer
      end

      def add_first_footer(footer, options={})
        footer.id = footers.sequence
        rel = relationships.add(footer.content_type, footer.target)
        footer.sequence = rel.id
        footers.add(footer)
        @first_footer = footer
      end

      def add_first_header(header, options={})
        header.id = headers.sequence
        rel = relationships.add(header.content_type, header.target)
        header.sequence = rel.id
        headers.add(header)
        @first_header = header
      end

      def add_header(header, options={})
        header.id = headers.sequence
        rel = relationships.add(header.content_type, header.target)
        header.sequence = rel.id
        headers.add(header)
        @header = header
      end

      def add_hyperlink(link, options={})
        hyperlink = Hyperlinks::Hyperlink.new(hyperlinks.sequence, options)
        rel = relationships.add(hyperlink.content_type, link, 'External')
        hyperlink.set_sequence(rel.id)
        hyperlinks.add(hyperlink)
        return rel.id
      end 

      def fonts
        @fonts ||= Fonts.new
      end

      def settings
        @settings ||= Settings.new
      end

      def styles
        @styles ||= Styles.new
      end

      def numbering
        @numbering ||= Numbering.new
      end

      def effects
        @effects ||= Effects.new
      end

      def web_settings
        @web_settings ||= WebSettings.new({})
      end

      def themes
        @themes ||= Themes.new
      end

      def medias
        @medias ||= Medias.new
      end

      def relationships
        @relationships ||= Relationships.new
      end

      def footnotes
        @footnotes ||= Footnotes.new
      end

      def endnotes
        @endnotes ||= Endnotes.new
      end

      def footers
        @footers ||= Footers.new
      end

      def first_footer
        @first_footer ||= Footers.new
      end

      def first_header
        @first_header ||= Headers.new
      end

      def headers
        @headers ||= Headers.new
      end

      def hyperlinks
        @hyperlinks ||= Hyperlinks.new
      end

      def render(zip)
        zip.put_next_entry('word/document.xml')
        zip.write(Docxi.to_xml(document))

        fonts.render(zip)
        settings.render(zip)
        styles.render(zip)
        effects.render(zip)
        numbering.render(zip)
        web_settings.render(zip)
        themes.render(zip)
        footnotes.render(zip)
        endnotes.render(zip)
        headers.render(zip)
        footers.render(zip)
        medias.render(zip)
        relationships.render(zip)
      end

      private
      def document
        Nokogiri::XML::Builder.with(Nokogiri::XML('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')) do |xml|
          xml.document('xmlns:o' => 'urn:schemas-microsoft-com:office:office', 'xmlns:r' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships' ,'xmlns:v' => 'urn:schemas-microsoft-com:vml', 'xmlns:w' => 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'xmlns:w10' =>'urn:schemas-microsoft-com:office:word', 'xmlns:wp' =>'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing') do
            xml.parent.namespace = xml.parent.namespace_definitions.find{|ns| ns.prefix == "w" }

            xml['w'].body do

              content.each do |element|
                element.render(xml)
              end

              xml['w'].p do
                xml['w'].bookmarkStart( 'w:id' => "0", 'w:name' => "_GoBack")
                xml['w'].bookmarkEnd('w:id' => "0")
              end
              xml['w'].sectPr do
                xml['w'].headerReference( 'w:type' => "default", 'r:id' => @header.sequence ) if @header
                xml['w'].headerReference( 'w:type' => "first", 'r:id' => @first_header.sequence ) if @first_header                
                xml['w'].footerReference( 'w:type' => "default", 'r:id' => @footer.sequence ) if @footer
                xml['w'].footerReference( 'w:type' => "first", 'r:id' => @first_footer.sequence ) if @first_footer
                xml['w'].titlePg
                xml['w'].pgSz( 'w:w' => 12240, 'w:h' => 15840 )
                xml['w'].pgMar( 'w:top' => @options[:top], 'w:right' => @options[:right], 'w:bottom' => @options[:bottom], 'w:left' => @options[:left], 'w:header' => @options[:header].present? ? @options[:header] : 190, 'w:footer' => @options[:footer].present? ? @options[:footer] : 200 , 'w:gutter' => 0)
                xml['w'].cols( 'w:space' => 720)
                xml['w'].docGrid( 'w:linePitch' => "360")
              end
            end
          end
        end
      end

      def content
        @content
      end

    end
  end
end
