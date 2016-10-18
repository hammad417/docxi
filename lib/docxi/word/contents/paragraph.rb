module Docxi
  module Word
    module Contents

      class Paragraph

        attr_accessor :content, :options, :relationships, :id
        def initialize(options={})
          @content = []
          @options = options
          @relationships = []

          if block_given?
            yield self
          else

          end
        end

        def content_type
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
        end

        def target
          "document.xml"
        end


        def hyperlink(text,link,id)
          text = Hyperlink.new(text, link, id)
          @content << text
        end

        class Hyperlink
          attr_accessor :text, :link, :id 
          def initialize(text, link, id )
            @text = text
            @link = link
            @id = id
          end

          def content_type
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
          end

          def render(xml)
            xml['w'].pPr do
              xml['w'].pStyle('w:val' => "Normal")
              xml['w'].rPr do 
                xml['w'].rStyle('w:val' => "InternetLink")
              end
            end
            xml['w'].hyperlink('r:id' => @id) do 
              xml['w'].r do
                xml['w'].rPr do
                  xml['w'].rStyle( 'w:val' => "InternetLink")
                end
                xml['w'].t @text
              end
            end
          end
        end
        
        def render(xml)
          xml['w'].p do
            xml['w'].pPr do
              xml['w'].jc( 'w:val' => @options[:align]) if @options[:align]
              xml['w'].shd( 'w:val' => 'clear','w:fill' => @options[:fill] ) if @options[:fill]
              xml['w'].ind( 'w:left' => @options[:left], 'w:right' => @options[:right] ) if @options[:left] and @options[:right]
              if options[:ul]
                xml['w'].numPr do
                  xml['w'].ilvl( 'w:val' => 0 )
                  xml['w'].numId( 'w:val' => 1 )
                end
              end
              if options[:bottom]
                xml['w'].pBdr do
                  xml['w'].top( 'w:val' => "nil")
                  xml['w'].left('w:val' => "nil")
                  xml['w'].bottom('w:val'=>"single", 'w:sz'=>"4", 'w:space'=>"1", 'w:color'=>"000000")
                  xml['w'].right('w:val'=>"nil")
                end
              end
            end
            @content.each do |element|
              element.render(xml)
            end
          end
        end

        def text(text, options={})
          options = @options.merge(options)
          text = Docxi::Word::Contents::Text.new(text, options)
          @content << text
          text
        end

        def frame(options={}, &block)
          element = Docxi::Word::Contents::Frame.new(options, &block)
          @content << element
          element
        end

        def br(options={})
          br = Docxi::Word::Contents::Break.new(options)
          @content << br
          br
        end

        def tab(options={})
          tab = Docxi::Word::Contents::Tab.new(options)
          @content << tab
          tab
        end

        def image(image, options={})
          img = Docxi::Word::Contents::Image.new(image, options)
          @content << img
          img
        end

      end

    end
  end
end
