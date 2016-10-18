module Docxi
  module Word
    module Contents

      class Frame

        attr_accessor :content, :options
        def initialize(options={})
          @content = []
          @options = options
          if block_given?
            yield self
          else

          end
        end

        def render(xml)
          xml['w'].r do
            xml['w'].pict do
              xml['v'].rect("fillcolor"=>"#bfbfbf[2412]", "strokecolor"=>"000000","strokeweight"=>"0pt", "style"=> @options[:style]) do
                xml['v'].fill("opacity" =>"52429f")
                xml['v'].textbox("inset" => "0in,0in,0in,0in") do 
                  xml['w'].txbxContent do
                    xml['w'].p do
                      xml['w'].pPr do
                      xml['w'].pStyle('w:val'=>"Normal")
                        xml['w'].pBdr do
                          xml['w'].top('w:val'=>"nil")
                          xml['w'].left('w:val'=>"nil")
                          xml['w'].bottom('w:val'=>"nil")
                          xml['w'].right('w:val'=>"nil")
                        end
                        if @options[:spacing]
                          xml['w'].spacing('w:before'=>"240")
                        end
                      end
                      @content.each do |element|
                        element.render(xml)
                      end
                    end
                  end
                end
              xml['w10'].wrap('type'=>"topAndBottom")
              end
            end
          end
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

        def frame(options={}, &block)
          element = Docxi::Word::Contents::Frame.new(options, &block)
          @content << element
          element
        end

        def text(text, options={})
          options = @options.merge(options)
          text = Docxi::Word::Contents::Text.new(text, options)
          @content << text
          text
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
