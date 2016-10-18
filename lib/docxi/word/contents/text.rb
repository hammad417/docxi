module Docxi
  module Word
    module Contents

      class Text

        attr_accessor :text, :options
        def initialize(text, options={})
          @text = text
          @options = options
        end

        def render(xml)
          if !@text.nil?
            xml['w'].r do
              xml['w'].rPr do
                xml['w'].rFonts( 'w:cs'=> @options[:font], 'w:ascii'=> @options[:font], 'w:hAnsi' => @options[:font] ) if @options[:font]
                xml['w'].b if @options[:bold]
                xml['w'].i if @options[:italic]
                xml['w'].u( 'w:val' => "single" ) if options[:underline]
                xml['w'].color( 'w:val' => @options[:color] ) if @options[:color]
                xml['w'].sz( 'w:val' => @options[:size].to_i * 2 ) if @options[:size]
                xml['w'].shd( 'w:val' => 'clear','w:fill' => @options[:fill] ) if @options[:fill]
              end
              if @options[:space]
                xml['w'].t( @text, 'xml:space' => "preserve" )
              else
                xml['w'].t @text
              end
            end
            if options[:br]
              br = Docxi::Word::Contents::Break.new
              br.render(xml)
            end
          end
        end
      end
    end
  end
end
