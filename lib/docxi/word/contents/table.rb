module Docxi
  module Word
    module Contents

      class Table

        attr_accessor :options, :rows
        def initialize(options={})
          @options = options
          @rows = []

          if block_given?
            yield self
          else

          end
        end

        def tr(options={}, &block)
          options = @options.merge(options)
          row = TableRow.new(options, &block)
          @rows << row
          row
        end

        def render(xml)
          xml['w'].tbl do
            xml['w'].tblPr do
              xml['w'].tblStyle( 'w:val' => options[:style] ) if options[:style]
              xml['w'].tblW( 'w:w' => options[:width], 'w:type' => "auto" ) if options[:width]
              xml['w'].jc( 'w:val' => options[:align] ) if options[:align]
              xml['w'].tblInd( 'w:w' => options[:iwidth], 'w:type' => "dxa" ) if options[:iwidth]
              xml['w'].tblLook( 'w:val' => "04A0", 'w:firstRow' => 1, 'w:lastRow' => 0, 'w:firstColumn' => 1, 'w:lastColumn' => 0, 'w:noHBand' => 0, 'w:noVBand' => 1 )
              if @options[:borders]
                xml['w'].tblBorders do
                  @options[:borders].each do |k, v|
                    if v.nil?
                      xml['w'].send(k, 'w:val' => "none", 'w:sz' => "0", 'w:space' => "0", 'w:color' => "auto" )
                    else
                      xml['w'].send(k, 'w:val' => "none", 'w:sz' => v, 'w:space' => "0", 'w:color' => "auto" )
                    end
                  end
                end
              end
              if @options[:padding]
                xml['w'].tblCellMar do
                  xml['w'].top( 'w:w' => "288", 'w:type' => "dxa" )
                  xml['w'].left( 'w:w' => "573", 'w:type' => "dxa" )
                  xml['w'].bottom( 'w:w' => "288", 'w:type' => "dxa" )
                  xml['w'].right( 'w:w' => "573", 'w:type' => "dxa" )
                end
              end

            end
            xml['w'].tblGrid do
              if options[:columns_width]
                options[:columns_width].each do |width|
                  xml['w'].gridCol( 'w:w' => width.to_i * 14.92 )
                end
              end
            end
            @rows.each do |row|
              row.render(xml)
            end
          end
        end

        class TableRow

          attr_accessor :options, :cells
          def initialize(options={})
            @options = options
            @cells = []

            if block_given?
              yield self
            else

            end
          end

          def tc(options={}, &block)
            if @options[:columns_width]
              width = @options[:columns_width][@cells.size]
              options[:width] ||= width if width
            end
            options = @options.merge(options)
            cell = TableCell.new(options, &block)
            @cells << cell
            cell
          end

          def render(xml)
            xml['w'].tr do
              if options[:height].present?
                xml['w'].trPr do
                  xml['w'].trHeight('w:val' => options[:height], 'w:hRule' => 'atLeast')
                  xml['w'].cantSplit('w:val' => 'false')
                end
              end
              @cells.each do |cell|
                cell.render(xml)
              end
            end
          end

          class TableCell

            attr_accessor :options, :content
            def initialize(options={})
              @options = options
              @content = []

              if block_given?
                yield self
              else

              end
            end

            def render(xml)
              if options[:fill].blank?
                options[:fill] = 'FFFFFF'
              end
              xml['w'].tc do
                xml['w'].tcPr do
                  xml['w'].tcW( 'w:w' => ( options[:width].to_i * 14.92 ).to_i, 'w:type' => "dxa" ) if options[:width]
                  xml['w'].shd( 'w:val' => "clear", 'w:color' => "auto", 'w:fill' => options[:fill] ) if options[:fill]
                  if options[:borders]
                    xml['w'].tcBorders do
                      options[:borders].each do |k, v|
                        if v.nil?
                          xml['w'].send(k, 'w:val' => "nil" )
                        else
                          xml['w'].send(k, 'w:val' => "single", 'w:sz' => v, 'w:space' => "0", 'w:color' => "auto" )
                        end
                      end
                    end
                  end
                  if options[:merged]
                    xml['w'].vMerge
                  else
                    xml['w'].vAlign( 'w:val' => options[:valign] ) if options[:valign]
                    xml['w'].gridSpan( 'w:val' => options[:colspan] ) if options[:colspan]
                    xml['w'].vMerge( 'w:val' => "restart" ) if options[:rowspan]
                  end
                end
                if options[:merged]
                  xml['w'].p
                else
                  @content.each do |element|
                    element.render(xml)
                  end
                end
              end
            end

            def text(text, options={})
              options = @options.merge(options)
              element = Docxi::Word::Contents::Paragraph.new(options) do |p|
                p.text(text)
              end
              @content << element
              element
            end

            def image(image, options={})
              img = Image.new(image, options)
              @content << img
              img
            end

            def page_numbers
              numbers = PageNumbers.new
              @content << numbers
              numbers
            end

            def p(options={}, &block)
              element = Docxi::Word::Contents::Paragraph.new(options, &block)
              @content << element
              element
            end

            class Image
              attr_accessor :media, :options
              def initialize(media, options={})
                @media = media
                @options = options
              end

              def render(xml)
                xml['w'].p do
                  xml['w'].r do
                    xml['w'].rPr do
                      xml['w'].noProof
                    end
                    xml['w'].drawing do
                      xml['wp'].inline( 'distT' => 0, 'distB' => 0, 'distL' => 0, 'distR' => 0 ) do
                        xml['wp'].extent( 'cx' => ( options[:width] * options[:height] * 14.92 ).to_i, 'cy' => ( options[:width] * options[:height] * 14.92 ).to_i ) if options[:width] && options[:height]
                        xml['wp'].effectExtent( 'l' => 0, 't' => 0, 'r' => 0, 'b' => 0 )
                        xml['wp'].docPr( 'id' => @media.id, 'name'=> "Image", 'descr' => "image")
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
                                xml['pic'].cNvPr( 'id' => 0, 'name' => "Image", 'descr' => "description" )
                                xml['pic'].cNvPicPr do
                                  xml['a'].picLocks( 'noChangeAspect' => "1", 'noChangeArrowheads' => "1" )
                                end
                              end
                              xml['pic'].blipFill do
                                xml['a'].blip( 'r:embed' => @media.sequence ) do
                                  xml['a'].extLst
                                end
                                xml['a'].srcRect
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
                                xml['a'].noFill
                                xml['a'].ln do
                                  xml['a'].noFill
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
            end

            class PageNumbers

              attr_accessor :options
              def initialize(options={})
                @options = options
              end

              def render(xml)
                xml['w'].sdt do
                  xml['w'].sdtPr do
                    xml['w'].id( 'w:val' => "-472213903" )
                    xml['w'].docPartObj do
                      xml['w'].docPartGallery( 'w:val' => "Page Numbers (Bottom of Page)" )
                      xml['w'].docPartUnique
                    end
                  end
                  xml['w'].sdtContent do
                    xml['w'].p do
                      xml['w'].pPr do
                        xml['w'].jc( 'w:val' => @options[:align] || 'right' )
                      end
                      xml['w'].r do
                        xml['w'].rPr do
                          xml['w'].rFonts( 'w:cs'=> 'Arial', 'w:ascii'=> 'Arial', 'w:hAnsi' => 'Arial' )
                          xml['w'].color( 'w:val' => '404040')
                          xml['w'].sz( 'w:val' => '20' )
                        end
                        xml['w'].t 'GlobalOptions   '
                      end
                      xml['w'].r do
                        xml['w'].fldChar( 'w:fldCharType' => "begin" )
                      end
                      xml['w'].r do
                        xml['w'].instrText "PAGE   \* MERGEFORMAT"
                      end
                      xml['w'].r do
                        xml['w'].fldChar( 'w:fldCharType' => "separate" )
                      end
                      xml['w'].r do
                        xml['w'].fldChar( 'w:fldCharType' => "end" )
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
  end
end
