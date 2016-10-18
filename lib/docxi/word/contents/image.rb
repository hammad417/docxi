module Docxi
  module Word
    module Contents

      class Image

        attr_accessor :media, :options
        def initialize(media, options={})
          @media = media
          # @media.file.rewind
          @options = options
        end

        def styles
          if @options[:style]
            @options[:style].collect{|k, v| [k, v].join(':')}.join(';')
          end
        end

        def render(xml)
          xml['w'].r do
            if options[:pH] and options[:pV] and options[:wrap]
              xml['w'].drawing do
                xml['wp'].anchor( "behindDoc" => 1 ,"distT" => 0,"distB" => 0, "distL" => 0, "distR" => 0 ,"simplePos" => 0 ,"locked" => 0, "layoutInCell" => 1, "allowOverlap" => 1, "relativeHeight" => 3) do
                  xml['wp'].simplePos('x'=> 0 , 'y' => 0 )
                  xml['wp'].positionH("relativeFrom" => "column") do
                    xml['wp'].posOffset options[:pH] 
                  end
                  xml['wp'].positionV("relativeFrom" => "paragraph") do
                    xml['wp'].posOffset options[:pV] 
                  end
                  xml['wp'].extent( 'cx' => ( options[:width] * options[:height] * 14.92 ).to_i, 'cy' => ( options[:width] * options[:height] * 14.92 ).to_i ) if options[:width] && options[:height]
                  xml['wp'].effectExtent( 'l' => 0, 't' => 0, 'r' => 0, 'b' => 0 )
                  xml['wp'].wrapNone
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
                          xml['pic'].cNvPicPr do
                            xml['a'].picLocks( 'noChangeAspect' => "1", 'noChangeArrowheads' => "1" )
                          end
                        end
                        xml['pic'].blipFill do
                          xml['a'].blip( 'r:embed' => @media.sequence )
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
                          xml['a'].ln('w' => "9525") do
                            xml['a'].noFill
                            xml['a'].miter('lim'=>"800000")
                            xml['a'].headEnd
                            xml['a'].tailEnd
                          end
                        end
                      end
                    end
                  end
                end
              end
            else  
              xml['w'].pict do
                unless options[:fill_color]
                  xml['v'].rect( 'id' => @media.uniq_id, 'type' => @media.type, 'style' => styles, :stroked => "f" ) do
                    xml['v'].imagedata( 'r:id' => @media.sequence, 'o:title' => @options[:title] )
                    xml['v'].wrap('v:type' => 'none') if options[:wrap]
                  end
                else
                  xml['v'].rect("fillcolor"=> options[:fill_color], 'id' => @media.uniq_id, 'type' => @media.type, 'style' => styles, :stroked => "f" ) do
                    if options[:opacity].present?
                      xml['v'].fill("opacity" => options[:opacity])
                    end
                    xml['v'].imagedata( 'r:id' => @media.sequence, 'o:title' => @options[:title] )
                    xml['v'].wrap('v:type' => 'none') if options[:wrap]
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
