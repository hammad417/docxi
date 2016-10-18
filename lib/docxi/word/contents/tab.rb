module Docxi
  module Word
    module Contents

      class Tab

        attr_accessor :options
        def initialize(options={})
          @options = options
          @options[:times] ||= 1
        end

        def render(xml)
          if @options[:page]
            xml['w'].p do
              xml['w'].r do
                xml['w'].tab
              end
            end
          else
            @options[:times].times do
              xml['w'].r do
                xml['w'].tab
              end
            end
          end
        end

      end

    end
  end
end
