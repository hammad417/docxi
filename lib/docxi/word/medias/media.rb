module Docxi
  module Word
    class Medias
      class Media

        URL = %r{\A(?:https?|ftp)://}i

        attr_accessor :id, :sequence, :file
        def initialize(id, file, options)
          @id = id
          @file = file
          self
        end

        # Strings are URLs or file paths. Kernel#open, and File.binread on
        # Ruby 2.5, run strings starting with "|" as shell commands, and
        # Kernel#open stopped fetching URLs in Ruby 3.
        def read
          if @file.is_a?(String)
            if @file =~ URL
              require 'open-uri'
              URI.parse(@file).open { |io| io.read }
            else
              File.open(@file, 'rb') { |file| file.read }
            end
          else
            @file.read
          end
        end

        def set_sequence(sequence)
          @sequence = sequence
          self
        end

        def type
          "#_x0000_t75"
        end

        def target
          "media/image#{id}.png"
        end

        def content_type
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
        end

        def uniq_id
          "_x0000_i102#{id}"
        end

      end
    end
  end
end