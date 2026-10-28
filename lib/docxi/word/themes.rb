require 'docxi/word/themes/theme'

module Docxi
  module Word
    class Themes

      attr_accessor :themes, :counter
      def initialize
        @themes = []
        @counter = 0
        
        add(Theme.new)
      end

      def add(theme)
        @counter += 1
        @themes << theme.set_sequence(@counter)
        theme
      end

      def render(zip)
        @themes.each do |theme|
          Docxi.write_entry(zip, "word/theme/theme#{theme.sequence}.xml", theme.render)
        end
      end

    end
  end
end