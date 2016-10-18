require 'docxi/word/hyperlinks/hyperlink'

module Docxi
  module Word
    class Hyperlinks

      attr_accessor :hyperlinks, :counter
      def initialize
        @hyperlinks = []
        @counter = 0
      end

      def add(hyperlink)
        @hyperlinks << hyperlink
        hyperlink
      end

      def sequence
        @counter += 1
      end

      def render(zip)
        @hyperlinks.each do |hyperlink|
          hyperlink.render(zip)
        end
      end

    end
  end
end