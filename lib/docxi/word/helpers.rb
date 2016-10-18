module Docxi
  module Word
    module Helpers

      def text(text, options={})
        p(options) do |p|
          p.text text
        end
      end

      def frame(text, options={})
        p(options) do |p|
          p.frame text
        end
      end

      def br(options={})
        element = Docxi::Word::Contents::Break.new(options)
        @content << element
        element
      end

      def tab(options={})
        element = Docxi::Word::Contents::Tab.new(options)
        @content << element
        element
      end

      def p(options={}, &block)
        element = Docxi::Word::Contents::Paragraph.new(options, &block)
        @content << element
        element
      end

      def table_of_content(options={}, &block)
        toc = Docxi::Word::Contents::TableOfContent.new(options, &block)
        @content << toc
        toc
      end

      def table(options={}, &block)
        table = Docxi::Word::Contents::Table.new(options, &block)
        @content << table
        table
      end

    end
  end
end