module Docxi
  class Document

    attr_accessor :document, :properties, :relationships, :content_types, :options

    def initialize(options = {top: 720,right: 1440,left: 1440,bottom: 720,footer: 200, header: 190})
      @options = options
    end

    def relationships
      @relationships ||= Relationships.new
    end

    def properties
      @properties ||= Properties.new({})
    end

    def content_types
      @content_types ||= ContentTypes.new
    end

    def document
      @document ||= Word::Document.new(options)
    end

    def render
      Zip::ZipOutputStream.write_buffer do |zip|
        document.render(zip)
        properties.render(zip)
        relationships.render(zip)
        content_types.render(zip)
      end
    end

  end
end