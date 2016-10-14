require 'docxi/properties/app'
require 'docxi/properties/core'

module Docxi
  class Properties
    attr_accessor :options, :app, :core

    def initialize(options)
      @options = options
    end

    def app
      @app ||= App.new(@options)
    end

    def core
      @core ||= Core.new(@options)
    end

    def render(zip)
      zip.put_next_entry('docProps/app.xml')
      zip.write(Docxi.to_xml(app.document))

      zip.put_next_entry('docProps/core.xml')
      zip.write(Docxi.to_xml(core.document))
    end

  end
end