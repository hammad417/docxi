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
      Docxi.write_entry(zip, 'docProps/app.xml', Docxi.to_xml(app.document))

      Docxi.write_entry(zip, 'docProps/core.xml', Docxi.to_xml(core.document))
    end

  end
end