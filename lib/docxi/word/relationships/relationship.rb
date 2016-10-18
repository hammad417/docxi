module Docxi
  module Word
    class Relationships
      class Relationship

        attr_accessor :id, :type, :target, :tm

        def initialize(id, type, target, tm = '')
          @id = id
          @type = type
          @target = target
          @tm = tm
        end

        def build(xml)
          if @tm.blank?
            xml.Relationship('Id' => @id, 'Type' => @type, 'Target' => @target)
          else
            xml.Relationship('Id' => @id, 'Type' => @type, 'Target' => @target, 'TargetMode' => @tm)
          end
        end

      end
    end
  end
end