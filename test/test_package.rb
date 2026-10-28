require 'test_helper'

class TestPackage < Minitest::Test
  def setup
    @bytes = SampleDocuments.kitchen_sink.render.string
    @parts = Package.parts(@bytes)
  end

  def test_render_returns_a_string_io
    assert_instance_of StringIO, SampleDocuments.kitchen_sink.render
  end

  def test_every_xml_part_is_well_formed
    malformed = @parts.select { |name, _| name.end_with?('.xml', '.rels') }.reject do |_, data|
      begin
        Nokogiri::XML(data) { |config| config.strict }
      rescue Nokogiri::XML::SyntaxError
        false
      end
    end
    assert_empty malformed.keys
  end

  def test_every_internal_relationship_points_to_a_part
    missing = []
    @parts.each do |name, data|
      next unless name.end_with?('.rels')
      base = File.dirname(File.dirname(name))
      xml(data).remove_namespaces!.xpath('//Relationship').each do |relationship|
        next if relationship['TargetMode'] == 'External'
        target = base == '.' ? relationship['Target'] : File.join(base, relationship['Target'])
        missing << "#{name} -> #{relationship['Target']}" unless @parts.key?(target)
      end
    end
    assert_empty missing
  end

  def test_every_part_has_a_content_type
    types = xml(@parts['[Content_Types].xml']).remove_namespaces!
    defaults = types.xpath('//Default').map { |node| node['Extension'].downcase }
    overrides = types.xpath('//Override').map { |node| node['PartName'] }
    untyped = @parts.keys.reject do |name|
      name == '[Content_Types].xml' || overrides.include?("/#{name}") ||
        defaults.include?(File.extname(name).delete('.').downcase)
    end
    assert_empty untyped
  end

  # rubyzip 3 writes Zip64 records unless entry sizes are known up front, and
  # Google Docs and older LibreOffice versions refuse to open such files.
  def test_archive_has_no_zip64_records
    entries = Package.entries(@bytes)
    zip64_versions = entries.reject { |entry| entry.versions_needed.all? { |version| version < 45 } }
    zip64_fields = entries.select { |entry| entry.extra_field_ids.include?(Package::ZIP64_EXTRA_FIELD) }

    assert_empty zip64_versions.map(&:name), 'entries that need zip version 4.5'
    assert_empty zip64_fields.map(&:name), 'entries with a Zip64 extra field'
    eocd = @bytes.b.rindex("PK\x05\x06".b)
    refute_equal "PK\x06\x07".b, @bytes.b[eocd - 20, 4], 'archive has a Zip64 end of central directory locator'
  end
end
