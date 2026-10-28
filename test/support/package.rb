require 'fileutils'
require 'zlib'
require 'nokogiri'

# Reads a rendered .docx straight from its zip structures, so the tests don't
# depend on whichever rubyzip version is installed.
module Package
  ZIP64_EXTRA_FIELD = 0x0001
  TIMESTAMP = /\d{4}-\d\d-\d\dT\d\d:\d\d:\d\dZ/

  Entry = Struct.new(:name, :versions_needed, :extra_field_ids, :data)

  def self.entries(bytes)
    bytes = bytes.b
    eocd = bytes.rindex("PK\x05\x06".b) or raise 'end of central directory record not found'
    count, _size, offset = bytes[eocd + 10, 10].unpack('vVV')

    Array.new(count) do
      raise 'corrupt central directory' unless bytes[offset, 4] == "PK\x01\x02".b
      version_needed, _flags, compression = bytes[offset + 6, 6].unpack('vvv')
      compressed_size, size = bytes[offset + 20, 8].unpack('VV')
      name_length, extra_length, comment_length = bytes[offset + 28, 6].unpack('vvv')
      local_offset = bytes[offset + 42, 4].unpack1('V')
      name = bytes[offset + 46, name_length].force_encoding('UTF-8')
      central_extra = extra_fields(bytes[offset + 46 + name_length, extra_length])
      offset += 46 + name_length + extra_length + comment_length

      # Zip64 entries store 0xFFFFFFFF here and the real values in their extra field.
      if [size, compressed_size, local_offset].include?(0xFFFFFFFF)
        values = central_extra.assoc(ZIP64_EXTRA_FIELD).last.unpack('Q<*')
        size = values.shift if size == 0xFFFFFFFF
        compressed_size = values.shift if compressed_size == 0xFFFFFFFF
        local_offset = values.shift if local_offset == 0xFFFFFFFF
      end

      local_version_needed = bytes[local_offset + 4, 2].unpack1('v')
      local_name_length, local_extra_length = bytes[local_offset + 26, 4].unpack('vv')
      local_extra = extra_fields(bytes[local_offset + 30 + local_name_length, local_extra_length])
      raw = bytes[local_offset + 30 + local_name_length + local_extra_length, compressed_size]
      data = compression == 8 ? inflate(raw) : raw
      raise "#{name}: expected #{size} bytes, read #{data.bytesize}" unless data.bytesize == size

      Entry.new(name, [version_needed, local_version_needed],
                (central_extra + local_extra).map(&:first), data)
    end
  end

  def self.parts(bytes)
    entries(bytes).each_with_object({}) { |entry, parts| parts[entry.name] = entry.data }
  end

  def self.canonical(xml)
    Nokogiri::XML(xml).canonicalize.gsub(TIMESTAMP, 'TIMESTAMP')
  end

  def self.write_golden(dir, bytes)
    FileUtils.rm_rf(dir)
    entries = entries(bytes)
    entries.each do |entry|
      path = File.join(dir, entry.name)
      FileUtils.mkdir_p(File.dirname(path))
      File.binwrite(path, entry.data)
    end
    File.write(File.join(dir, 'entries.txt'), entries.map(&:name).join("\n") + "\n")
  end

  def self.extra_fields(extra)
    fields = []
    position = 0
    while position + 4 <= extra.bytesize
      id, length = extra[position, 4].unpack('vv')
      fields << [id, extra[position + 4, length]]
      position += 4 + length
    end
    fields
  end

  def self.inflate(raw)
    zstream = Zlib::Inflate.new(-Zlib::MAX_WBITS)
    data = zstream.inflate(raw)
    data << zstream.finish
    zstream.close
    data
  end
end
