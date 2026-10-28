require 'test_helper'
require 'socket'
require 'tmpdir'

class TestMedia < Minitest::Test
  PNG = File.join(SampleDocuments::FILES, 'red.png')

  def image_bytes(source)
    parts = render_parts do |document|
      media = document.add_media(source)
      document.p { |p| p.image media, style: { width: '10pt', height: '10pt' } }
    end
    parts['word/media/image1.png']
  end

  def test_reads_a_file_path
    assert_equal File.binread(PNG), image_bytes(PNG)
  end

  def test_reads_an_io
    assert_equal File.binread(PNG), image_bytes(StringIO.new(File.binread(PNG)))
  end

  def test_downloads_a_url
    png = File.binread(PNG)
    server = TCPServer.new('127.0.0.1', 0)
    thread = Thread.new do
      client = server.accept
      nil while (line = client.gets) && line != "\r\n"
      client.write "HTTP/1.1 200 OK\r\nContent-Type: image/png\r\nContent-Length: #{png.bytesize}\r\nConnection: close\r\n\r\n"
      client.write png
      client.close
    end
    proxies = %w[http_proxy HTTP_PROXY].map { |name| [name, ENV.delete(name)] }

    assert_equal png, image_bytes("http://127.0.0.1:#{server.addr[1]}/red.png")
  ensure
    proxies.each { |name, value| ENV[name] = value if value } if proxies
    thread.join(5) if thread
    server.close if server
  end

  def test_does_not_run_strings_starting_with_a_pipe
    marker = File.join(Dir.tmpdir, "docxi-pipe-test-#{Process.pid}")

    assert_raises(Errno::ENOENT) { image_bytes("|touch #{marker}") }
    refute File.exist?(marker), 'the string was run as a shell command'
  ensure
    File.delete(marker) if marker && File.exist?(marker)
  end
end
