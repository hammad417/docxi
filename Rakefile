#!/usr/bin/env rake
require 'bundler/gem_tasks'
require 'rake/testtask'

Rake::TestTask.new do |t|
  t.libs << 'test'
  t.pattern = 'test/**/test_*.rb'
  t.warning = false
end

desc 'Rewrite test/golden from the current code (only after an intentional output change)'
task :golden do
  $LOAD_PATH.unshift File.expand_path('lib', __dir__)
  require 'docxi'
  require_relative 'test/support/package'
  require_relative 'test/support/sample_documents'

  SampleDocuments.all.each do |name, word|
    Package.write_golden(File.join(__dir__, 'test', 'golden', name), word.render.string)
  end
end

desc 'Write the sample documents to tmp/samples for opening in Word, LibreOffice or Google Docs'
task :samples do
  $LOAD_PATH.unshift File.expand_path('lib', __dir__)
  require 'docxi'
  require 'fileutils'
  require_relative 'test/support/sample_documents'

  FileUtils.mkdir_p('tmp/samples')
  SampleDocuments.all.each do |name, word|
    File.binwrite("tmp/samples/#{name}.docx", word.render.string)
    puts "tmp/samples/#{name}.docx"
  end
end

desc "Run tests"
task :default => :test
