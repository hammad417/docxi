#!/usr/bin/env ruby
# Builds the GitHub wiki from the wiki/ source folder into flat, published
# page files, then checks every link against the published page names.
#
#   ruby wiki/scripts/publish.rb --check           # build into a temporary folder and validate only
#   ruby wiki/scripts/publish.rb --out ../docxi.wiki  # write the pages into a clone of the wiki repository
#   ruby wiki/scripts/publish.rb --out ../docxi.wiki --prune  # also delete .md pages the build no longer produces
#
# The script never runs git. Commit and push the wiki clone yourself after
# looking over the changes (see wiki/wiki-publishing-guide.md).

require "fileutils"
require "tmpdir"

WIKI = File.expand_path("..", __dir__)
REPO = File.expand_path("..", WIKI)
REPO_URL = "https://github.com/hammad417/docxi/blob/master/".freeze

LANGUAGES = [
  ["zh-CN", "简体中文"], ["zh-TW", "繁體中文"], ["es", "Español"], ["ar", "العربية"],
  ["fr", "Français"], ["pt-BR", "Português (Brasil)"], ["de", "Deutsch"], ["hi", "हिन्दी"],
  ["ja", "日本語"], ["ko", "한국어"], ["ru", "Русский"], ["id", "Bahasa Indonesia"],
  ["tr", "Türkçe"], ["it", "Italiano"], ["vi", "Tiếng Việt"]
].freeze

NOT_PAGES = %w[README.md TRANSLATION_GUIDE.md wiki-publishing-guide.md _Sidebar.md _Footer.md].freeze
PAGES = Dir.children(WIKI).select { |f| f.end_with?(".md") && !NOT_PAGES.include?(f) }.map { |f| f.delete_suffix(".md") }.sort.freeze
LINK = /\]\(([^)\s]+)\)/

def published_name(page, code)
  code ? "#{page}-#{code}" : page
end

# A language is published only when every English page has a translation.
def language_status
  LANGUAGES.map do |code, name|
    dir = File.join(WIKI, "i18n", code)
    present = PAGES.select { |page| File.exist?(File.join(dir, "#{page}.md")) }
    [code, name, present.size == PAGES.size, present.size]
  end
end

def switcher(page, current, complete)
  entries = [[nil, "English"]] + complete.map { |code, name| [code, name] }
  return "" if entries.size == 1
  entries.map { |code, name| code == current ? "**#{name}**" : "[#{name}](#{published_name(page, code)})" }.join(" · ") + "\n\n"
end

# Inside a translation, links to English page names point to the same language.
def localize_links(text, code)
  text.gsub(LINK) do
    target = Regexp.last_match(1)
    page, anchor = target.split("#", 2)
    PAGES.include?(page) ? "](#{published_name(page, code)}#{anchor ? "##{anchor}" : ""})" : "](#{target})"
  end
end

def build(out)
  complete = language_status.select { |_, _, done, _| done }.map { |code, name, _, _| [code, name] }
  written = []

  PAGES.each do |page|
    source = File.read(File.join(WIKI, "#{page}.md"))
    File.write(File.join(out, "#{page}.md"), switcher(page, nil, complete) + source)
    written << "#{page}.md"
    complete.each do |code, _|
      text = File.read(File.join(WIKI, "i18n", code, "#{page}.md"))
      File.write(File.join(out, "#{published_name(page, code)}.md"), switcher(page, code, complete) + localize_links(text, code))
      written << "#{published_name(page, code)}.md"
    end
  end

  languages = (["- [English](Home)"] + complete.map { |code, name| "- [#{name}](#{published_name("Home", code)})" }).join("\n")
  File.write(File.join(out, "_Sidebar.md"), File.read(File.join(WIKI, "_Sidebar.md")).sub("<!-- languages -->", languages))
  FileUtils.cp(File.join(WIKI, "_Footer.md"), File.join(out, "_Footer.md"))
  written + %w[_Sidebar.md _Footer.md]
end

def slug(heading)
  heading.strip.downcase.gsub(/<[^>]+>/, "").gsub(/[^\p{Word}\s-]/u, "").gsub(/\s/, "-")
end

def anchors(text)
  prose = text.gsub(/^```.*?^```/m, "")
  explicit = prose.scan(/<a id="([^"]+)"><\/a>/).flatten
  headings = prose.scan(/^#+ (.+)$/).flatten.map { |h| slug(h) }
  explicit + headings
end

# Inline code is blanked out unless it is the whole text of a link. Spans are
# matched left to right, so a closing backtick is never paired with the next
# span's opening one.
def without_code(text)
  text.gsub(/`[^`\n]+`/) do
    m = Regexp.last_match
    link_text = m.begin(0).positive? && text[m.begin(0) - 1] == "[" && text[m.end(0), 2] == "]("
    link_text ? m[0] : "``"
  end
end

def validate(out, written)
  problems = []
  names = written.map { |f| f.delete_suffix(".md") }

  bad_names = written.reject { |f| f =~ /\A(_Sidebar|_Footer|[A-Za-z][A-Za-z0-9-]*)\.md\z/ }
  problems << "file names GitHub can't use: #{bad_names}" unless bad_names.empty?
  dupes = names.group_by(&:downcase).select { |_, v| v.size > 1 }.keys
  problems << "page names that collide: #{dupes}" unless dupes.empty?

  written.each do |file|
    text = File.read(File.join(out, file))
    # Case-sensitive, because "todo" is an ordinary word in Spanish and Portuguese.
    problems << "#{file}: placeholder text" if text =~ /\b(TODO|TBD|FIXME)\b/ || text =~ /lorem ipsum/i
    problems << "#{file}: unbalanced code fences" if text.scan(/^```/).size.odd?
    prose = without_code(text.gsub(/^```.*?^```/m, ""))
    prose.scan(LINK).flatten.each do |target|
      if target.start_with?(REPO_URL)
        path = target.delete_prefix(REPO_URL).split("#").first
        problems << "#{file}: repository file not found: #{path}" unless File.exist?(File.join(REPO, path))
      elsif target =~ %r{\A(https?|mailto):}
        next
      elsif target.start_with?("#")
        problems << "#{file}: missing anchor #{target}" unless anchors(text).include?(target[1..-1])
      else
        page, anchor = target.split("#", 2)
        if !names.include?(page)
          problems << "#{file}: link to missing page #{target}"
        elsif anchor && !anchors(File.read(File.join(out, "#{page}.md"))).include?(anchor)
          problems << "#{file}: missing anchor #{target}"
        end
      end
    end
  end
  problems
end

# A translation is stale when its line-1 comment names a different version of the English page.
def stale_translations
  LANGUAGES.flat_map do |code, _|
    PAGES.map do |page|
      path = File.join(WIKI, "i18n", code, "#{page}.md")
      next unless File.exist?(path)
      source = `git -C #{REPO} hash-object #{File.join(WIKI, "#{page}.md")}`.strip
      recorded = File.foreach(path).first.to_s[/git blob ([0-9a-f]{40})/, 1]
      "#{code}/#{page}.md (translated from #{recorded || "unknown"}, English is now #{source[0, 10]})" unless recorded == source
    end.compact
  end
end

check = ARGV.include?("--check")
prune = ARGV.include?("--prune")
out = ARGV[ARGV.index("--out") + 1] if ARGV.include?("--out")
abort "usage: ruby wiki/scripts/publish.rb --check | --out WIKI_CLONE [--prune]" unless check || out
abort "#{out} is not a directory. Clone the wiki repository first." if out && !File.directory?(out)

target = out || Dir.mktmpdir("docxi-wiki")
written = build(target)

if out && prune
  stale = Dir.children(target).select { |f| f.end_with?(".md") } - written
  stale.each { |f| File.delete(File.join(target, f)) }
  puts "pruned #{stale.size} page(s) no longer in the source: #{stale.join(", ")}" unless stale.empty?
elsif out
  extra = Dir.children(target).select { |f| f.end_with?(".md") } - written
  puts "note: #{extra.size} page(s) in #{target} aren't produced by the source (use --prune to delete): #{extra.join(", ")}" unless extra.empty?
end

puts "English pages: #{PAGES.size}"
language_status.each do |code, _, done, count|
  puts "  #{code.ljust(6)} #{done ? "complete, published" : "#{count}/#{PAGES.size} pages, not published"}"
end
stale = stale_translations
puts stale.empty? ? "stale translations: none" : "stale translations (#{stale.size}):\n  " + stale.join("\n  ")

problems = validate(target, written)
puts "#{written.size} files written to #{target}"
if problems.empty?
  puts "links: all internal links, anchors and repository links resolve"
else
  puts "problems:\n  " + problems.join("\n  ")
end
puts "\nNothing was committed or pushed. Look over the folder, then commit and push it from the wiki clone." if out
exit(problems.empty? ? 0 : 1)
