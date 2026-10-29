#!/usr/bin/env ruby
# Checks the README, its translations, docs/ and the wiki source. Needs git but
# no network access and no gems:
#
#   ruby docs/scripts/check_docs.rb
#
# It reports:
# - required documentation files that are missing
# - languages without a README translation or without every wiki page
# - translations made from an older English version that aren't marked
#   "outdated", and status files that don't match the files on disk
# - translations whose headings, code blocks or links differ from English
# - relative links and #anchors that don't resolve, and wiki links to pages
#   that don't exist or are in the wrong language
# - pages without exactly one H1, or that skip a heading level
# - Phaedra Solutions links outside the README and the wiki Home and Credits
#   pages, and misspellings of the name
#
# Exits 1 on any error. Warnings are printed but don't fail the run. Links
# between wiki pages are checked by `ruby wiki/scripts/publish.rb --check`.

require "json"

ROOT = File.expand_path("../..", __dir__)
WIKI_URL = "https://github.com/hammad417/docxi/wiki".freeze
PHAEDRA_URL = "https://www.phaedrasolutions.com/".freeze
STATUSES = %w[current outdated].freeze

REQUIRED = %w[
  README.md LICENSE CHANGELOG.md examples/full_document.rb
  docs/i18n/README.md docs/i18n/TRANSLATION_GUIDE.md docs/i18n/translation-status.json
  wiki/README.md wiki/TRANSLATION_GUIDE.md wiki/wiki-publishing-guide.md
  wiki/_Sidebar.md wiki/_Footer.md wiki/Home.md wiki/scripts/publish.rb
  wiki/i18n/translation-status.json
].freeze

WIKI_GUIDES = %w[README.md TRANSLATION_GUIDE.md wiki-publishing-guide.md].freeze
WIKI_SHARED = %w[_Sidebar.md _Footer.md].freeze

FENCE = /^( *)```([^\n]*)\n(.*?)^\1```[ \t]*$/m
LINK = /\]\(([^)\s]+)\)/

@errors = []
@warnings = []

def error(file, message, line = nil)
  @errors << [file, line, message]
end

def warning(file, message, line = nil)
  @warnings << [file, line, message]
end

def read(path)
  File.read(File.join(ROOT, path), encoding: "UTF-8")
end

def exist?(path)
  File.exist?(File.join(ROOT, path))
end

def blobs(paths)
  out = IO.popen(["git", "-C", ROOT, "hash-object", "--stdin-paths"], "r+") do |io|
    io.puts(paths)
    io.close_write
    io.read
  end
  paths.zip(out.split("\n")).to_h
end

def prose(text)
  text.gsub(FENCE, "")
end

def code_blocks(text)
  text.scan(FENCE).map { |_, lang, body| [lang.strip, body] }
end

# Inline code is blanked out unless it is the whole text of a link, as in
# [`examples/full_document.rb`](...). Spans are matched left to right, so a
# closing backtick is never paired with the next span's opening one.
def without_code(text)
  text.gsub(/`[^`\n]+`/) do
    m = Regexp.last_match
    link_text = m.begin(0).positive? && text[m.begin(0) - 1] == "[" && text[m.end(0), 2] == "]("
    link_text ? m[0] : "``"
  end
end

def links(text)
  without_code(prose(text)).scan(LINK).flatten
end

def headings(text)
  prose(text).scan(/^(#+) (.+)$/).map { |hashes, title| [hashes.size, title.strip] }
end

def line_of(text, fragment)
  index = text.lines.index { |line| line.include?(fragment) }
  index && index + 1
end

# GitHub's heading ids: lower case, punctuation dropped, spaces become hyphens.
def slug(title)
  title.gsub(/<[^>]+>/, "").strip.downcase.gsub(/[^\p{Word}\s-]/u, "").gsub(/\s/, "-")
end

def anchors(text)
  ids = prose(text).scan(/<a id="([^"]+)"><\/a>/).flatten
  seen = Hash.new(0)
  headings(text).each do |_, title|
    id = slug(title)
    ids << (seen[id].zero? ? id : "#{id}-#{seen[id]}")
    seen[id] += 1
  end
  ids
end

def external?(target)
  target.match?(/\A[a-z][a-z0-9+.-]*:/i)
end

# Repository-relative form of a relative link, so links written from different
# folders can be compared.
def resolve(path, target)
  file, anchor = target.split("#", 2)
  resolved = File.expand_path(file, File.join("/", File.dirname(path))).delete_prefix("/")
  anchor ? "#{resolved}##{anchor}" : resolved
end

# --- languages --------------------------------------------------------------

readme_status = JSON.parse(read("docs/i18n/translation-status.json"))
wiki_status = JSON.parse(read("wiki/i18n/translation-status.json"))
english = read("README.md")

selector = english.lines.find { |line| line.start_with?("[English](README.md) · ") }.to_s.strip
entries = selector.split(" · ").map { |entry| entry.match(/\A\[(.+)\]\((.+)\)\z/)&.captures }
error("README.md", "the language selector line is missing or malformed") if selector.empty? || entries.include?(nil)
entries = entries.compact
languages = entries.drop(1).map { |name, file| [file[/\Adocs\/i18n\/README\.(.+)\.md\z/, 1], name] }
codes = languages.map(&:first)

publisher = read("wiki/scripts/publish.rb")[/LANGUAGES = \[(.*?)\]\.freeze/m, 1].to_s.scan(/\["([^"]+)", "([^"]+)"\]/)
{
  "docs/i18n/translation-status.json" => readme_status["translations"].map { |t| t["code"] },
  "wiki/i18n/translation-status.json" => wiki_status["translations"].map { |t| t["code"] },
  "docs/i18n/README.<code>.md files" => Dir[File.join(ROOT, "docs/i18n/README.*.md")].map { |f| File.basename(f)[/\AREADME\.(.+)\.md\z/, 1] },
  "wiki/i18n/<code>/ folders" => Dir[File.join(ROOT, "wiki/i18n/*/")].map { |d| File.basename(d) }
}.each do |source, list|
  next if list.sort == codes.sort
  error(source, "languages differ from the README selector: missing #{(codes - list).inspect}, extra #{(list - codes).inspect}")
end
error("wiki/scripts/publish.rb", "LANGUAGES must list the README selector's languages, names and order") unless publisher == languages

REQUIRED.each { |path| error(path, "required file is missing") unless exist?(path) }

# --- wiki page names --------------------------------------------------------

pages = Dir.children(File.join(ROOT, "wiki")).select { |f| f.end_with?(".md") }.reject { |f| (WIKI_GUIDES + WIKI_SHARED).include?(f) }.map { |f| f.delete_suffix(".md") }.sort
CODES = codes.freeze
PAGES = pages.freeze
PUBLISHED = (pages + codes.product(pages).map { |code, page| "#{page}-#{code}" }).freeze

def wiki_page_name(url)
  page = url.delete_prefix(WIKI_URL).delete_prefix("/").split("#").first.to_s
  page.empty? ? "Home" : page
end

# Wiki URLs in a file in language `code` (nil for English) must name a page the
# publisher writes, in that file's language.
def check_wiki_urls(path, text, code)
  links(text).select { |t| t.start_with?(WIKI_URL) }.each do |target|
    page = wiki_page_name(target)
    line = line_of(text, target)
    if !PUBLISHED.include?(page)
      error(path, "wiki link to a page the wiki source doesn't produce: #{target}", line)
    elsif code ? !page.end_with?("-#{code}") : !PAGES.include?(page)
      error(path, "wiki link should point to the #{code || "English"} page: #{target}", line)
    end
  end
end

# --- every Markdown file ----------------------------------------------------

markdown = Dir.chdir(ROOT) { Dir["*.md"] + Dir["docs/**/*.md"] + Dir["wiki/**/*.md"] }.sort
link_count = 0

markdown.each do |path|
  text = read(path)
  wiki_page = path.start_with?("wiki/") && !WIKI_GUIDES.include?(File.basename(path)) || path.start_with?("wiki/i18n/")

  # Heading outline. _Sidebar.md and _Footer.md have no title.
  unless WIKI_SHARED.include?(File.basename(path))
    outline = headings(text)
    h1 = outline.count { |level, _| level == 1 }
    error(path, "expected one H1, found #{h1}") unless h1 == 1
    outline.each_cons(2) do |(above, _), (level, title)|
      error(path, "heading skips from H#{above} to H#{level}: #{title}", line_of(text, title)) if level > above + 1
    end
  end

  # Case-sensitive, because "todo" is an ordinary word in Spanish and Portuguese.
  placeholder = prose(text)[/\b(?:TODO|TBD|FIXME)\b/] || prose(text)[/lorem ipsum/i]
  error(path, "placeholder text: #{placeholder}", line_of(text, placeholder)) if placeholder
  error(path, "unbalanced code fences") if text.scan(/^ *```/).size.odd?

  # Phaedra Solutions: one official URL and one spelling.
  links(text).select { |t| t.include?("phaedrasolutions.com") }.each do |target|
    error(path, "link to #{target}; use #{PHAEDRA_URL}", line_of(text, target)) unless target == PHAEDRA_URL
  end
  prose(text).gsub(/<[^>]+>/, "").gsub(LINK, "]()").gsub(/[\w.@-]*phaedrasolutions\.com\S*/i, "").scan(/phaedra[\s-]*solutions?\b/i).uniq.each do |name|
    error(path, "write the company name as \"Phaedra Solutions\", not \"#{name}\"", line_of(text, name)) unless name == "Phaedra Solutions"
  end

  # Relative links and anchors. Links between wiki pages use page names, which
  # only resolve on GitHub, so publish.rb checks those.
  next if wiki_page

  links(text).uniq.each do |target|
    next if external?(target)

    link_count += 1
    line = line_of(text, "](#{target})")
    file, anchor = target.split("#", 2)
    if file.to_s.empty?
      error(path, "missing anchor ##{anchor}", line) unless anchors(text).include?(anchor)
      next
    end
    resolved = resolve(path, file)
    if !exist?(resolved)
      error(path, "link to a missing file: #{target}", line)
    elsif anchor && resolved.end_with?(".md") && !anchors(read(resolved)).include?(anchor)
      error(path, "missing anchor #{target}", line)
    end
  end
end

# --- translations keep the English structure ---------------------------------

def outline_ids(text, explicit)
  headings(text).reject { |level, _| level == 1 }.map do |level, title|
    [level, explicit ? title[/\A<a id="([^"]+)"><\/a>/, 1] : slug(title)]
  end
end

# Comments in code may be translated, and Mermaid labels are, so compare the
# code without them.
def comparable_code(text)
  code_blocks(text).map do |lang, body|
    lines = body.lines.map(&:rstrip)
    lines = lines.grep_v(/\A\s*#/) if %w[ruby sh].include?(lang)
    lines = lines.map { |line| line.gsub(/"[^"]*"/, '""') } if lang == "mermaid"
    [lang, lines]
  end
end

# Translations link to the wiki page in their own language. README links are
# file paths, so they're compared from the repository root. Wiki pages link by
# page name, which is the same in every language.
def comparable_links(path, text, files)
  links(text).map do |target|
    if target.start_with?(WIKI_URL)
      page = wiki_page_name(target)
      code = CODES.find { |c| page.end_with?("-#{c}") }
      "wiki:" + (code ? page.delete_suffix("-#{code}") : page)
    elsif external?(target) || target.start_with?("#") || !files
      target
    else
      resolve(path, target)
    end
  end.sort
end

def compare(path, text, source_path, source, readme)
  ids = outline_ids(text, true)
  expected_ids = outline_ids(source, false)
  error(path, "headings differ from #{source_path}: expected #{expected_ids.map(&:last).inspect}, found #{ids.map(&:last).inspect}") unless ids == expected_ids

  ours = comparable_code(text)
  theirs = comparable_code(source)
  if ours.size != theirs.size
    error(path, "has #{ours.size} code blocks, #{source_path} has #{theirs.size}")
  else
    ours.zip(theirs).each_with_index do |(a, b), i|
      error(path, "code block #{i + 1} differs from #{source_path}") unless a == b
    end
  end

  expected = comparable_links(source_path, source, readme)
  expected = (expected + ["README.md"]).sort if readme # the note saying the English README is authoritative
  found = comparable_links(path, text, readme)
  return if found == expected

  missing = expected.dup
  extra = []
  found.each { |t| (i = missing.index(t)) ? missing.delete_at(i) : extra << t }
  error(path, "links differ from #{source_path}: missing #{missing.uniq.inspect}, extra #{extra.uniq.inspect}")
end

# Arabic is written right to left. GitHub needs the wrapper and the
# left-to-right spans to show code and paths in the right order.
def check_rtl(path, text)
  body = text.lines.drop(1).join.strip
  error(path, "Arabic pages must be wrapped in <div dir=\"rtl\"> ... </div>") unless body.start_with?("<div dir=\"rtl\">") && body.end_with?("</div>")
  spans = prose(text).to_enum(:scan, /`[^`\n]+`/).map { Regexp.last_match }
  bare = spans.count do |m|
    opened = m.pre_match.rindex('<span dir="ltr">')
    closed = m.pre_match.rindex("</span>")
    opened.nil? || (closed && closed > opened)
  end
  error(path, "#{bare} inline code span(s) not wrapped in <span dir=\"ltr\">") if bare.positive?
end

# --- README translations -------------------------------------------------------

readme_blob = blobs(["README.md"])["README.md"]
if readme_status.dig("source", "blob") != readme_blob
  error("docs/i18n/translation-status.json", "source.blob is #{readme_status.dig("source", "blob")}, but README.md is now #{readme_blob}")
end

expected_selector = selector.gsub("](README.md)", "](../../README.md)").gsub("](docs/i18n/", "](")
readme_files = readme_status["translations"].map { |t| t["file"] }.select { |f| exist?(f) }
readme_blobs = blobs(readme_files)
readme_counts = Hash.new(0)

readme_status["translations"].each do |t|
  path = t["file"]
  unless exist?(path)
    error(path, "missing translation for #{t["language"]}")
    next
  end
  text = read(path)
  recorded = text.lines.first.to_s[/\A<!-- Translated from README\.md \(git blob ([0-9a-f]{40})\)\. See docs\/i18n\/TRANSLATION_GUIDE\.md\. -->/, 1]
  status = t["status"]
  readme_counts[status] += 1

  error(path, "line 1 must name the README.md blob it was translated from", 1) unless recorded
  error(path, "line 1 names #{recorded}, translation-status.json has #{t["source_blob"]}", 1) if recorded && recorded != t["source_blob"]
  error("docs/i18n/translation-status.json", "translation_blob of #{t["code"]} is #{t["translation_blob"]}, but #{path} is now #{readme_blobs[path]}") unless t["translation_blob"] == readme_blobs[path]
  error("docs/i18n/translation-status.json", "#{t["code"]} has an unknown status #{status.inspect}") unless STATUSES.include?(status)
  if recorded && recorded != readme_blob
    if status == "outdated"
      warning(path, "marked outdated: translated from README.md #{recorded[0, 10]}, README.md is now #{readme_blob[0, 10]}")
    else
      error(path, "translated from README.md #{recorded[0, 10]}, but README.md is now #{readme_blob[0, 10]}. Update the translation, or set its status to \"outdated\"", 1)
    end
  end

  error(path, "the language selector must be the README's, with paths relative to docs/i18n/") unless text.lines.map(&:strip).include?(expected_selector)
  compare(path, text, "README.md", english, true)
  check_wiki_urls(path, text, t["code"])
  check_rtl(path, text) if t["code"] == "ar"
  count = links(text).count { |link| link == PHAEDRA_URL }
  error(path, "should link to #{PHAEDRA_URL} once, found #{count}") unless count == 1
end

check_wiki_urls("README.md", english, nil)
count = links(english).count { |link| link == PHAEDRA_URL }
error("README.md", "should link to #{PHAEDRA_URL} once, found #{count}") unless count == 1

# --- wiki translations -------------------------------------------------------

wiki_paths = pages.map { |page| "wiki/#{page}.md" }
translated_paths = codes.flat_map { |code| pages.map { |page| "wiki/i18n/#{code}/#{page}.md" } }.select { |f| exist?(f) }
wiki_blobs = blobs(wiki_paths + translated_paths)

pages.each do |page|
  recorded = wiki_status.dig("source", "pages", page)
  current = wiki_blobs["wiki/#{page}.md"]
  error("wiki/i18n/translation-status.json", "source.pages.#{page} is #{recorded.inspect}, but wiki/#{page}.md is now #{current}") unless recorded == current
  text = read("wiki/#{page}.md")
  error("wiki/#{page}.md", "should link back to the wiki Home page") unless page == "Home" || links(text).include?("Home")
end
(wiki_status.dig("source", "pages") || {}).each_key do |page|
  error("wiki/i18n/translation-status.json", "lists #{page}, which has no English page") unless pages.include?(page)
end
error("wiki/_Sidebar.md", "the <!-- languages --> line is missing") unless read("wiki/_Sidebar.md").include?("<!-- languages -->")

wiki_counts = Hash.new(0)
complete = 0

wiki_status["translations"].each do |t|
  code = t["code"]
  dir = "wiki/i18n/#{code}"
  present = pages.select { |page| exist?("#{dir}/#{page}.md") }
  status = t["status"]
  wiki_counts[status] += 1
  complete += 1 if present.size == pages.size

  extra = exist?(dir) ? Dir.children(File.join(ROOT, dir)).map { |f| f.delete_suffix(".md") } - pages : []
  error(dir, "pages without an English page: #{extra.inspect}") unless extra.empty?
  if present.size < pages.size
    message = "#{present.size}/#{pages.size} pages, so the language isn't published. Missing: #{(pages - present).inspect}"
    t["complete"] ? error(dir, "translation-status.json says complete, but it has #{message}") : warning(dir, message)
  end
  error("wiki/i18n/translation-status.json", "#{code}: complete is #{t["complete"]}, but #{present.size}/#{pages.size} pages exist") if t["complete"] != (present.size == pages.size)
  error("wiki/i18n/translation-status.json", "#{code}: pages_translated is #{t["pages_translated"]}, but #{present.size} pages exist") if t["pages_translated"] != present.size
  error("wiki/i18n/translation-status.json", "#{code} has an unknown status #{status.inspect}") unless STATUSES.include?(status)

  present.each do |page|
    path = "#{dir}/#{page}.md"
    text = read(path)
    entry = (t["pages"] || {})[page] || {}
    recorded = text.lines.first.to_s[/\A<!-- Translated from wiki\/#{Regexp.escape(page)}\.md \(git blob ([0-9a-f]{40})\)\. See wiki\/TRANSLATION_GUIDE\.md\. -->/, 1]
    current = wiki_blobs["wiki/#{page}.md"]

    error(path, "line 1 must name the wiki/#{page}.md blob it was translated from", 1) unless recorded
    error(path, "line 1 names #{recorded}, translation-status.json has #{entry["source_blob"].inspect}", 1) if recorded && recorded != entry["source_blob"]
    error("wiki/i18n/translation-status.json", "#{code}/#{page}: translation_blob is #{entry["translation_blob"].inspect}, but the file is now #{wiki_blobs[path]}") unless entry["translation_blob"] == wiki_blobs[path]
    if recorded && recorded != current
      if status == "outdated"
        warning(path, "marked outdated: translated from #{recorded[0, 10]}, wiki/#{page}.md is now #{current[0, 10]}")
      else
        error(path, "translated from wiki/#{page}.md #{recorded[0, 10]}, but the English page is now #{current[0, 10]}. Update the translation, or set the language's status to \"outdated\"", 1)
      end
    end

    compare(path, text, "wiki/#{page}.md", read("wiki/#{page}.md"), false)
    check_wiki_urls(path, text, code)
    check_rtl(path, text) if code == "ar"
  end
end

# Only the wiki Home and Credits pages link to the company website.
(wiki_paths + translated_paths).each do |path|
  count = links(read(path)).count { |link| link == PHAEDRA_URL }
  expected = %w[Home Credits].include?(File.basename(path, ".md")) ? 1 : 0
  error(path, "should link to #{PHAEDRA_URL} #{expected} time(s), found #{count}") unless count == expected
end

# --- report ------------------------------------------------------------------

def counts(hash)
  hash.sort.map { |status, n| "#{n} #{status}" }.join(", ")
end

puts "Required files: #{REQUIRED.size} checked"
puts "README translations: #{readme_status["translations"].size} languages (#{counts(readme_counts)})"
puts "Wiki: #{pages.size} English pages, #{complete}/#{codes.size} languages complete (#{counts(wiki_counts)})"
puts "Markdown files: #{markdown.size} checked for headings, attribution and placeholders"
puts "Relative links and anchors: #{link_count} checked outside the wiki pages"

github = ENV["GITHUB_ACTIONS"] == "true"
[[@warnings, "warning"], [@errors, "error"]].each do |list, kind|
  list.each do |file, line, message|
    if github
      puts "::#{kind} file=#{file}#{line ? ",line=#{line}" : ""}::#{message}"
    else
      puts "#{kind}: #{file}#{line ? ":#{line}" : ""}: #{message}"
    end
  end
end

puts @errors.empty? ? "No problems found." : "#{@errors.size} error(s) found."
exit(@errors.empty? ? 0 : 1)
