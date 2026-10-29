# SEO notes

Facts about how docxi's pages can be found, and the terms each language uses. They were checked in October 2026; GitHub's behaviour can change.

## What GitHub allows

- The repository page's `<title>` is `GitHub - hammad417/docxi: <description> · GitHub`, and its meta description is the repository description followed by "Contribute to hammad417/docxi development…". The description is the most important search text docxi controls, and it's a setting, not a file. Don't end it with a period, or the meta description gets two.
- A file page such as `docs/i18n/README.ja.md` is titled by its path and gets the same English meta description. Only the content of a translated README is localized.
- GitHub strips `<meta>` and `<link>` tags from Markdown and shows `<script>` and `<title>` as text, so a README can't carry canonical URLs, `hreflang`, structured data or its own title.
- Search engines only index GitHub wikis with 500 or more stars that don't allow public editing ([GitHub docs](https://docs.github.com/en/communities/documenting-your-project-with-wikis/about-wikis)). Until then the wiki is for readers who come from the README, and the README and its translations are what search engines see.
- Links in rendered Markdown get `rel="nofollow"`.

## Where the keywords are

The README title says "create Word (.docx) documents with Ruby". The first paragraph covers Ruby, Microsoft Word, `.docx`, Office Open XML and Rails, and says Word, LibreOffice and templates aren't needed. The feature list names tables, images, headers and footers, page numbers and the table of contents. Troubleshooting quotes the exact error messages people paste into a search engine, and the FAQ answers how-to, requirement and "can it…" questions. Each "no" answer says what docxi doesn't do, so the page doesn't attract searches it can't help with.

"Word", "Ruby" and "Rails" each appear 12 to 18 times in the README's prose, about once per section, and the translations use them the same number of times.

## Terms by language

The translations use the terms from Microsoft Word's interface in each language, which is also how developers name these features. This is terminology, not measured search demand.

| Code | Title phrase (H1) | "Word document" | Headers and footers | Table of contents | FAQ: Word document in Rails |
|---|---|---|---|---|---|
| en | create Word (.docx) documents with Ruby | Word document | Headers and footers | Table of contents | How do I generate a Word document in Ruby on Rails? |
| zh-CN | 使用 Ruby 生成 Word（.docx）文档 | Word 文档 | 页眉与页脚 | 目录 | 如何在 Ruby on Rails 中生成 Word 文档？ |
| zh-TW | 使用 Ruby 產生 Word（.docx）文件 | Word 文件 | 頁首與頁尾 | 目錄 | 如何在 Ruby on Rails 中產生 Word 文件？ |
| es | genera documentos Word (.docx) con Ruby | documento de Word | Encabezados y pies de página | Tabla de contenido | ¿Cómo genero un documento de Word en Ruby on Rails? |
| ar | إنشاء مستندات Word (.docx) باستخدام Ruby | مستند Word | رؤوس الصفحات وتذييلاتها | جدول المحتويات | كيف أُنشئ مستند Word في Ruby on Rails؟ |
| fr | générer des documents Word (.docx) avec Ruby | document Word | En-têtes et pieds de page | Table des matières | Comment générer un document Word avec Ruby on Rails ? |
| pt-BR | gere documentos Word (.docx) com Ruby | documento do Word | Cabeçalhos e rodapés | Sumário | Como gerar um documento do Word no Ruby on Rails? |
| de | Word-Dokumente (.docx) mit Ruby erstellen | Word-Dokument | Kopf- und Fußzeilen | Inhaltsverzeichnis | Wie erzeuge ich ein Word-Dokument in Ruby on Rails? |
| hi | Ruby से Word (.docx) डॉक्यूमेंट बनाएं | Word डॉक्यूमेंट | हेडर और फ़ुटर | विषय-सूची | Ruby on Rails में Word डॉक्यूमेंट कैसे बनाएं? |
| ja | Ruby で Word（.docx）文書を生成 | Word 文書 | ヘッダーとフッター | 目次 | Ruby on Rails で Word 文書を生成するにはどうすればよいですか？ |
| ko | Ruby로 Word(.docx) 문서 생성하기 | Word 문서 | 머리글과 바닥글 | 목차 | Ruby on Rails에서 Word 문서를 생성하려면 어떻게 하나요? |
| ru | создание документов Word (.docx) на Ruby | документ Word | Колонтитулы | Оглавление | Как создать документ Word в Ruby on Rails? |
| id | membuat dokumen Word (.docx) dengan Ruby | dokumen Word | Header dan footer | Daftar isi | Bagaimana cara membuat dokumen Word di Ruby on Rails? |
| tr | Ruby ile Word (.docx) belgeleri oluşturun | Word belgesi | Üst ve alt bilgiler | İçindekiler tablosu | Ruby on Rails'te nasıl Word belgesi oluşturabilirim? |
| it | generare documenti Word (.docx) con Ruby | documento Word | Intestazioni e piè di pagina | Sommario | Come si genera un documento Word in Ruby on Rails? |
| vi | tạo tài liệu Word (.docx) bằng Ruby | tài liệu Word | Đầu trang và chân trang | Mục lục | Làm thế nào để tạo tài liệu Word trong Ruby on Rails? |

Product names, code, error messages, commands and the MIME type stay in English everywhere, because that's how people search for them.

## A documentation site

docxi has no documentation site. If search traffic to the full guide ever matters, the wiki source in `wiki/` could also be built into a GitHub Pages site. Only a site you host can use the things GitHub doesn't allow: per-language URLs with `hreflang` alternates and `x-default`, self-referencing canonical URLs, an XML sitemap, a `SoftwareSourceCode` JSON-LD block, and an [`llms.txt`](https://llmstxt.org/) file at the root. `llms.txt` is a proposal that some tools read; it doesn't guarantee anything. If a site is built, decide which copy is the main one so the site and the wiki don't compete.
