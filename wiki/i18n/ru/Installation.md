<!-- Translated from wiki/Installation.md (git blob 0bd43a0434cbf17fa4dcd95d837e40ca50c52119). See wiki/TRANSLATION_GUIDE.md. -->
# Установка docxi: требования, настройка Bundler и обновление

docxi — гем, написанный только на Ruby. При установке подтягиваются всего две зависимости, nokogiri и rubyzip, и больше ничего. Офисные программы не нужны.

## <a id="requirements"></a>Требования

| Требование | Версия |
|---|---|
| Ruby | 2.5 или новее |
| nokogiri | `>= 1.8, < 2`, устанавливается автоматически |
| rubyzip | `>= 1.3, < 4`, устанавливается автоматически |
| Rails | Не требуется |

## <a id="install-with-bundler"></a>Установка через Bundler

Добавьте гем в Gemfile:

```ruby
gem "docxi", "~> 0.1"
```

Затем установите его:

```sh
bundle install
```

## <a id="install-without-bundler"></a>Установка без Bundler

```sh
gem install docxi
```

## <a id="check-the-installation"></a>Проверка установки

```sh
ruby -e 'require "docxi"; puts Docxi::VERSION'
```

В проекте с Bundler запускайте команду как `bundle exec ruby -e 'require "docxi"; puts Docxi::VERSION'`. Она выводит установленную версию, например `0.1.0`.

## <a id="upgrading-from-00x"></a>Обновление с 0.0.x

Если Gemfile фиксирует docxi на версии 0.0.x, например через `"~> 0.0.6"`, замените строку на `gem "docxi", "~> 0.1"`. Затем выполните `bundle update docxi`. Менять код не нужно, и 0.1.0 создаёт те же документы, что и 0.0.6: набор тестов сравнивает результат с файлами, созданными версией 0.0.6.

Что изменилось внутри:

- nokogiri и rubyzip теперь могут быть любой версии из указанных выше диапазонов, поэтому docxi устанавливается рядом с актуальными версиями Rails.
- docxi больше не требует, чтобы был загружен ActiveSupport.
- Архивы записываются без записей Zip64.
- Изображения загружаются без `Kernel#open`, поэтому URL работают в Ruby 3.
- `tab`, `frame` и номера страниц в верхних колонтитулах работают.

Приложения на Ruby старше 2.5 по-прежнему получают 0.0.6, потому что 0.1.0 объявляет Ruby 2.5 минимальной версией. Все изменения перечислены в [журнале изменений](Changelog).

## <a id="versions-tested-in-ci"></a>Версии, проверяемые в CI

Каждая версия Rails проверяется с полностью загруженным фреймворком рядом с docxi.

| Ruby | Rails |
|---|---|
| 2.5 | 5.2 |
| 2.7 | 6.0 |
| 3.0 | 6.1 |
| 3.1 | 7.0 |
| 3.2 | 7.1 |
| 3.3 | 7.2 |
| 3.4 | 8.0 |
| 4.0 | 8.1 |

CI также запускает:

- тесты без Rails на Ruby 2.5, 2.6 и 4.0
- тесты с rubyzip 1.3 (Ruby 2.5), rubyzip 2.4 (Ruby 3.3) и nokogiri 1.11 (Ruby 2.5)
- проверку того, что LibreOffice может конвертировать примеры документов

## <a id="platform-notes"></a>Замечания о платформах

- nokogiri поставляет предкомпилированные гемы для распространённых платформ Linux, macOS и Windows. Если подходящего нет, он компилируется из исходников, и для этого нужен компилятор C. См. [руководство по установке nokogiri](https://nokogiri.org/tutorials/installing_nokogiri.html).
- CI работает на Ubuntu 24.04, тесты также запускались на macOS. Windows не тестируется.

## <a id="if-installation-fails"></a>Если установка не удалась

Конфликты Bundler вокруг nokogiri или rubyzip почти всегда означают, что подбирается старая версия docxi 0.0.6. См. [Устранение неполадок](Troubleshooting#installation).

## <a id="related-pages"></a>Связанные страницы

- [Начало работы](Getting-Started)
- [Настройка](Configuration)
- [Развёртывание](Deployment)

---

[← Главная вики](Home) · [README](https://github.com/hammad417/docxi#readme)
