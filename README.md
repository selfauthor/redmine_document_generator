# Redmine Document Generator

Плагин для генерации документов Word (.docx) и Excel (.xlsx) по выборке задач проекта с использованием настраиваемых шаблонов.

![Версия](https://img.shields.io/badge/version-1.0.0-blue)
![Redmine](https://img.shields.io/badge/Redmine-5.0%2B-green)
![Лицензия](https://img.shields.io/badge/license-GPLv2-orange)


## Совместимость
- Redmine: 5.x, 6.x
- Ruby: 3.2+
- База данных: MySQL/MariaDB, PostgreSQL, SQLite

## Зависимости
Плагин требует установки следующих gem-ов, которых нет в стандартной поставке Redmine:
- `rubyXL` (~> 3.4) — для чтения и модификации шаблонов Excel с сохранением форматирования.

*Примечание: `rubyzip` и `nokogiri` уже входят в состав Redmine и дополнительно не требуются.*

## Установка
1. Скопируйте файлы плагина в директорию `plugins/redmine_document_generator` вашего Redmine.
2. Установите зависимости плагина из корневой директории Redmine:
   ```bash
   /opt/ruby-3.2/bin/bundle install
   ```
3. Перезапустите приложение Redmine (например, Passenger или Puma).

## Документация
- 📘 [Руководство пользователя](docs/USER_MANUAL.md) — как создавать шаблоны и пользоваться плагином.
- ⚙️ [Конфигурация](docs/CONFIGURATION.md) — описание параметров настройки плагина.
- 🛠️ [Техническая документация](docs/TECHNICAL.md) — архитектура и правила разработки для участников проекта.
