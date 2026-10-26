# frozen_string_literal: true

# Gem-зависимости плагина "Генератор документов"
# Устанавливаются командой: /opt/ruby-3.2/bin/bundle install (в корне Redmine)

source 'https://rubygems.org'

# === Excel (.xlsx) ===
# Чтение шаблонов Excel с сохранением форматирования
# Этот gem отсутствует в стандартной поставке Redmine
gem 'rubyXL', '~> 3.4'

# === Гемы, которые УЖЕ ЕСТЬ в Redmine ===
# rubyzip (~> 2.3.0) — уже подключён в основном Gemfile Redmine
# caxlsx — уже подключён в основном Gemfile Redmine для экспорта CSV/Excel
# Не дублируем их здесь, чтобы избежать конфликта версий