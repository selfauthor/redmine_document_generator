# frozen_string_literal: true

require 'redmine'

# Проверка наличия всех необходимых gem-ов.
# Если какой-то gem отсутствует, плагин регистрируется, но выводит предупреждение
# в журнал и не активирует функциональность генерации.
DOCUMENT_GENERATOR_GEMS_LOADED = begin
  # Основные гемы (Word и Excel)
  require 'rubyXL'

  # Вспомогательные гемы (могут быть уже установлены в Redmine)
  begin
    require 'caxlsx'
  rescue LoadError
    # Пробуем старый axlsx для обратной совместимости
    begin
      require 'axlsx'
    rescue LoadError
      Rails.logger.warn "[DocumentGenerator] Gem 'caxlsx' or 'axlsx' not found. Excel export will be unavailable. Please run: gem install caxlsx"
    end
  end

  begin
    require 'zip' # rubyzip загружается как 'zip'
  rescue LoadError
    Rails.logger.warn "[DocumentGenerator] Gem 'rubyzip' not found. ZIP packaging will be unavailable. Please run: gem install rubyzip"
  end

  true
rescue LoadError => e
  Rails.logger.error "[DocumentGenerator] Critical error: #{e.message}. Plugin registered, but document generation will be unavailable. Please run 'bundle install' in the Redmine root directory."
  false
end

Rails.configuration.to_prepare do
  require_dependency 'document_generator/hooks'
end

Redmine::Plugin.register :redmine_document_generator do
  name 'Генератор документов'
  author 'Андрей Якушев'
  author_url 'https://a2ya.ru'
  description 'Генератор документов Word и Excel по выборке записей проекта с использованием шаблонов'
  version '0.1.2'
  url 'https://github.com/selfauthor/redmine_document_generator'
  requires_redmine version_or_higher: '6.0.0'

  # Модуль проекта
  project_module :document_generator do
    permission :use_document_generator,
               { document_generator: [:dialog, :export] },
               require: :member
  end
end