# frozen_string_literal: true

require 'redmine'

# Проверка наличия всех необходимых gem-ов.
# Если какой-то gem отсутствует, плагин регистрируется, но выводит предупреждение
# в журнал и не активирует функциональность.
DOCUMENT_GENERATOR_GEMS_LOADED = begin
  # Основные гемы (Word и Excel)
  require 'sablon'
  require 'rubyXL'
  
  # Вспомогательные гемы (могут быть в Redmine)
  begin
    require 'caxlsx'
  rescue LoadError
    # Пробуем старый axlsx (для обратной совместимости)
    begin
      require 'axlsx'
    rescue LoadError
      Rails.logger.warn "[DocumentGenerator] Gem 'caxlsx' или 'axlsx' не найден. " \
                        "Экспорт в Excel будет недоступен. " \
                        "Установите: gem install caxlsx"
    end
  end
  
  begin
    require 'zip'  # rubyzip загружается как 'zip'
  rescue LoadError
    Rails.logger.warn "[DocumentGenerator] Gem 'rubyzip' не найден. " \
                      "Упаковка файлов в ZIP будет недоступна. " \
                      "Установите: gem install rubyzip"
  end
  
  true
rescue LoadError => e
  Rails.logger.error "[DocumentGenerator] Критическая ошибка: #{e.message}. " \
                     "Плагин зарегистрирован, но генерация документов будет недоступна. " \
                     "Выполните 'bundle install' в корне Redmine."
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
  version '1.0.0'
  url 'https://github.com/selfauthor/redmine_document_generator'

  requires_redmine version_or_higher: '6.0.0'

  # Модуль проекта
  project_module :document_generator do
    permission :use_document_generator,
               { document_generator: [:dialog, :export] },
               require: :member
  end
end