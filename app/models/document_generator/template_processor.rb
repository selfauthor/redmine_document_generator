# frozen_string_literal: true

require 'zip'
require 'nokogiri'

# v2609141321
module DocumentGenerator
  # TemplateProcessor отвечает за общую логику обработки шаблонов.
  # Он абстрагирует работу с ZIP-архивами (форматы .docx и .xlsx)
  # и предоставляет универсальные методы для замены маркеров в тексте.
  class TemplateProcessor
    class << self
      # Извлекает значение из контекста по ключу, поддерживая вложенность через точку.
      # Например, ключ "Parent.Тема" извлечет значение из context['Parent']['Тема'].
      #
      # @param key [String] Имя ключа (например, "Тема" или "Parent.Статус")
      # @param context [Hash] Хэш данных, переданный из ContextBuilder
      # @return [String, nil] Найденное значение в виде строки или nil, если не найдено
      def resolve_value(key, context)
        return nil if key.blank? || context.nil?

        parts = key.split('.')
        current = context

        parts.each do |part|
          return nil unless current.is_a?(Hash)
          # Ищем ключ без учета регистра для большей устойчивости
          matched_key = current.keys.find { |k| k.to_s.downcase == part.downcase }
          return nil unless matched_key
          
          current = current[matched_key]
        end

        # Форматируем итоговое значение (массивы в строку, даты и т.д.)
        format_value(current)
      end

      # Заменяет все маркеры <%...%> в переданной строке на значения из контекста.
      #
      # @param text [String] Исходный текст, содержащий маркеры
      # @param context [Hash] Хэш данных для подстановки
      # @return [String] Текст с замененными маркерами
      def substitute_markers(text, context)
        return text if text.blank? || !text.include?('<%')

        text.gsub(/<%\s*([^%]+?)\s*%>/) do |match|
          marker_content = $1.strip
          
          # Игнорируем управляющие маркеры, они обрабатываются на уровне XML-узлов
          next match if marker_content.match?(/^(BEGIN_|END_|IF|ELSE|GROUP_BY)/i)

          resolved_value = resolve_value(marker_content, context)
          resolved_value.to_s
        end
      end

      # Универсальный метод для обработки ZIP-архива шаблона.
      # Открывает архив, находит указанные XML-файлы, передает их в блок для модификации
      # через Nokogiri, а затем сохраняет измененный архив в новый файл.
      #
      # @param template_path [String] Путь к исходному файлу шаблона (.docx или .xlsx)
      # @param output_path [String] Путь для сохранения результирующего файла
      # @param xml_paths [Array<String>] Массив путей к XML-файлам внутри архива для обработки
      # @yield [Nokogiri::XML::Document, String] Передает документ Nokogiri и имя файла в блок
      def process_archive(template_path, output_path, xml_paths)
        temp_output = "#{output_path}.tmp"
        FileUtils.cp(template_path, temp_output)

        Zip::File.open(temp_output) do |zip_file|
          xml_paths.each do |xml_path|
            # Поддержка wildcard (например, 'word/header*.xml')
            entries = xml_path.include?('*') ? zip_file.glob(xml_path) : [zip_file.find_entry(xml_path)]
            
            entries.compact.each do |entry|
              xml_content = entry.get_input_stream.read
              doc = Nokogiri::XML(xml_content)
              
              # Передаем документ в рендерер для специфичной модификации
              yield(doc, entry.name) if block_given?
              
              # Перезаписываем измененный XML обратно в архив
              zip_file.get_output_stream(entry.name) { |os| os.write(doc.to_xml(indent: 0, save_with: 0)) }
            end
          end
        end

        FileUtils.mv(temp_output, output_path)
      end

      private

      # Вспомогательный метод для приведения различных типов данных к строке.
      #
      # @param val [Object] Исходное значение
      # @return [String] Отформатированная строка
      def format_value(val)
        return '' if val.nil?
        return val.join(', ') if val.is_a?(Array)
        return val.strftime('%d.%m.%Y') if val.is_a?(Date) || val.is_a?(Time)
        
        val.to_s
      end
    end
  end
end