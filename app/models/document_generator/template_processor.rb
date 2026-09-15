# frozen_string_literal: true

module DocumentGenerator
  class TemplateProcessor
    # Подставляет значения из контекста в текст, заменяя маркеры вида <%...%>
    #
    # @param text [String] Исходный текст с маркерами
    # @param context [Hash] Данные для подстановки
    # @return [String] Текст с подставленными значениями
    def self.substitute_markers(text, context)
      return text unless text.is_a?(String)

      text.gsub(/<%\s*(.*?)\s*%>/) do |match|
        key = $1.strip
        
        # Поддержка вложенных ключей (например, "Parent.Тема" или "Subtask.ID")
        value = if key.include?('.')
                  parts = key.split('.')
                  parts.inject(context) { |h, k| h.is_a?(Hash) ? h[k] : nil }
                else
                  context[key]
                end

        # Если значение nil или пустая строка, возвращаем пустую строку.
        # Это гарантирует, что плейсхолдер исчезнет, если данных нет.
        if value.nil? || value.to_s.strip.empty?
          ""
        else
          value.to_s
        end
      end
    end

    # Обрабатывает архив .docx, извлекая указанные XML-файлы, 
    # передавая их в блок для модификации, и сохраняя обратно.
    #
    # @param input_path [String] Путь к исходному шаблону
    # @param output_path [String] Путь для сохранения результата
    # @param xml_targets [Array<String>] Маски имен файлов внутри архива (например, 'word/document.xml')
    # @yield [Nokogiri::XML::Document, String] Блок получает документ и имя файла для обработки
    def self.process_archive(input_path, output_path, xml_targets)
      require 'zip'

      Zip::File.open(input_path) do |zip_file|
        Zip::File.open(output_path, Zip::File::CREATE) do |out_zip|
          zip_file.each do |entry|
            # Копируем все файлы по умолчанию
            out_zip.get_output_stream(entry.name) do |os|
              os.write(zip_file.read(entry.name))
            end

            # Если файл подходит под маску, обрабатываем его
            if xml_targets.any? { |target| File.fnmatch(target, entry.name) }
              # Извлекаем, обрабатываем и перезаписываем в новом архиве
              xml_content = zip_file.read(entry.name)
              doc = Nokogiri::XML(xml_content)
              
              yield(doc, entry.name)
              
              out_zip.get_output_stream(entry.name) do |os|
                # Сохраняем с исходной кодировкой и без деклараций, чтобы не сломать структуру Word
                os.write(doc.to_xml(indent: 0, save_with: Nokogiri::XML::Node::SaveOptions::AS_XML))
              end
            end
          end
        end
      end
    end
  end
  # v2609150945
end