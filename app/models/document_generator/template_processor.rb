# frozen_string_literal: true

module DocumentGenerator
  class TemplateProcessor
    # ==========================================================================
    # ПОДСТАНОВКА ЗНАЧЕНИЙ (общая логика для Word и Excel)
    # ==========================================================================
    
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
        
        # Если значение nil или пустая строка, возвращаем пустую строку
        if value.nil? || value.to_s.strip.empty?
          ""
        else
          value.to_s
        end
      end
    end
    
    # ==========================================================================
    # ОБРАБОТКА УСЛОВНЫХ БЛОКОВ IF/ELSE/END (общая логика)
    # ==========================================================================
    
    # Обрабатывает условные блоки в тексте, оставляя только подходящую ветку
    #
    # @param text [String] Текст с маркерами IF/ELSE/END
    # @param context [Hash] Данные для вычисления условий
    # @return [String] Текст с обработанными условиями
    def self.resolve_conditionals(text, context)
      return text unless text.is_a?(String)
      
      # Обрабатываем в цикле, пока есть условия
      loop do
        match = text.match(/<%\s*IF\s*\(\s*(.*?)\s*\)\s*%>/im)
        break unless match
        
        condition_str = match[1].strip
        is_true = evaluate_condition(condition_str, context)
        
        if_start = match.begin(0)
        if_end = match.end(0)
        
        # Ищем ELSE и END после IF
        rest = text[if_end..-1]
        else_match = rest.match(/<%\s*ELSE\s*%>/im)
        end_match = rest.match(/<%\s*END\s*%>/im)
        
        break unless end_match
        
        end_pos = if_end + end_match.begin(0)
        end_len = end_match[0].length
        
        # Определяем, какие части удалить
        if else_match && else_match.begin(0) < end_match.begin(0)
          else_pos = if_end + else_match.begin(0)
          else_len = else_match[0].length
          
          if is_true
            # Оставляем IF ветку, удаляем ELSE и END
            text = text[0, if_start] + text[if_end, else_pos - if_end] + text[end_pos + end_len..-1]
          else
            # Оставляем ELSE ветку, удаляем IF и ELSE
            text = text[0, if_start] + text[else_pos + else_len, end_pos - else_pos - else_len] + text[end_pos + end_len..-1]
          end
        else
          # Нет ELSE
          if is_true
            # Оставляем IF ветку, удаляем маркеры
            text = text[0, if_start] + text[if_end, end_pos - if_end] + text[end_pos + end_len..-1]
          else
            # Удаляем всю IF ветку
            text = text[0, if_start] + text[end_pos + end_len..-1]
          end
        end
      end
      
      text
    end
    
    # Вычисляет условие (поддерживает ==, != и проверку на непустоту)
    #
    # @param condition_str [String] Строка условия
    # @param context [Hash] Контекст данных
    # @return [Boolean] Результат вычисления
    def self.evaluate_condition(condition_str, context)
      condition_str = condition_str.strip
      
      if condition_str.include?('==')
        left, right = condition_str.split('==', 2).map(&:strip)
        right = right.gsub(/^['"]|['"]$/, '')
        left_val = get_context_value(left, context)
        return left_val.to_s.strip == right.to_s.strip
      elsif condition_str.include?('!=')
        left, right = condition_str.split('!=', 2).map(&:strip)
        right = right.gsub(/^['"]|['"]$/, '')
        left_val = get_context_value(left, context)
        return left_val.to_s.strip != right.to_s.strip
      else
        # Проверка на непустоту
        val = get_context_value(condition_str, context)
        return !val.to_s.strip.empty?
      end
    end
    
    # Получает значение из контекста по ключу (поддерживает вложенность через точку)
    #
    # @param key [String] Ключ (например, "Parent.Тема")
    # @param context [Hash] Контекст данных
    # @return [Object, nil] Значение или nil
    def self.get_context_value(key, context)
      if key.include?('.')
        parts = key.split('.')
        parts.inject(context) { |h, k| h.is_a?(Hash) ? h[k] : nil }
      else
        context[key]
      end
    end
    
    # ==========================================================================
    # ОЧИСТКА УПРАВЛЯЮЩИХ МАРКЕРОВ (общая логика)
    # ==========================================================================
    
    # Удаляет управляющие маркеры из текста
    #
    # @param text [String] Текст с маркерами
    # @return [String] Очищенный текст
    def self.clean_control_markers(text)
      return text unless text.is_a?(String)
      
      text.gsub(/<%\s*(BEGIN_ROW|END_ROW|BEGIN_SUBTASKS|END_SUBTASKS|BEGIN_WATCHERS|END_WATCHERS|BEGIN_RELATIONS|END_RELATIONS|BEGIN_GROUP_HEADER|END_GROUP_HEADER|BEGIN_GROUP_HEADER_2|END_GROUP_HEADER_2|BEGIN_GROUP_FOOTER|END_GROUP_FOOTER|BEGIN_GROUP_FOOTER_2|END_GROUP_FOOTER_2|BEGIN_TOTAL|END_TOTAL|GROUP_BY|GROUP_BY_2|IF|ELSE|END)\s*%>/i, '')
    end
    
    # ==========================================================================
    # РАБОТА С АРХИВАМИ (общая утилита)
    # ==========================================================================
    
    # Обрабатывает архив .docx/.xlsx, извлекая указанные XML-файлы,
    # передавая их в блок для модификации, и сохраняя обратно.
    #
    # @param input_path [String] Путь к исходному шаблону
    # @param output_path [String] Путь для сохранения результата
    # @param xml_targets [Array<String>] Маски имен файлов внутри архива
    # @yield [Nokogiri::XML::Document, String] Блок получает документ и имя файла
    def self.process_archive(input_path, output_path, xml_targets)
      require 'zip'
      Zip::File.open(input_path) do |zip_file|
        Zip::File.open(output_path, Zip::File::CREATE) do |out_zip|
          zip_file.each do |entry|
            out_zip.get_output_stream(entry.name) do |os|
              os.write(zip_file.read(entry.name))
            end
            if xml_targets.any? { |target| File.fnmatch(target, entry.name) }
              xml_content = zip_file.read(entry.name)
              doc = Nokogiri::XML(xml_content)
              yield(doc, entry.name)
              out_zip.get_output_stream(entry.name) do |os|
                os.write(doc.to_xml(indent: 0, save_with: Nokogiri::XML::Node::SaveOptions::AS_XML))
              end
            end
          end
        end
      end
    end
  end
end
# v2609151129