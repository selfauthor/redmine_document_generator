# frozen_string_literal: true

module DocumentGenerator
  # ============================================================================
  # ОБРАБОТКА ШАБЛОНОВ - БАЗОВЫЙ ПРОЦЕССОР
  # ============================================================================
  # Этот класс отвечает за обработку шаблонов Word и Excel.
  # 
  # ЯЗЫК ШАБЛОНОВ:
  # ================
  # 
  # 1. УПРАВЛЯЮЩИЕ КОМАНДЫ:
  # ------------------------
  # <%IF(условие)%> ... <%ELSE%> ... <%END%>
  #   Условное выполнение. Если условие истинно - выполняется первая ветка,
  #   иначе - вторая (после ELSE). ELSE необязателен.
  #   Пример: <%IF(Статус == 'Закрыт')%>Закрыта<%ELSE%>Открыта<%END%>
  #
  # <%BEGIN_ROW%> ... <%END_ROW%>
  #   Цикл по записям. Блок между маркерами повторяется для каждой записи.
  #   Для таблиц: если маркеры в ячейках - клонируется вся строка таблицы.
  #   Для текста: если маркеры в абзацах - клонируются абзацы.
  #
  # <%BEGIN_SUBTASKS%> ... <%END_SUBTASKS%>
  #   Цикл по подзадачам текущей задачи.
  #   Внутри блока доступны поля с префиксом Subtask.: 
  #   <%Subtask.Тема%>, <%Subtask.Статус%> и т.д.
  #
  # <%BEGIN_WATCHERS%> ... <%END_WATCHERS%>
  #   Цикл по наблюдателям задачи.
  #   Внутри: <%Watcher.Имя%>
  #
  # <%BEGIN_RELATIONS%> ... <%END_RELATIONS%>
  #   Цикл по связанным задачам.
  #   Внутри: <%Relation.Тип%>, <%Relation.Тема%>, <%Relation.Статус%>
  #
  # <%BEGIN_RELATIONS:тип%> ... <%END_RELATIONS%>
  #   Цикл по связям конкретного типа (blocks, relates, и т.д.)
  #
  # 2. ПОЛЯ ДАННЫХ:
  # ---------------
  # <%ИмяПоля%> - подстановка значения поля
  #   Стандартные поля: ID, Тема, Описание, Статус, Приоритет, Автор, 
  #   Назначенный, Дата начала, Срок, Готовность, Оценка времени, 
  #   Фактическое время, Дата создания, Дата обновления, Дата закрытия,
  #   Проект, Трекер, Категория, Версия, Parent (ID родителя)
  #
  # <%Parent.ИмяПоля%> - поле родительской задачи
  #   Пример: <%Parent.Тема%>, <%Parent.Статус%>
  #
  # <%Subtask.ИмяПоля%> - поле подзадачи (внутри цикла BEGIN_SUBTASKS)
  # <%Watcher.Имя%> - имя наблюдателя (внутри цикла BEGIN_WATCHERS)
  # <%Relation.ИмяПоля%> - поле связанной задачи (внутри цикла BEGIN_RELATIONS)
  #
  # <%ИмяПользовательскогоПоля%> - обращение к пользовательскому полю
  #   Пример: <%Сложность%>, <%Срочность%>
  #
  # <%CF:ИмяПоля%> - явное указание пользовательского поля
  #   (если имя конфликтует со стандартным полем)
  #
  # 3. СПЕЦИАЛЬНЫЕ ПЕРЕМЕННЫЕ:
  # --------------------------
  # <%row_number%> - порядковый номер записи в выборке (начиная с 1)
  # <%row_number_in_group%> - номер записи внутри текущей группы
  # <%GroupValue%> - значение поля группировки для текущего блока
  # <%GroupValue2%> - значение поля второго уровня группировки
  #
  # 4. АГРЕГАТНЫЕ ФУНКЦИИ (для групповых и итоговых строк):
  # ------------------------------------------------------
  # <%count%> - количество записей в группе/выборке
  # <%sum(ИмяПоля)%> - сумма значений числового поля
  # <%avg(ИмяПоля)%> - среднее значение
  # <%min(ИмяПоля)%> - минимальное значение
  # <%max(ИмяПоля)%> - максимальное значение
  # <%concat(ИмяПоля, ', ')%> - перечисление значений через разделитель
  #
  # Префикс total_ для общих итогов:
  # <%total_count%> - общее количество записей
  # <%total_sum(ИмяПоля)%> - общая сумма
  # <%total_avg(ИмяПоля)%> - общее среднее
  #
  # 5. ФУНКЦИИ ФОРМАТИРОВАНИЯ:
  # --------------------------
  # <%date(Дата создания, 'DD.MM.YYYY')%> - форматирование даты
  # <%now('DD.MM.YYYY HH:mm')%> - текущая дата/время
  # <%upper(Тема)%> - ВЕРХНИЙ РЕГИСТР
  # <%lower(Тема)%> - нижний регистр
  # <%capitalize(Тема)%> - Первая буква заглавная
  # <%truncate(Описание, 100)%> - обрезка до N символов
  # <%strip_html(Описание)%> - удаление HTML-тегов
  # <%nl2br(Описание)%> - переносы строк в <br>
  # <%replace(Тема, 'старое', 'новое')%> - замена подстроки
  # <%number(Оценка времени, 2)%> - форматирование числа (2 знака после запятой)
  # <%default(Назначенный, 'не назначен')%> - значение по умолчанию
  # <%length(Тема)%> - длина строки
  # <%concat(Тема, ' (', ID, ')')%> - конкатенация строк
  #
  # 6. МЕТАДАННЫЕ ВЫГРУЗКИ:
  # -----------------------
  # <%ExportDate%> - дата/время формирования отчёта
  # <%ExportUser%> - имя пользователя, запустившего выгрузку
  # <%ProjectName%> - название проекта
  # <%QueryName%> - имя сохранённого запроса (если есть)
  # <%FilterDescription%> - текстовое описание условий фильтра
  #
  # 7. ГРУППИРОВКА (для Excel):
  # ---------------------------
  # <%GROUP_BY:Статус%> - директива группировки (первый уровень)
  # <%GROUP_BY_2:Назначенный%> - второй уровень группировки
  #
  # Для Excel маркеры размещаются в первой ячейке строки:
  # GROUP_HEADER - заголовок группы (1-й уровень)
  # GROUP_HEADER_2 - заголовок группы (2-й уровень)
  # ROW - строка данных (повторяется для каждой записи)
  # GROUP_FOOTER_2 - итоги группы (2-й уровень)
  # GROUP_FOOTER - итоги группы (1-й уровень)
  # TOTAL - общие итоги по выборке
  #
  # ============================================================================

  class TemplateProcessor

    # ==========================================================================
    # НОРМАЛИЗАЦИЯ XML - ОБЪЕДИНЕНИЕ РАЗБИТЫХ МАРКЕРОВ
    # ==========================================================================
    # Объединяет разбитые на несколько <w:t> маркеры в один узел.
    # Разделяет узлы, содержащие и маркеры, и обычный текст.
    # Работает на уровне <w:r> (прогонов) для сохранения форматирования.
    #
    # @param parent_node [Nokogiri::XML::Node] Родительский узел (абзац или строка)
    # @param ns [Hash] Пространства имен XML
    # @return [void]
    def self.normalize_xml_nodes(parent_node, ns)
      # Собираем все текстовые узлы
      text_nodes = parent_node.xpath('.//w:t', ns).to_a
      return if text_nodes.empty?
      
      # Проходим по всем текстовым узлам и объединяем/разделяем их
      i = 0
      while i < text_nodes.length
        node = text_nodes[i]
        text = node.text
        
        # Пропускаем пустые узлы
        if text.strip.empty?
          i += 1
          next
        end
        
        # Проверяем, является ли узел полным маркером (начинается с <% и заканчивается %>)
        is_full_marker = text.start_with?('<%') && text.end_with?('%>')
        
        # Если это полный маркер - пропускаем его, не трогаем
        if is_full_marker
          i += 1
          next
        end
        
        # Проверяем, содержит ли узел начало маркера
        if text.include?('<%') && !text.end_with?('%>')
          # Ищем конец маркера в следующих узлах
          j = i + 1
          full_marker = text
          nodes_to_merge = [node]
          
          while j < text_nodes.length && !full_marker.end_with?('%>')
            next_text = text_nodes[j].text
            
            # Если следующий узел тоже начинается с <% - это отдельный маркер, останавливаемся
            break if next_text.start_with?('<%') && !full_marker.end_with?('%>')
            
            full_marker += next_text
            nodes_to_merge << text_nodes[j]
            j += 1
          end
          
          # Если нашли полный маркер
          if full_marker.match?(/<%.*?%>/)
            # Получаем родительский <w:r> узел для первого текстового узла
            first_run = nodes_to_merge.first.parent
            
            # Создаем новый <w:r> элемент
            new_run = Nokogiri::XML::Node.new('w:r', parent_node.document)
            
            # Копируем форматирование из первого <w:r> (если есть)
            rpr = first_run.at_xpath('w:rPr', ns)
            new_run.add_child(rpr.dup) if rpr
            
            # Создаем новый <w:t> с объединённым текстом
            new_text_node = Nokogiri::XML::Node.new('w:t', parent_node.document)
            new_text_node.content = full_marker
            
            # Добавляем <w:t> в новый <w:r>
            new_run.add_child(new_text_node)
            
            # Вставляем новый <w:r> перед первым <w:r>
            first_run.add_previous_sibling(new_run)
            
            # Удаляем старые <w:r> узлы
            nodes_to_merge.each do |text_node|
              run = text_node.parent
              run.remove if run && run.parent
            end
            
            # Обновляем список узлов
            text_nodes = parent_node.xpath('.//w:t', ns).to_a
            i += 1
            next
          end
        end
        
        # Проверяем, содержит ли узел маркер и текст (но не является полным маркером)
        if text.match?(/<%.*?%>/) && !is_full_marker
          # Разделяем узел на части
          parts = text.split(/(<%.*?%>)/).reject(&:empty?)
          
          if parts.length > 1
            # Получаем родительский <w:r> узел
            run = node.parent
            rpr = run.at_xpath('w:rPr', ns)
            
            # Создаем новые <w:r> узлы для каждой части
            new_runs = parts.map do |part|
              new_run = Nokogiri::XML::Node.new('w:r', parent_node.document)
              new_run.add_child(rpr.dup) if rpr
              
              new_text = Nokogiri::XML::Node.new('w:t', parent_node.document)
              new_text.content = part
              new_run.add_child(new_text)
              
              new_run
            end
            
            # Вставляем новые <w:r> узлы перед старым
            new_runs.reverse_each do |new_run|
              run.add_next_sibling(new_run)
            end
            
            # Удаляем старый <w:r>
            run.remove if run && run.parent
            
            # Обновляем список узлов
            text_nodes = parent_node.xpath('.//w:t', ns).to_a
            i += new_runs.length - 1
            next
          end
        end
        
        i += 1
      end
    end

    # ==========================================================================
    # ПОДСТАНОВКА ЗНАЧЕНИЙ (общая логика для Word и Excel)
    # ==========================================================================
    # Подставляет значения из контекста в текст, заменяя маркеры вида <%...%>
    #
    # @param text [String] Исходный текст с маркерами
    # @param context [Hash] Данные для подстановки
    # @param error_behavior [String] Поведение при ошибках ('abort', 'skip_field', 'skip_record')
    # @return [String] Текст с подставленными значениями
    # @raise [TemplateError] если error_behavior='abort' и поле не найдено
    def self.substitute_markers(text, context, error_behavior = 'abort')
      return text unless text.is_a?(String)
      text.gsub(/<%\s*(.*?)\s*%>/) do |match|
        key = $1.strip
        value = if key.include?('.')
                  parts = key.split('.')
                  parts.inject(context) { |h, k| h.is_a?(Hash) ? h[k] : nil }
                else
                  context[key]
                end
        if value.nil? || value.to_s.strip.empty?
          if error_behavior == 'abort'
            raise TemplateError, I18n.t('document_generator.error_field_not_found', field: key)
          else
            ""
          end
        else
          value.to_s
        end
      end
    end

    # ==========================================================================
    # ОБРАБОТКА УСЛОВИЙ С ОТСЛЕЖИВАНИЕМ ПОЗИЦИЙ
    # ==========================================================================
    # Обрабатывает IF/ELSE/END блоки, определяя какие узлы удалить
    #
    # @param full_text [String] Полный текст блока
    # @param node_map [Array<Hash>] Карта узлов с позициями
    # @param context [Hash] Контекст данных
    # @param error_behavior [String] Поведение при ошибках ('abort', 'skip_field', 'skip_record')
    # @return [Hash] Результат с информацией об узлах для удаления и оставшихся
    def self.resolve_conditionals_with_positions(full_text, node_map, context, error_behavior = 'abort')
      nodes_to_remove = []
      remaining_nodes = node_map.dup
      loop do
        match = full_text.match(/<%\s*IF\s*\(\s*(.*?)\s*\)\s*%>/im)
        break unless match
        begin
          condition_str = match[1].strip
          is_true = evaluate_condition(condition_str, context)
        rescue => e
          if error_behavior == 'abort'
            raise TemplateError, I18n.t('document_generator.error_condition_evaluation',
                                         condition: condition_str, message: e.message)
          elsif error_behavior == 'skip_field'
            Rails.logger.warn "[DocumentGenerator] Skipping condition due to error: #{e.message}"
            full_text = full_text[0, match.begin(0)] + full_text[match.end(0)..-1]
            next
          else
            Rails.logger.warn "[DocumentGenerator] Skipping record due to condition error: #{e.message}"
            return { nodes_to_remove: node_map.map { |n| n[:node] }, remaining_nodes: [] }
          end
        end
        if_start = match.begin(0)
        if_end = match.end(0)
        rest = full_text[if_end..-1]
        else_match = rest.match(/<%\s*ELSE\s*%>/im)
        end_match = rest.match(/<%\s*END\s*%>/im)
        unless end_match
          if error_behavior == 'abort'
            raise TemplateError, I18n.t('document_generator.error_missing_end_marker',
                                         condition: condition_str)
          elsif error_behavior == 'skip_field'
            Rails.logger.warn "[DocumentGenerator] Missing END marker, skipping condition block"
            full_text = full_text[0, match.begin(0)] + full_text[match.end(0)..-1]
            next
          else
            Rails.logger.warn "[DocumentGenerator] Missing END marker, skipping record"
            return { nodes_to_remove: node_map.map { |n| n[:node] }, remaining_nodes: [] }
          end
        end
        end_pos = if_end + end_match.begin(0)
        end_len = end_match[0].length
        remove_ranges = []
        keep_range = nil
        if else_match && else_match.begin(0) < end_match.begin(0)
          else_pos = if_end + else_match.begin(0)
          else_len = else_match[0].length
          if is_true
            keep_range = (if_end...else_pos)
            remove_ranges << (if_start...if_end)
            remove_ranges << (else_pos...(else_pos + else_len))
            remove_ranges << (end_pos...(end_pos + end_len))
          else
            keep_range = (else_pos + else_len...end_pos)
            remove_ranges << (if_start...if_end)
            remove_ranges << (else_pos...(else_pos + else_len))
            remove_ranges << (end_pos...(end_pos + end_len))
          end
        else
          if is_true
            keep_range = (if_end...end_pos)
            remove_ranges << (if_start...if_end)
            remove_ranges << (end_pos...(end_pos + end_len))
          else
            remove_ranges << (if_start...(end_pos + end_len))
          end
        end
        remaining_nodes.each do |node_info|
          node_start = node_info[:start_pos]
          node_end = node_info[:end_pos]
          should_remove = false
          remove_ranges.each do |range|
            if node_start >= range.begin && node_end <= range.end
              should_remove = true
              break
            end
          end
          if should_remove
            nodes_to_remove << node_info[:node] unless nodes_to_remove.include?(node_info[:node])
          elsif keep_range
            if node_end <= keep_range.begin || node_start >= keep_range.end
              nodes_to_remove << node_info[:node] unless nodes_to_remove.include?(node_info[:node])
            elsif node_start < keep_range.begin || node_end > keep_range.end
              start_offset = [node_start, keep_range.begin].max - node_start
              end_offset = [node_end, keep_range.end].min - node_start
              new_text = node_info[:text][start_offset...end_offset] || ''
              node_info[:node].content = new_text
              node_info[:text] = new_text
            end
          end
        end
        remaining_nodes = remaining_nodes.reject { |n| nodes_to_remove.include?(n[:node]) }
        current_pos = 0
        remaining_nodes.each do |info|
          info[:start_pos] = current_pos
          info[:end_pos] = current_pos + info[:text].length
          current_pos = info[:end_pos]
        end
        new_full_text = remaining_nodes.map { |n| n[:text] }.join
        break if new_full_text == full_text  # Прогресса нет — выходим
        
        full_text = new_full_text
      end
      {
        nodes_to_remove: nodes_to_remove.uniq,
        remaining_nodes: remaining_nodes
      }
    end

    # ==========================================================================
    # ОБРАБОТКА УСЛОВИЙ НА УРОВНЕ XML-УЗЛОВ
    # ==========================================================================
    # Обрабатывает IF/ELSE/END блоки на уровне XML-узлов, удаляя ненужные узлы
    #
    # @param block_node [Nokogiri::XML::Node] XML-узел (абзац или строка)
    # @param context [Hash] Контекст данных
    # @param ns [Hash] Пространства имен XML
    # @param error_behavior [String] Поведение при ошибках ('abort', 'skip_field', 'skip_record')
    # @return [Nokogiri::XML::Node] Обработанный узел
    def self.process_conditionals_in_block(block_node, context, ns, error_behavior = 'abort')
      text_nodes = block_node.xpath('.//w:t', ns).to_a
      return block_node if text_nodes.empty?
      
      # Собираем полный текст с позициями узлов
      node_map = build_node_map(text_nodes)
      full_text = node_map.map { |n| n[:text] }.join
      
      # Обрабатываем условия
      processed_result = resolve_conditionals_with_positions(full_text, node_map, context, error_behavior)
      
      # Удаляем узлы из отброшенных веток
      processed_result[:nodes_to_remove].each do |node|
        node.remove
      end
      
      # Очищаем оставшиеся узлы от маркеров условий
      processed_result[:remaining_nodes].each do |node_info|
        cleaned = clean_control_markers(node_info[:node].text)
        node_info[:node].content = cleaned
      end
      
      block_node
    end

    # ==========================================================================
    # ПОСТРОЕНИЕ КАРТЫ УЗЛОВ
    # ==========================================================================
    # Создает структуру с информацией о позициях текстовых узлов
    #
    # @param text_nodes [Array<Nokogiri::XML::Node>] Массив текстовых узлов
    # @return [Array<Hash>] Массив хэшей с информацией об узлах
    def self.build_node_map(text_nodes)
      node_map = []
      current_pos = 0
      text_nodes.each do |node|
        text = node.text
        node_map << {
          node: node,
          text: text,
          start_pos: current_pos,
          end_pos: current_pos + text.length
        }
        current_pos += text.length
      end
      node_map
    end

    # ==========================================================================
    # ПОДСТАНОВКА ЗНАЧЕНИЙ В УЗЛЫ
    # ==========================================================================
    # Подставляет значения полей в текстовые узлы
    #
    # @param block_node [Nokogiri::XML::Node] XML-узел
    # @param context [Hash] Контекст данных
    # @param ns [Hash] Пространства имен XML
    # @param error_behavior [String] Поведение при ошибках ('abort', 'skip_field', 'skip_record')
    def self.substitute_in_block(block_node, context, ns, error_behavior = 'abort')
      text_nodes = block_node.xpath('.//w:t', ns)
      normalize_xml_nodes(block_node, ns)
      block_node.xpath('.//w:t', ns).each_with_index do |text_node, idx|
        original = text_node.text
        cleaned = clean_control_markers(original)
        substituted = substitute_markers(cleaned, context, error_behavior)
        text_node.content = substituted if original != substituted
      end
    end

    # ==========================================================================
    # ВЫЧИСЛЕНИЕ УСЛОВИЯ
    # ==========================================================================
    # Вычисляет истинность логического условия на основе переданного контекста данных.
    # Поддерживает операторы сравнения (==, !=), проверку на наличие непустого значения,
    # а также специальную проверку на отсутствие элементов в коллекциях (NOT Subtasks, NOT Watchers, NOT Relations).
    #
    # @param condition_str [String] Строка условия для вычисления 
    #   (например, "Статус == 'Закрыт'", "Назначенный != ''" или "NOT Subtasks")
    # @param context [Hash] Хэш контекста данных, содержащий значения полей текущей записи и вложенных сущностей
    # @return [Boolean] true, если условие истинно; false в противном случае
    def self.evaluate_condition(condition_str, context)
      condition_str = condition_str.strip
      
      # Поддержка проверки на пустоту коллекций (например, <%IF(NOT Subtasks)%>)
      if condition_str =~ /^NOT\s+(.+)/i
        key = $1.strip
        if key == 'Subtasks' || key == 'subtasks'
          return (context['subtasks'] || []).empty?
        elsif key == 'Watchers' || key == 'watchers'
          return (context['watchers'] || []).empty?
        elsif key == 'Relations' || key == 'relations'
          return (context['relations'] || []).empty?
        else
          val = get_context_value(key, context)
          return val.to_s.strip.empty?
        end
      end
      
      # Поддержка оператора равенства (==)
      if condition_str.include?('==')
        left, right = condition_str.split('==', 2).map(&:strip)
        right = right.gsub(/^['"]|['"]$/, '') # Удаляем кавычки из строкового литерала
        left_val = get_context_value(left, context)
        return left_val.to_s.strip == right.to_s.strip
        
      # Поддержка оператора неравенства (!=)
      elsif condition_str.include?('!=')
        left, right = condition_str.split('!=', 2).map(&:strip)
        right = right.gsub(/^['"]|['"]$/, '') # Удаляем кавычки из строкового литерала
        left_val = get_context_value(left, context)
        return left_val.to_s.strip != right.to_s.strip
        
      # Проверка на наличие любого непустого значения (если нет операторов)
      else
        val = get_context_value(condition_str, context)
        return !val.to_s.strip.empty?
      end
    end

    # ==========================================================================
    # ПОЛУЧЕНИЕ ЗНАЧЕНИЯ ИЗ КОНТЕКСТА
    # ==========================================================================
    # Получает значение из контекста по ключу
    #
    # @param key [String] Ключ (поддерживает вложенность через точку)
    # @param context [Hash] Контекст данных
    # @return [Object, nil] Значение или nil
    def self.get_context_value(key, context)
      return nil if key.blank? || context.nil?
      if key.include?('.')
        parts = key.split('.')
        parts.inject(context) { |h, k| h.is_a?(Hash) ? h[k] : nil }
      else
        context[key]
      end
    end

    # ==========================================================================
    # ОЧИСТКА УПРАВЛЯЮЩИХ МАРКЕРОВ
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
    # РАБОТА С АРХИВАМИ
    # ==========================================================================
    # Обрабатывает архив .docx/.xlsx, извлекая указанные XML-файлы
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

    # ==========================================================================
    # ОБРАБОТКА КОЛЛЕКЦИЙ (ПОДЗАДАЧИ, НАБЛЮДАТЕЛИ, СВЯЗИ)
    # ==========================================================================
    # Унифицированный метод для Word и Excel. Разворачивает блоки BEGIN_.../END_...
    # Принимает массив узлов и работает с ним как с единым целым (аналогично process_row_blocks).
    #
    # @param block_nodes [Array<Nokogiri::XML::Node>] Массив XML-узлов (абзацы, строки таблиц)
    # @param context [Hash] Контекст данных
    # @param ns [Hash] Пространства имен XML
    # @param error_behavior [String] Поведение при ошибках ('abort', 'skip_field', 'skip_record')
    # @return [Array<Nokogiri::XML::Node>] Массив узлов с развернутыми коллекциями
    def self.process_collection_blocks(block_nodes, context, ns, error_behavior)
      collections_config = {
        'SUBTASKS' => { context_key: 'subtasks', item_prefix: 'Subtask' },
        'WATCHERS' => { context_key: 'watchers', item_prefix: 'Watcher' },
        'RELATIONS' => { context_key: 'relations', item_prefix: 'Relation' }
      }

      is_word = ns.key?('w')
      text_xpath = is_word ? './/w:t' : './/xmlns:t | .//xmlns:v'

      collections_config.each do |marker_base, config|
        begin_marker = "<%BEGIN_#{marker_base}%>"
        end_marker = "<%END_#{marker_base}%>"

        # Ищем begin и end блоки в пределах ВСЕХ узлов (как в process_row_blocks)
        begin_idx = block_nodes.index { |b| b.xpath(text_xpath, ns).map(&:text).join.include?(begin_marker) }
        end_idx = block_nodes.index { |b| b.xpath(text_xpath, ns).map(&:text).join.include?(end_marker) }

        next unless begin_idx && end_idx

        if begin_idx == end_idx
          # Случай 1: Маркеры находятся внутри одного блока (ячейки или абзаца)
          block = block_nodes[begin_idx]
          block_text = block.xpath(text_xpath, ns).map(&:text).join
          regex = /<%\s*BEGIN_#{marker_base}\s*%>(.*?)<%\s*END_#{marker_base}\s*%>/im
          
          new_block_text = block_text.gsub(regex) do |match|
            inner_template = $1.strip
            collection_data = context[config[:context_key]] || []
            
            if collection_data.empty?
              "" # Блок просто не выводится, если коллекция пуста
            else
              expanded_items = collection_data.map do |item|
                item_context = context.merge(config[:item_prefix] => item)
                substitute_markers(inner_template, item_context, error_behavior)
              end
              expanded_items.join("\n")
            end
          end
          
          # Очищаем старые текстовые узлы и записываем развернутый текст в первый
          block.xpath(text_xpath, ns).each { |n| n.content = '' }
          first_text_node = block.xpath(text_xpath, ns).first
          first_text_node.content = new_block_text if first_text_node
          
        elsif end_idx > begin_idx
          # Случай 2: Маркеры в разных блоках (клонируем целые строки/абзацы)
          template_blocks = block_nodes[(begin_idx + 1)...end_idx]
          
          # Очищаем маркеры из граничных блоков
          begin_block = block_nodes[begin_idx]
          end_block = block_nodes[end_idx]
          begin_block.xpath(text_xpath, ns).each { |n| n.content = clean_control_markers(n.content) }
          end_block.xpath(text_xpath, ns).each { |n| n.content = clean_control_markers(n.content) }
          
          # Удаляем шаблонные блоки и граничные маркеры из массива
          block_nodes = block_nodes[0..begin_idx] + block_nodes[(end_idx + 1)..-1]
          
          collection_data = context[config[:context_key]] || []
          
          if collection_data.empty?
            next
          end
          
          # Разворачиваем коллекцию
          expanded_blocks = []
          collection_data.each do |item|
            item_context = context.merge(config[:item_prefix] => item)
            template_blocks.each do |tmpl_block|
              clone = tmpl_block.dup
              # ВАЖНО: Вызываем process_conditionals_in_block и substitute_in_block ЗДЕСЬ,
              # потому что только здесь доступен правильный item_context с ключом Subtask/Watcher/Relation
              process_conditionals_in_block(clone, item_context, ns, error_behavior)
              substitute_in_block(clone, item_context, ns, error_behavior)
              expanded_blocks << clone
            end
          end
          
          # Вставляем развернутые блоки после begin
          insert_idx = begin_idx + 1
          block_nodes = block_nodes[0...insert_idx] + expanded_blocks + block_nodes[insert_idx..-1]
        end
      end
      
      block_nodes
    end

  end
end
# v2609281204