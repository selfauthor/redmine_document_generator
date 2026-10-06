# frozen_string_literal: true

module DocumentGenerator
  # ============================================================================
  # ОБРАБОТКА ШАБЛОНОВ - БАЗОВЫЙ ПРОЦЕССОР
  # ============================================================================
  # Этот класс отвечает за обработку шаблонов Word и Excel.
  # 
  # ЯЗЫК ШАБЛОНОВ:
  # ==============
  #
  # Все конструкции шаблона имеют вид:
  #
  #   <%выражение%>
  #
  # Имена полей могут содержать символы Unicode. Специального префикса
  # для пользовательских полей нет.
  #
  #
  # 1. ПОЛЯ ЗАДАЧИ
  # ==============
  #
  # <%ИмяПоля%>
  #   Подставляет значение поля текущей задачи.
  #
  # Примеры:
  #
  #   <%ID%>
  #   <%Тема%>
  #   <%Описание%>
  #   <%Статус%>
  #   <%Приоритет%>
  #   <%Автор%>
  #   <%Назначенный%>
  #   <%Дата начала%>
  #   <%Срок%>
  #   <%Готовность%>
  #   <%Оценка времени%>
  #   <%Фактическое время%>
  #   <%Дата создания%>
  #   <%Дата обновления%>
  #   <%Дата закрытия%>
  #   <%Проект%>
  #   <%Трекер%>
  #   <%Категория%>
  #   <%Версия%>
  #
  # Доступны стандартные поля Redmine и пользовательские поля задачи.
  #
  # Пользовательское поле указывается непосредственно по его названию:
  #
  #   <%Сложность%>
  #   <%Срочность%>
  #   <%Почтовый адрес%>
  #
  # Если название пользовательского поля совпадает с названием стандартного
  # поля, приоритет имеет стандартное поле.
  #
  #
  # 2. ПОЛЕ РОДИТЕЛЬСКОЙ ЗАДАЧИ
  # ============================
  #
  # <%Parent.ИмяПоля%>
  #   Подставляет значение указанного поля родительской задачи.
  #
  # Примеры:
  #
  #   <%Parent.ID%>
  #   <%Parent.Тема%>
  #   <%Parent.Статус%>
  #
  #
  # 3. ПОДЗАДАЧИ
  # ============
  #
  # <%BEGIN_SUBTASKS%>
  #   Начало блока повторения по подзадачам текущей задачи.
  #
  # <%END_SUBTASKS%>
  #   Конец блока повторения по подзадачам.
  #
  # Блок между этими командами повторяется для каждой подзадачи.
  #
  # Внутри блока доступны поля с префиксом Subtask.:
  #
  #   <%Subtask.ID%>
  #   <%Subtask.Тема%>
  #   <%Subtask.Статус%>
  #   <%Subtask.Автор%>
  #
  #
  # 4. НАБЛЮДАТЕЛИ
  # ==============
  #
  # <%BEGIN_WATCHERS%>
  #   Начало блока повторения по наблюдателям текущей задачи.
  #
  # <%END_WATCHERS%>
  #   Конец блока повторения по наблюдателям.
  #
  # Внутри блока доступны поля текущего наблюдателя:
  #
  #   <%Watcher.ID%>
  #   <%Watcher.Имя%>
  #   <%Watcher.Фамилия%>
  #   <%Watcher.Email%>
  #
  #
  # 5. СВЯЗИ ЗАДАЧ
  # ==============
  #
  # <%BEGIN_RELATIONS%>
  #   Начало блока повторения по связанным задачам.
  #
  # <%END_RELATIONS%>
  #   Конец блока повторения по связанным задачам.
  #
  # Внутри блока доступны поля связанной задачи:
  #
  #   <%Relation.ID%>
  #   <%Relation.Тема%>
  #   <%Relation.Статус%>
  #   <%Relation.Тип%>
  #
  # Для выборки связей только определённого типа используется:
  #
  # <%BEGIN_RELATIONS:тип%>
  #   Начало блока связей указанного типа.
  #
  # <%END_RELATIONS%>
  #   Конец блока.
  #
  # Примеры:
  #
  #   <%BEGIN_RELATIONS:blocks%>
  #   <%Relation.Тема%>
  #   <%END_RELATIONS%>
  #
  #   <%BEGIN_RELATIONS:relates%>
  #   <%Relation.Тема%>
  #   <%END_RELATIONS%>
  #
  #
  # 6. ПОВТОРЕНИЕ ОСНОВНЫХ ЗАПИСЕЙ
  # ==============================
  #
  # <%BEGIN_ROW%>
  #   Начало блока основной записи.
  #
  # <%END_ROW%>
  #   Конец блока основной записи.
  #
  # Блок между BEGIN_ROW и END_ROW повторяется для каждой основной задачи
  # результата выборки.
  #
  # В Excel при расположении управляющих маркеров в строках таблицы
  # повторяется соответствующая строка.
  #
  # В Word повторяется соответствующий блок документа.
  #
  #
  # 7. УСЛОВИЯ
  # ===========
  #
  # <%IF(условие)%>
  #   Начало условного блока.
  #
  # <%ELSE%>
  #   Необязательная альтернативная ветка.
  #
  # <%END%>
  #   Конец условного блока.
  #
  # Пример:
  #
  #   <%IF(Статус == 'Закрыт')%>
  #   Задача закрыта
  #   <%ELSE%>
  #   Задача открыта
  #   <%END%>
  #
  # В условии можно использовать значения полей текущего контекста.
  #
  #
  # 8. НУМЕРАЦИЯ ЗАПИСЕЙ
  # ====================
  #
  # <%row_number%>
  #   Глобальный порядковый номер основной записи.
  #
  # Нумерация начинается с 1.
  #
  # Номер увеличивается только для основных записей выборки.
  # Подзадачи, наблюдатели и связанные задачи не изменяют row_number.
  #
  # <%row_number_in_group%>
  #   Порядковый номер основной записи внутри группы первого уровня.
  #
  # <%row_number_in_group_2%>
  #   Порядковый номер основной записи внутри группы второго уровня.
  #
  #
  # 9. ГРУППИРОВКА
  # ==============
  #
  # Группировка предназначена для Excel-шаблонов.
  #
  # Первый уровень:
  #
  # <%GROUP_BY:ИмяПоля%>
  #
  #   Задаёт поле первого уровня группировки.
  #
  # Второй уровень:
  #
  # <%GROUP_BY_2:ИмяПоля%>
  #
  #   Задаёт поле второго уровня вложенной группировки.
  #
  # Второй уровень группировки является вложенным в первый.
  # Он НЕ является отдельной группировкой всего набора записей.
  #
  # Допустимы оба варианта расположения команд:
  #
  #   <%GROUP_BY:Поле1%>
  #   <%GROUP_BY_2:Поле2%>
  #   <%GROUP_HEADER%>
  #   <%GROUP_HEADER_2%>
  #
  # и:
  #
  #   <%GROUP_BY:Поле1%>
  #   <%GROUP_HEADER%>
  #   <%GROUP_BY_2:Поле2%>
  #   <%GROUP_HEADER_2%>
  #
  #
  # 10. ЗАГОЛОВОК ГРУППЫ ПЕРВОГО УРОВНЯ
  # ====================================
  #
  # <%GROUP_HEADER%>
  #
  #   Маркер начала содержимого заголовка группы первого уровня.
  #
  # Строка с самим маркером является управляющей и удаляется из результата.
  #
  # В заголовке доступны:
  #
  #   <%GroupValue%>
  #   <%count%>
  #   агрегаты текущей группы
  #   обычные поля контекста группы
  #
  # Пример:
  #
  #   <%GROUP_HEADER%>
  #   Статус: <%GroupValue%>
  #   Количество: <%count%>
  #
  #
  # 11. ЗАГОЛОВОК ГРУППЫ ВТОРОГО УРОВНЯ
  # ====================================
  #
  # <%GROUP_HEADER_2%>
  #
  #   Маркер начала содержимого заголовка вложенной группы второго уровня.
  #
  # В заголовке доступны:
  #
  #   <%GroupValue%>
  #   <%GroupValue2%>
  #   <%count%>
  #   агрегаты текущей группы второго уровня
  #
  #
  # 12. ИТОГ ГРУППЫ ВТОРОГО УРОВНЯ
  # ================================
  #
  # <%GROUP_FOOTER_2%>
  #
  #   Необязательный блок итогов группы второго уровня.
  #
  # Внутри доступны значения и агрегаты соответствующей группы второго уровня.
  #
  #
  # 13. ИТОГ ГРУППЫ
  # ===============
  #
  # <%GROUP_FOOTER%>
  #
  #   Обязательный завершающий блок при использовании любой группировки.
  #
  # GROUP_FOOTER закрывает все уровни текущей группировки.
  #
  # Если используется GROUP_BY или GROUP_BY_2, наличие GROUP_FOOTER обязательно.
  #
  # После GROUP_FOOTER разрешён произвольный статический текст.
  #
  #
  # 14. ОБЩИЕ ИТОГИ
  # ===============
  #
  # <%BEGIN_TOTAL%>
  #   Начало блока общих итогов по всей выборке.
  #
  # <%END_TOTAL%>
  #   Конец блока общих итогов.
  #
  # Общие агрегаты в этом блоке используют префикс total_.
  #
  #
  # 15. КОЛИЧЕСТВО ЗАПИСЕЙ
  # =======================
  #
  # <%count%>
  #
  #   Количество основных записей текущего контейнера повторения.
  #
  # В GROUP_HEADER/GROUP_FOOTER:
  #   количество записей текущей группы первого уровня.
  #
  # В GROUP_HEADER_2/GROUP_FOOTER_2:
  #   количество записей текущей группы второго уровня.
  #
  # В BEGIN_ROW, BEGIN_SUBTASKS, BEGIN_WATCHERS, BEGIN_RELATIONS и
  # вне контейнера группировки использование count запрещено и считается
  # ошибкой шаблона.
  #
  # <%total_count%>
  #
  #   Общее количество основных записей всей выборки.
  #
  #
  # 16. АГРЕГАТНЫЕ ФУНКЦИИ
  # =======================
  #
  # Агрегаты первого уровня:
  #
  # <%sum(ИмяПоля)%>
  #   Сумма значений указанного поля в текущем контейнере.
  #
  # <%avg(ИмяПоля)%>
  #   Среднее арифметическое значений указанного поля.
  #
  # <%min(ИмяПоля)%>
  #   Минимальное значение указанного поля.
  #
  # <%max(ИмяПоля)%>
  #   Максимальное значение указанного поля.
  #
  # Общие агрегаты:
  #
  # <%total_sum(ИмяПоля)%>
  #   Сумма значений указанного поля по всей выборке.
  #
  # <%total_avg(ИмяПоля)%>
  #   Среднее значение по всей выборке.
  #
  # <%total_min(ИмяПоля)%>
  #   Минимальное значение по всей выборке.
  #
  # <%total_max(ИмяПоля)%>
  #   Максимальное значение по всей выборке.
  #
  # count и total_count являются специальными переменными и не являются
  # агрегатами с аргументом поля.
  #
  # Каждая агрегатная команда может присутствовать в шаблоне только один раз.
  # Повторное использование одной и той же агрегатной команды является
  # ошибкой шаблона.
  #
  #
  # 17. ФУНКЦИИ ФОРМАТИРОВАНИЯ
  # ===========================
  #
  # Функции форматирования работают с отдельными значениями.
  # Они НЕ являются агрегатами.
  #
  # <%date(ИмяПоля, 'формат')%>
  #   Форматирует дату или дату/время.
  #
  #   Пример:
  #   <%date(Дата создания, 'DD.MM.YYYY')%>
  #
  # <%now('формат')%>
  #   Возвращает текущие дату и время.
  #
  #   Пример:
  #   <%now('DD.MM.YYYY HH:mm')%>
  #
  # <%upper(ИмяПоля)%>
  #   Переводит значение в верхний регистр.
  #
  # <%lower(ИмяПоля)%>
  #   Переводит значение в нижний регистр.
  #
  # <%capitalize(ИмяПоля)%>
  #   Делает первую букву строки прописной.
  #
  # <%truncate(ИмяПоля, N)%>
  #   Ограничивает строку длиной N символов.
  #
  # <%strip_html(ИмяПоля)%>
  #   Удаляет HTML-теги из значения.
  #
  # <%nl2br(ИмяПоля)%>
  #   Преобразует переводы строк в HTML-теги <br>.
  #
  # <%replace(ИмяПоля, 'старое', 'новое')%>
  #   Заменяет указанную подстроку.
  #
  # <%number(ИмяПоля, N)%>
  #   Форматирует числовое значение с N знаками после десятичного разделителя.
  #
  # <%default(ИмяПоля, 'значение')%>
  #   Возвращает значение поля, если оно заполнено, либо указанное значение
  #   по умолчанию, если поле пустое.
  #
  # <%length(ИмяПоля)%>
  #   Возвращает длину строкового значения.
  #
  # <%concat(аргумент1, аргумент2, ...)%>
  #   Объединяет несколько значений в одну строку.
  #
  # Аргументом concat может быть поле или строковый литерал.
  #
  # Пример:
  #
  #   <%concat(Тема, ' (', ID, ')')%>
  #
  # concat является функцией форматирования и не выполняет группового
  # агрегирования.
  #
  #
  # 18. МЕТАДАННЫЕ ВЫГРУЗКИ
  # ========================
  #
  # <%ExportDate%>
  #   Дата и время формирования документа.
  #
  # <%ExportUser%>
  #   Пользователь Redmine, запустивший выгрузку.
  #
  # <%ProjectName%>
  #   Название текущего проекта.
  #
  # <%QueryName%>
  #   Название сохранённого запроса, если выгрузка выполняется по сохранённому
  #   запросу.
  #
  # <%FilterDescription%>
  #   Текстовое описание применённых условий фильтрации.
  #
  #
  # 19. СТАТИЧЕСКИЙ ТЕКСТ
  # =====================
  #
  # Любой текст, не заключённый в <% ... %>, считается обычным текстом
  # шаблона и переносится в результирующий документ без обработки.
  #
  # После <%GROUP_FOOTER%> разрешён произвольный статический текст.
  #
  #
  # 20. ОГРАНИЧЕНИЯ И ПРАВИЛА
  # ==========================
  #
  # - Стандартное поле имеет приоритет перед пользовательским полем
  #   с таким же названием.
  # - GROUP_BY_2 используется только вместе с GROUP_BY.
  # - При наличии GROUP_BY или GROUP_BY_2 GROUP_FOOTER обязателен.
  # - GROUP_FOOTER закрывает все уровни группировки.
  # - GROUP_FOOTER_2 является необязательным.
  # - count допустим только в контексте группы.
  # - total_count относится ко всей выборке.
  # - row_number относится только к основным записям.
  # - Подзадачи, наблюдатели и связи не изменяют row_number.
  # - concat является функцией форматирования, а не агрегатом.
  # - Каждая агрегатная команда может использоваться в шаблоне только один раз.
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

            # Сохраняем граничные пробелы, если они присутствуют в объединённом тексте.
            new_text_node['xml:space'] = 'preserve' if full_marker.match?(/\A\s|\s\z/)
            
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
              # Сохраняем пробелы в начале и конце узла, чтобы Word не удалял их при отображении.
              new_text['xml:space'] = 'preserve' if part.match?(/\A\s|\s\z/)
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

    # Подставляет значения в маркеры шаблона.
    #
    # Метод различает:
    # - обычные поля;
    # - специальные переменные;
    # - count/total_count;
    # - агрегатные функции;
    # - функции форматирования.
    #
    # @param text [String] Текст, содержащий маркеры.
    # @param context [Hash] Контекст текущей записи или группы.
    # @param error_behavior [String] Стратегия обработки ошибок.
    # @return [String] Текст с подставленными значениями.
    def self.substitute_markers(text, context, error_behavior = 'abort')
      return text unless text.is_a?(String)

      text.gsub(/<%\s*(.*?)\s*%>/m) do
        expression = Regexp.last_match(1).strip

        begin
          # Сначала обрабатываем специальные функции, поскольку их синтаксис
          # отличается от обычного имени поля.
          if expression.match?(/\A[a-z_][a-z0-9_]*\s*\(/i)
            evaluate_template_function(
              expression,
              context,
              error_behavior
            )
          elsif expression.casecmp?('count')
            evaluate_count(
              context,
              error_behavior
            )
          elsif expression.casecmp?('total_count')
            evaluate_total_count(
              context,
              error_behavior
            )
          else
            # Обычный маркер поля.
            value = get_context_value(
              expression,
              context
            )

            if context_value_exists?(expression, context)
              value.nil? ? '' : value.to_s
            else
              handle_missing_template_field(
                expression,
                context,
                error_behavior
              )
            end
          end
        rescue DocumentGenerator::SkipRecordError
          raise
        rescue DocumentGenerator::RenderError
          raise
        rescue DocumentGenerator::TemplateError
          raise
        rescue StandardError => e
          message = I18n.t(
            'document_generator.error_function_failed',
            func: expression,
            message: e.message
          )

          handle_template_error(
            message,
            context,
            error_behavior
          )

          ''
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
      # Выбираем XPath в зависимости от формата документа.
      # Для Word сохраняем прежний поиск по пространству имён w.
      # Для Excel используем локальное имя элемента, не зависящее от префикса.
      text_xpath = ns.key?('w') ? './/w:t' : ".//*[local-name()='t']"

      text_nodes = block_node.xpath(text_xpath, ns).to_a
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
    # Подставляет значения полей в текстовые узлы Word или Excel.
    #
    # @param block_node [Nokogiri::XML::Node] XML-узел обрабатываемого блока.
    # @param context [Hash] Контекст текущей записи.
    # @param ns [Hash] Пространства имён XML документа.
    # @param error_behavior [String] Стратегия обработки ошибок:
    #   'abort', 'skip_field' или 'skip_record'.
    # @return [Nokogiri::XML::Node] Обработанный XML-узел.
    def self.substitute_in_block(block_node, context, ns, error_behavior = 'abort')
      # Word и Excel используют разные XML-пространства имён.
      # Для Word сохраняем существующую нормализацию разбитых маркеров.
      if ns.key?('w')
        normalize_xml_nodes(
          block_node,
          ns
        )

        text_nodes = block_node.xpath(
          './/w:t',
          ns
        ).to_a
      else
        # Excel не должен проходить через Word-нормализацию.
        # После преобразования sharedStrings текст находится непосредственно
        # в элементах <t>.
        text_nodes = block_node.xpath(
          ".//*[local-name()='t']"
        ).to_a
      end

      # Последовательно обрабатываем каждый текстовый узел.
      text_nodes.each do |text_node|
        original = text_node.text

        # Удаляем управляющие маркеры, которые не должны попасть
        # в конечный документ.
        cleaned = clean_control_markers(
          original
        )

        # Подставляем значения обычных полей.
        substituted = substitute_markers(
          cleaned,
          context,
          error_behavior
        )

        next if original == substituted

        text_node.content = substituted

        # Для Word сохраняем правила XML-пробелов.
        # Для Excel этот атрибут также допустим и безвреден.
        if substituted.match?(/\A\s|\s\z/)
          text_node['xml:space'] = 'preserve'
        else
          text_node.remove_attribute('xml:space')
        end
      end

      block_node
    end

    # Разбирает список аргументов функции с учётом кавычек.
    #
    # Обычный String#split(',') здесь не подходит, поскольку запятая
    # может находиться внутри строкового литерала.
    #
    # @param arguments [String] Строка аргументов функции.
    # @return [Array<String>] Массив отдельных аргументов.
    def self.split_function_arguments(arguments)
      result = []
      current = +''
      quote = nil
      escaped = false

      arguments.each_char do |char|
        if escaped
          current << char
          escaped = false
          next
        end

        if char == '\\' && quote
          current << char
          escaped = true
          next
        end

        if quote
          current << char

          if char == quote
            quote = nil
          end

          next
        end

        if char == "'" || char == '"'
          quote = char
          current << char
        elsif char == ','
          result << current.strip
          current = +''
        else
          current << char
        end
      end

      result << current.strip unless current.empty?

      result
    end

    # Выполняет функцию шаблона.
    #
    # @param expression [String] Выражение функции без внешних маркеров <% %>.
    # @param context [Hash] Контекст текущей записи или группы.
    # @param error_behavior [String] Стратегия обработки ошибок.
    # @return [String] Результат выполнения функции.
    def self.evaluate_template_function(expression, context, error_behavior)
      match = expression.match(
        /\A([a-z_][a-z0-9_]*)\s*\((.*)\)\z/im
      )

      unless match
        message = I18n.t(
          'document_generator.error_function_failed',
          func: expression,
          message: I18n.t(
            'document_generator.error_invalid_template',
            message: expression
          )
        )

        return handle_template_error(
          message,
          context,
          error_behavior
        )
      end

      function_name = match[1].downcase
      arguments = split_function_arguments(match[2])

      # Агрегаты обрабатываются отдельно от функций форматирования.
      if %w[sum avg min max total_sum total_avg total_min total_max].include?(
        function_name
      )
        return evaluate_aggregate_function(
          function_name,
          arguments,
          context,
          error_behavior
        )
      end

      case function_name
      when 'date'
        evaluate_date_function(arguments, context)
      when 'now'
        evaluate_now_function(arguments)
      when 'upper'
        evaluate_unary_string_function(arguments, context) { |value| value.upcase }
      when 'lower'
        evaluate_unary_string_function(arguments, context) { |value| value.downcase }
      when 'capitalize'
        evaluate_unary_string_function(arguments, context) { |value| value.capitalize }
      when 'truncate'
        evaluate_truncate_function(arguments, context)
      when 'strip_html'
        evaluate_unary_string_function(arguments, context) do |value|
          ActionController::Base.helpers.strip_tags(value)
        end
      when 'nl2br'
        evaluate_unary_string_function(arguments, context) do |value|
          value.gsub(/\r\n|\r|\n/, '<br>')
        end
      when 'replace'
        evaluate_replace_function(arguments, context)
      when 'number'
        evaluate_number_function(arguments, context)
      when 'default'
        evaluate_default_function(arguments, context)
      when 'length'
        evaluate_unary_string_function(arguments, context) { |value| value.length }
      when 'concat'
        evaluate_concat_function(arguments, context)
      else
        message = I18n.t(
          'document_generator.error_function_failed',
          func: function_name,
          message: I18n.t(
            'document_generator.error_invalid_template',
            message: expression
          )
        )

        handle_template_error(
          message,
          context,
          error_behavior
        )
      end
    end

    # Возвращает количество записей текущего контейнера.
    #
    # count разрешён только там, где ContextBuilder явно сформировал
    # соответствующий контейнерный контекст, например в заголовке или
    # подвале группы.
    #
    # @param context [Hash] Текущий контекст.
    # @param error_behavior [String] Стратегия обработки ошибок.
    # @return [String] Количество записей.
    def self.evaluate_count(context, error_behavior)
      unless context.is_a?(Hash) && context.key?('count')
        message = I18n.t(
          'document_generator.error_field_not_found',
          field: 'count'
        )

        return handle_template_error(
          message,
          context,
          error_behavior
        )
      end

      context['count'].to_i.to_s
    end

    # Возвращает общее количество основных записей выборки.
    #
    # @param context [Hash] Контекст итогового блока.
    # @param error_behavior [String] Стратегия обработки ошибок.
    # @return [String] Общее количество записей.
    def self.evaluate_total_count(context, error_behavior)
      unless context.is_a?(Hash) && context.key?('total_count')
        message = I18n.t(
          'document_generator.error_field_not_found',
          field: 'total_count'
        )

        return handle_template_error(
          message,
          context,
          error_behavior
        )
      end

      context['total_count'].to_i.to_s
    end

    # Выполняет функцию, принимающую один аргумент-значение.
    #
    # @param arguments [Array<String>] Один аргумент функции.
    # @param context [Hash] Текущий контекст.
    # @yield [String] Значение поля для форматирования.
    # @return [String] Отформатированное значение.
    def self.evaluate_unary_string_function(arguments, context)
      return '' if arguments.empty?

      value = resolve_function_value(
        arguments.first,
        context
      )

      value = '' if value.nil?

      result = yield(value.to_s)

      result.to_s
    end

    # Форматирует дату или время согласно формату шаблона.
    #
    # @param arguments [Array<String>] Поле даты и строка формата.
    # @param context [Hash] Текущий контекст.
    # @return [String] Отформатированная дата.
    def self.evaluate_date_function(arguments, context)
      value = resolve_function_value(
        arguments[0],
        context
      )

      return '' if value.nil?

      format = unquote_function_argument(
        arguments[1] || "'DD.MM.YYYY'"
      )

      value = Time.zone.parse(value.to_s) if value.is_a?(String)

      unless value.respond_to?(:strftime)
        return value.to_s
      end

      value.strftime(
        convert_date_format(format)
      )
    end

    # Возвращает текущую дату и время в формате шаблона.
    #
    # @param arguments [Array<String>] Аргументы функции now.
    # @return [String] Текущая дата и время.
    def self.evaluate_now_function(arguments)
      format = unquote_function_argument(
        arguments.first || "'DD.MM.YYYY HH:mm'"
      )

      Time.current.strftime(
        convert_date_format(format)
      )
    end

    # Ограничивает строку заданным количеством символов.
    #
    # @param arguments [Array<String>] Значение и максимальная длина.
    # @param context [Hash] Текущий контекст.
    # @return [String] Обрезанная строка.
    def self.evaluate_truncate_function(arguments, context)
      value = resolve_function_value(
        arguments[0],
        context
      )

      length = arguments[1].to_i

      return '' if value.nil?
      return value.to_s if length <= 0

      value.to_s.truncate(length)
    end

    # Заменяет одну подстроку другой.
    #
    # @param arguments [Array<String>] Значение, искомая и новая строки.
    # @param context [Hash] Текущий контекст.
    # @return [String] Результат замены.
    def self.evaluate_replace_function(arguments, context)
      value = resolve_function_value(
        arguments[0],
        context
      )

      old_value = unquote_function_argument(
        arguments[1] || ''
      )

      new_value = unquote_function_argument(
        arguments[2] || ''
      )

      return '' if value.nil?

      value.to_s.gsub(
        old_value,
        new_value
      )
    end

    # Форматирует числовое значение с указанным количеством знаков.
    #
    # @param arguments [Array<String>] Значение и количество знаков после запятой.
    # @param context [Hash] Текущий контекст.
    # @return [String] Отформатированное число.
    def self.evaluate_number_function(arguments, context)
      value = resolve_function_value(
        arguments[0],
        context
      )

      return '' if value.nil? || value.to_s.strip.empty?

      precision = arguments[1].to_i

      format(
        "%.#{precision}f",
        value.to_f
      )
    end

    # Возвращает значение поля либо значение по умолчанию.
    #
    # @param arguments [Array<String>] Основное значение и значение по умолчанию.
    # @param context [Hash] Текущий контекст.
    # @return [String] Исходное либо резервное значение.
    def self.evaluate_default_function(arguments, context)
      value = resolve_function_value(
        arguments[0],
        context
      )

      default_value = unquote_function_argument(
        arguments[1] || ''
      )

      if value.nil? || value.to_s.strip.empty?
        default_value
      else
        value.to_s
      end
    end

    # Объединяет произвольное количество полей и строковых литералов.
    #
    # @param arguments [Array<String>] Значения и строковые литералы.
    # @param context [Hash] Текущий контекст.
    # @return [String] Объединённая строка.
    def self.evaluate_concat_function(arguments, context)
      arguments.map do |argument|
        argument = argument.strip

        if quoted_function_argument?(argument)
          unquote_function_argument(argument)
        else
          value = get_context_value(
            argument,
            context
          )

          value.nil? ? '' : value.to_s
        end
      end.join
    end

    # Разрешает аргумент функции как литерал или поле контекста.
    #
    # @param argument [String] Аргумент функции.
    # @param context [Hash] Текущий контекст.
    # @return [Object, nil] Значение аргумента.
    def self.resolve_function_value(argument, context)
      argument = argument.to_s.strip

      return unquote_function_argument(argument) if quoted_function_argument?(argument)

      get_context_value(
        argument,
        context
      )
    end

    # Проверяет, является ли аргумент строковым литералом.
    #
    # @param argument [String] Аргумент функции.
    # @return [Boolean] true, если аргумент заключён в одинарные или двойные кавычки.
    def self.quoted_function_argument?(argument)
      argument.match?(/\A(['"]).*\1\z/m)
    end

    # Удаляет внешние кавычки строкового аргумента функции.
    #
    # @param argument [String] Аргумент функции.
    # @return [String] Значение без внешних кавычек.
    def self.unquote_function_argument(argument)
      value = argument.to_s.strip

      if value.length >= 2 &&
         ((value.start_with?("'") && value.end_with?("'")) ||
          (value.start_with?('"') && value.end_with?('"')))
        value[1...-1]
      else
        value
      end
    end

    # Преобразует формат даты из синтаксиса шаблона в формат strftime Ruby.
    #
    # @param format [String] Формат даты в синтаксисе шаблона.
    # @return [String] Формат, совместимый с strftime.
    def self.convert_date_format(format)
      format.to_s
        .gsub('YYYY', '%Y')
        .gsub('YY', '%y')
        .gsub('MM', '%m')
        .gsub('DD', '%d')
        .gsub('HH', '%H')
        .gsub('mm', '%M')
        .gsub('SS', '%S')
    end

    # Проверяет наличие поля в контексте без оценки его значения.
    #
    # @param key [String] Имя поля, включая возможный вложенный путь.
    # @param context [Hash] Текущий контекст.
    # @return [Boolean] true, если поле существует.
    def self.context_value_exists?(key, context)
      parts = key.to_s.split('.')

      current = context

      parts.each do |part|
        return false unless current.is_a?(Hash) && current.key?(part)

        current = current[part]
      end

      true
    end

    # Обрабатывает обращение к отсутствующему полю согласно выбранной стратегии.
    #
    # @param field [String] Имя отсутствующего поля.
    # @param context [Hash] Текущий контекст.
    # @param error_behavior [String] Стратегия обработки ошибок.
    # @return [String] Пустая строка при пропуске поля.
    def self.handle_missing_template_field(field, context, error_behavior)
      issue = context.is_a?(Hash) ? context['__issue'] : nil

      issue_label =
        if issue.respond_to?(:id)
          "##{issue.id} — #{issue.subject}"
        elsif issue.is_a?(Hash)
          "##{issue['id']} — #{issue['subject']}"
        else
          I18n.t('document_generator.unknown_record')
        end

      message = I18n.t(
        'document_generator.error_field_missing',
        issue: issue_label,
        field: field
      )

      handle_template_error(
        message,
        context,
        error_behavior
      )
    end

    # Применяет стратегию обработки ошибки шаблона.
    #
    # @param message [String] Локализованное сообщение об ошибке.
    # @param context [Hash] Текущий контекст.
    # @param error_behavior [String] Стратегия обработки ошибки.
    # @return [String, nil] Пустое значение при пропуске поля.
    # @raise [DocumentGenerator::RenderError] При режиме abort.
    # @raise [DocumentGenerator::SkipRecordError] При режиме skip_record.
    def self.handle_template_error(message, context, error_behavior)
      case error_behavior
      when 'abort'
        raise DocumentGenerator::RenderError, message

      when 'skip_field'
        warnings = context.is_a?(Hash) ? context['__warnings'] : nil
        warnings << message if warnings.is_a?(Array)

        Rails.logger.warn(
          "[DocumentGenerator] Template field was skipped: #{message}"
        )

        ''

      when 'skip_record'
        warnings = context.is_a?(Hash) ? context['__warnings'] : nil
        warnings << message if warnings.is_a?(Array)

        raise DocumentGenerator::SkipRecordError, message

      else
        raise DocumentGenerator::RenderError, message
      end
    end

    # Получает предварительно рассчитанное агрегатное значение из контекста.
    #
    # @param function_name [String] Имя агрегата.
    # @param arguments [Array<String>] Аргументы функции.
    # @param context [Hash] Контекст текущей записи или группы.
    # @param error_behavior [String] Стратегия обработки ошибок.
    # @return [String] Значение агрегата.
    def self.evaluate_aggregate_function(function_name, arguments, context, error_behavior)
      unless arguments.length == 1 && !arguments.first.to_s.strip.empty?
        message = I18n.t(
          'document_generator.error_function_failed',
          func: function_name,
          message: I18n.t(
            'document_generator.error_invalid_template',
            message: function_name
          )
        )

        return handle_template_error(
          message,
          context,
          error_behavior
        )
      end

      field_name = arguments.first.strip
      is_total = function_name.start_with?('total_')
      aggregate_name = function_name.sub(/\Atotal_/, '')

      # Общие агрегаты используют total_agg_.
      # Вложенная группа второго уровня использует group_2_agg_.
      # Первая группа использует group_agg_.
      prefix =
        if is_total
          'total_agg_'
        elsif context.key?('GroupValue2')
          'group_2_agg_'
        else
          'group_agg_'
        end

      key = "#{prefix}#{aggregate_name}_#{field_name}"

      unless context.key?(key)
        message = I18n.t(
          'document_generator.error_unknown_field_in_function',
          field: field_name
        )

        return handle_template_error(
          message,
          context,
          error_behavior
        )
      end

      value = context[key]

      value.nil? ? '' : value.to_s
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

    # Получает значение из текущего контекста по имени поля.
    #
    # Поддерживает вложенные ключи через точку:
    # Parent.Тема, Subtask.Тема, Watcher.Имя и Relation.Тема.
    #
    # @param key [String] Имя поля или вложенного значения.
    # @param context [Hash] Текущий контекст.
    # @return [Object, nil] Найденное значение.
    def self.get_context_value(key, context)
      return nil if key.blank? || context.nil?

      if key.include?('.')
        parts = key.split('.')

        parts.inject(context) do |current, part|
          current.is_a?(Hash) ? current[part] : nil
        end
      else
        context[key]
      end
    end

    # Удаляет управляющие команды из текста после их обработки.
    #
    # @param text [String] Текст XML-узла.
    # @return [String] Текст без управляющих команд.
    def self.clean_control_markers(text)
      return text unless text.is_a?(String)

      text.gsub(
        /<%\s*(
          BEGIN_ROW|
          END_ROW|
          BEGIN_SUBTASKS|
          END_SUBTASKS|
          BEGIN_WATCHERS|
          END_WATCHERS|
          BEGIN_RELATIONS(?:\s*:\s*[^%]+)?|
          END_RELATIONS|
          GROUP_BY(?:_2)?\s*:\s*[^%]+|
          GROUP_HEADER(?:_2)?|
          GROUP_FOOTER(?:_2)?|
          BEGIN_TOTAL|
          END_TOTAL|
          IF(?:\s*\([^%]*\))?|
          ELSE|
          END
        )\s*%>/ix,
        ''
      )
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
    # Разворачивает блоки коллекций подзадач, наблюдателей и связей.
    #
    # @param block_nodes [Array<Nokogiri::XML::Node>, Nokogiri::XML::Node] XML-узлы блока.
    # @param context [Hash] Контекст текущей задачи.
    # @param ns [Hash] Пространства имён XML.
    # @param error_behavior [String] Стратегия обработки ошибок.
    # @return [Array<Nokogiri::XML::Node>] Обработанные XML-узлы.
    def self.process_collection_blocks(block_nodes, context, ns, error_behavior)
      # Приводим одиночный XML-узел к массиву.
      block_nodes = [block_nodes] unless block_nodes.is_a?(Array)

      # Описываем поддерживаемые коллекции.
      collections_config = {
        'SUBTASKS' => {
          context_key: 'subtasks',
          item_prefix: 'Subtask'
        },
        'WATCHERS' => {
          context_key: 'watchers',
          item_prefix: 'Watcher'
        },
        'RELATIONS' => {
          context_key: 'relations',
          item_prefix: 'Relation'
        }
      }

      # Выбираем XPath текстовых узлов в зависимости от формата документа.
      is_word = ns.key?('w')

      text_xpath =
        if is_word
          './/w:t'
        else
          ".//*[local-name()='t']"
        end

      collections_config.each do |marker_base, config|
        # Для отношений допускается дополнительный фильтр:
        # BEGIN_RELATIONS:blocks ... END_RELATIONS
        if marker_base == 'RELATIONS'
          begin_regex =
            /<%\s*BEGIN_RELATIONS(?:\s*:\s*([^%]+?))?\s*%>/i

          end_regex =
            /<%\s*END_RELATIONS\s*%>/i
        else
          begin_regex =
            /<%\s*BEGIN_#{Regexp.escape(marker_base)}\s*%>/i

          end_regex =
            /<%\s*END_#{Regexp.escape(marker_base)}\s*%>/i
        end

        # Собираем текст каждого XML-узла.
        node_texts = block_nodes.map do |node|
          node.xpath(
            text_xpath,
            ns
          ).map(&:text).join
        end

        full_text = node_texts.join

        # Ищем начало коллекционного блока.
        begin_match = full_text.match(
          begin_regex
        )

        next unless begin_match

        # Ищем конец только после начала блока.
        end_match = full_text.match(
          end_regex,
          begin_match.end(0)
        )

        next unless end_match

        # Для BEGIN_RELATIONS запоминаем необязательный тип связи.
        relation_type =
          if marker_base == 'RELATIONS'
            begin_match[1]&.strip
          end

        # Определяем диапазон текста каждого XML-узла.
        node_ranges = []
        current_position = 0

        node_texts.each do |node_text|
          node_ranges << {
            start: current_position,
            finish: current_position + node_text.length
          }

          current_position += node_text.length
        end

        # Определяем XML-узел, содержащий BEGIN.
        begin_idx = node_ranges.index do |range|
          begin_match.begin(0) >= range[:start] &&
            begin_match.begin(0) < range[:finish]
        end

        # Определяем XML-узел, содержащий END.
        end_idx = node_ranges.index do |range|
          end_match.begin(0) >= range[:start] &&
            end_match.begin(0) < range[:finish]
        end

        next unless begin_idx && end_idx

        # Получаем исходную коллекцию текущей задачи.
        collection_data =
          context[config[:context_key]] || []

        # Для BEGIN_RELATIONS:тип оставляем только связи указанного типа.
        if marker_base == 'RELATIONS' && relation_type.present?
          relation_type_key = I18n.t(
            'document_generator.relation_type',
            default: 'Relation Type'
          )

          collection_data = collection_data.select do |item|
            item.is_a?(Hash) &&
              item[relation_type_key].to_s.casecmp?(relation_type)
          end
        end

        # Если BEGIN и END находятся в одном XML-узле,
        # обрабатываем коллекцию непосредственно внутри текста.
        if begin_idx == end_idx
          block = block_nodes[begin_idx]
          block_text = node_texts[begin_idx]

          collection_regex =
            if marker_base == 'RELATIONS'
              /<%\s*BEGIN_RELATIONS(?:\s*:\s*[^%]+?)?\s*%>(.*?)<%\s*END_RELATIONS\s*%>/im
            else
              /<%\s*BEGIN_#{Regexp.escape(marker_base)}\s*%>(.*?)<%\s*END_#{Regexp.escape(marker_base)}\s*%>/im
            end

          new_block_text = block_text.gsub(
            collection_regex
          ) do
            inner_template = Regexp.last_match(1)

            collection_data.map do |item|
              # Формируем контекст текущего элемента коллекции.
              item_context = context.merge(
                config[:item_prefix] => item
              )

              begin
                # Подставляем поля текущего элемента.
                substitute_markers(
                  inner_template,
                  item_context,
                  error_behavior
                )
              rescue DocumentGenerator::SkipRecordError
                # Пропускаем только текущий элемент коллекции.
                item_issue =
                  item.is_a?(Hash) ? item['__issue'] : nil

                Rails.logger.warn(
                  "[DocumentGenerator] Collection item was skipped; issue ##{item_issue&.id || 'unknown'}."
                )

                ''
              end
            end.join("\n")
          end

          # Сохраняем результат в первом текстовом узле.
          text_nodes = block.xpath(
            text_xpath,
            ns
          )

          text_nodes.each do |node|
            node.content = ''
          end

          text_nodes.first.content = new_block_text if text_nodes.first

          next
        end

        # END не может находиться раньше BEGIN.
        if end_idx < begin_idx
          raise TemplateError,
                I18n.t('document_generator.error_row_block_mismatch')
        end

        begin_node = block_nodes[begin_idx]
        end_node = block_nodes[end_idx]

        begin_range = node_ranges[begin_idx]
        end_range = node_ranges[end_idx]

        # Ограничивает текст XML-узла указанным диапазоном.
        trim_node_to_range = lambda do |node, node_start, range_start, range_end|
          local_position = 0
          retained_text = +''

          node.xpath(
            text_xpath,
            ns
          ).each do |text_node|
            original_text = text_node.text

            text_start = node_start + local_position
            text_end = text_start + original_text.length

            keep_start = [
              text_start,
              range_start
            ].max

            keep_end = [
              text_end,
              range_end
            ].min

            if keep_end > keep_start
              fragment =
                original_text[
                  (keep_start - text_start)...(keep_end - text_start)
                ] || ''
            else
              fragment = ''
            end

            text_node.content = fragment
            retained_text << fragment

            local_position += original_text.length
          end

          retained_text.strip
        end

        # Сохраняем текст до BEGIN, если он находится
        # в том же XML-узле.
        prefix_node = begin_node.dup

        prefix_text = trim_node_to_range.call(
          prefix_node,
          begin_range[:start],
          begin_range[:start],
          begin_match.begin(0)
        )

        # Сохраняем текст после END, если он находится
        # в том же XML-узле.
        suffix_node = end_node.dup

        suffix_text = trim_node_to_range.call(
          suffix_node,
          end_range[:start],
          end_match.end(0),
          end_range[:finish]
        )

        prefix_nodes =
          prefix_text.empty? ? [] : [prefix_node]

        suffix_nodes =
          suffix_text.empty? ? [] : [suffix_node]

        # Формируем расширенный блок для всех элементов коллекции.
        expanded_blocks = []

        collection_data.each do |item|
          # Формируем контекст текущего элемента коллекции.
          item_context = context.merge(
            config[:item_prefix] => item
          )

          item_blocks = []

          begin
            (begin_idx..end_idx).each do |node_idx|
              template_node = block_nodes[node_idx]
              original_range = node_ranges[node_idx]

              # Клонируем XML-узел.
              clone = template_node.dup

              # Оставляем только содержимое между BEGIN и END.
              retained_text = trim_node_to_range.call(
                clone,
                original_range[:start],
                begin_match.end(0),
                end_match.begin(0)
              )

              # Пустые граничные узлы не добавляем.
              if (node_idx == begin_idx || node_idx == end_idx) &&
                 retained_text.empty?
                next
              end

              # Обрабатываем вложенные условия.
              process_conditionals_in_block(
                clone,
                item_context,
                ns,
                error_behavior
              )

              # Подставляем поля текущего элемента.
              substitute_in_block(
                clone,
                item_context,
                ns,
                error_behavior
              )

              item_blocks << clone
            end

            # Добавляем элемент только после полной успешной обработки.
            expanded_blocks.concat(
              item_blocks
            )
          rescue DocumentGenerator::SkipRecordError
            # Ошибка относится только к текущему элементу коллекции.
            item_issue =
              item.is_a?(Hash) ? item['__issue'] : nil

            Rails.logger.warn(
              "[DocumentGenerator] Collection item was skipped; issue ##{item_issue&.id || 'unknown'}."
            )

            next
          end
        end

        # Заменяем исходный блок расширенным содержимым.
        block_nodes =
          block_nodes[0...begin_idx] +
          prefix_nodes +
          expanded_blocks +
          suffix_nodes +
          Array(block_nodes[(end_idx + 1)..-1])
      end

      block_nodes
    end

  end
end
# v2610061505