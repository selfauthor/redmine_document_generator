# frozen_string_literal: true

require 'sablon'
require 'zip'
require 'fileutils'
require 'nokogiri'

module DocumentGenerator
  # Класс отвечает за генерацию документа Word (.docx) на основе шаблона и данных.
  # Использует библиотеку Sablon для сохранения форматирования OOXML.
  class WordRenderer
    # Инициализация рендерера
    # @param template_path [String] Путь к временному файлу шаблона
    # @param issues [ActiveRecord::Relation] Выборка задач для выгрузки
    # @param parser_config [Hash] Конфигурация, полученная от TemplateParser
    # @param error_behavior [String] Стратегия обработки ошибок ('abort', 'skip_field', 'skip_record')
    def initialize(template_path, issues, parser_config, error_behavior)
      @template_path = template_path
      @issues = issues
      @parser_config = parser_config
      @error_behavior = error_behavior
    end

    # Основной метод генерации документа
    # @return [String] Бинарные данные сгенерированного документа .docx
    def render
      # Шаг 1: Формируем хэш данных (контекст), понятный библиотеке Sablon
      context = build_context

      # Шаг 2: Создаем временную копию шаблона, заменяя наши маркеры <%...%> на маркеры Sablon «...»
      processed_template_path = preprocess_template

      # Шаг 3: Выполняем рендеринг через Sablon
      template = Sablon.template(processed_template_path)
      doc = template.render(context)

      doc
    ensure
      # Гарантированная очистка временного файла после рендеринга
      FileUtils.rm_f(processed_template_path) if processed_template_path && File.exist?(processed_template_path)
    end

    private

    # Определяет, нужно ли строить контекст с группировкой или плоский
    # @return [Hash] Контекст данных для Sablon
    def build_context
      if @parser_config[:group_by]
        build_grouped_context
      else
        build_flat_context
      end
    end

    # Строит плоский контекст данных (режим "Один файл на запись" или без группировки)
    # @return [Hash] Контекст с массивом записей и общими итогами
    def build_flat_context
      records = @issues.map { |issue| build_issue_hash(issue) }
      base_context(records).merge(
        records: records,
        total_count: records.size,
        totals: [{ total_count: records.size }.merge(calculate_aggregates(@issues, 'total_'))]
      )
    end

    # Строит иерархический контекст данных для режимов с группировкой (1 или 2 уровня)
    # @return [Hash] Контекст с вложенной структурой groups -> records
    def build_grouped_context
      group_field = @parser_config[:group_by]
      group_field_2 = @parser_config[:group_by_2]

      resolved_group = FieldResolver.resolve(group_field)
      resolved_group_2 = group_field_2 ? FieldResolver.resolve(group_field_2) : nil

      # Группируем задачи по значениям указанных полей
      groups_data = @issues.group_by do |issue|
        val1 = FieldResolver.get_value(issue, resolved_group)
        if resolved_group_2
          val2 = FieldResolver.get_value(issue, resolved_group_2)
          [val1, val2]
        else
          val1
        end
      end

      groups_array = []
      total_count = 0

      groups_data.each do |group_key, group_issues|
        # Нормализуем ключ группы в массив (для поддержки 1 или 2 уровней)
        group_key_array = group_key.is_a?(Array) ? group_key : [group_key, nil]
        
        records = group_issues.map { |issue| build_issue_hash(issue) }
        # Добавляем порядковый номер записи внутри группы
        records = records.map.with_index(1) { |r, i| r.merge('row_number_in_group' => i) }

        group_aggregates = calculate_aggregates(group_issues, 'group_')

        groups_array << {
          'GroupValue' => group_key_array[0],
          'GroupValue2' => group_key_array[1],
          'count' => records.size,
          'records' => records
        }.merge(group_aggregates)

        total_count += records.size
      end

      base_context(groups_array.flatten).merge(
        groups: groups_array,
        total_count: total_count,
        totals: [{ total_count: total_count }.merge(calculate_aggregates(@issues, 'total_'))]
      )
    end

    # Формирует базовый контекст с мета-данными выгрузки, общий для всех режимов
    # @param _data_sample [Array] Пример данных (не используется, оставлен для совместимости сигнатуры)
    # @return [Hash] Базовые мета-данные
    def base_context(_data_sample)
      first_issue = @issues.first
      {
        'ExportDate' => Time.now.strftime('%d.%m.%Y %H:%M'),
        'ExportUser' => User.current.name,
        'ProjectName' => first_issue&.project&.name || '',
        'QueryName' => '',
        'FilterDescription' => ''
      }
    end

    # Преобразует объект Issue в хэш, ключи которого соответствуют именам полей в шаблоне
    # @param issue [Issue] Объект задачи Redmine
    # @return [Hash] Хэш с данными задачи
    def build_issue_hash(issue)
      hash = {}
      
      # Словарь стандартных полей: ключ - внутренний идентификатор Redmine (латиница)
      standard_fields_map = {
        'id' => issue.id,
        'subject' => issue.subject,
        'description' => issue.description ? Sablon::Content::Html.new(issue.description) : '',
        'status' => issue.status&.name,
        'priority' => issue.priority&.name,
        'author' => issue.author&.name,
        'assigned_to' => issue.assigned_to&.name,
        'start_date' => issue.start_date,
        'due_date' => issue.due_date,
        'done_ratio' => "#{issue.done_ratio}%",
        'estimated_hours' => issue.estimated_hours,
        'spent_hours' => issue.spent_hours,
        'created_on' => issue.created_on,
        'updated_on' => issue.updated_on,
        'closed_on' => issue.closed_on,
        'project' => issue.project&.name,
        'tracker' => issue.tracker&.name,
        'category' => issue.category&.name,
        'fixed_version' => issue.fixed_version&.name,
        'parent_id' => issue.parent_id
      }

      # Заполняем хэш, используя локализованные имена полей из языкового файла Redmine в качестве ключей.
      # Это гарантирует, что ключ в хэше совпадет с тем, что пользователь написал в шаблоне (например, "Тема").
      standard_fields_map.each do |field_key, value|
        localized_name = I18n.t("field_#{field_key}", default: field_key.humanize)
        hash[localized_name] = value
      end

      # Поля родительской задачи. Префикс 'Parent.' остается на латинице, имя поля берется из локализации.
      if issue.parent
        parent_prefix = 'Parent.'
        hash["#{parent_prefix}#{I18n.t('field_subject', default: 'Subject')}"] = issue.parent.subject
        hash["#{parent_prefix}#{I18n.t('field_status', default: 'Status')}"] = issue.parent.status&.name
        hash["#{parent_prefix}#{I18n.t('field_assigned_to', default: 'Assigned to')}"] = issue.parent.assigned_to&.name
      end

      # Пользовательские поля: имя берется напрямую из custom_field.name (оно задается администратором в интерфейсе)
      issue.custom_field_values.each do |cfv|
        hash[cfv.custom_field.name] = format_custom_value(cfv)
      end

      # Связанные сущности: ключи массивов на английском (согласно спецификации шаблонов), 
      # внутренние ключи полей берутся из локализации для единообразия.
      hash['subtasks'] = issue.children.map do |child|
        {
          'ID' => child.id,
          I18n.t('field_subject', default: 'Subject') => child.subject,
          I18n.t('field_status', default: 'Status') => child.status&.name,
          I18n.t('field_assigned_to', default: 'Assigned to') => child.assigned_to&.name
        }
      end

      hash['relations'] = issue.relations.map do |rel|
        target = rel.issue_to_id == issue.id ? rel.issue_from : rel.issue_to
        next nil unless target
        {
          I18n.t('document_generator.relation_type', default: 'Relation Type') => rel.relation_type_for(issue),
          I18n.t('field_subject', default: 'Subject') => target.subject,
          I18n.t('field_status', default: 'Status') => target.status&.name
        }
      end.compact

      hash['watchers'] = issue.watchers.map do |w|
        { I18n.t('document_generator.watcher_name', default: 'Name') => w.user&.name }
      end

      # Вычисляем значения для функций форматирования, упомянутых в тексте шаблона
      evaluate_template_functions(@parser_config[:template_text] || '', issue, hash)

      hash
    end

    # Анализирует текст шаблона и вычисляет значения функций форматирования (date, upper, default и т.д.)
    # @param template_text [String] Исходный текст шаблона
    # @param issue [Issue] Объект задачи
    # @param hash [Hash] Хэш данных, который дополняется вычисленными значениями
    def evaluate_template_functions(template_text, issue, hash)
      # Регулярное выражение ищет конструкции вида <%func(arg1, arg2)%>
      template_text.scan(/<%\s*([a-z_]+)\s*\((.*?)\)\s*%>/i) do |func, args_str|
        args = args_str.split(',').map { |a| a.strip.gsub(/^['"]|['"]$/, '') }
        func_name = func.downcase
        
        case func_name
        when 'date'
          val = get_field_value(issue, args[0])
          hash["#{args[0]}_formatted"] = val.is_a?(Date) || val.is_a?(Time) ? val.strftime(args[1]) : val
        when 'now'
          hash["now_formatted"] = Time.now.strftime(args[0])
        when 'upper'
          val = get_field_value(issue, args[0])
          hash["#{args[0]}_upper"] = val.to_s.upcase
        when 'lower'
          val = get_field_value(issue, args[0])
          hash["#{args[0]}_lower"] = val.to_s.downcase
        when 'default'
          val = get_field_value(issue, args[0])
          hash["#{args[0]}_default"] = val.present? ? val : args[1]
        when 'strip_html'
          val = get_field_value(issue, args[0])
          hash["#{args[0]}_stripped"] = val.to_s.gsub(/<[^>]*>/, '')
        end
      end
    end

    # Получает значение поля из задачи по его имени из шаблона
    # @param issue [Issue] Объект задачи
    # @param field_name [String] Имя поля из шаблона
    # @return [Object] Значение поля
    def get_field_value(issue, field_name)
      resolved = FieldResolver.resolve(field_name)
      FieldResolver.get_value(issue, resolved)
    end

    # Форматирует значение пользовательского поля (например, объединяет массивы через запятую)
    # @param cfv [CustomValue] Объект значения пользовательского поля
    # @return [String, Object] Отформатированное значение
    def format_custom_value(cfv)
      value = cfv.value
      value.is_a?(Array) ? value.join(', ') : value
    end

    # Вычисляет агрегатные функции (sum, avg, min, max, count, concat) для группы или всей выборки.
    # Ключи хэша формируются с использованием оригинальных имён полей — Ruby-хэши корректно
    # работают с любыми строковыми ключами, включая кириллицу и пробелы.
    # @param issues [Array<Issue>] Массив задач для агрегации
    # @param prefix [String] Префикс для ключа в хэше ('group_' или 'total_')
    # @return [Hash] Хэш с вычисленными агрегатами
    def calculate_aggregates(issues, prefix)
      aggregates = {}
      (@parser_config[:aggregates] || []).each do |req|
        func = req[:func]
        field = req[:field]
        
        # Используем оригинальное имя поля без санитизации.
        # Ruby-хэши и Sablon корректно обрабатывают любые Unicode-символы в ключах.
        key = "#{prefix}agg_#{func}_#{field}"
        
        aggregates[key] = AggregateCalculator.new(issues).calculate(func, field)
      end
      aggregates
    end

    # Предварительная обработка шаблона: замена кастомных маркеров <%...%> на маркеры Sablon «...».
    # Создаёт временную копию файла шаблона и модифицирует XML-содержимое документов,
    # колонтитулов Word, чтобы Sablon смог корректно их обработать.
    # @return [String] Путь к обработанному временному файлу шаблона
    def preprocess_template
      # Создаём временную копию шаблона, чтобы не модифицировать исходный файл
      temp_path = "#{@template_path}.preprocessed.docx"
      FileUtils.cp(@template_path, temp_path)
      
      # Собираем весь текстовый контент шаблона для последующего анализа функций форматирования
      template_text = ""

      Zip::File.open(temp_path) do |zip|
        # Находим все XML-файлы, содержащие текст: основной документ и все колонтитулы.
        # Колонтитулы обрабатываются отдельно, так как они хранятся в отдельных XML-файлах.
        xml_entries = zip.glob('word/document.xml') +
                      zip.glob('word/header*.xml') +
                      zip.glob('word/footer*.xml')
        
        xml_entries.each do |entry|
          xml_content = entry.get_input_stream.read
          
          # Извлекаем текст для анализа функций форматирования (date, upper и т.д.)
          template_text += " " + extract_xml_text(xml_content)

          # Главная замена: преобразуем кастомные маркеры <%...%> в маркеры Sablon «...».
          # Регулярное выражение захватывает всё содержимое между <% и %>, игнорируя пробелы по краям.
          xml_content = xml_content.gsub(/<%\s*([^%]+?)\s*%>/) do |match|
            inner = $1.strip
            
            # --- Блок 1: Функции форматирования ---
            # Преобразуем функции вида <%date(Дата создания, 'DD.MM.YYYY')%> в «date_Дата создания, 'DD.MM.YYYY'».
            # Ключ хэша формируется в evaluate_template_functions, здесь должно быть точное соответствие.
            if inner.match?(/^(date|now|upper|lower|default|strip_html)\s*\(/i)
              func = inner[/^([a-z_]+)/i, 1].downcase
              args = inner[/\((.*)\)$/, 1]
              "«#{func}_#{args}»"
            
            # --- Блок 2: Агрегатные функции ---
            # Преобразуем <%sum(Оценка времени)%> в «agg_sum_Оценка времени».
            # Ключ хэша формируется в calculate_aggregates, здесь должно быть точное соответствие.
            elsif inner.match?(/^(sum|avg|min|max|count|concat)\s*\(/i)
              func = inner[/^([a-z_]+)/i, 1].downcase
              field = inner[/\((.*?)\)$/, 1].strip
              "«agg_#{func}_#{field}»"
            
            # --- Блок 3: Условные блоки ---
            # Преобразуем <%IF(Назначенный)%> в «if Назначенный» (синтаксис Sablon).
            elsif inner.match?(/^IF\((.*?)\)$/i)
              "«if #{$1.strip}»"
            elsif inner.match?(/^ELSE$/i)
              "«else»"
            elsif inner.match?(/^END$/i)
              "«end»"
            
            # --- Блок 4: Циклы и блоки данных ---
            # Преобразуем блочные маркеры в синтаксис Sablon для итерации по массивам контекста.
            # «tr records» — команда Sablon для клонирования строки таблицы для каждого элемента массива.
            
            # Основная строка данных (повторяется для каждой записи)
            elsif inner.match?(/^BEGIN_ROW$/i)
              "«tr records»"
            elsif inner.match?(/^END_ROW$/i)
              "«end»"
            
            # Цикл по подзадачам
            elsif inner.match?(/^BEGIN_SUBTASKS$/i)
              "«tr subtasks»"
            elsif inner.match?(/^END_SUBTASKS$/i)
              "«end»"
            
            # Цикл по наблюдателям
            elsif inner.match?(/^BEGIN_WATCHERS$/i)
              "«tr watchers»"
            elsif inner.match?(/^END_WATCHERS$/i)
              "«end»"
            
            # Цикл по связанным записям
            elsif inner.match?(/^BEGIN_RELATIONS$/i)
              "«tr relations»"
            elsif inner.match?(/^END_RELATIONS$/i)
              "«end»"
            
            # --- Блок 5: Блоки группировки ---
            # Группировка обрабатывается как цикл по массиву groups в контексте.
            # Заголовок группы — начало цикла, итоги группы — конец цикла.
            
            # Заголовок группы 1-го уровня (начало цикла по группам)
            elsif inner.match?(/^BEGIN_GROUP_HEADER$/i)
              "«tr groups»"
            elsif inner.match?(/^END_GROUP_HEADER$/i)
              ""
            
            # Итоги группы 1-го уровня (конец цикла по группам)
            elsif inner.match?(/^BEGIN_GROUP_FOOTER$/i)
              ""
            elsif inner.match?(/^END_GROUP_FOOTER$/i)
              "«end»"
            
            # --- Блок 6: Общие итоги ---
            # Обрабатываются как цикл по массиву totals (всегда 1 элемент).
            elsif inner.match?(/^BEGIN_TOTAL$/i)
              "«tr totals»"
            elsif inner.match?(/^END_TOTAL$/i)
              "«end»"
            
            # --- Блок 7: Вложенные сущности ---
            # Удаляем префикс (Subtask., Relation., Watcher.), чтобы Sablon искал ключ
            # внутри текущего элемента цикла. Например, <%Subtask.Тема%> становится «Тема»,
            # и Sablon ищет ключ 'Тема' в хэше текущей подзадачи из массива subtasks.
            elsif inner.match?(/^Subtask\.(.+)$/i)
              "«#{$1.strip}»"
            elsif inner.match?(/^Relation\.(.+)$/i)
              "«#{$1.strip}»"
            elsif inner.match?(/^Watcher\.(.+)$/i)
              "«#{$1.strip}»"
            
            # --- Блок 8: Обычные поля данных ---
            # Все остальные маркеры просто преобразуются в маркеры Sablon без изменений содержимого.
            else
              "«#{inner}»"
            end
          end
          
          # Перезаписываем обработанный XML обратно в архив (временный файл шаблона)
          entry.get_output_stream.write(xml_content)
        end
      end
      
      # Сохраняем извлечённый текст для последующего анализа функций форматирования
      # в методе build_issue_hash (evaluate_template_functions)
      @parser_config[:template_text] = template_text
      temp_path
    end

    # Извлекает текстовое содержимое из XML-узлов Word для анализа маркеров
    # @param xml_content [String] Исходная строка XML
    # @return [String] Извлеченный текст
    def extract_xml_text(xml_content)
      doc = Nokogiri::XML(xml_content)
      doc.xpath('//w:t', 'w' => 'http://schemas.openxmlformats.org/wordprocessingml/2006/main').map(&:text).join(' ')
    end
  end
end