# frozen_string_literal: true

module DocumentGenerator
  # Класс отвечает за подготовку структурированных данных (контекста) 
  # из массива записей для последующей передачи в рендереры (Word/Excel).
  class ContextBuilder
    # @param issues [ActiveRecord::Relation] Выборка записей
    # @param parser_config [Hash] Конфигурация из TemplateParser
    # @param error_behavior [String] Стратегия обработки ошибок ('abort', 'skip_field', 'skip_record')
    def initialize(issues, parser_config, error_behavior)
      @issues = issues
      @parser_config = parser_config
      @error_behavior = error_behavior
      @template_text = @parser_config[:template_text] || ''
      # Общий массив предупреждений для всех обрабатываемых записей.
      @warnings = []
    end

    # Основной метод: возвращает готовый хэш данных
    # @return [Hash] Структурированный контекст для шаблонизатора
    def build
      if @parser_config[:group_by]
        build_grouped_context
      else
        build_flat_context
      end
    rescue DocumentGenerator::TemplateError => e
      raise e
    rescue StandardError => e
      Rails.logger.error "[DocumentGenerator] Context build failed: #{e.message}"
      handle_error(I18n.t('document_generator.error_context_build_failed', message: e.message))
    end

    # Форматирует значение для безопасного вывода в шаблон
    # @param val [Object] Исходное значение
    # @return [String] Отформатированная строка
    def self.format_value(val)
      return '' if val.nil?
      if val.is_a?(Array)
        val.join(', ')
      elsif val.is_a?(Date) || val.is_a?(Time)
        val.strftime('%d.%m.%Y')
      else
        val.to_s
      end
    end

    private

    # Формирует плоский контекст без группировки.
    #
    # Для каждой основной задачи формируется отдельный контекст записи.
    # row_number является сквозной нумерацией основных записей и не зависит
    # от количества подзадач, наблюдателей или связей.
    #
    # @return [Hash] Контекст выгрузки без группировки.
    def build_flat_context
      records = []
      valid_issues = []

      @issues.each do |issue|
        begin
          record = build_issue_hash(issue)

          # При отсутствии GROUP_BY вся выборка считается одной группой.
          # Поэтому сквозной номер записи одновременно является номером
          # записи внутри этой единственной группы.
          row_number = records.length + 1
          record['row_number'] = row_number
          record['row_number_in_group'] = row_number

          records << record
          valid_issues << issue
        rescue DocumentGenerator::SkipRecordError
          next
        end
      end

      {
        'records' => records,
        'total_count' => records.size,
        'totals' => [
          {
            'total_count' => records.size
          }.merge(calculate_aggregates(valid_issues, 'total_'))
        ],
        'ExportDate' => Time.current.strftime('%d.%m.%Y %H:%M'),
        'ExportUser' => User.current.name,
        'ProjectName' => @issues.first&.project&.name || '',
        '__warnings' => @warnings
      }
    end

    # Формирует вложенный контекст двухуровневой группировки.
    #
    # Первый уровень определяется GROUP_BY.
    # Второй уровень, если указан, формируется внутри каждой группы первого уровня
    # по GROUP_BY_2.
    #
    # Для каждой основной записи формируются:
    # - row_number — сквозной номер записи во всей выборке;
    # - row_number_in_group — номер записи внутри группы первого уровня.
    #
    # Для второго уровня дополнительно формируется row_number_in_group_2.
    #
    # @return [Hash] Контекст выгрузки с группами первого и второго уровня.
    def build_grouped_context
      group_field = @parser_config[:group_by]
      group_field_2 = @parser_config[:group_by_2]

      resolved_group = FieldResolver.resolve(group_field)
      resolved_group_2 = group_field_2 ? FieldResolver.resolve(group_field_2) : nil

      unless resolved_group && resolved_group[:type] != :unknown
        handle_error(
          I18n.t(
            'document_generator.error_unknown_group_field',
            field: group_field
          )
        )
      end

      if group_field_2 && (!resolved_group_2 || resolved_group_2[:type] == :unknown)
        handle_error(
          I18n.t(
            'document_generator.error_unknown_group_field',
            field: group_field_2
          )
        )
      end

      # Сначала формируем группы первого уровня.
      first_level_groups = @issues.group_by do |issue|
        FieldResolver.get_value(issue, resolved_group)
      end

      groups_array = []
      all_valid_issues = []
      global_row_number = 0

      first_level_groups.each do |group_value, group_issues|
        valid_group_issues = []
        records = []

        # Формируем основные записи текущей группы.
        group_issues.each do |issue|
          begin
            record = build_issue_hash(issue)

            global_row_number += 1
            record['row_number'] = global_row_number

            records << record
            valid_group_issues << issue
            all_valid_issues << issue
          rescue DocumentGenerator::SkipRecordError
            next
          end
        end

        # Нумерация строк первого уровня начинается заново для каждой группы.
        records.each_with_index do |record, index|
          record['row_number_in_group'] = index + 1
        end

        group_context = {
          'count' => records.size,
          'records' => records
        }

        # Значение группировки сохраняем под именем соответствующего поля.
        # Имя приводится к нижнему регистру для совместимости с TemplateProcessor.
        group_context[group_field.to_s.strip.downcase] = group_value

        # Если объявлен второй уровень, создаём вложенные группы только
        # из записей текущей группы первого уровня.
        if resolved_group_2
          second_level_groups = valid_group_issues.group_by do |issue|
            FieldResolver.get_value(issue, resolved_group_2)
          end

          group_context['groups_2'] = []
          second_level_records_by_issue_id = {}

          second_level_groups.each do |group_value_2, second_group_issues|
            second_records = []

            second_group_issues.each do |issue|
              record = records.find do |candidate|
                candidate['__issue']&.id == issue.id
              end

              next unless record

              # Используем отдельную ссылку на тот же контекст записи,
              # чтобы row_number оставался глобальным, а номер второго уровня
              # рассчитывался независимо.
              second_records << record
            end

            second_records.each_with_index do |record, index|
              record['row_number_in_group_2'] = index + 1
            end

            # Формируем контекст внутренней группы.
            # Значение группировки доступно по имени поля, по которому выполняется группировка.
            second_group_context = {
              'count' => second_records.size,
              'records' => second_records
            }

            # Значение второй группировки сохраняем под именем соответствующего поля.
            # Имя приводится к нижнему регистру для совместимости с TemplateProcessor.
            second_group_context[group_field_2.to_s.strip.downcase] = group_value_2

            # Добавляем агрегаты второй группы.
            second_group_context.merge!(
              calculate_aggregates(second_group_issues, 'group_2_')
            )

            group_context['groups_2'] << second_group_context

            second_records.each do |record|
              second_level_records_by_issue_id[record['__issue'].id] = true
            end
          end
        end

        # Агрегаты первого уровня рассчитываются только по успешно
        # обработанным основным задачам этой группы.
        group_context.merge!(
          calculate_aggregates(valid_group_issues, 'group_')
        )

        groups_array << group_context
      end

      {
        'groups' => groups_array,
        'total_count' => all_valid_issues.size,
        'totals' => [
          {
            'total_count' => all_valid_issues.size
          }.merge(calculate_aggregates(all_valid_issues, 'total_'))
        ],
        'ExportDate' => Time.current.strftime('%d.%m.%Y %H:%M'),
        'ExportUser' => User.current.name,
        'ProjectName' => @issues.first&.project&.name || '',
        '__warnings' => @warnings
      }
    end

    # Формирует контекст одной основной задачи.
    #
    # @param issue [Issue] Основная задача Redmine.
    # @return [Hash] Контекст задачи со стандартными, пользовательскими полями
    #   и вложенными коллекциями.
    def build_issue_hash(issue)
      hash = {}

      # Сохраняем объект задачи для формирования диагностических сообщений.
      hash['__issue'] = issue

      # Все записи используют общий массив предупреждений текущей выгрузки.
      hash['__warnings'] = @warnings

      standard_fields_map = {
        'id' => issue.id,
        'subject' => issue.subject,
        'description' => issue.description,
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

      # Сначала добавляем стандартные поля.
      # Это обеспечивает их приоритет над пользовательскими полями
      # с совпадающим отображаемым названием.
      standard_fields_map.each do |field_key, value|
        localized_name = I18n.t(
          "field_#{field_key}",
          default: field_key.humanize
        )

        # Все имена полей в контексте хранятся в нижнем регистре.
        # Это делает стандартные поля регистронезависимыми при обращении из шаблона.
        hash[localized_name.to_s.downcase] = self.class.format_value(value)
      end

      # Формируем контекст родительской задачи.
      if issue.parent
        # Имена полей родительской задачи также хранятся в нижнем регистре,
        # чтобы обращения Parent.ID, Parent.Id, Parent.id и Parent.iD
        # обрабатывались одинаково.
        hash['parent'] = {
          I18n.t('field_subject', default: 'Subject').to_s.downcase => issue.parent.subject,
          I18n.t('field_status', default: 'Status').to_s.downcase => issue.parent.status&.name,
          I18n.t('field_assigned_to', default: 'Assigned to').to_s.downcase => issue.parent.assigned_to&.name,
          'id' => issue.parent.id
        }
      end

      # Добавляем пользовательские поля только при отсутствии стандартного
      # поля с таким же отображаемым названием.
      issue.custom_field_values.each do |cfv|
        field_name = cfv.custom_field.name.to_s.downcase

        # Стандартное поле с таким отображаемым именем имеет приоритет
        # над пользовательским полем.
        next if hash.key?(field_name)

        value = cfv.value.is_a?(Array) ? cfv.value.join(', ') : cfv.value

        hash[field_name] = self.class.format_value(value)
      end

      # Формируем контекст подзадач.
      hash['subtasks'] = Issue.where(parent_id: issue.id)
                              .includes(
                                :status,
                                :assigned_to,
                                custom_values: :custom_field
                              )
                              .map do |child|
        subtask_hash = {
          '__issue' => child,
          'id' => child.id,
          I18n.t('field_subject', default: 'Subject').to_s.downcase => child.subject,
          I18n.t('field_status', default: 'Status').to_s.downcase => child.status&.name
        }

        # Пользовательские поля подзадачи также не должны перезаписывать
        # стандартные поля с тем же названием.
        child.custom_field_values.each do |cfv|
          field_name = cfv.custom_field.name.to_s.downcase

          # Стандартное поле подзадачи имеет приоритет
          # над пользовательским полем с таким же названием.
          next if subtask_hash.key?(field_name)

          value = cfv.value.is_a?(Array) ? cfv.value.join(', ') : cfv.value

          subtask_hash[field_name] = self.class.format_value(value)
        end

        subtask_hash
      end

      # Формируем контекст связанных задач.
      hash['relations'] = issue.relations.map do |relation|
        target =
          if relation.issue_to_id == issue.id
            relation.issue_from
          else
            relation.issue_to
          end

        next nil unless target

      {
        I18n.t(
          'document_generator.relation_type',
          default: 'Relation Type'
        ).to_s.downcase => relation.relation_type_for(issue),
        I18n.t(
          'field_subject',
          default: 'Subject'
        ).to_s.downcase => target.subject,
        I18n.t(
          'field_status',
          default: 'Status'
        ).to_s.downcase => target.status&.name
      }
      end.compact

      # Формируем контекст наблюдателей.
      hash['watchers'] = issue.watchers.map do |watcher|
      {
        I18n.t(
          'document_generator.watcher_name',
          default: 'Name'
        ).to_s.downcase => watcher.user&.name
      }
      end

      hash
    rescue StandardError => e
      error_msg = I18n.t(
        'document_generator.error_record_processing_failed',
        id: issue.id,
        message: e.message
      )

      handle_error(error_msg)

      raise DocumentGenerator::SkipRecordError, error_msg if @error_behavior == 'skip_record'

      {}
    end

    # Вычисляет агрегаты для указанного набора основных задач.
    #
    # @param issues [Array<Issue>, ActiveRecord::Relation] Набор задач,
    #   для которого вычисляются агрегаты.
    # @param prefix [String] Префикс внутреннего ключа:
    #   group_, group_2_ или total_.
    # @return [Hash] Хэш вычисленных агрегатов.
    def calculate_aggregates(issues, prefix)
      aggregates = {}

      extract_aggregate_requests.each do |request|
        # count и total_count не являются агрегатами с аргументом поля.
        # Их значение формируется непосредственно из размера контейнера.
        next if request[:func] == 'count'

        # total_* должен вычисляться только для общего набора.
        if request[:is_total] && prefix != 'total_'
          next
        end

        # Обычный агрегат группы не должен попадать в общий итог.
        if !request[:is_total] && prefix == 'total_'
          next
        end

        func = request[:func]
        field = request[:field]

        key = "#{prefix}agg_#{func}_#{field}"

        aggregates[key] =
          AggregateCalculator.new(issues).calculate(func, field)
      end

      aggregates
    end

    # Извлекает все агрегатные команды из текста шаблона и проверяет,
    # что каждая команда встречается только один раз.
    #
    # Поддерживаются:
    # - <%sum(Поле)%>
    # - <%avg(Поле)%>
    # - <%min(Поле)%>
    # - <%max(Поле)%>
    # - <%total_sum(Поле)%>
    # - <%total_avg(Поле)%>
    # - <%total_min(Поле)%>
    # - <%total_max(Поле)%>
    # - <%count%>
    #
    # count обрабатывается отдельно от агрегатов с аргументом поля, поскольку
    # его значение определяется размером текущего контейнера.
    #
    # @return [Array<Hash>] Описания агрегатных команд.
    # @raise [DocumentGenerator::TemplateError] Если одна команда используется
    #   более одного раза.
    def extract_aggregate_requests
      requests = []
      seen = {}

      # Извлекаем агрегаты, принимающие имя поля.
      @template_text.scan(
        /<%\s*(total_)?(sum|avg|min|max)\s*\(\s*([^%]+?)\s*\)\s*%>/i
      ) do |is_total, function_name, field_name|
        function = function_name.downcase
        field = field_name.strip.downcase
        total = !is_total.nil?

        command_name =
          "#{total ? 'total_' : ''}#{function}(#{field})"

        command_key = command_name.downcase

        if seen.key?(command_key)
          raise DocumentGenerator::TemplateError,
                I18n.t(
                  'document_generator.error_duplicate_aggregate',
                  aggregate: command_name
                )
        end

        seen[command_key] = true

        requests << {
          is_total: total,
          func: function,
          field: field
        }
      end

      # count является специальной контекстной переменной.
      # Она может использоваться произвольное количество раз в разных
      # группах и подвалах, поскольку её значение определяется текущим
      # контекстом рендеринга.
      if @template_text.match?(/<%\s*count\s*%>/i)
        requests << {
          is_total: false,
          func: 'count',
          field: nil
        }
      end

      # total_count является отдельной командой общего итога.
      total_count_matches = @template_text.scan(
        /<%\s*total_count\s*%>/i
      ).size

      if total_count_matches > 1
        raise DocumentGenerator::TemplateError,
              I18n.t(
                'document_generator.error_duplicate_aggregate',
                aggregate: 'total_count'
              )
      end

      if total_count_matches == 1
        requests << {
          is_total: true,
          func: 'count',
          field: nil
        }
      end

      requests
    end

    # Универсальный обработчик ошибок, действующий согласно выбранной стратегии
    # @param message [String] Текст ошибки (уже локализованный)
    def handle_error(message)
      case @error_behavior
      when 'abort'
        raise DocumentGenerator::TemplateError, message
      when 'skip_field'
        Rails.logger.warn "[DocumentGenerator] Skipping field: #{message}"
        nil
      when 'skip_record'
        Rails.logger.warn "[DocumentGenerator] Skipping record: #{message}"
        raise DocumentGenerator::SkipRecordError, message
      end
    end
  end
end
# v2610071413