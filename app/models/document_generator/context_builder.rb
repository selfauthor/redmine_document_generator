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

    def build_flat_context
      records = []
      @issues.each do |issue|
        records << build_issue_hash(issue)
      rescue DocumentGenerator::SkipRecordError => e
        Rails.logger.debug "[DocumentGenerator] Skipping record ##{issue.id}: #{e.message}"
        next
      end
      
      {
        'records' => records,
        'total_count' => records.size,
        'totals' => [{ 'total_count' => records.size }.merge(calculate_aggregates(@issues, 'total_'))],
        'ExportDate' => Time.now.strftime('%d.%m.%Y %H:%M'),
        'ExportUser' => User.current.name,
        'ProjectName' => @issues.first&.project&.name || ''
      }
    end

    def build_grouped_context
      group_field = @parser_config[:group_by]
      group_field_2 = @parser_config[:group_by_2]
      
      resolved_group = FieldResolver.resolve(group_field)
      resolved_group_2 = group_field_2 ? FieldResolver.resolve(group_field_2) : nil

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
        group_key_array = group_key.is_a?(Array) ? group_key : [group_key, nil]
        records = []
        
        group_issues.each do |issue|
          records << build_issue_hash(issue)
        rescue DocumentGenerator::SkipRecordError => e
          Rails.logger.debug "[DocumentGenerator] Skipping record in group ##{issue.id}: #{e.message}"
          next
        end
        
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

      {
        'groups' => groups_array,
        'total_count' => total_count,
        'totals' => [{ 'total_count' => total_count }.merge(calculate_aggregates(@issues, 'total_'))],
        'ExportDate' => Time.now.strftime('%d.%m.%Y %H:%M'),
        'ExportUser' => User.current.name,
        'ProjectName' => @issues.first&.project&.name || ''
      }
    end

    def build_issue_hash(issue)
      hash = {}
      
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

      standard_fields_map.each do |field_key, value|
        localized_name = I18n.t("field_#{field_key}", default: field_key.humanize)
        hash[localized_name] = self.class.format_value(value)
      end

      if issue.parent
        parent_prefix = 'Parent.'
        hash["#{parent_prefix}#{I18n.t('field_subject', default: 'Subject')}"] = issue.parent.subject
        hash["#{parent_prefix}#{I18n.t('field_status', default: 'Status')}"] = issue.parent.status&.name
        hash["#{parent_prefix}#{I18n.t('field_assigned_to', default: 'Assigned to')}"] = issue.parent.assigned_to&.name
      end

      issue.custom_field_values.each do |cfv|
        val = cfv.value.is_a?(Array) ? cfv.value.join(', ') : cfv.value
        hash[cfv.custom_field.name] = self.class.format_value(val)
      end

      hash['subtasks'] = issue.children.map do |child|
        {
          'ID' => child.id,
          I18n.t('field_subject', default: 'Subject') => child.subject,
          I18n.t('field_status', default: 'Status') => child.status&.name
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

      evaluate_template_functions(hash, issue)
      hash
    rescue StandardError => e
      error_msg = I18n.t('document_generator.error_record_processing_failed', id: issue.id, message: e.message)
      handle_error(error_msg)
      raise DocumentGenerator::SkipRecordError, error_msg if @error_behavior == 'skip_record'
      {}
    end

    def evaluate_template_functions(hash, issue)
      @template_text.scan(/<%\s*([a-z_]+)\s*\((.*?)\)\s*%>/i) do |func, args_str|
        args = args_str.split(',').map { |a| a.strip.gsub(/^['"]|['"]$/, '') }
        func_name = func.downcase
        
        begin
          case func_name
          when 'date'
            val = get_field_value(issue, args[0])
            hash["#{args[0]}_formatted"] = (val.is_a?(Date) || val.is_a?(Time)) ? val.strftime(args[1]) : val
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
            hash["#{args[0]}_default"] = val.to_s.presence || args[1]
          when 'strip_html'
            val = get_field_value(issue, args[0])
            hash["#{args[0]}_stripped"] = val.to_s.gsub(/<[^>]*>/, '')
          end
        rescue StandardError => e
          error_msg = I18n.t('document_generator.error_function_failed', func: func_name, message: e.message)
          handle_error(error_msg)
        end
      end
    end

    def get_field_value(issue, field_name)
      resolved = FieldResolver.resolve(field_name)
      if resolved[:type] == :unknown
        error_msg = I18n.t('document_generator.error_unknown_field_in_function', field: field_name)
        handle_error(error_msg)
        return nil
      end
      FieldResolver.get_value(issue, resolved)
    end

    def calculate_aggregates(issues, prefix)
      aggregates = {}
      aggregate_requests = extract_aggregate_requests
      
      aggregate_requests.each do |req|
        is_total = req[:is_total]
        func = req[:func]
        field = req[:field]
        
        target_issues = is_total ? @issues : issues
        key = "#{prefix}agg_#{func}_#{field}"
        aggregates[key] = AggregateCalculator.new(target_issues).calculate(func, field)
      end
      
      aggregates
    end

    def extract_aggregate_requests
      requests = []
      @template_text.scan(/<%\s*(total_)?(count|sum|avg|min|max|concat)\s*\(\s*([^%]+?)\s*\)\s*%>/i) do |is_total, func, field|
        requests << { is_total: is_total.present?, func: func.downcase, field: field.strip }
      end
      requests.uniq
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