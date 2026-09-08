# frozen_string_literal: true

module DocumentGenerator
  class FieldResolver
    @@field_name_cache = nil
    @@custom_fields_cache = nil

    # Сброс кэша полезен при разработке или если администратор добавил новое поле
    def self.reset_cache!
      @@field_name_cache = nil
      @@custom_fields_cache = nil
    end

    def self.field_name_cache
      @@field_name_cache ||= build_field_name_cache
    end

    def self.custom_fields_cache
      @@custom_fields_cache ||= build_custom_fields_cache
    end

    def self.build_field_name_cache
      cache = {}
      
      # Список внутренних ключей стандартных полей Redmine
      standard_fields = %w[
        project tracker status priority author assigned_to category fixed_version 
        subject description start_date due_date done_ratio estimated_hours 
        spent_hours created_on updated_on closed_on
      ]

      standard_fields.each do |field|
        # Получаем локализованное название из файлов locales (например, "Тема" или "Subject")
        # Если перевод не найден, используем humanized ключ (например, "Subject")
        begin
          localized_name = I18n.t("field_#{field}", default: field.humanize).downcase
        rescue I18n::MissingTranslationData
          localized_name = field.humanize.downcase
        end
        
        # Добавляем в кэш вариант с пробелами и без них
        cache[localized_name] = field
        cache[localized_name.gsub(/\s+/, '')] = field
        
        # Также разрешаем использование латинского ключа напрямую (для универсальности)
        cache[field.downcase] = field
      end
      
      # Специальный случай для ID
      cache['id'] = 'id'
      begin
        cache[I18n.t('field_id', default: 'ID').downcase] = 'id'
      rescue I18n::MissingTranslationData
        cache['id'] = 'id'
      end

      cache
    end

    def self.build_custom_fields_cache
      cache = {}
      # Загружаем все пользовательские поля для задач. 
      # Проверка прав на конкретное поле происходит позже, при вызове custom_field_value.
      IssueCustomField.all.each do |cf|
        name = cf.name.downcase
        cache[name] = cf.id
        cache[name.gsub(/\s+/, '')] = cf.id
      end
      cache
    end

    # Разрешает строку имени поля из шаблона в структурированный хэш
    def self.resolve(field_name)
      return nil if field_name.blank?

      original = field_name.strip
      name = original.downcase

      # 1. Явное указание пользовательского поля: CF:ИмяПоля
      if name.start_with?('cf:')
        cf_name = name.sub(/^cf:\s*/, '')
        cf_id = custom_fields_cache[cf_name] || custom_fields_cache[cf_name.gsub(/\s+/, '')]
        return { type: :custom, key: "cf_#{cf_id}", cf_id: cf_id } if cf_id
      end

      # 2. Поле родительской задачи: Parent.ИмяПоля
      if name.start_with?('parent.')
        sub_field_name = original.sub(/^parent\.\s*/i, '').strip
        resolved = resolve(sub_field_name)
        if resolved && resolved[:type] != :parent
          resolved[:type] = :parent
          return resolved
        end
      end

      # 3. Стандартное поле (по локализованному или латинскому имени)
      if field_name_cache.key?(name)
        return { type: :standard, key: field_name_cache[name] }
      end

      # 4. Пользовательское поле по имени (без префикса CF:)
      cf_id = custom_fields_cache[name] || custom_fields_cache[name.gsub(/\s+/, '')]
      return { type: :custom, key: "cf_#{cf_id}", cf_id: cf_id } if cf_id

      # 5. Неизвестное поле (может быть спец. переменной или функцией, обработает рендерер)
      { type: :unknown, key: original }
    end

    # Извлекает значение из объекта Issue на основе разрешённого поля
    def self.get_value(issue, resolved_field)
      return nil unless issue && resolved_field

      target = issue
      if resolved_field[:type] == :parent
        target = issue.parent
        return nil unless target
      end

      case resolved_field[:type]
      when :standard then get_standard_value(target, resolved_field[:key])
      when :custom then target.custom_field_value(resolved_field[:cf_id])
      else nil
      end
    end

    def self.get_standard_value(issue, key)
      case key
      when 'id' then issue.id
      when 'subject' then issue.subject
      when 'description' then issue.description
      when 'status' then issue.status&.name
      when 'priority' then issue.priority&.name
      when 'author' then issue.author&.name
      when 'assigned_to' then issue.assigned_to&.name
      when 'start_date' then issue.start_date
      when 'due_date' then issue.due_date
      when 'done_ratio' then "#{issue.done_ratio}%"
      when 'estimated_hours' then issue.estimated_hours
      when 'spent_hours' then issue.spent_hours
      when 'created_on' then issue.created_on
      when 'updated_on' then issue.updated_on
      when 'closed_on' then issue.closed_on
      when 'project' then issue.project&.name
      when 'tracker' then issue.tracker&.name
      when 'category' then issue.category&.name
      when 'fixed_version' then issue.fixed_version&.name
      when 'parent_id' then issue.parent_id
      else 
        # Fallback для любых других стандартных методов объекта issue
        issue.send(key) if issue.respond_to?(key)
      end
    end
  end
end