# frozen_string_literal: true

module DocumentGenerator
  class DataProvider
    attr_reader :project, :user, :params

    def initialize(project, user, params)
      @project = project
      @user = user
      @params = params
    end

    def fetch_issues
      Rails.logger.info "[DOC_GEN_DEBUG] === DataProvider#fetch_issues НАЧАЛО ==="
      
      query = build_query
      Rails.logger.info "[DOC_GEN_DEBUG] Query создан. Фильтры: #{query.filters.inspect}"
      Rails.logger.info "[DOC_GEN_DEBUG] SQL-условие: #{query.statement}"
      Rails.logger.info "[DOC_GEN_DEBUG] sort_criteria: #{query.sort_criteria.inspect}"
      
      # ИСПОЛЬЗУЕМ query.issues НАПРЯМУЮ
      # Он уже:
      # 1. Применяет правильную сортировку (включая 'parent' через JOIN)
      # 2. Учитывает права доступа (Issue.visible)
      # 3. Возвращает массив задач в нужном порядке
      issues_from_query = query.issues
      Rails.logger.info "[DOC_GEN_DEBUG] query.issues вернул #{issues_from_query.size} задач"
      
      if issues_from_query.empty?
        Rails.logger.warn "[DOC_GEN_DEBUG] query.issues вернул пустой массив."
        return []
      end
      
      # Сохраняем порядок ID из query.issues
      ordered_ids = issues_from_query.map(&:id)
      Rails.logger.info "[DOC_GEN_DEBUG] Порядок ID: #{ordered_ids.inspect}"
      
      # Перезагружаем задачи с нужными includes (без .order — чтобы не ломать 'parent')
      issues = Issue.where(id: ordered_ids)
                    .includes(
                      :status, :priority, :author, :assigned_to, :project, :tracker, :category, :fixed_version, :parent,
                      custom_values: :custom_field,
                      watchers: :user
                    )
                    .to_a
      
      # Восстанавливаем порядок из query.issues
      issues_by_id = issues.index_by(&:id)
      issues = ordered_ids.map { |id| issues_by_id[id] }.compact
      
      Rails.logger.info "[DOC_GEN_DEBUG] Загружено объектов Issue с includes: #{issues.size}"
      Rails.logger.info "[DOC_GEN_DEBUG] Вызов Issue.load_relations..."
      Issue.load_relations(issues)
      Rails.logger.info "[DOC_GEN_DEBUG] === DataProvider#fetch_issues УСПЕШНО ЗАВЕРШЕНО ==="
      
      issues
      
    rescue ActiveRecord::RecordNotFound => e
      Rails.logger.error "[DOC_GEN_DEBUG] RecordNotFound: #{e.message}"
      Rails.logger.error "[DOC_GEN_DEBUG] Backtrace: #{e.backtrace[0..5].join("\n[DOC_GEN_DEBUG] ")}"
      raise
      
    rescue => e
      Rails.logger.error "[DOC_GEN_DEBUG] КРИТИЧЕСКАЯ ОШИБКА в DataProvider: #{e.class} - #{e.message}"
      Rails.logger.error "[DOC_GEN_DEBUG] Полный backtrace:\n#{e.backtrace.join("\n")}"
      raise DocumentGenerator::RenderError, "Ошибка получения данных: #{e.message}"
    end

    private

    def build_query
      query = IssueQuery.new(name: '_document_generator_temp')
      query.project = @project if @project
      query.build_from_params(@params, {})
      query.user = @user
      
      Rails.logger.info "[DOC_GEN_DEBUG] build_query: user = #{@user.login}, project = #{@project&.identifier}"
      Rails.logger.info "[DOC_GEN_DEBUG] build_query: @params = #{@params.inspect}"
      
      # Убираем группировку фильтра (группировка управляется шаблоном)
      query.group_by = nil
      
      query
    end
  end
  # v2609141619
end