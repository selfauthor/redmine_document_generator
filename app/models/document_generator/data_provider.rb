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
      
      query = build_query
      
      # ИСПОЛЬЗУЕМ query.issues НАПРЯМУЮ
      # Он уже:
      # 1. Применяет правильную сортировку (включая 'parent' через JOIN)
      # 2. Учитывает права доступа (Issue.visible)
      # 3. Возвращает массив задач в нужном порядке
      issues_from_query = query.issues
      
      if issues_from_query.empty?
        return []
      end
      
      # Сохраняем порядок ID из query.issues
      ordered_ids = issues_from_query.map(&:id)
      
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
      
      Issue.load_relations(issues)
      
      issues
      
    rescue ActiveRecord::RecordNotFound => e
      raise
      
    rescue => e
      raise DocumentGenerator::RenderError, "Ошибка получения данных: #{e.message}"
    end

    private

    def build_query
      query = IssueQuery.new(name: '_document_generator_temp')
      query.project = @project if @project
      query.build_from_params(@params, {})
      query.user = @user
      
      # Убираем группировку фильтра (группировка управляется шаблоном)
      query.group_by = nil
      
      query
    end
  end
end
# v2609151543