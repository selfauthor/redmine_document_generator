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
      issue_ids = query.issues.map(&:id)
      return [] if issue_ids.empty?

      sort_order = if query.sort_criteria.present?
                     query.sort_criteria.map { |c, d| "#{c} #{d.to_s.downcase == 'desc' ? 'DESC' : 'ASC'}" }.join(', ')
                   else
                     "#{Issue.table_name}.id DESC"
                   end

      # Загружаем задачи БЕЗ relations в includes, т.к. relations — это метод, а не ассоциация
      issues = Issue.where(id: issue_ids)
                    .includes(
                      :status, :priority, :author, :assigned_to, :project, :tracker, :category, :fixed_version, :parent,
                      custom_values: :custom_field,
                      watchers: :user
                    )
                    .order(Arel.sql(sort_order))
                    .to_a

      # Предзагружаем связи через специальный метод Redmine
      Issue.load_relations(issues)

      issues
    rescue ActiveRecord::RecordNotFound, StandardError => e
      Rails.logger.error "[DocumentGenerator] DataProvider error: #{e.message}"
      []
    end

    private

    def build_query
      query = IssueQuery.new(name: '_document_generator_temp')
      query.project = @project if @project
      query.build_from_params(params, {})
      query.user = @user
      
      # Убираем группировку фильтра (группировка управляется шаблоном)
      query.group_by = nil
      
      query
    end
  end
  # v2609111123
end